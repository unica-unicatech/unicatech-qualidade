"""Core automation loop: for each record, search CEP + Número in the remote app and
classify the result as KEEP (no data found -> eligible) or ELIMINATE (data found).
CNPJ is carried along purely as the output's identifying/reference column -- the
site itself is searched by CEP + Número, not by Documento/CNPJ."""
import difflib
import json
import random
import re
import threading
import time
from pathlib import Path

import pyautogui
import pytesseract

import auto_detect
import citrix_utils
import ocr_utils

try:
    import win32clipboard
    _HAS_CLIPBOARD = True
except ImportError:
    _HAS_CLIPBOARD = False


def _set_clipboard_text(text: str):
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text, win32clipboard.CF_TEXT)
    finally:
        win32clipboard.CloseClipboard()

STATE_DIR = Path(__file__).parent / "state"
STATE_DIR.mkdir(exist_ok=True)
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

CHECKPOINT_PATH = STATE_DIR / "checkpoint.json"

STATUS_KEEP = "mantido"
STATUS_ELIMINATE = "eliminado"
STATUS_ERROR = "erro"
STATUS_INVALID_INPUT = "cep_ou_numero_invalido"

MOUSE_MOVE_SECONDS = 0.25  # visible glide instead of an instant teleport, but still quick

MAX_LOGIN_RETRIES = 2  # how many consecutive "not logged in" checks before auto-pausing
MAX_TEMPLATE_DRIFT_PX = 60  # see _locate_click_point's sanity check against the manual calibration point
OCR_RETRIES = 4
LOGIN_CHECK_INTERVAL = 6  # in dual-window mode, re-check login every N fills per lane (not every single one)


class AutomationRunner:
    def __init__(self, config: dict, records, log=print, on_progress=None, on_logout=None,
                 input_signature="default"):
        self.config = config
        pytesseract.pytesseract.tesseract_cmd = config.get(
            "tesseract_cmd", r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        )
        # records: list of {"cnpj": str, "cep": str, "numero": str}
        self.records = [r for r in records if r.get("cep") or r.get("numero")]
        self.cnpj_list = [r["cnpj"] for r in self.records]
        self._records_by_cnpj = {r["cnpj"]: r for r in self.records}
        self.log = log
        self.on_progress = on_progress or (lambda *a, **k: None)
        self.on_logout = on_logout or (lambda: None)
        self.input_signature = input_signature

        self.results = {}  # cnpj -> status
        self.index = 0

        self._pause_event = threading.Event()
        self._pause_event.set()  # not paused
        self._stop_event = threading.Event()
        self._consecutive_login_failures = 0
        self._last_window_pos = None
        self._click_cache = {}  # (win.left, win.top, element_name) -> (x, y), see _locate_click_point
        self._region_cache = {}  # (win.left, win.top, element_name) -> (x1,y1,x2,y2), see _locate_region_box
        self._force_recheck_login = True  # re-check on the very first item / right after a resume
        self._data_signatures = []  # cleaned "data" OCR texts seen so far, see _read_result
        self._pre_search_snapshot = {}  # (win.left, win.top) -> pre-click table snapshot, see _start_search

        self._load_checkpoint_if_matching()

    # ---------- checkpoint ----------
    def _load_checkpoint_if_matching(self):
        if not CHECKPOINT_PATH.exists():
            return
        try:
            data = json.loads(CHECKPOINT_PATH.read_text(encoding="utf-8"))
        except Exception:
            return
        if data.get("input_signature") != self.input_signature:
            return
        self.results = data.get("results", {})
        self.index = data.get("index", 0)
        self.log(f"Checkpoint encontrado: retomando do item {self.index + 1}/{len(self.records)}.")

    def _save_checkpoint(self):
        data = {
            "input_signature": self.input_signature,
            "index": self.index,
            "results": self.results,
        }
        CHECKPOINT_PATH.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")

    def clear_checkpoint(self):
        if CHECKPOINT_PATH.exists():
            CHECKPOINT_PATH.unlink()

    # ---------- controls ----------
    def pause(self):
        self._pause_event.clear()
        self.log("Pausado. Clique em Retomar quando estiver pronto (relogue no Citrix se precisar).")

    def resume(self):
        self._consecutive_login_failures = 0
        self._force_recheck_login = True
        self._pause_event.set()
        self.log("Retomando...")

    def stop(self):
        self._stop_event.set()
        self._pause_event.set()  # unblock if paused, so it can exit the loop

    # ---------- window/session helpers ----------
    def _get_window(self):
        # Primary: relocate the DOCUMENTO label by its saved appearance and find the
        # window underneath that exact point. This is far more reliable than matching
        # by window title -- title-substring lookups can pick the wrong window when
        # more than one shares similar text (e.g. a stale/minimized duplicate). The
        # label stays on screen and unchanged regardless of which fields (CEP/Número)
        # we're actually typing into.
        if auto_detect.has_templates():
            marker_box, _variant = auto_detect._locate_box("app_marker_region", lambda *a, **k: None)
            if marker_box is not None:
                anchor_x = marker_box.left + marker_box.width // 2
                anchor_y = marker_box.top + marker_box.height // 2
                win = citrix_utils.find_window_at_point(anchor_x, anchor_y)
                if win is not None:
                    self._last_window_pos = (win.left, win.top)
                    return win

        win = citrix_utils.find_window(self.config["window_title_contains"])
        if win is None:
            time.sleep(0.5)  # tolerate a transient enumeration glitch before giving up
            win = citrix_utils.find_window(self.config["window_title_contains"])
        if win is None and self._last_window_pos is not None:
            # Fallback: title lookup can occasionally miss (timing/pygetwindow quirk).
            # Try the last known window position + the cep_field offset, which
            # should still land inside the window if it hasn't actually moved/closed.
            lx, ly = self._last_window_pos
            fx = lx + self.config["cep_field"]["x"]
            fy = ly + self.config["cep_field"]["y"]
            win = citrix_utils.find_window_at_point(fx, fy)
        if win is not None:
            self._last_window_pos = (win.left, win.top)
        return win

    def _is_logged_in(self, win) -> bool:
        try:
            x1, y1, x2, y2 = self._locate_region_box(win, "app_marker_region")
            text = ocr_utils.read_text(x1, y1, x2, y2)
            ok = ocr_utils.fuzzy_contains(text, "documento", threshold=0.6)
            if not ok:
                self.log(f"   (debug) OCR da região de verificação leu: '{text}' "
                         f"(região abs: {x1},{y1} -> {x2},{y2}; janela em {win.left},{win.top})")
            return ok
        except Exception as exc:
            self.log(f"Aviso: falha ao checar sessão ({exc}).")
            return False

    def _locate_click_point(self, win, name: str):
        """Where to click `name` (cep_field/numero_field/pesquisar_button) for this window.

        Detecting by appearance on every single search (region-restricted to this
        window, so 2-window mode can't cross-match the other lane's element) is what
        makes the system adapt automatically to a layout change -- but doing it EVERY
        time is expensive and eats into the speedup 2-window mode is supposed to give.
        So: detect once per window and cache it; only re-detect if the cache is
        explicitly invalidated (see _invalidate_click_cache, called when a search
        comes back as an error -- a good sign the layout may have shifted).
        """
        key = (win.left, win.top, name)
        cached = self._click_cache.get(key)
        if cached is not None:
            return cached

        config_point = None
        if name in self.config:
            config_point = citrix_utils.absolute_point(win, self.config[name]["x"], self.config[name]["y"])

        point = None
        if auto_detect.has_templates():
            region = (win.left, win.top, win.width, win.height)
            pos = auto_detect._locate_point(name, lambda *a, **k: None, region=region)
            if pos is not None:
                point = (pos.x, pos.y)
                # Sanity check against the manually-calibrated point: cep_field,
                # numero_field and documento_field are near-identical blank input
                # boxes, so appearance matching can confidently lock onto the
                # WRONG one of the three (seen in practice: CEP typed into the
                # Documento field). The manual calibration point is ground truth
                # right after being (re)captured, so if the image match landed
                # far from it, trust the manual point instead of the match.
                if config_point is not None:
                    dx = point[0] - config_point[0]
                    dy = point[1] - config_point[1]
                    if (dx * dx + dy * dy) ** 0.5 > MAX_TEMPLATE_DRIFT_PX:
                        self.log(f"   Aviso: detecção por imagem de '{name}' achou {point}, longe "
                                 f"do ponto calibrado manualmente {config_point} -- usando o calibrado.")
                        point = None
        if point is None:
            point = config_point

        self._click_cache[key] = point
        return point

    def _field_ocr_box(self, win, name: str):
        """Region to OCR-read back a filled field's value, for verification.

        Deliberately crops a tight band around the CLICK point, not the matched
        template's full box: some field templates are captured with extra context
        above the input (e.g. its label -- see auto_detect._click_ratio), so the
        matched box can span both the label line and the input line. Debug
        screenshots showed that a crop spanning both lines reads back as EMPTY
        every time, because the digit-only OCR call uses psm 7 (single text line)
        -- it doesn't ever misread the label, it just fails outright on a
        two-line image. The click point itself is calibrated to sit inside the
        actual input, so cropping around just that stays single-line.

        Horizontal pad is deliberately modest: on a tighter layout (seen on a
        1600x900 machine vs. the 1920x1080 one this was tuned on), the fields
        sit close enough together that a wide crop pulls in the neighboring
        input box and even part of the Pesquisar button -- debug screenshots
        showed exactly that, with OCR then garbling the reading (e.g. reading
        '45102606' for an 8-digit CEP, or empty) despite the correct value
        being clearly visible in the crop."""
        x, y = self._locate_click_point(win, name)
        # Asymmetric vertical pad: on a tighter layout the label sits close enough
        # above the input that even +/-20 reaches up into it (confirmed by debug
        # screenshots showing "Numero"/label text captured alongside the correct
        # value, reading back empty for the same two-line reason as the box-based
        # crop this replaced). Padding up less than down keeps the label out while
        # still covering the full digit height below the click point.
        return x - 70, y - 10, x + 70, y + 18

    def _locate_region_box(self, win, name: str, pad: int = 8):
        """Where `name` (app_marker_region/table_row_region) is for this window, as
        (x1, y1, x2, y2). Same appearance-detection-with-cache strategy as
        _locate_click_point -- fixes the marker/table check reading the wrong spot
        when the page has scrolled/reflowed since config.json's offsets were saved,
        even though window-finding itself (which already re-detects fresh every
        time) still works fine."""
        key = (win.left, win.top, name)
        cached = self._region_cache.get(key)
        if cached is not None:
            return cached

        config_box = None
        if name in self.config:
            config_box = citrix_utils.absolute_region(win, self.config[name], pad=pad)

        box = None
        if auto_detect.has_templates():
            region = (win.left, win.top, win.width, win.height)
            b = auto_detect._locate_region(name, lambda *a, **k: None, region=region)
            if b is not None:
                box = (b.left - pad, b.top - pad, b.left + b.width + pad, b.top + b.height + pad)
                # Same sanity check as _locate_click_point: don't trust an
                # appearance match that landed far from the manually calibrated
                # region -- seen in practice as table-read OCR returning the
                # same garbage every attempt because the crop was pinned to the
                # wrong static spot on screen instead of the results table.
                if config_box is not None:
                    dx = box[0] - config_box[0]
                    dy = box[1] - config_box[1]
                    if (dx * dx + dy * dy) ** 0.5 > MAX_TEMPLATE_DRIFT_PX:
                        self.log(f"   Aviso: detecção por imagem de '{name}' achou {box}, longe "
                                 f"da região calibrada manualmente {config_box} -- usando a calibrada.")
                        box = None
        if box is None:
            box = config_box

        self._region_cache[key] = box
        return box

    def _invalidate_click_cache(self, win):
        key_prefix = (win.left, win.top)
        for key in [k for k in self._click_cache if k[:2] == key_prefix]:
            del self._click_cache[key]
        for key in [k for k in self._region_cache if k[:2] == key_prefix]:
            del self._region_cache[key]

    # ---------- per-record processing ----------
    def _read_field_digits(self, win, field_name: str) -> str:
        x1, y1, x2, y2 = self._field_ocr_box(win, field_name)
        text = ocr_utils.read_digits(x1, y1, x2, y2)
        return re.sub(r"\D", "", text)

    def _fill_field(self, win, field_name: str, value: str):
        """Fill the field via a single clipboard paste (one paste action for the
        whole value) when available -- across every debug screenshot gathered
        while chasing field-corruption reports, paste never once actually failed
        to land the correct value. Typing (simulating individual keystrokes) is
        only used as a fallback when clipboard isn't available at all, since it
        has an actual proven failure mode: on some PCs/sessions, typed keystrokes
        get registered twice by the remote app regardless of speed (e.g. '776'
        arrives as '777766').

        There used to be an OCR readback here to verify the fill and retry on
        mismatch, but it was doing more harm than good: it never once caught a
        genuine paste failure, it regularly misread short values (e.g. read
        '196' for a correct, clearly-legible '26'), and "retrying" on those
        false alarms -- by re-clicking/re-pasting -- was the actual source of
        every real corruption seen (the typing-duplication bug above, and
        re-clicks landing wrong mid-fill)."""
        x, y = self._locate_click_point(win, field_name)
        pyautogui.click(x, y, duration=MOUSE_MOVE_SECONDS)
        time.sleep(0.15)
        pyautogui.hotkey("ctrl", "a")
        pyautogui.press("delete")
        time.sleep(0.1)
        if not value:
            return

        if _HAS_CLIPBOARD:
            _set_clipboard_text(value)
            time.sleep(0.05)
            pyautogui.hotkey("ctrl", "v")
        else:
            pyautogui.typewrite(value, interval=0.05)
        time.sleep(0.15)

    def _start_search(self, cnpj: str, win):
        """Fill CEP + Número and hit Pesquisar. Does NOT wait for the result --
        split out so a second window can be worked while this one is still loading
        (see run_dual)."""
        record = self._records_by_cnpj[cnpj]

        self._fill_field(win, "cep_field", record["cep"])
        self._fill_field(win, "numero_field", record["numero"])

        # Snapshot the table region right before clicking Pesquisar, so
        # _read_result can tell whether it's still looking at the PREVIOUS
        # search's table (page hasn't refreshed yet) instead of trusting
        # whatever's on screen the instant the fixed post-click wait elapses.
        rx1, ry1, rx2, ry2 = self._locate_region_box(win, "table_row_region")
        self._pre_search_snapshot[(win.left, win.top)] = ocr_utils.grab_region_gray(rx1, ry1, rx2, ry2)

        # Detected right before clicking (not up front) -- some layouts style the
        # button differently while the fields are still empty vs. filled in.
        btn_x, btn_y = self._locate_click_point(win, "pesquisar_button")
        pyautogui.click(btn_x, btn_y, duration=MOUSE_MOVE_SECONDS)

    def _verify_cep_field(self, win, cnpj: str) -> bool:
        """Extra safety net (mainly for 2-window mode): confirm the CEP field in
        `win` still shows the CEP we searched for THIS record, before trusting
        whatever the table shows. Catches any window mixup -- wrong lane, a stale
        cached click position, etc. -- instead of silently reporting a result for
        the wrong CNPJ."""
        record = self._records_by_cnpj.get(cnpj)
        if not record or not record.get("cep"):
            return True
        got = self._read_field_digits(win, "cep_field")
        expected = record["cep"]
        match = bool(got) and (got == expected or got in expected or expected in got)
        if not match:
            self.log(f"   [{cnpj}] (debug) campo CEP leu '{got}' (esperado '{expected}').")
        return match

    @staticmethod
    def _clean_for_signature(text: str) -> str:
        return re.sub(r"[^a-z0-9]", "", text.lower())

    def _looks_like_watermark_noise(self, text: str) -> bool:
        """A 'data' classification should mean genuine, per-record table content --
        but a session watermark (username/IP overlaid on the page, drifting across
        it -- see auto_detect.py's CONFIDENCE comment) sitting inside the table
        crop can read as legible OCR text that happens not to match the 'no data'
        phrase, wrongly defaulting to 'data' (seen in practice: eliminating a
        client based on watermark garbage). Real data legitimately varies company
        to company; the watermark doesn't. So: if this reading is nearly
        identical to a 'data' reading already seen for a DIFFERENT record in this
        same run, it's almost certainly static noise, not this record's data."""
        cleaned = self._clean_for_signature(text)
        if len(cleaned) < 6:
            return False
        for seen in self._data_signatures:
            if difflib.SequenceMatcher(None, cleaned, seen).ratio() >= 0.88:
                return True
        return self._has_internal_repeat(cleaned)

    @staticmethod
    def _has_internal_repeat(cleaned: str, chunk_len: int = 7) -> bool:
        """Catches watermark noise on its very FIRST appearance (before there's
        any prior reading to compare against): a watermark tiles/repeats within
        the page, so it tends to show the same chunk of garbled text twice
        within the SAME crop (e.g. literally 'ODA Seal) (1) ODA Seal) (1)').
        Genuine table content (column headers, an address, a name) doesn't
        repeat a 7+ character run of itself like that."""
        n = len(cleaned)
        if n < chunk_len * 2:
            return False
        seen_chunks = set()
        for i in range(n - chunk_len + 1):
            chunk = cleaned[i:i + chunk_len]
            if chunk in seen_chunks:
                return True
            seen_chunks.add(chunk)
        return False

    def _classify_table_by_image(self, win):
        """Recognize whether the table shows 'no data' by the on-screen
        APPEARANCE of that banner (see calibration.calibrate_result_banners)
        instead of OCR-reading the row's text. Much more tolerant of the session
        watermark than character-level OCR -- it only needs to match the overall
        shape (same appearance-matching-with-retries as auto_detect already uses
        for buttons/fields, which already proved workable against this same
        watermark), not make out individual letters.

        Only distinguishes empty vs. has-data -- whether the CEP/Número itself
        was invalid doesn't change the KEEP/ELIMINATE decision, so that state
        isn't checked here (it still gets detected by the OCR fallback below
        when these banner templates aren't calibrated at all).

        Returns "empty", "data", or None if the banner template hasn't been
        calibrated -- callers should fall back to OCR-based
        classify_table_region in that case."""
        if not auto_detect.variant_paths("no_data_banner"):
            return None
        region = (win.left, win.top, win.width, win.height)
        no_data_box, _ = auto_detect._locate_box("no_data_banner", lambda *a, **k: None, region=region)
        return "empty" if no_data_box is not None else "data"

    def _read_result(self, cnpj: str, win, verify: bool = False) -> str:
        """OCR-classify the results table for the search already triggered on `win`.
        Caller is responsible for having waited long enough beforehand.

        `verify=True` first confirms the CEP field still shows this record's CEP --
        only meaningful (and only worth the extra OCR call) in 2-window mode, where
        a mixup between lanes is actually possible; single-window mode can't have
        that problem, so it skips this check entirely."""
        if verify and not self._verify_cep_field(win, cnpj):
            self.log(f"   [{cnpj}] (debug) ALERTA: o campo CEP dessa janela não bate com o esperado "
                     f"na hora de ler o resultado -- pode ter havido mistura de janela. Marcando como "
                     f"erro pra reprocessar depois, em vez de arriscar um resultado errado.")
            self._invalidate_click_cache(win)
            return STATUS_ERROR

        rx1, ry1, rx2, ry2 = self._locate_region_box(win, "table_row_region")

        before = self._pre_search_snapshot.get((win.left, win.top))
        if before is not None:
            for _ in range(4):
                after = ocr_utils.grab_region_gray(rx1, ry1, rx2, ry2)
                if ocr_utils.region_changed(before, after):
                    break
                self.log(f"   [{cnpj}] (debug) tabela ainda igual à busca anterior -- "
                         f"aguardando a página atualizar...")
                time.sleep(1.0)
            # Not blocking indefinitely: if it genuinely never changes (e.g. two
            # consecutive records that both happen to have no data, which render
            # pixel-identical), fall through and classify whatever's there once
            # this bounded wait is spent, same as before this check existed.

        if auto_detect.variant_paths("no_data_banner"):
            # Don't trust "banner not found" on the very first look -- a slow
            # page load can still be rendering (spinner, blank transition) at
            # that instant, which looks identical to "not found" and would
            # otherwise get read as "has data" and wrongly eliminate the
            # record. Keep checking for a while before concluding that.
            for attempt in range(1, OCR_RETRIES + 1):
                image_kind = self._classify_table_by_image(win)
                if image_kind == "empty":
                    self.log(f"   [{cnpj}] (debug) tabela lida como VAZIA (reconhecida por imagem).")
                    return STATUS_KEEP
                if attempt < OCR_RETRIES:
                    self.log(f"   [{cnpj}] (debug) aviso de vazio ainda não apareceu na tela "
                             f"(tentativa {attempt}/{OCR_RETRIES}) -- pode estar carregando. "
                             f"Aguardando mais um pouco...")
                    time.sleep(1.5)
            self.log(f"   [{cnpj}] (debug) tabela lida como COM DADOS (aviso de vazio não "
                     f"reconhecido na tela após {OCR_RETRIES} tentativas).")
            return STATUS_ELIMINATE

        best_text = ""
        for attempt in range(1, OCR_RETRIES + 1):
            # Fast (fewer OCR combinations) on the first, most-likely-to-succeed
            # attempt; fall back to the thorough sweep only once it's proven harder.
            kind, text = ocr_utils.classify_table_region(rx1, ry1, rx2, ry2, fast=(attempt == 1))
            if len(text) > len(best_text):
                best_text = text
            if kind == "empty":
                self.log(f"   [{cnpj}] (debug) tabela lida como VAZIA. Texto OCR: '{text}'")
                return STATUS_KEEP
            if kind == "invalid_input":
                self.log(f"   [{cnpj}] (debug) CEP/Número inválido ou vazio nessa linha da planilha. "
                         f"Texto OCR: '{text}'")
                return STATUS_INVALID_INPUT
            if kind == "data":
                if self._looks_like_watermark_noise(text):
                    self.log(f"   [{cnpj}] (debug) tentativa {attempt}/{OCR_RETRIES} leu 'com dados', mas "
                             f"o texto parece ser a marca d'água da sessão, não dado real da tabela "
                             f"(tentativa {attempt}/{OCR_RETRIES}). Texto OCR: '{text}'. Aguardando a "
                             f"marca d'água se mover e tentando de novo...")
                    # Don't give up on the FIRST watermark-looking read -- the watermark
                    # drifts across the page over time (see auto_detect.py's CONFIDENCE
                    # comment), so waiting and re-reading often gets a clean shot at the
                    # real content within the same search, instead of burning a whole
                    # extra reprocessing pass later for something a few more seconds
                    # here would have resolved.
                    time.sleep(1.5)
                    continue
                self.log(f"   [{cnpj}] (debug) tabela lida como COM DADOS. Texto OCR: '{text}'")
                self._data_signatures.append(self._clean_for_signature(text))
                return STATUS_ELIMINATE
            self.log(f"   [{cnpj}] Tabela ainda sem texto legível (tentativa {attempt}/{OCR_RETRIES}), "
                     f"melhor leitura: '{text}'. Aguardando mais um pouco...")
            time.sleep(1.5)

        self.log(f"   [{cnpj}] Falha final de OCR (marca d'água persistente ou tabela ilegível). "
                 f"Última leitura: '{best_text}'")
        return STATUS_ERROR

    def _process_one(self, cnpj: str, win) -> str:
        """Single-lane version: search and wait for the fixed delay before reading."""
        self._start_search(cnpj, win)
        time.sleep(self.config.get("wait_after_search_seconds", 3.5))
        return self._read_result(cnpj, win)

    def _ensure_ready(self):
        """Get a focused, logged-in window. Returns None (and pauses/notifies) if not ready."""
        win = self._get_window()
        if win is None:
            self.log("Janela do Citrix/site não encontrada. Pausando automaticamente.")
            self.pause()
            self.on_logout()
            return None

        citrix_utils.focus_window(win)
        time.sleep(0.3)

        if not self._is_logged_in(win):
            self._consecutive_login_failures += 1
            self.log(f"Aviso: não encontrei o campo DOCUMENTO na tela "
                     f"({self._consecutive_login_failures}/{MAX_LOGIN_RETRIES}). "
                     f"Pode ter deslogado ou a página não carregou.")
            if self._consecutive_login_failures >= MAX_LOGIN_RETRIES:
                self.log("Pausando automaticamente para você relogar.")
                self.pause()
                self.on_logout()
            else:
                time.sleep(2)
            return None

        self._consecutive_login_failures = 0
        return win

    # ---------- main loop ----------
    def run(self):
        total = len(self.records)
        self.log(f"Iniciando processamento de {total} registros (a partir do item {self.index + 1}).")

        while self.index < total:
            if self._stop_event.is_set():
                self.log("Parado pelo usuário.")
                break

            self._pause_event.wait()
            if self._stop_event.is_set():
                break

            cnpj = self.cnpj_list[self.index]

            win = self._ensure_ready()
            if win is None:
                continue

            try:
                status = self._process_one(cnpj, win)
            except Exception as exc:
                self.log(f"   [{cnpj}] Erro inesperado: {exc}")
                status = STATUS_ERROR

            self.results[cnpj] = status
            self.log(f"   [{cnpj}] -> {status.upper()}")
            self.index += 1
            self._save_checkpoint()
            self.on_progress(self.index, total, cnpj, status)

            delay = self.config.get("delay_between_cnpjs_seconds", 0.7)
            time.sleep(delay + random.uniform(0, 0.5))

        finished_all = self.index >= total
        if finished_all and not self._stop_event.is_set():
            self._retry_errors()

        self.log("Processamento finalizado." if finished_all else "Processamento interrompido.")
        return self.results

    # ---------- dual-window pipelined loop (~2x throughput) ----------
    def run_dual(self, win_a, win_b):
        """Interleave two windows: while one is waiting for its result to render,
        the other is already typing/searching its next record. Roughly halves the
        wall-clock time versus running the two windows one after another."""
        total = len(self.records)
        wait_s = self.config.get("wait_after_search_seconds", 3.5)

        if (win_a.left, win_a.top) == (win_b.left, win_b.top):
            # The two "windows" the caller found are actually the same physical
            # window (a detection mixup) -- running dual mode on this would type
            # both lanes' CNPJs into the SAME field and read mixed-up results.
            # Refuse outright instead of risking silently wrong answers.
            self.log("ERRO: as duas janelas detectadas para o modo 2 janelas são, na prática, "
                     "a mesma janela (mesma posição). Abortando o modo 2 janelas para não misturar "
                     "resultados -- rode no modo normal (1 janela) ou recalibre.")
            return self.results

        self.log(f"Iniciando processamento em modo 2 janelas de {total} registros "
                 f"(a partir do item {self.index + 1}). Janela A em ({win_a.left},{win_a.top}), "
                 f"janela B em ({win_b.left},{win_b.top}).")

        lanes = [
            {"win": win_a, "cnpj": None, "started": None, "login_failures": 0, "fills_since_check": 0},
            {"win": win_b, "cnpj": None, "started": None, "login_failures": 0, "fills_since_check": 0},
        ]

        while self.index < total or any(lane["cnpj"] for lane in lanes):
            if self._stop_event.is_set():
                self.log("Parado pelo usuário.")
                break
            self._pause_event.wait()
            if self._stop_event.is_set():
                break

            # Start a search on any idle lane that still has work available.
            paused_mid_fill = False
            for lane in lanes:
                if lane["cnpj"] is None and self.index < total:
                    citrix_utils.focus_window(lane["win"])
                    time.sleep(0.05)

                    # Checking login status is a real OCR call -- doing it on every
                    # single item eats into the speedup this mode exists for. Only
                    # do it occasionally, plus always right after start/resume.
                    due_for_check = (
                        self._force_recheck_login
                        or lane["fills_since_check"] >= LOGIN_CHECK_INTERVAL
                    )
                    if due_for_check and not self._is_logged_in(lane["win"]):
                        lane["login_failures"] += 1
                        lane["fills_since_check"] = 0
                        self.log(f"Aviso: uma janela parece deslogada "
                                 f"({lane['login_failures']}/{MAX_LOGIN_RETRIES}).")
                        if lane["login_failures"] >= MAX_LOGIN_RETRIES:
                            self.log("Pausando automaticamente para você relogar.")
                            self.pause()
                            self.on_logout()
                            paused_mid_fill = True
                            break
                        continue  # give this lane another shot next loop tick, other lane keeps going
                    if due_for_check:
                        lane["login_failures"] = 0
                        lane["fills_since_check"] = 0
                    else:
                        lane["fills_since_check"] += 1

                    cnpj = self.cnpj_list[self.index]
                    self.index += 1
                    self.log(f"   [{cnpj}] buscando na janela em {lane['win'].left},{lane['win'].top}...")
                    try:
                        self._start_search(cnpj, lane["win"])
                    except Exception as exc:
                        self.log(f"   [{cnpj}] Erro ao iniciar busca: {exc}")
                        self.results[cnpj] = STATUS_ERROR
                        self._save_checkpoint()
                        self.on_progress(self.index, total, cnpj, STATUS_ERROR)
                        continue
                    lane["cnpj"] = cnpj
                    lane["started"] = time.time()

            self._force_recheck_login = False

            if paused_mid_fill:
                continue

            # Collect any lane whose wait has elapsed.
            progressed = False
            for lane in lanes:
                if lane["cnpj"] is not None and time.time() - lane["started"] >= wait_s:
                    cnpj = lane["cnpj"]
                    try:
                        status = self._read_result(cnpj, lane["win"], verify=True)
                    except Exception as exc:
                        self.log(f"   [{cnpj}] Erro inesperado: {exc}")
                        status = STATUS_ERROR
                    if status == STATUS_ERROR:
                        # Might mean the layout shifted under us -- force a fresh
                        # detection (and an earlier login re-check) next time
                        # instead of trusting the possibly-stale cached position.
                        self._invalidate_click_cache(lane["win"])
                        lane["fills_since_check"] = LOGIN_CHECK_INTERVAL
                    self.results[cnpj] = status
                    self.log(f"   [{cnpj}] (janela em {lane['win'].left},{lane['win'].top}) -> {status.upper()}")
                    self._save_checkpoint()
                    self.on_progress(self.index, total, cnpj, status)
                    lane["cnpj"] = None
                    lane["started"] = None
                    progressed = True

            if not progressed:
                time.sleep(0.3)

        finished_all = self.index >= total and not any(lane["cnpj"] for lane in lanes)
        if finished_all and not self._stop_event.is_set():
            self._retry_errors()

        self.log("Processamento finalizado." if finished_all else "Processamento interrompido.")
        return self.results

    def _retry_errors(self):
        """Second pass: give records that errored out one more shot before calling it done."""
        error_cnpjs = [c for c, s in self.results.items() if s == STATUS_ERROR]
        if not error_cnpjs:
            return

        self.log(f"\nReprocessando {len(error_cnpjs)} registro(s) que deram erro na primeira passada...")
        for cnpj in error_cnpjs:
            if self._stop_event.is_set():
                break
            self._pause_event.wait()
            if self._stop_event.is_set():
                break

            win = self._ensure_ready()
            if win is None:
                self.log("Não foi possível continuar a nova tentativa (sessão indisponível). "
                         "Os itens restantes continuam marcados como erro.")
                break

            try:
                status = self._process_one(cnpj, win)
            except Exception as exc:
                self.log(f"   [{cnpj}] Erro na nova tentativa: {exc}")
                status = STATUS_ERROR

            self.results[cnpj] = status
            self.log(f"   [{cnpj}] (nova tentativa) -> {status.upper()}")
            self._save_checkpoint()

            delay = self.config.get("delay_between_cnpjs_seconds", 0.7)
            time.sleep(delay + random.uniform(0, 0.5))
