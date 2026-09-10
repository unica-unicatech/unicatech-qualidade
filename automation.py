"""Core automation loop: for each record, search CEP + Número in the remote app and
classify the result as KEEP (no data found -> eligible) or ELIMINATE (data found).
CNPJ is carried along purely as the output's identifying/reference column -- the
site itself is searched by CEP + Número, not by Documento/CNPJ."""
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
        actual input, so cropping around just that stays single-line."""
        x, y = self._locate_click_point(win, name)
        return x - 145, y - 22, x + 145, y + 22

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

        box = None
        if auto_detect.has_templates():
            region = (win.left, win.top, win.width, win.height)
            b = auto_detect._locate_region(name, lambda *a, **k: None, region=region)
            if b is not None:
                box = (b.left - pad, b.top - pad, b.left + b.width + pad, b.top + b.height + pad)
        if box is None:
            box = citrix_utils.absolute_region(win, self.config[name], pad=pad)

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

    def _read_field_digits_settled(self, win, field_name: str, expected: str) -> str:
        """Poll the field a few times instead of reading once immediately after
        filling it. On a laggier Citrix session, the remote screen can take longer
        than our fixed post-fill pause to visually update, so a single read right
        away can see a stale/blank field even though the paste/type already
        landed correctly -- this waits for the read to either match or stabilize
        instead of trusting the very first snapshot."""
        got = ""
        for _ in range(8):
            got = self._read_field_digits(win, field_name)
            if got == expected:
                return got
            time.sleep(0.35)
        return got

    def _fill_field(self, win, field_name: str, value: str):
        """Fill the field and OCR-read it back, retrying if it doesn't match.

        Always fills via clipboard paste (one paste action for the whole value)
        when available -- proven reliable across machines/sessions, unlike
        simulating individual keystrokes: on some PCs (seen with Citrix over a
        laggier connection), typed keystrokes get registered twice by the remote
        app regardless of typing speed (e.g. '776' arrives as '777766'), so typing
        is only ever used as a last-resort fallback when clipboard isn't
        available, never as a "retry" for a paste that just needed more time to
        show up on screen (see _read_field_digits_settled)."""
        x, y = self._locate_click_point(win, field_name)
        if not value:
            pyautogui.click(x, y, duration=MOUSE_MOVE_SECONDS)
            time.sleep(0.15)
            pyautogui.hotkey("ctrl", "a")
            pyautogui.press("delete")
            time.sleep(0.1)
            return

        attempts = 3
        for attempt in range(attempts):
            pyautogui.click(x, y, duration=MOUSE_MOVE_SECONDS)
            time.sleep(0.15)
            pyautogui.hotkey("ctrl", "a")
            pyautogui.press("delete")
            time.sleep(0.1)
            if _HAS_CLIPBOARD:
                _set_clipboard_text(value)
                time.sleep(0.05)
                pyautogui.hotkey("ctrl", "v")
            else:
                interval = 0.05 + attempt * 0.05
                pyautogui.typewrite(value, interval=interval)
            time.sleep(0.15)
            got = self._read_field_digits_settled(win, field_name, value)
            if got == value:
                return
            debug_path = self._save_field_debug_shot(win, field_name, attempt)
            self.log(f"   Aviso: campo {field_name} leu '{got}' (esperado '{value}'), "
                     f"tentativa {attempt + 1}/{attempts}. Print salvo em: {debug_path}")
        self.log(f"   Aviso: campo {field_name} pode ter ficado incorreto após {attempts} tentativas.")

    def _save_field_debug_shot(self, win, field_name: str, attempt: int) -> str:
        """Save exactly the region the OCR verification is reading, so a failed
        readback can be diagnosed from the actual pixels instead of guesswork
        (e.g. a click/coordinate offset -- possible on a PC with display scaling
        different from 100% -- would show up here as the crop landing on the
        wrong spot, or the field looking empty/blank in the image itself)."""
        try:
            debug_dir = OUTPUT_DIR / "debug_field_reads"
            debug_dir.mkdir(exist_ok=True)
            x1, y1, x2, y2 = self._field_ocr_box(win, field_name)
            img = ocr_utils.grab_region(x1, y1, x2, y2)
            path = debug_dir / f"{field_name}_{int(time.time())}_{attempt}.png"
            img.save(path)
            return str(path)
        except Exception as exc:
            return f"(falha ao salvar print: {exc})"

    def _start_search(self, cnpj: str, win):
        """Fill CEP + Número and hit Pesquisar. Does NOT wait for the result --
        split out so a second window can be worked while this one is still loading
        (see run_dual)."""
        record = self._records_by_cnpj[cnpj]

        self._fill_field(win, "cep_field", record["cep"])
        self._fill_field(win, "numero_field", record["numero"])

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
                self.log(f"   [{cnpj}] (debug) tabela lida como COM DADOS. Texto OCR: '{text}'")
                return STATUS_ELIMINATE
            self.log(f"   [{cnpj}] Tabela ainda sem texto legível (tentativa {attempt}/{OCR_RETRIES}), "
                     f"melhor leitura: '{text}'. Aguardando mais um pouco...")
            time.sleep(1.5)

        self.log(f"   [{cnpj}] Falha final de OCR. Última leitura: '{best_text}'")
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
