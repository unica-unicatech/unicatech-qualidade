"""Tries to relocate the calibrated UI elements automatically by matching template
images saved by calibration.py against the current screen. Used by the 'Testar
calibração automática' button, and by automation.py on every search (to click the
right spot even if the layout has shifted a bit).

Each element can have MULTIPLE saved appearance variants (e.g. the site's normal
wide layout vs. its narrow/responsive layout, which changes how some buttons look).
Variants are stored as templates/<name>/1.png, 2.png, ... and are all tried in
turn -- new variants only ever get ADDED (see calibration.py), so recalibrating
for a new layout never throws away one that already works for another layout.
"""
import json
import time
from pathlib import Path

import pyautogui

import citrix_utils

TEMPLATES_DIR = Path(__file__).parent / "templates"
CONFIG_PATH = Path(__file__).parent / "config.json"

# The session watermark (timestamp/IP text) drifts across the page and occasionally
# overlaps one of these elements right at screenshot time, which can drop the match
# confidence just enough to miss. A lower threshold + a couple of quick retries
# (the watermark won't be in the exact same spot a few hundred ms later) makes this
# reliable without needing a stricter/slower detection method.
CONFIDENCE = 0.82
RETRIES = 3
RETRY_DELAY_SECONDS = 0.4

POINT_ELEMENTS = ["cep_field", "numero_field", "pesquisar_button"]
REGION_ELEMENTS = ["app_marker_region", "table_row_region"]


def _variant_dir(name) -> Path:
    return TEMPLATES_DIR / name


def variant_paths(name):
    """All saved appearance variants for an element, e.g. templates/pesquisar_button/1.png, 2.png..."""
    d = _variant_dir(name)
    if not d.exists():
        return []
    return sorted(d.glob("*.png"), key=lambda p: p.stem)


def next_variant_path(name) -> Path:
    """Where the NEXT new variant for this element should be saved (never overwrites)."""
    d = _variant_dir(name)
    d.mkdir(parents=True, exist_ok=True)
    existing = [int(p.stem) for p in d.glob("*.png") if p.stem.isdigit()]
    n = (max(existing) + 1) if existing else 1
    return d / f"{n}.png"


def _click_ratio(png_path: Path):
    """Where inside the matched template box the actual click point sits, as a
    (x_ratio, y_ratio) fraction of (width, height) from the top-left. Defaults to
    dead-center (0.5, 0.5), which is correct when the template was cropped
    symmetrically around the click point (calibration.py's own convention).
    Templates that include extra context (e.g. a label above the input box) can
    ship a sidecar '<n>.click.json' next to '<n>.png' with a different ratio."""
    sidecar = png_path.with_suffix(".click.json")
    if sidecar.exists():
        try:
            data = json.loads(sidecar.read_text(encoding="utf-8"))
            return data["x_ratio"], data["y_ratio"]
        except Exception:
            pass
    return 0.5, 0.5


def _locate_box(name, log, region=None):
    """Try every saved appearance variant of `name`, retrying a few full passes
    (transient watermark overlap can drop a match momentarily). Pass `region`
    (left, top, width, height) to constrain the search to a specific window --
    important in 2-window mode so lane A's search can't match lane B's element.
    Returns (pyscreeze Box, variant_path) or (None, None)."""
    paths = variant_paths(name)
    if not paths:
        log(f"  Sem template salvo para '{name}'.")
        return None, None
    for attempt in range(1, RETRIES + 1):
        for path in paths:
            try:
                kwargs = {"confidence": CONFIDENCE}
                if region is not None:
                    kwargs["region"] = region
                box = pyautogui.locateOnScreen(str(path), **kwargs)
            except Exception:
                box = None
            if box is not None:
                return box, path
        if attempt < RETRIES:
            time.sleep(RETRY_DELAY_SECONDS)
    log(f"  Não encontrei '{name}' na tela após tentar {len(paths)} variante(s) de aparência "
        f"({RETRIES}x cada). Pode ser um layout novo -- rode a calibração manual pra ensinar essa aparência.")
    return None, None


def _locate_point(name, log, region=None):
    box, path = _locate_box(name, log, region=region)
    if box is None:
        return None
    rx, ry = _click_ratio(path)
    x = box.left + box.width * rx
    y = box.top + box.height * ry
    return pyautogui.Point(int(x), int(y))


def _locate_region(name, log, region=None):
    box, _path = _locate_box(name, log, region=region)
    return box


def has_templates() -> bool:
    if not TEMPLATES_DIR.exists():
        return False
    return all(variant_paths(n) for n in POINT_ELEMENTS + REGION_ELEMENTS)


def try_auto_calibrate(log=print) -> bool:
    """Returns True if every element was found and config.json was (re)written."""
    if not has_templates():
        log("Ainda não há templates salvos (rode a calibração manual pelo menos uma vez primeiro).")
        return False

    log("Procurando os elementos na tela pela aparência salva (sem precisar clicar)...")

    cep_pos = _locate_point("cep_field", log)
    if cep_pos is None:
        return False
    log(f"  Campo CEP encontrado em {cep_pos}.")

    win = citrix_utils.find_window_at_point(cep_pos.x, cep_pos.y)
    if win is None:
        log("  Encontrei o campo CEP, mas não identifiquei a janela por trás dele.")
        return False
    log(f"  Janela identificada: '{win.title}' em {win.left},{win.top}.")

    numero_pos = _locate_point("numero_field", log)
    if numero_pos is None:
        return False
    log(f"  Campo Número encontrado em {numero_pos}.")

    btn_pos = _locate_point("pesquisar_button", log)
    if btn_pos is None:
        return False
    log(f"  Botão Pesquisar encontrado em {btn_pos}.")

    marker_box = _locate_region("app_marker_region", log)
    if marker_box is None:
        return False
    log(f"  Marcador de sessão (rótulo DOCUMENTO) encontrado em {marker_box}.")

    row_box = _locate_region("table_row_region", log)
    if row_box is None:
        return False
    log(f"  Região da tabela encontrada em {row_box}.")

    old_config = {}
    if CONFIG_PATH.exists():
        try:
            old_config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass

    config = {
        "window_title_contains": citrix_utils.stable_title_anchor(win.title),
        "wait_after_search_seconds": old_config.get("wait_after_search_seconds", 3.5),
        "delay_between_cnpjs_seconds": old_config.get("delay_between_cnpjs_seconds", 0.7),
        "tesseract_cmd": old_config.get("tesseract_cmd", r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
        "cep_field": {"x": int(cep_pos.x - win.left), "y": int(cep_pos.y - win.top)},
        "numero_field": {"x": int(numero_pos.x - win.left), "y": int(numero_pos.y - win.top)},
        "pesquisar_button": {"x": int(btn_pos.x - win.left), "y": int(btn_pos.y - win.top)},
        "app_marker_region": {
            "x1": int(marker_box.left - win.left), "y1": int(marker_box.top - win.top),
            "x2": int(marker_box.left + marker_box.width - win.left),
            "y2": int(marker_box.top + marker_box.height - win.top),
        },
        "table_row_region": {
            "x1": int(row_box.left - win.left), "y1": int(row_box.top - win.top),
            "x2": int(row_box.left + row_box.width - win.left),
            "y2": int(row_box.top + row_box.height - win.top),
        },
    }

    CONFIG_PATH.write_text(json.dumps(config, indent=2, ensure_ascii=False), encoding="utf-8")
    log(f"Calibração automática concluída e salva em {CONFIG_PATH}.")
    return True


def locate_all_windows(log=print):
    """Find every on-screen instance of the app (matched by the DOCUMENTO label
    template, trying all its saved appearance variants) and return their underlying
    windows, sorted left-to-right.

    Used for the 2-window ('pipeline') mode: since both windows have identical
    layout, the single calibrated set of relative offsets in config.json applies
    to each window found here -- we just need each window's own top-left corner.
    """
    paths = variant_paths("app_marker_region")
    if not paths:
        log("Sem template salvo para 'app_marker_region' (rode a calibração ao menos uma vez).")
        return []

    all_boxes = []
    for path in paths:
        try:
            all_boxes.extend(pyautogui.locateAllOnScreen(str(path), confidence=CONFIDENCE))
        except Exception as exc:
            log(f"Erro ao procurar instâncias da janela com '{path.name}': {exc}")

    # The scan can yield several overlapping boxes for the same real match (and
    # different variants can both match the same window); cluster by proximity and
    # keep one representative center per cluster.
    centers = []
    for b in all_boxes:
        cx, cy = b.left + b.width / 2, b.top + b.height / 2
        if not any(abs(cx - ex) < 60 and abs(cy - ey) < 60 for ex, ey in centers):
            centers.append((cx, cy))

    windows = []
    seen = set()
    for cx, cy in centers:
        win = citrix_utils.find_window_at_point(int(cx), int(cy))
        if win is None:
            continue
        key = (win.left, win.top, win.width, win.height, win.title)
        if key in seen:
            continue
        seen.add(key)
        windows.append(win)

    windows.sort(key=lambda w: w.left)
    return windows
