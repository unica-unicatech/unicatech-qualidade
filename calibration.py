"""Calibration wizard.

Walks the user through pointing the mouse at each required screen element inside
the remote (Citrix) window and pressing F8 to capture it. All captured points are
stored RELATIVE to the window's top-left corner, so the automation keeps working
even if the window later opens in a different screen position (e.g. after a
re-login), as long as its internal layout/zoom doesn't change.

As a side effect, this also saves small screenshot "templates" of each element to
templates/*.png. Those are used by auto_detect.py to try to relocate the elements
automatically next time (by appearance), so you don't have to manually recalibrate
after every re-login -- manual calibration here is both the primary setup step and
the way to (re)generate the backup templates for auto-detection.
"""
import json
import time
from pathlib import Path

import keyboard
import pyautogui
from PIL import ImageGrab

import auto_detect
import citrix_utils

CONFIG_PATH = Path(__file__).parent / "config.json"
CAPTURE_KEY = "f8"
CANCEL_KEY = "esc"

POINT_PAD_X = 90
POINT_PAD_Y = 16


def _wait_for_capture(log):
    """Block until the user presses F8 (capture) or ESC (cancel). Returns mouse pos or None."""
    log(f"   Posicione o mouse e pressione [{CAPTURE_KEY.upper()}] para capturar "
        f"(ou [{CANCEL_KEY.upper()}] para cancelar a calibração).")
    while True:
        if keyboard.is_pressed(CAPTURE_KEY):
            pos = pyautogui.position()
            while keyboard.is_pressed(CAPTURE_KEY):
                time.sleep(0.02)
            return pos
        if keyboard.is_pressed(CANCEL_KEY):
            while keyboard.is_pressed(CANCEL_KEY):
                time.sleep(0.02)
            return None
        time.sleep(0.02)


def _save_template(name: str, box, log, click_ratio=None):
    """Save this capture as a NEW appearance variant for `name` -- never overwrites
    an existing one, so calibrating for a different layout (e.g. a narrower window)
    doesn't throw away the variant that already works for the normal layout."""
    path = auto_detect.next_variant_path(name)
    img = ImageGrab.grab(bbox=box, all_screens=True)
    img.save(path)
    if click_ratio is not None:
        sidecar = path.with_suffix(".click.json")
        sidecar.write_text(
            json.dumps({"x_ratio": click_ratio[0], "y_ratio": click_ratio[1]}), encoding="utf-8"
        )
    log(f"   Template salvo: {name}/{path.name} (variante nova, as anteriores continuam valendo)")


def run_calibration(log=print):
    """Run the interactive calibration wizard. `log` receives status strings.

    Returns True on success (config.json written), False if cancelled.
    """
    log("=== CALIBRAÇÃO ===")
    log("Deixe a janela do Citrix com o site Vivo Qualidade aberta e visível.")
    log("No primeiro passo, a janela será identificada pela posição do MOUSE no "
        "momento em que você apertar F8 (não importa qual janela está em foco agora).")

    config = {
        "wait_after_search_seconds": 3.5,
        "delay_between_cnpjs_seconds": 0.7,
        "tesseract_cmd": "C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
    }

    window_left = window_top = None
    window_title = None

    steps = [
        ("cep_field", "point", "Posicione o mouse sobre o CAMPO CEP."),
        ("numero_field", "point", "Posicione o mouse sobre o CAMPO NÚMERO."),
        ("pesquisar_button", "point", "Posicione o mouse sobre o BOTÃO PESQUISAR."),
        ("row_region_tl", "point", "Posicione o mouse no CANTO SUPERIOR ESQUERDO da PRIMEIRA LINHA da tabela "
                                    "(logo abaixo dos cabeçalhos MOVIMENTO/DATA/... — onde aparece "
                                    "'No data available in table' quando vazio, ou o primeiro resultado). "
                                    "Deixe uma margem folgada acima/à esquerda."),
        ("row_region_br", "point", "Posicione o mouse no CANTO INFERIOR DIREITO dessa mesma primeira linha "
                                    "da tabela. Deixe uma margem folgada abaixo/à direita."),
        ("marker_tl", "point", "Posicione o mouse no CANTO SUPERIOR ESQUERDO do rótulo 'DOCUMENTO' "
                                "(usado para detectar se a sessão deslogou)."),
        ("marker_br", "point", "Posicione o mouse no CANTO INFERIOR DIREITO desse mesmo rótulo 'DOCUMENTO'."),
    ]

    captured = {}
    captured_abs = {}
    for i, (key, kind, instruction) in enumerate(steps):
        log(f"\n> {instruction}")
        pos = _wait_for_capture(log)
        if pos is None:
            log("Calibração cancelada pelo usuário.")
            return False

        if i == 0:
            win = citrix_utils.find_window_at_point(pos.x, pos.y)
            if win is None:
                log("Não consegui identificar a janela sob o mouse. Tente novamente, "
                    "certifique-se de que o mouse está sobre a janela do Citrix.")
                return False
            window_left, window_top = win.left, win.top
            window_title = citrix_utils.stable_title_anchor(win.title)
            log(f"   Janela identificada: '{win.title}' (canto superior-esquerdo em "
                f"{window_left},{window_top}, tamanho {win.width}x{win.height})")
            log(f"   Usando '{window_title}' como referência estável do título.")
            config["window_title_contains"] = window_title

        captured[key] = (pos.x - window_left, pos.y - window_top)
        captured_abs[key] = (pos.x, pos.y)
        log(f"   Capturado: offset relativo à janela = {captured[key]}")

    config["cep_field"] = {"x": captured["cep_field"][0], "y": captured["cep_field"][1]}
    config["numero_field"] = {"x": captured["numero_field"][0], "y": captured["numero_field"][1]}
    config["pesquisar_button"] = {"x": captured["pesquisar_button"][0], "y": captured["pesquisar_button"][1]}
    config["table_row_region"] = {
        "x1": captured["row_region_tl"][0], "y1": captured["row_region_tl"][1],
        "x2": captured["row_region_br"][0], "y2": captured["row_region_br"][1],
    }
    config["app_marker_region"] = {
        "x1": captured["marker_tl"][0], "y1": captured["marker_tl"][1],
        "x2": captured["marker_br"][0], "y2": captured["marker_br"][1],
    }

    CONFIG_PATH.write_text(json.dumps(config, indent=2, ensure_ascii=False), encoding="utf-8")
    log(f"\nCalibração salva em {CONFIG_PATH}")

    log("\nSalvando templates de imagem (backup para tentativa de calibração automática)...")
    try:
        cx, cy = captured_abs["cep_field"]
        _save_template("cep_field", (cx - POINT_PAD_X, cy - POINT_PAD_Y, cx + POINT_PAD_X, cy + POINT_PAD_Y),
                        log, click_ratio=(0.5, 0.5))

        nx, ny = captured_abs["numero_field"]
        _save_template("numero_field", (nx - POINT_PAD_X, ny - POINT_PAD_Y, nx + POINT_PAD_X, ny + POINT_PAD_Y),
                        log, click_ratio=(0.5, 0.5))

        bx, by = captured_abs["pesquisar_button"]
        _save_template("pesquisar_button", (bx - POINT_PAD_X, by - POINT_PAD_Y, bx + POINT_PAD_X, by + POINT_PAD_Y),
                        log, click_ratio=(0.5, 0.5))

        rx1, ry1 = captured_abs["row_region_tl"]
        rx2, ry2 = captured_abs["row_region_br"]
        _save_template("table_row_region", (min(rx1, rx2), min(ry1, ry2), max(rx1, rx2), max(ry1, ry2)), log)

        mx1, my1 = captured_abs["marker_tl"]
        mx2, my2 = captured_abs["marker_br"]
        _save_template("app_marker_region", (min(mx1, mx2), min(my1, my2), max(mx1, mx2), max(my1, my2)), log)
    except Exception as exc:
        log(f"   Aviso: não consegui salvar todos os templates ({exc}). "
            f"A calibração manual continua válida, só a automática pode não funcionar depois.")

    return True


def load_config():
    if not CONFIG_PATH.exists():
        return None
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))


def calibrate_result_banners(log=print) -> bool:
    """Optional, separate mini-calibration for the two result banners ('No data
    available in table' / the red invalid-CEP-Número banner). Kept apart from
    run_calibration() -- it requires the site to actually be SHOWING each state
    at capture time (a real empty search, then a real invalid one), which isn't
    true during normal calibration, and skipping it shouldn't force the whole
    main wizard to be redone.

    Once captured, automation.py recognizes these states by their on-screen
    APPEARANCE (like it already does for buttons/fields) instead of trying to
    OCR-read the row's text -- much more tolerant of a watermark overlapping it,
    since it doesn't need to make out individual characters, just the overall
    shape. Without these templates, automation.py falls back to the older
    OCR-based classification, so this is safe to skip or redo at any time.
    """
    log("=== CALIBRAÇÃO DOS AVISOS DE RESULTADO (opcional) ===")
    log("Isso ensina o sistema a RECONHECER (por aparência, não por texto) os avisos "
        "que a tabela mostra quando não há dado ou quando o CEP/Número é inválido -- "
        "mais confiável que ler o texto quando a marca d'água da sessão atrapalha.")

    log("\n> Primeiro, faça uma busca com um CEP/Número que você sabe que NÃO retorna "
        "dado nenhum (mostra 'No data available in table'). Deixe essa tela visível.")
    log("> Posicione o mouse no CANTO SUPERIOR ESQUERDO do aviso 'No data available in table'.")
    tl = _wait_for_capture(log)
    if tl is None:
        log("Cancelado.")
        return False
    log("> Agora o CANTO INFERIOR DIREITO desse mesmo aviso.")
    br = _wait_for_capture(log)
    if br is None:
        log("Cancelado.")
        return False
    box = (min(tl.x, br.x), min(tl.y, br.y), max(tl.x, br.x), max(tl.y, br.y))
    _save_template("no_data_banner", box, log)

    log("\n> Agora faça uma busca com um CEP ou Número INVÁLIDO (ex: com letras), "
        "pra ver o aviso vermelho de 'preencha o documento... /número inválido'. "
        "Deixe essa tela visível.")
    log("> Posicione o mouse no CANTO SUPERIOR ESQUERDO desse aviso vermelho.")
    tl2 = _wait_for_capture(log)
    if tl2 is None:
        log("Cancelado (o aviso de 'sem dados' já foi salvo, só esse ficou de fora).")
        return False
    log("> Agora o CANTO INFERIOR DIREITO desse aviso vermelho.")
    br2 = _wait_for_capture(log)
    if br2 is None:
        log("Cancelado (o aviso de 'sem dados' já foi salvo, só esse ficou de fora).")
        return False
    box2 = (min(tl2.x, br2.x), min(tl2.y, br2.y), max(tl2.x, br2.x), max(tl2.y, br2.y))
    _save_template("invalid_input_banner", box2, log)

    log("\nCalibração dos avisos concluída.")
    return True


if __name__ == "__main__":
    run_calibration()
