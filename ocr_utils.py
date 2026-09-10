"""Small OCR helpers tuned for reading short UI strings off a screenshot region."""
import difflib
import re
from collections import Counter

import numpy as np
import pytesseract
from PIL import Image, ImageGrab, ImageOps

ENTRIES_RE = re.compile(r"[o0]f\D{0,4}(\d+)\D{0,6}entries", re.IGNORECASE)
PSM_MODES = (7, 6, 11, 13)
PSM_MODES_FAST = (7, 6)  # the two that matter most in practice; used on the first (most likely to succeed) attempt
NO_DATA_PHRASE = "no data available in table"
# The site shows a red banner instead of a table when CEP/Número is missing or
# invalid (e.g. letters in Número) -- these records aren't "no data" or "has data",
# they're a bad INPUT and should be flagged separately so they don't get silently
# misread as either. Using the two banner lines joined into ONE long target (instead
# of matching several short phrases separately) matters: short phrases occasionally
# false-positive-match garbled OCR noise from a REAL data row (observed ~0.55-0.58
# ratio, right at the old threshold); the long combined phrase keeps genuine matches
# (even a partial, one-line capture) around 0.6-1.0 while real-data noise drops to
# ~0.3, giving a much safer margin at the same 0.55 threshold.
INVALID_INPUT_PHRASE = (
    "preencha o documento do cliente ou o cep numero do logradouro "
    "numero do endereco invalido o numero do endereco nao pode conter letras"
)


def grab_region(x1, y1, x2, y2) -> Image.Image:
    if x2 < x1:
        x1, x2 = x2, x1
    if y2 < y1:
        y1, y2 = y2, y1
    # all_screens=True is required on Windows to capture monitors positioned with
    # negative coordinates (e.g. a secondary monitor to the left of the primary) --
    # without it, regions outside the primary monitor come back blank.
    return ImageGrab.grab(bbox=(x1, y1, x2, y2), all_screens=True)


def grab_region_gray(x1, y1, x2, y2) -> np.ndarray:
    """Cheap grayscale snapshot of a region, for detecting whether it visibly
    changed between two points in time (see region_changed) -- no OCR involved,
    just pixel comparison."""
    img = grab_region(x1, y1, x2, y2)
    return np.array(ImageOps.grayscale(img), dtype=np.float32)


def region_changed(before: np.ndarray, after: np.ndarray, threshold: float = 8.0) -> bool:
    """True if `after` looks meaningfully different from `before`. Used to detect
    that the results table actually refreshed after a new search, instead of
    trusting whatever's on screen the instant our fixed post-click wait elapses
    -- on a slower page load, that can still be the PREVIOUS search's table,
    read and classified as if it were this record's result."""
    if before is None or after is None or before.shape != after.shape:
        return True
    return bool(np.abs(before - after).mean() > threshold)


def _upscale(img: Image.Image, scale: int = 4) -> Image.Image:
    return img.resize((max(1, img.width * scale), max(1, img.height * scale)), Image.LANCZOS)


def _preprocess_variants(img: Image.Image, fast: bool = False):
    """Yield several preprocessed versions of the same crop. Small, low-contrast UI
    text (especially with something like a watermark crossing it) often needs more
    than one binarization strategy before Tesseract reads it cleanly.

    `fast=True` yields only the 2 variants that empirically cover the vast majority
    of reads, for the common/easy case; the full set (4 variants) is the thorough
    fallback for retries once something's already proven harder to read."""
    big = _upscale(img)
    gray = ImageOps.grayscale(big)

    autocon = ImageOps.autocontrast(gray)
    yield ImageOps.expand(autocon, border=16, fill=255)

    arr = np.array(gray, dtype=np.float32)
    thresh = arr.mean()
    binarized = gray.point(lambda p: 255 if p > thresh else 0)
    yield ImageOps.expand(binarized, border=16, fill=255)

    if fast:
        return

    binarized_inv = gray.point(lambda p: 0 if p > thresh else 255)
    yield ImageOps.expand(binarized_inv, border=16, fill=255)

    # slightly stricter threshold, helps when text is thin/light
    thresh_strict = arr.mean() + (arr.max() - arr.mean()) * 0.15
    strict = gray.point(lambda p: 255 if p > thresh_strict else 0)
    yield ImageOps.expand(strict, border=16, fill=255)


def read_text(x1, y1, x2, y2, psm: int = 7) -> str:
    """Single-pass read (used for the login/session marker check)."""
    img = grab_region(x1, y1, x2, y2)
    img = _upscale(img)
    img = ImageOps.grayscale(img)
    img = ImageOps.autocontrast(img)
    return pytesseract.image_to_string(img, config=f"--psm {psm}").strip()


def read_digits(x1, y1, x2, y2, psm: int = 7) -> str:
    """Read a region that should contain ONLY digits (e.g. a CEP field), restricting
    Tesseract's character set to 0-9. Much more reliable than general text OCR for
    this -- it can't confuse a digit for a stray letter/symbol."""
    img = grab_region(x1, y1, x2, y2)
    img = _upscale(img)
    img = ImageOps.grayscale(img)
    img = ImageOps.autocontrast(img)
    config = f"--psm {psm} -c tessedit_char_whitelist=0123456789"
    return pytesseract.image_to_string(img, config=config).strip()


def read_entries_total(x1, y1, x2, y2):
    """Try several preprocessing variants x OCR modes to reliably parse the
    'Showing X to Y of Z entries' counter. Returns (total_or_None, best_raw_text)."""
    img = grab_region(x1, y1, x2, y2)
    best_text = ""
    for variant in _preprocess_variants(img):
        for psm in PSM_MODES:
            text = pytesseract.image_to_string(variant, config=f"--psm {psm}").strip()
            if len(text) > len(best_text):
                best_text = text
            total = parse_entries_total(text)
            if total is not None:
                return total, text
    return None, best_text


def _best_substring_match(text: str, target: str):
    """Like _best_substring_ratio, but also returns the (start, end) window that
    scored best, so the caller can strip it out of `text` afterward."""
    n = len(target)
    if n == 0:
        return 0.0, 0, 0
    if len(text) <= n:
        return difflib.SequenceMatcher(None, text, target).ratio(), 0, len(text)
    best, best_start = 0.0, 0
    step = max(1, n // 4)
    for i in range(0, len(text) - n + 1, step):
        ratio = difflib.SequenceMatcher(None, text[i:i + n], target).ratio()
        if ratio > best:
            best, best_start = ratio, i
    return best, best_start, best_start + n


def _best_substring_ratio(text: str, target: str) -> float:
    """Best similarity of `target` against any same-length window of `text`.
    Comparing the whole (possibly noisy) text against a short target phrase dilutes
    the ratio when there's a lot of surrounding junk (e.g. watermark garbage); this
    instead slides a target-sized window across the text and keeps the best match,
    so an embedded match isn't drowned out by noise elsewhere in the string."""
    ratio, _, _ = _best_substring_match(text, target)
    return ratio


# The results table's column header row ("MOVIMENTO DATA NOME DO CLIENTE...") is
# static -- always on screen regardless of whether the table actually has data --
# and long enough on its own to pass classify_table_region's min_chars gate. If
# the crop also picks up some watermark noise but the real content underneath
# (a genuine data row, or "No data available") doesn't come through legibly, the
# header alone would otherwise be enough to wrongly default to "data" (seen in
# practice: a record eliminated based on header + noise, with no real evidence of
# an actual result). It's stripped out before that final decision so only
# genuinely per-record content counts.
HEADER_PHRASE = (
    "movimento data nome do cliente endereco complemento bairro cidade cep "
    "parque plano origem tecnologia"
)


def classify_table_region(x1, y1, x2, y2, min_chars: int = 6, match_threshold: float = 0.55,
                           fast: bool = False):
    """Classify the first row of the results table.

    Returns (kind, best_text) where kind is:
      - "empty"        -> matched 'No data available in table' (no result -> keep CNPJ)
      - "invalid_input" -> matched the red 'preencha/inválido' banner (bad CEP/Número
                            in the source spreadsheet, not a real search outcome)
      - "data"         -> readable text that does NOT match either phrase above (data found -> eliminate)
      - "inconclusive" -> not enough legible text yet (probably still loading) -> caller should retry

    `fast=True` tries fewer preprocessing/PSM combinations -- use it for the first,
    most-likely-to-succeed attempt; let retries fall back to the thorough (default)
    sweep once something's already proven harder to read.
    """
    img = grab_region(x1, y1, x2, y2)
    empty_target = re.sub(r"[^a-z]", "", NO_DATA_PHRASE)
    invalid_target = re.sub(r"[^a-z]", "", INVALID_INPUT_PHRASE)
    psm_modes = PSM_MODES_FAST if fast else PSM_MODES
    best_text = ""
    # (newline_count, -length): the table's real row is genuinely one line, so a
    # cleanly-read single-line result is trusted over a longer but messier one.
    # This matters because a scattered/sparse PSM mode can dump the SAME content
    # one word per line (many newlines) -- and that shape, despite being real
    # data, can coincidentally score close to the "no data available in table"
    # phrase (shared words like "data"/"table" landing on their own lines), while
    # the clean single-line read of the exact same row scores clearly low. Picking
    # "longest text wins" as the candidate to trust was choosing the noisier,
    # more misleading reading over the cleaner one.
    best_key = None
    for variant in _preprocess_variants(img, fast=fast):
        for psm in psm_modes:
            text = pytesseract.image_to_string(variant, config=f"--psm {psm}").strip()
            key = (text.count("\n"), -len(text))
            if best_key is None or key < best_key:
                best_key = key
                best_text = text
            cleaned = re.sub(r"[^a-z]", "", text.lower())
            if len(cleaned) >= min_chars:
                if _best_substring_ratio(cleaned, empty_target) >= match_threshold:
                    return "empty", text
                if _best_substring_ratio(cleaned, invalid_target) >= match_threshold:
                    return "invalid_input", text

    # A "suspiciously close but not quite confirmed" match against the empty
    # phrase (seen in practice: 0.45, e.g. a watermark partially obscuring "No
    # data available in table") is much more likely to mean "actually empty,
    # just hard to read" than "actually real data" -- so it should NOT fall
    # through to the "data" default below. Only checked against the single
    # best_text chosen above (not the max over every variant tried) -- checking
    # every variant let one unlucky noisy reading veto an otherwise-clean "data"
    # result (see the newline-preference comment above for the concrete case).
    cleaned_best = re.sub(r"[^a-z]", "", best_text.lower())
    if len(cleaned_best) >= min_chars and _best_substring_ratio(cleaned_best, empty_target) >= 0.40:
        return "inconclusive", best_text

    residual = cleaned_best
    header_target = re.sub(r"[^a-z]", "", HEADER_PHRASE)
    header_ratio, hstart, hend = _best_substring_match(cleaned_best, header_target)
    if header_ratio >= match_threshold:
        residual = cleaned_best[:hstart] + cleaned_best[hend:]
    if len(residual) >= min_chars:
        return "data", best_text
    return "inconclusive", best_text


def parse_entries_total(text: str):
    """Extract the Z in 'Showing X to Y of Z entries'. Returns int or None if not found."""
    match = ENTRIES_RE.search(text.replace(",", ""))
    if match:
        return int(match.group(1))
    return None


def fuzzy_contains(text: str, expected: str, threshold: float = 0.7) -> bool:
    """Order-independent, noise-tolerant check: are most letters of `expected`
    present in `text`? Handles OCR letter-scrambling caused by things like a
    diagonal watermark crossing through the text (e.g. 'DNOCLIMENTO' vs 'DOCUMENTO').
    """
    t = re.sub(r"[^a-z]", "", text.lower())
    e = re.sub(r"[^a-z]", "", expected.lower())
    if not e:
        return False
    overlap = sum((Counter(t) & Counter(e)).values())
    return (overlap / len(e)) >= threshold
