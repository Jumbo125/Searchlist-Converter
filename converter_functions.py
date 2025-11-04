import csv
import os
import re
import sys
import tempfile
from typing import Iterable, List, Optional, Sequence, Tuple

from PIL import Image, ImageDraw, ImageFont
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    import pyphen  # type: ignore
except Exception:  # pragma: no cover - optional dependency
    pyphen = None  # läuft auch ohne, dann keine Silbentrennung

DPI = 300
A4_MM = (210, 297)
HEADER_HARD_WRAP_CHARS = 5

COLOR_PRESETS = [
    ("Hellgrau", "#F5F5F5"),
    ("Wolkenweiß", "#FFFFFF"),
    ("Blassblau", "#EEF4FB"),
    ("Eisblau", "#F7FBFF"),
    ("Mint", "#ECF7F2"),
    ("Salbei", "#F1F7F2"),
    ("Lavendel", "#F3F0FA"),
    ("Altrosa", "#FDF0F3"),
    ("Sand", "#F6F1EB"),
]
PRESET_NAMES = [name for name, _ in COLOR_PRESETS] + ["Benutzerdefiniert…"]
NAME_TO_HEX = {name: hex_value for name, hex_value in COLOR_PRESETS}

HYPHENATOR: Optional["pyphen.Pyphen"] = None
HYPHENATE_HEADERS = False


def configure_hyphenation(enable_headers: bool, hyphenator: Optional["pyphen.Pyphen"]) -> None:
    global HYPHENATOR, HYPHENATE_HEADERS
    HYPHENATOR = hyphenator
    HYPHENATE_HEADERS = enable_headers


def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def mm_to_px(mm: float, dpi: int = DPI) -> int:
    return int(round(mm / 25.4 * dpi))


def a4_pixels(orientation: str = "portrait", dpi: int = DPI) -> Tuple[int, int]:
    w_mm, h_mm = A4_MM
    if orientation == "landscape":
        w_mm, h_mm = h_mm, w_mm
    return mm_to_px(w_mm, dpi), mm_to_px(h_mm, dpi)


def hex_to_rgb(hex_value: str) -> Tuple[int, int, int]:
    try:
        cleaned = hex_value.strip()
        if cleaned.startswith("#"):
            cleaned = cleaned[1:]
        if len(cleaned) == 3:
            cleaned = "".join(ch * 2 for ch in cleaned)
        return tuple(int(cleaned[i:i + 2], 16) for i in (0, 2, 4))
    except Exception:
        return 245, 245, 245


def _choose_existing(paths: Iterable[str]) -> Optional[str]:
    return next((p for p in paths if os.path.exists(p)), None)


def load_fonts(size_body: int = 24, size_header: int = 28) -> Tuple[ImageFont.FreeTypeFont, ImageFont.FreeTypeFont]:
    candidates = [
        (r"C:\Windows\Fonts\segoeui.ttf", r"C:\Windows\Fonts\segoeuib.ttf"),
        (r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\arialbd.ttf"),
        (r"C:\Windows\Fonts\calibri.ttf", r"C:\Windows\Fonts\calibrib.ttf"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/System/Library/Fonts/Supplemental/Arial.ttf",
         "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
        ("/System/Library/Fonts/Helvetica.ttc", "/System/Library/Fonts/Helvetica.ttc"),
    ]
    for normal, bold in candidates:
        normal_path = _choose_existing([normal])
        bold_path = _choose_existing([bold])
        if normal_path and bold_path:
            try:
                return (
                    ImageFont.truetype(normal_path, size_body),
                    ImageFont.truetype(bold_path, size_header),
                )
            except Exception:
                continue
    return ImageFont.load_default(), ImageFont.load_default()


def text_size(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont) -> Tuple[int, int]:
    bbox = draw.textbbox((0, 0), str(text), font=font)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_width: int,
              hard_chunk: Optional[int] = None) -> List[str]:
    content = str(text or "")
    if not content:
        return [""]
    inner = max(1, max_width)
    use_hyphenation = bool(hard_chunk) and HYPHENATE_HEADERS and HYPHENATOR is not None

    lines: List[str] = []
    current = ""

    def text_width(fragment: str) -> int:
        return text_size(draw, fragment, font)[0]

    def hyphen_split_once(token: str) -> Tuple[Optional[str], str]:
        if not use_hyphenation or not token:
            return None, token
        try:
            syllables = (HYPHENATOR.inserted(token) or token).split("\u00AD")  # type: ignore[arg-type]
            if len(syllables) <= 1:
                return None, token
            prefix = ""
            consumed = 0
            for idx in range(len(syllables) - 1):
                candidate = prefix + syllables[idx]
                if text_width(candidate + "-") <= inner:
                    prefix = candidate
                    consumed = idx + 1
                else:
                    break
            if consumed > 0:
                return prefix + "\u00AD", "".join(syllables[consumed:])
            return None, token
        except Exception:
            return None, token

    def split_token(token: str) -> List[str]:
        parts: List[str] = []
        remaining = token
        while remaining:
            if text_width(remaining) <= inner:
                parts.append(remaining)
                break
            prefix, rest = hyphen_split_once(remaining)
            if prefix is not None:
                parts.append(prefix)
                remaining = rest
                continue
            limit = min(len(remaining), hard_chunk) if hard_chunk and hard_chunk > 0 else len(remaining)
            lo, hi, fit = 1, limit, 0
            while lo <= hi:
                mid = (lo + hi) // 2
                if text_width(remaining[:mid] + "-") <= inner:
                    fit = mid
                    lo = mid + 1
                else:
                    hi = mid - 1
            if fit == 0:
                fit = 1
            parts.append(remaining[:fit] + "\u00AD")
            remaining = remaining[fit:]
        return parts

    for token in content.split(" "):
        token = token or ""
        trial = (current + " " + token) if current else token
        if current and text_width(trial) <= inner:
            current = trial
            continue
        if current:
            lines.append(current)
            current = ""
        for idx, piece in enumerate(split_token(token)):
            separator = " " if (idx == 0 and current) else ""
            trial = (current + separator + piece) if current else piece
            if current and text_width(trial) > inner:
                lines.append(current)
                current = piece
            else:
                current = trial
            if piece.endswith("\u00AD"):
                lines.append(current)
                current = ""
    if current:
        lines.append(current)
    return lines


def ellipsize(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_width: int) -> str:
    value = str(text or "")
    if not value:
        return ""
    if text_size(draw, value, font)[0] <= max_width:
        return value
    ellipsis = "…"
    lo, hi = 0, len(value)
    while lo < hi:
        mid = (lo + hi) // 2
        candidate = value[:mid] + ellipsis
        if text_size(draw, candidate, font)[0] <= max_width:
            lo = mid + 1
        else:
            hi = mid
    return (value[:max(lo - 1, 0)] + ellipsis) if lo > 0 else ellipsis


def read_csv(path: str, delimiter: str, encoding_utf8: bool = True) -> List[List[str]]:
    encoding = "utf-8-sig" if encoding_utf8 else "cp1252"
    rows: List[List[str]] = []
    with open(path, "r", encoding=encoding, newline="") as handle:
        reader = csv.reader(handle, delimiter=delimiter)
        rows.extend(reader)
    return rows


def normalize_rows(rows: Sequence[Sequence[str]]) -> List[List[str]]:
    if not rows:
        return []
    cleaned: List[List[str]] = []
    for row in rows:
        if row is None:
            continue
        normalized = [("" if cell is None else str(cell)) for cell in row]
        if any(fragment.strip() for fragment in normalized):
            cleaned.append(normalized)
    if not cleaned:
        return []
    if cleaned[0] and isinstance(cleaned[0][0], str):
        cleaned[0][0] = cleaned[0][0].lstrip("\ufeff")
    max_cols = max(len(r) for r in cleaned)
    return [(r + [""] * (max_cols - len(r))) if len(r) < max_cols else list(r[:max_cols]) for r in cleaned]


def create_temp_csv_without_empty_columns(path: str, delimiter: str, encoding_utf8: bool = True) -> str:
    rows = normalize_rows(read_csv(path, delimiter, encoding_utf8=encoding_utf8))
    if not rows:
        raise ValueError("Die CSV-Datei enthält keine verwertbaren Daten.")
    max_cols = max(len(r) for r in rows)
    padded = [(r + [""] * (max_cols - len(r))) if len(r) < max_cols else r[:max_cols] for r in rows]
    header, body = padded[0], padded[1:]
    keep_indices = [idx for idx in range(max_cols) if any(str(row[idx]).strip() for row in body)]
    if not keep_indices:
        raise ValueError("Alle Spalten sind im Body leer – nichts zu behalten.")
    fd, temp_path = tempfile.mkstemp(prefix="csv_clean_", suffix=".csv")
    os.close(fd)
    encoding = "utf-8-sig" if encoding_utf8 else "cp1252"
    with open(temp_path, "w", encoding=encoding, newline="") as handle:
        writer = csv.writer(handle, delimiter=delimiter)
        for row in padded:
            writer.writerow([row[idx] for idx in keep_indices])
    return temp_path


def read_cochrane_txt(path: str) -> Tuple[List[str], List[List[str]], dict]:
    text = _read_text_with_fallback(path)
    return _parse_cochrane_search_manager_txt(text)


def _read_text_with_fallback(path: str) -> str:
    for encoding in ("utf-8", "utf-16", "utf-16-le", "utf-16-be", "latin-1", "cp1252"):
        try:
            with open(path, "r", encoding=encoding, newline="") as handle:
                return handle.read()
        except Exception:
            continue
    with open(path, "rb") as handle:
        return handle.read().decode("utf-8", errors="ignore")


def _parse_cochrane_search_manager_txt(text: str) -> Tuple[List[str], List[List[str]], dict]:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    meta: dict = {}
    for key in ("Search Name", "Date Run", "Comment"):
        match = re.search(rf"(?m)^{re.escape(key)}:\t(.*)$", text)
        if match:
            meta[key] = match.group(1).strip()
    header_match = re.search(r"(?m)^ID\s*\t\s*Search\s*\t\s*Hits\s*$", text)
    if not header_match:
        raise ValueError("Konnte den Tabellenkopf 'ID\\tSearch\\tHits' nicht finden.")
    lines = text[header_match.end():].strip("\n").split("\n")
    header = ["ID", "Search", "Hits"]
    rows: List[List[str]] = []
    current_id: Optional[str] = None
    buffer: List[str] = []

    def flush_row(final_chunk: Optional[str], hits_str: str) -> None:
        nonlocal rows, buffer, current_id
        parts = buffer + ([final_chunk] if final_chunk is not None else [])
        collapsed = re.sub(r"\s+", " ", " ".join(p.replace("\t", " ").strip() for p in parts if p)).strip()
        disp_id = f"#{current_id}" if current_id and not str(current_id).startswith("#") else (current_id or "")
        rows.append([disp_id, collapsed, hits_str.strip()])
        buffer.clear()
        current_id = None

    id_line_re = re.compile(r"^#?(\d+)\t(.*)$")
    for line in lines:
        if not line.strip():
            continue
        match = id_line_re.match(line)
        if match and current_id is None:
            current_id = match.group(1)
            rest = match.group(2)
            hits_match = re.search(r"\t(\d+)\s*$", rest)
            if hits_match:
                flush_row(rest[:hits_match.start()], hits_match.group(1))
            else:
                buffer = [rest]
            continue
        if current_id is not None:
            hits_match = re.search(r"\t(\d+)\s*$", line)
            if hits_match:
                flush_row(line[:hits_match.start()], hits_match.group(1))
            else:
                buffer.append(line)
            continue
        if match:
            current_id = match.group(1)
            rest = match.group(2)
            hits_match = re.search(r"\t(\d+)\s*$", rest)
            if hits_match:
                flush_row(rest[:hits_match.start()], hits_match.group(1))
            else:
                buffer = [rest]
    if current_id is not None and buffer:
        flush_row(" ", "")
    return header, rows, meta


def _guess_lang_from_header(header_list: Sequence[str]) -> str:
    text = " ".join(map(str, header_list or []))
    if any(ch in text for ch in "äöüÄÖÜß"):
        return "de_DE"
    return "en_US"


def make_hyphenator(choice: str, header_for_auto: Sequence[str]) -> Optional["pyphen.Pyphen"]:
    if pyphen is None:
        return None
    if choice.startswith("Aus"):
        return None
    if choice.startswith("Auto"):
        lang = _guess_lang_from_header(header_for_auto)
    elif "de_DE" in choice or "Deutsch" in choice:
        lang = "de_DE"
    else:
        lang = "en_US"
    try:
        return pyphen.Pyphen(lang=lang)  # type: ignore[call-arg]
    except Exception:
        return None


def measure_header_and_words(draw: ImageDraw.ImageDraw, header: Sequence[str], rows: Sequence[Sequence[str]],
                             font_header: ImageFont.ImageFont, font_body: ImageFont.ImageFont, pad_x: int
                             ) -> Tuple[List[int], List[int], List[int]]:
    n = len(header)
    word_re = re.compile(r"\S+")
    header_piece_longest = [0] * n
    body_longest_word = [0] * n
    natural = [0] * n

    for j in range(n):
        header_text = str(header[j])
        natural[j] = max(natural[j], text_size(draw, header_text, font_header)[0] + 2 * pad_x)
        for token in word_re.findall(header_text):
            if HEADER_HARD_WRAP_CHARS and len(token) > HEADER_HARD_WRAP_CHARS:
                for idx in range(0, len(token), HEADER_HARD_WRAP_CHARS):
                    piece = token[idx:idx + HEADER_HARD_WRAP_CHARS]
                    width, _ = text_size(draw, piece, font_header)
                    header_piece_longest[j] = max(header_piece_longest[j], width + 2 * pad_x)
            else:
                width, _ = text_size(draw, token, font_header)
                header_piece_longest[j] = max(header_piece_longest[j], width + 2 * pad_x)

    for row in rows:
        for j in range(min(len(row), n)):
            text = str(row[j])
            natural[j] = max(natural[j], text_size(draw, text, font_body)[0] + 2 * pad_x)
            for token in word_re.findall(text):
                width, _ = text_size(draw, token, font_body)
                body_longest_word[j] = max(body_longest_word[j], width + 2 * pad_x)
    return header_piece_longest, body_longest_word, natural


def min_floor_from_font(draw: ImageDraw.ImageDraw, font_body: ImageFont.ImageFont, pad_x: int,
                        min_chars: int = 3) -> int:
    width, _ = text_size(draw, "W" * min_chars, font_body)
    return width + 2 * pad_x


def compute_wrap_score(draw: ImageDraw.ImageDraw, col_texts: Sequence[str], font: ImageFont.ImageFont,
                       col_width: int, pad_x: int, line_gap: int, is_header: bool = False) -> int:
    inner = max(1, col_width - 2 * pad_x)
    score = 0
    _, one_h = text_size(draw, "Ag", font)
    for text in col_texts:
        lines = wrap_text(draw, str(text), font, inner, hard_chunk=HEADER_HARD_WRAP_CHARS if is_header else None)
        score += max(0, len(lines) - 1) * one_h
    return score


def flex_fit_widths(base: Sequence[int], minw: Sequence[int], target_width: int) -> List[int]:
    n = len(base)
    if n == 0:
        return []
    total = sum(base) or 1
    widths = [max(minw[i], int(round(base[i] * (target_width / total)))) for i in range(n)]
    diff = target_width - sum(widths)
    order = sorted(range(n), key=lambda idx: base[idx], reverse=True) or list(range(n))
    guard = 0
    while diff != 0 and guard < 10000:
        changed = False
        for idx in order:
            if diff > 0:
                widths[idx] += 1
                diff -= 1
                changed = True
                if diff == 0:
                    break
            else:
                if widths[idx] > minw[idx]:
                    widths[idx] -= 1
                    diff += 1
                    changed = True
                    if diff == 0:
                        break
        if not changed:
            break
        guard += 1
    leftover = target_width - sum(widths)
    if leftover != 0:
        widths[-1] = max(minw[-1], widths[-1] + leftover)
    return widths


def equal_width_with_buffer(place_w: int, header: Sequence[str], rows: Sequence[Sequence[str]],
                            fonts: Tuple[ImageFont.ImageFont, ImageFont.ImageFont],
                            pad_x: int, line_gap: int) -> List[int]:
    font_body, font_header = fonts
    probe = Image.new("RGB", (100, 100), "white")
    drawer = ImageDraw.Draw(probe)
    header_piece_longest, _, natural = measure_header_and_words(drawer, header, rows, font_header, font_body, pad_x)
    min_floor = min_floor_from_font(drawer, font_body, pad_x, min_chars=3)

    n = len(header)
    minw = [max(min_floor, header_piece_longest[j]) for j in range(n)]
    equal = max(1, place_w // n)
    widths = [max(minw[j], min(natural[j], equal)) for j in range(n)]
    total = sum(widths)

    if total > place_w or sum(minw) > place_w:
        widths = flex_fit_widths(natural, minw, place_w)
    else:
        buffer_width = place_w - total
        if buffer_width > 0:
            columns = list(zip(*rows)) if rows else [[] for _ in range(n)]
            scores = []
            for j in range(n):
                header_only = [header[j]] if j < len(header) else []
                column_texts = columns[j] if j < len(columns) else []
                score = compute_wrap_score(drawer, header_only, font_header, widths[j], pad_x, line_gap, is_header=True)
                score += compute_wrap_score(drawer, column_texts, font_body, widths[j], pad_x, line_gap, is_header=False)
                scores.append(score)
            order = sorted(range(n), key=lambda j: scores[j], reverse=True)
            remaining = buffer_width
            for j in order:
                can_add = max(0, natural[j] - widths[j])
                add = min(remaining, can_add)
                widths[j] += add
                remaining -= add
                if remaining <= 0:
                    break
        diff = place_w - sum(widths)
        if diff != 0:
            columns = list(zip(*rows)) if rows else [[] for _ in range(n)]
            scores = []
            for j in range(n):
                header_only = [header[j]] if j < len(header) else []
                column_texts = columns[j] if j < len(columns) else []
                score = compute_wrap_score(drawer, header_only, font_header, max(widths[j], 1), pad_x, line_gap, is_header=True)
                score += compute_wrap_score(drawer, column_texts, font_body, max(widths[j], 1), pad_x, line_gap, is_header=False)
                scores.append(score)
            order = sorted(range(n), key=lambda j: scores[j], reverse=True) or list(range(n))
            guard = 0
            while diff != 0 and guard < 10000:
                changed = False
                for j in order:
                    if diff > 0:
                        widths[j] += 1
                        diff -= 1
                        changed = True
                        if diff == 0:
                            break
                    else:
                        if widths[j] > minw[j]:
                            widths[j] -= 1
                            diff += 1
                            changed = True
                            if diff == 0:
                                break
                if not changed:
                    break
                guard += 1
        leftover = place_w - sum(widths)
        if leftover != 0:
            widths[-1] = max(minw[-1], widths[-1] + leftover)
    return widths


def row_height_for_widths(draw: ImageDraw.ImageDraw, cells: Sequence[str], col_widths: Sequence[int],
                          font: ImageFont.ImageFont, pad_x: int, pad_y: int, line_gap: int,
                          is_header: bool = False) -> int:
    max_lines = 1
    for idx, text in enumerate(cells):
        col_width = max(1, col_widths[idx] - 2 * pad_x)
        lines = wrap_text(draw, str(text), font, col_width,
                          hard_chunk=HEADER_HARD_WRAP_CHARS if is_header else None)
        max_lines = max(max_lines, len(lines))
    _, one_h = text_size(draw, "Ag", font)
    return one_h * max_lines + 2 * pad_y + (max_lines - 1) * line_gap


def paginate_rows_by_height(draw: ImageDraw.ImageDraw, header: Sequence[str], rows: Sequence[Sequence[str]],
                            col_widths: Sequence[int], fonts: Tuple[ImageFont.ImageFont, ImageFont.ImageFont],
                            page_body_height: int, pad_x: int, pad_y: int, line_gap: int
                            ) -> List[Tuple[int, int, int]]:
    font_body, font_header = fonts
    header_height = row_height_for_widths(draw, header, col_widths, font_header, pad_x, pad_y, line_gap, is_header=True)
    heights = [
        row_height_for_widths(
            draw,
            [row[j] if j < len(row) else "" for j in range(len(header))],
            col_widths,
            font_body,
            pad_x,
            pad_y,
            line_gap,
            is_header=False,
        )
        for row in rows
    ]
    pages: List[Tuple[int, int, int]] = []
    i = 0
    while i < len(rows):
        y = 0
        start = i
        while i < len(rows):
            row_height = heights[i]
            if y + row_height <= page_body_height:
                y += row_height
                i += 1
            else:
                break
        if start == i:
            i += 1
        pages.append((start, i, header_height))
    return pages


def draw_page(canvas: Image.Image, margin: int, place_w: int, place_h: int,
              header: Sequence[str], rows: Sequence[Sequence[str]], col_widths: Sequence[int],
              fonts: Tuple[ImageFont.ImageFont, ImageFont.ImageFont], pad_x: int, pad_y: int, line_gap: int,
              start: int, end: int, page_idx: int, page_count: int,
              header_rgb: Tuple[int, int, int], zebra_rgb: Tuple[int, int, int],
              top_note_text: Optional[str] = None, note_h: int = 0) -> None:
    font_body, font_header = fonts
    drawer = ImageDraw.Draw(canvas)
    grid = (200, 200, 200)
    text_color = (0, 0, 0)
    y_cursor = margin
    if top_note_text:
        inner_w = max(1, place_w - 2 * pad_x)
        note_text = ellipsize(drawer, top_note_text, font_body, inner_w)
        drawer.text((margin + pad_x, y_cursor + pad_y), note_text, font=font_body, fill=text_color)
        drawer.line([(margin, y_cursor + note_h - 1), (margin + place_w - 1, y_cursor + note_h - 1)], fill=grid)
        y_cursor += note_h
    header_h = row_height_for_widths(drawer, header, col_widths, font_header, pad_x, pad_y, line_gap, is_header=True)
    drawer.rectangle([margin, y_cursor, margin + place_w - 1, y_cursor + header_h - 1], fill=header_rgb)
    x_cursor = margin
    for col_idx in range(len(header)):
        col_width = col_widths[col_idx]
        drawer.line([(x_cursor, y_cursor), (x_cursor, y_cursor + place_h)], fill=grid)
        inner_w = max(1, col_width - 2 * pad_x)
        lines = wrap_text(drawer, str(header[col_idx]), font_header, inner_w, hard_chunk=HEADER_HARD_WRAP_CHARS)
        _, one_h = text_size(drawer, "Ag", font_header)
        total_h = len(lines) * one_h + (len(lines) - 1) * line_gap
        y_text = y_cursor + (header_h - total_h) // 2
        for line in lines:
            drawer.text((x_cursor + pad_x, y_text), line, font=font_header, fill=text_color)
            y_text += one_h + line_gap
        x_cursor += col_width
    drawer.line([(margin + place_w - 1, y_cursor), (margin + place_w - 1, y_cursor + place_h)], fill=grid)
    drawer.line([(margin, y_cursor + header_h - 1), (margin + place_w - 1, y_cursor + header_h - 1)], fill=grid)
    y = y_cursor + header_h
    for idx in range(start, end):
        row = rows[idx]
        row_safe = [row[j] if j < len(row) else "" for j in range(len(header))]
        row_height = row_height_for_widths(drawer, row_safe, col_widths, font_body, pad_x, pad_y, line_gap, is_header=False)
        if (idx - start) % 2 == 0:
            drawer.rectangle([margin, y, margin + place_w - 1, y + row_height - 1], fill=zebra_rgb)
        x_cursor = margin
        for col_idx in range(len(header)):
            col_width = col_widths[col_idx]
            inner_w = max(1, col_width - 2 * pad_x)
            lines = wrap_text(drawer, str(row_safe[col_idx]), font_body, inner_w, hard_chunk=None)
            _, one_h = text_size(drawer, "Ag", font_body)
            y_text = y + pad_y
            for line in lines:
                drawer.text((x_cursor + pad_x, y_text), line, font=font_body, fill=text_color)
                y_text += one_h + line_gap
            drawer.rectangle([x_cursor, y, x_cursor + col_width - 1, y + row_height - 1], outline=grid)
            x_cursor += col_width
        y += row_height
    page_text = f"Seite {page_idx}/{page_count}"
    tw, th = text_size(drawer, page_text, font_body)
    drawer.text((canvas.width // 2 - tw // 2, canvas.height - mm_to_px(12) - th),
                page_text, font=font_body, fill=(0, 0, 0))


def render_pages_dynamic(header: Sequence[str], rows: Sequence[Sequence[str]], orientation: str,
                         header_hex: str, zebra_hex: str, top_note: Optional[str] = None) -> List[Image.Image]:
    page_w, page_h = a4_pixels(orientation)
    margin = mm_to_px(12)
    place_w, place_h = page_w - 2 * margin, page_h - 2 * margin
    body_size, header_size, min_header = 24, 28, 16
    pad_x, pad_y, line_gap = 16, 12, 6

    probe = Image.new("RGB", (100, 100), "white")
    drawer = ImageDraw.Draw(probe)

    while True:
        fonts = load_fonts(size_body=body_size, size_header=header_size)
        font_body, font_header = fonts
        n = len(header)
        header_piece_longest = [0] * n
        word_re = re.compile(r"\S+")
        for j in range(n):
            for token in word_re.findall(str(header[j])):
                if HEADER_HARD_WRAP_CHARS and len(token) > HEADER_HARD_WRAP_CHARS:
                    for idx in range(0, len(token), HEADER_HARD_WRAP_CHARS):
                        piece = token[idx:idx + HEADER_HARD_WRAP_CHARS]
                        width, _ = text_size(drawer, piece, font_header)
                        header_piece_longest[j] = max(header_piece_longest[j], width + 2 * pad_x)
                else:
                    width, _ = text_size(drawer, token, font_header)
                    header_piece_longest[j] = max(header_piece_longest[j], width + 2 * pad_x)
        minw_try = [max(min_floor_from_font(drawer, font_body, pad_x, min_chars=3), header_piece_longest[j])
                    for j in range(len(header))]
        if sum(minw_try) <= place_w or header_size <= min_header:
            break
        header_size -= 2

    fonts = load_fonts(size_body=body_size, size_header=header_size)
    col_widths = equal_width_with_buffer(place_w, header, rows, fonts, pad_x, line_gap)
    font_body, font_header = fonts
    header_h_tmp = row_height_for_widths(drawer, header, col_widths, font_header, pad_x, pad_y, line_gap, is_header=True)
    _, one_h = text_size(drawer, "Ag", font_body)
    note_h = (2 * pad_y + one_h) if (top_note and str(top_note).strip()) else 0
    page_body_height = place_h - header_h_tmp - note_h
    if page_body_height < 80:
        pad_y = max(8, pad_y - 4)
        header_h_tmp = row_height_for_widths(drawer, header, col_widths, font_header, pad_x, pad_y, line_gap, is_header=True)
        page_body_height = place_h - header_h_tmp - note_h
    pages_layout = paginate_rows_by_height(drawer, header, rows, col_widths, fonts, page_body_height, pad_x, pad_y, line_gap)
    header_rgb = hex_to_rgb(header_hex)
    zebra_rgb = hex_to_rgb(zebra_hex)
    pages: List[Image.Image] = []
    page_count = len(pages_layout) if pages_layout else 1
    for idx, (start, end, _) in enumerate(pages_layout, start=1):
        canvas = Image.new("RGB", (page_w, page_h), "white")
        draw_page(
            canvas, margin, place_w, place_h, header, rows, col_widths,
            fonts, pad_x, pad_y, line_gap, start, end, idx, page_count,
            header_rgb, zebra_rgb, top_note_text=(top_note or ""), note_h=note_h,
        )
        pages.append(canvas)
    if not rows:
        canvas = Image.new("RGB", (page_w, page_h), "white")
        drawer = ImageDraw.Draw(canvas)
        drawer.text((margin, margin), "Keine Daten", font=font_header, fill=(0, 0, 0))
        pages = [canvas]
    return pages


def save_as_excel(header: Sequence[str], rows: Sequence[Sequence[str]], path: str,
                  orientation: str, fit_width: bool = True, meta_note: Optional[str] = None) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Tabelle"
    first_data_row = 1
    if meta_note:
        ws.append([meta_note])
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(header))
        ws["A1"].alignment = Alignment(vertical="center", wrap_text=True)
        ws["A1"].font = Font(bold=False, italic=True)
        first_data_row = 2
    ws.append(list(map(str, header)))
    for row in rows:
        ws.append([("" if idx >= len(row) else str(row[idx])) for idx in range(len(header))])

    header_font = Font(bold=True)
    align = Alignment(vertical="top", wrap_text=True)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    header_row_idx = first_data_row
    for col_idx in range(1, len(header) + 1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(ws[f"{col_letter}{header_row_idx}"].value or ""))
        for row_idx in range(header_row_idx + 1, ws.max_row + 1):
            val = ws[f"{col_letter}{row_idx}"].value
            if val is None:
                continue
            max_len = max(max_len, len(str(val)))
        ws.column_dimensions[col_letter].width = min(80, max(10, int(max_len * 1.2)))

    for row in ws.iter_rows(min_row=header_row_idx, max_row=ws.max_row, min_col=1, max_col=len(header)):
        for cell in row:
            cell.alignment = align
            cell.border = border

    fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
    for cell in ws[header_row_idx]:
        cell.font = header_font
        cell.fill = fill

    try:
        ws.page_setup.paperSize = 9  # A4
        ws.page_setup.orientation = "landscape" if orientation == "landscape" else "portrait"
        ws.page_setup.fitToWidth = 1 if fit_width else 0
        ws.page_setup.fitToHeight = 0
        ws.print_title_rows = f"{header_row_idx}:{header_row_idx}"
        ws.page_margins.left = 0.39
        ws.page_margins.right = 0.39
        ws.page_margins.top = 0.59
        ws.page_margins.bottom = 0.59
    except Exception:
        pass

    ws.freeze_panes = f"A{header_row_idx + 1}"
    wb.save(path)