import os
import re
import io
import uuid
import shutil
import argparse
import textwrap
from pathlib import Path

from PIL import Image
from pypdf import PdfReader, PdfWriter, Transformation

try:
    from docx2pdf import convert as docx2pdf_convert
except ImportError:
    raise SystemExit(
        "docx2pdf is required for DOCX → PDF conversion.\n"
        "Install it with:\n    pip install docx2pdf"
    )

try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import inch
    from reportlab.pdfbase.pdfmetrics import stringWidth
except ImportError:
    raise SystemExit(
        "reportlab is required for Bates stamping and text-based PDFs.\n"
        "Install it with:\n    pip install reportlab"
    )

try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None

# ======================================================
# Application Version (used by build.sh and GUI updater)
# ======================================================
APP_VERSION = "1.0.1"

# ========== CONFIG (overridden by run_pipeline) ==========

ROOT_FOLDER = r"/path/to/root/folder"

PREFIX = "CF"       # Prefix for filenames and Bates labels
DIGITS = 4          # Zero padding: 4 -> 0001
START_COUNTER = 1   # Starting sequence number
DRY_RUN = True      # True = preview only; False = actually modify files

BACKUP_BEFORE_BATES = True
BACKUP_FOLDER_NAME = "_bates_backups"   # created inside ROOT_FOLDER

# Toggle 1: include original filename after Bates label for files
# True  -> "CF 0001-0008 - Original Name.ext"
# False -> "CF 0001-0008.ext"
KEEP_ORIGINAL_NAME = True

# Toggle 2: rename folders based on Bates range of their contents
RENAME_FOLDERS = False

# Toggle 3: when renaming folders, append original folder name?
# True  -> "CF 0001-0244 - Folder Name"
# False -> "CF 0001-0244"
KEEP_FOLDER_NAME = True

# Toggle 4: number videos at end (after all other items) vs inline
# True  -> videos at end
# False -> videos inline in Finder order
NUMBER_VIDEOS_AT_END = True

# Toggle 5: create final combined PDF for full CF range
COMBINE_FINAL = False

# Toggle 6: conversion-only mode (no renaming, no Bates, just convert + letter-format)
CONVERSION_ONLY = False

# File type groups
PDF_EXT = ".pdf"
WORD_EXTS = {".docx"}  # .doc is blocked
EXCEL_EXTS = {".xls", ".xlsx", ".xlsm", ".xlsb"}
VIDEO_EXTS = {".mp4", ".mov", ".m4v", ".avi", ".mkv", ".wmv", ".flv"}
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".gif"}
HTML_EXTS = {".html", ".htm"}
TEXT_EXTS = {".txt"}

# Blocked types (not auto-handled yet)
BLOCKED_OTHER_EXTS = {".doc", ".eml", ".msg"}

# US Letter (points)
LETTER_PORTRAIT = (612, 792)      # 8.5 x 11
LETTER_LANDSCAPE = (792, 612)     # 11 x 8.5

# Bates appearance
BATES_FONT = "Times-Bold"
BATES_FONT_SIZE = 12
BATES_MARGIN_BOTTOM = 0.5 * inch
BATES_MARGIN_RIGHT = 1.0 * inch
BATES_FOOTER_BAND = 0.75 * inch

# Parse names like:
#   "CF 0001"
#   "CF 0001-0008"
#   "CF 0001-0008 - Original Name"
BATES_NAME_PATTERN = re.compile(
    r"^(?P<prefix>[A-Za-z0-9]+)\s+"
    r"(?P<start>\d+)"
    r"(?:-(?P<end>\d+))?"
    r"(?:\s*-\s*.+)?$"
)

# =======================================================


def natural_key(path: Path):
    """Finder-like natural sort with numeric awareness."""
    parts = re.split(r"(\d+)", path.name)
    return [int(p) if p.isdigit() else p.lower() for p in parts]


def iter_finder_order_files(root: Path):
    """
    Depth-first traversal in natural order.

    Skips:
      - Hidden files/folders starting with '.' or '~'
      - Common system junk (Thumbs.db, desktop.ini)
      - Backup folder tree
    """
    entries = sorted(root.iterdir(), key=natural_key)
    for entry in entries:
        name = entry.name

        if BACKUP_FOLDER_NAME in entry.parts:
            continue

        # Skip hidden / temp / junk
        if name.startswith(".") or name.startswith("~"):
            continue
        if name in {"Thumbs.db", "desktop.ini"}:
            continue

        if entry.is_dir():
            yield from iter_finder_order_files(entry)
        else:
            yield entry


def find_blocking_files(root: Path):
    """
    Return list of disallowed files:
      - .doc
      - .eml
      - .msg
    """
    blocking = []
    for path in iter_finder_order_files(root):
        if not path.is_file():
            continue
        if path.suffix.lower() in BLOCKED_OTHER_EXTS:
            blocking.append(path)
    return blocking


def get_pdf_page_count(path: Path) -> int:
    try:
        reader = PdfReader(str(path))
        return len(reader.pages)
    except Exception as e:
        print(f"⚠️  Skipping unreadable PDF {path}: {e}")
        return 0


# ---------- Image → PDF ----------

def convert_image_to_pdf(image_path: Path, pdf_path: Path):
    """Convert a single image to a single-page PDF."""
    with Image.open(image_path) as img:
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(pdf_path, "PDF")


def convert_images_in_tree(root: Path, delete_original: bool):
    """
    Recursively convert images under `root` to PDFs.

    - Honors DRY_RUN via delete_original flag and convert logic.
    - Avoids overwriting existing PDFs.
    """
    conversions = []
    errors = []

    for path in iter_finder_order_files(root):
        if not path.is_file():
            continue

        ext = path.suffix.lower()
        if ext not in IMAGE_EXTS:
            continue

        pdf_path = path.with_suffix(".pdf")
        counter = 1
        while pdf_path.exists():
            pdf_path = path.with_name(f"{path.stem}_{counter}.pdf")
            counter += 1

        if DRY_RUN:
            print(f"(DRY RUN) Would convert image to PDF: {path} -> {pdf_path}")
            conversions.append((str(path), str(pdf_path)))
            continue

        try:
            convert_image_to_pdf(path, pdf_path)
            conversions.append((str(path), str(pdf_path)))
            if delete_original:
                path.unlink(missing_ok=True)
        except Exception as e:
            msg = f"{path}: {e}"
            print(f"⚠️  Image→PDF conversion failed: {msg}")
            errors.append(msg)

    return conversions, errors


# ---------- DOCX → PDF (delete original) ----------

def convert_word_to_pdf(word_path: Path):
    """
    Convert .docx to .pdf via docx2pdf.
    - In DRY_RUN: log only, return None.
    - On success: delete original .docx, return pdf_path.
    """
    pdf_path = word_path.with_suffix(".pdf")

    if DRY_RUN:
        print(f"(DRY RUN) Would convert DOCX to PDF (and delete DOCX): {word_path} -> {pdf_path}")
        return None

    try:
        docx2pdf_convert(str(word_path), str(pdf_path))
        if pdf_path.exists():
            try:
                word_path.unlink()
            except FileNotFoundError:
                pass
            return pdf_path
        print(f"⚠️  docx2pdf did not create expected file: {pdf_path}")
    except Exception as e:
        print(f"⚠️  Failed DOCX→PDF conversion for {word_path}: {e}")
    return None


def convert_docx_in_tree(root: Path):
    """Convert all .docx in tree to PDFs, deleting originals on real run."""
    conversions = []
    errors = []

    for path in iter_finder_order_files(root):
        if not path.is_file():
            continue
        if path.suffix.lower() not in WORD_EXTS:
            continue

        pdf_path = path.with_suffix(".pdf")

        if DRY_RUN:
            print(f"(DRY RUN) Would convert DOCX to PDF (and delete DOCX): {path} -> {pdf_path}")
            conversions.append((str(path), str(pdf_path)))
            continue

        result = convert_word_to_pdf(path)
        if result:
            conversions.append((str(path), str(result)))
        else:
            msg = f"{path}: failed to convert DOCX to PDF"
            errors.append(msg)

    return conversions, errors


# ---------- HTML → PDF ----------

def html_to_text(html: str) -> str:
    if BeautifulSoup is not None:
        return BeautifulSoup(html, "html.parser").get_text(separator="\n")
    return re.sub(r"<[^>]+>", "", html)


def write_text_pdf(pdf_path: Path, title: str, body: str):
    """Write a simple text PDF with optional title and body."""
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(pdf_path), pagesize=LETTER_PORTRAIT)

    x_margin = 0.75 * inch
    y = LETTER_PORTRAIT[1] - 0.75 * inch
    line_height = 12

    def draw_line(text: str):
        nonlocal y
        if y < 1 * inch:
            c.showPage()
            y = LETTER_PORTRAIT[1] - 0.75 * inch
        c.drawString(x_margin, y, text)
        y -= line_height

    if title:
        c.setFont("Times-Bold", 14)
        draw_line(title)
        c.setFont("Times-Roman", 11)
        draw_line("")

    for line in (body or "").splitlines():
        max_width = LETTER_PORTRAIT[0] - 2 * x_margin
        words = line.split()
        current = ""
        for word in words:
            test = (current + " " + word).strip()
            if stringWidth(test, "Times-Roman", 11) <= max_width:
                current = test
            else:
                draw_line(current)
                current = word
        if current:
            draw_line(current)

    c.save()


def convert_html_to_pdf(html_path: Path, pdf_path: Path):
    """Convert .html/.htm to a text-based PDF snapshot."""
    if DRY_RUN:
        print(f"(DRY RUN) Would convert HTML to PDF: {html_path} -> {pdf_path}")
        return None

    try:
        text = html_path.read_text(encoding="utf-8", errors="ignore")
        body = html_to_text(text)
        title = f"HTML: {html_path.name}"
        write_text_pdf(pdf_path, title, body)
        return pdf_path
    except Exception as e:
        print(f"⚠️  Failed HTML→PDF conversion for {html_path}: {e}")
        return None


def convert_htmls_in_tree(root: Path, delete_original: bool):
    conversions = []
    errors = []

    for path in iter_finder_order_files(root):
        if not path.is_file():
            continue

        if path.suffix.lower() not in HTML_EXTS:
            continue

        pdf_path = path.with_suffix(".pdf")
        counter = 1
        while pdf_path.exists():
            pdf_path = path.with_name(f"{path.stem}_html_{counter}.pdf")
            counter += 1

        try:
            result = convert_html_to_pdf(path, pdf_path)
            if result:
                conversions.append((str(path), str(pdf_path)))
                if not DRY_RUN and delete_original:
                    path.unlink(missing_ok=True)
        except Exception as e:
            msg = f"{path}: {e}"
            print(f"⚠️  HTML→PDF conversion failed: {msg}")
            errors.append(msg)

    return conversions, errors


# ---------- TXT → PDF ----------

def convert_txt_to_pdf(txt_path: Path, pdf_path: Path):
    """Convert plain text file to a simple text PDF."""
    if DRY_RUN:
        print(f"(DRY RUN) Would convert TXT to PDF: {txt_path} -> {pdf_path}")
        return None

    try:
        body = txt_path.read_text(encoding="utf-8", errors="ignore")
    except UnicodeDecodeError:
        body = txt_path.read_text(errors="ignore")

    title = f"TXT: {txt_path.name}"
    write_text_pdf(pdf_path, title, body)
    return pdf_path


def convert_txts_in_tree(root: Path, delete_original: bool):
    conversions = []
    errors = []

    for path in iter_finder_order_files(root):
        if not path.is_file():
            continue

        if path.suffix.lower() not in TEXT_EXTS:
            continue

        pdf_path = path.with_suffix(".pdf")
        counter = 1
        while pdf_path.exists():
            pdf_path = path.with_name(f"{path.stem}_txt_{counter}.pdf")
            counter += 1

        try:
            result = convert_txt_to_pdf(path, pdf_path)
            if result:
                conversions.append((str(path), str(pdf_path)))
                if not DRY_RUN and delete_original:
                    path.unlink(missing_ok=True)
        except Exception as e:
            msg = f"{path}: {e}"
            print(f"⚠️  TXT→PDF conversion failed: {msg}")
            errors.append(msg)

    return conversions, errors


# ---------- Planning ----------

def plan_items(root: Path):
    """
    Build logical items in final processing order.

    Each item:
      - kind: 'pdf', 'word_no_pdf', 'excel', 'video'
      - pages: int (# Bates slots)
      - paths: dict of paths
    """
    items = []

    for path in iter_finder_order_files(root):
        if not path.is_file():
            continue
        if BACKUP_FOLDER_NAME in path.parts:
            continue

        suffix = path.suffix.lower()

        if suffix == PDF_EXT:
            pages = get_pdf_page_count(path)
            if pages > 0:
                items.append({"kind": "pdf", "pages": pages, "paths": {"pdf": path}})

        elif suffix in WORD_EXTS:
            # For normal pipeline, we still convert here
            pdf_path = convert_word_to_pdf(path)
            if pdf_path:
                pages = get_pdf_page_count(pdf_path) or 1
                items.append({"kind": "pdf", "pages": pages, "paths": {"pdf": pdf_path}})
            else:
                items.append({
                    "kind": "word_no_pdf",
                    "pages": 1,
                    "paths": {"word": path},
                })

        elif suffix in EXCEL_EXTS:
            items.append({
                "kind": "excel",
                "pages": 1,
                "paths": {"excel": path},
            })

        elif suffix in VIDEO_EXTS:
            items.append({
                "kind": "video",
                "pages": 1,
                "paths": {"video": path},
            })

    return items


def reorder_items_for_videos(items):
    """Optionally move videos to the end for numbering."""
    if not NUMBER_VIDEOS_AT_END:
        return items
    non_video = [it for it in items if it["kind"] != "video"]
    videos = [it for it in items if it["kind"] == "video"]
    return non_video + videos


def make_bates_filename(base: str, path: Path) -> str:
    """
    Build output filename according to KEEP_ORIGINAL_NAME:
      True:  '<base> - <original_stem><ext>'
      False: '<base><ext>'
    """
    if KEEP_ORIGINAL_NAME:
        return f"{base} - {path.stem}{path.suffix}"
    else:
        return f"{base}{path.suffix}"


def build_renames(items):
    """
    Assign Bates ranges and produce rename operations.

    Returns:
      operations: list[(src_path, dst_path)]
      excel_placeholders: [] (unused, kept for compatibility)
    """
    operations = []
    excel_placeholders = []
    counter = START_COUNTER

    for item in items:
        pages = item["pages"]
        start = counter
        end = counter + pages - 1

        if pages == 1:
            base = f"{PREFIX} {start:0{DIGITS}d}"
        else:
            base = f"{PREFIX} {start:0{DIGITS}d}-{end:0{DIGITS}d}"

        counter = end + 1
        kind = item["kind"]
        paths = item["paths"]

        if kind == "pdf":
            p = paths["pdf"]
            new_name = make_bates_filename(base, p)
            operations.append((p, p.with_name(new_name)))

        elif kind == "word_no_pdf":
            w = paths["word"]
            new_name = make_bates_filename(base, w)
            operations.append((w, w.with_name(new_name)))

        elif kind == "excel":
            e = paths["excel"]
            new_name = make_bates_filename(base, e)
            operations.append((e, e.with_name(new_name)))

        elif kind == "video":
            v = paths["video"]
            new_name = make_bates_filename(base, v)
            operations.append((v, v.with_name(new_name)))

    dests = [dst for _, dst in operations]
    if len(dests) != len(set(dests)):
        raise SystemExit("❌ Conflict: multiple files planned for same destination. Aborting.")

    return operations, excel_placeholders


# ---------- Renames ----------

def apply_renames(operations):
    """Safely apply renames using temp names. Honors DRY_RUN."""
    print("\n--- RENAME PLAN ---")
    for src, dst in operations:
        if src != dst:
            print(f"{src}  ->  {dst}")

    if DRY_RUN:
        print("\n(DRY RUN) No files were renamed.")
        return

    temp_map = {}

    # Step 1: move to unique temp names
    for src, dst in operations:
        if src == dst or not src.exists():
            continue
        tmp = src.with_name(f"__tmp__{uuid.uuid4().hex}__{src.name}")
        os.rename(src, tmp)
        temp_map[(src, dst)] = tmp

    # Step 2: move temps to final names
    for (src, dst), tmp in temp_map.items():
        dst.parent.mkdir(parents=True, exist_ok=True)
        os.rename(tmp, dst)

    print("✅ Renaming complete.")


# ---------- Letter Reformat ----------

def choose_letter_size(orig_width: float, orig_height: float):
    """Choose Letter portrait or landscape based on orientation."""
    return LETTER_LANDSCAPE if orig_width >= orig_height else LETTER_PORTRAIT


def reformat_pdf_to_letter_in_place(pdf_path: Path):
    """Reformat one PDF to Letter, preserving orientation, overwriting original."""
    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()

    for page in reader.pages:
        orig_w = float(page.mediabox.width)
        orig_h = float(page.mediabox.height)

        target_w, target_h = choose_letter_size(orig_w, orig_h)
        scale = min(target_w / orig_w, target_h / orig_h)
        new_w, new_h = orig_w * scale, orig_h * scale

        tx = (target_w - new_w) / 2.0
        ty = (target_h - new_h) / 2.0

        new_page = writer.add_blank_page(width=target_w, height=target_h)
        transform = Transformation().scale(scale).translate(tx, ty)
        new_page.merge_transformed_page(page, transform)

    temp_path = pdf_path.with_suffix(".tmp.pdf")
    with open(temp_path, "wb") as f_out:
        writer.write(f_out)

    os.replace(temp_path, pdf_path)
    print(f"✅ Reformatted to Letter: {pdf_path}")


# ---------- Backup originals ----------

def backup_originals(root: Path):
    """
    Backup ALL original files (any type) to ROOT/_bates_backups/,
    preserving relative paths and original names.

    Runs ONCE at the very start, before any conversion, renaming, or Bates.
    """
    backup_root = root / BACKUP_FOLDER_NAME
    print("\n--- BACKUP ORIGINAL TREE ---")

    if DRY_RUN:
        for path in iter_finder_order_files(root):
            if not path.is_file():
                continue
            rel = path.relative_to(root)
            dest = backup_root / rel
            print(f"(DRY RUN) Would backup: {path} -> {dest}")
        return

    for path in iter_finder_order_files(root):
        if not path.is_file():
            continue
        rel = path.relative_to(root)
        dest = backup_root / rel
        if dest.exists():
            continue
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, dest)

    print("✅ Original tree backup complete.")


# ---------- Bates stamping ----------

def create_bates_overlay(label: str, page_width: float, page_height: float):
    """Create a one-page overlay with Bates label in footer band."""
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))

    c.setFont(BATES_FONT, BATES_FONT_SIZE)
    text_width = stringWidth(label, BATES_FONT, BATES_FONT_SIZE)

    x_right = page_width - BATES_MARGIN_RIGHT
    x = x_right - text_width
    y = BATES_MARGIN_BOTTOM

    c.drawString(x, y, label)
    c.save()

    packet.seek(0)
    overlay_reader = PdfReader(packet)
    return overlay_reader.pages[0]


def apply_bates_to_pdf(pdf_path: Path):
    """
    Bates-stamp a single PDF based on filename:
      - 'CF 0001.pdf'
      - 'CF 0001-0008.pdf'
      - 'CF 0001-0008 - Original Name.pdf'
    """
    m = BATES_NAME_PATTERN.match(pdf_path.stem)
    if not m:
        print(f"ℹ️  Skipping Bates (name pattern mismatch): {pdf_path.name}")
        return

    prefix = m.group("prefix")
    start = int(m.group("start"))
    end_str = m.group("end")

    reader = PdfReader(str(pdf_path))
    num_pages = len(reader.pages)

    expected = (int(end_str) - start + 1) if end_str else 1
    if expected != num_pages:
        raise SystemExit(
            f"❌ Bates mismatch for {pdf_path.name}: "
            f"filename implies {expected} page(s), PDF has {num_pages}."
        )

    if DRY_RUN:
        last_num = start + num_pages - 1
        print(
            f"(DRY RUN) Would Bates-stamp {pdf_path.name} "
            f"from {prefix} {start:0{DIGITS}d} to {prefix} {last_num:0{DIGITS}d}"
        )
        return

    writer = PdfWriter()

    for i, original_page in enumerate(reader.pages):
        current_num = start + i
        label = f"{prefix} {current_num:0{DIGITS}d}"

        pw = float(original_page.mediabox.width)
        ph = float(original_page.mediabox.height)

        reserved = min(BATES_FOOTER_BAND, ph / 3)
        scale = min((ph - reserved) / ph, 1.0)

        scaled_w = pw * scale
        scaled_h = ph * scale

        tx = (pw - scaled_w) / 2.0
        ty = reserved + (ph - reserved - scaled_h) / 2.0

        new_page = writer.add_blank_page(width=pw, height=ph)

        transform = Transformation().scale(scale).translate(tx, ty)
        new_page.merge_transformed_page(original_page, transform)

        overlay = create_bates_overlay(label, pw, ph)
        new_page.merge_page(overlay)

    tmp = pdf_path.with_name(f"__bates__{uuid.uuid4().hex}__{pdf_path.name}")
    with open(tmp, "wb") as f:
        writer.write(f)
    os.replace(tmp, pdf_path)

    print(f"✅ Bates-stamped: {pdf_path.name}")


def apply_bates_to_all_pdfs(root: Path):
    """
    Reformat to Letter, then Bates-stamp all eligible PDFs.

    Returns:
        { "total_pages": int, "errors": [str, ...] }
    """
    print("\n--- BATES STAMP PLAN ---")

    pdfs = [
        p for p in iter_finder_order_files(root)
        if p.is_file()
        and p.suffix.lower() == PDF_EXT
        and BACKUP_FOLDER_NAME not in p.parts
    ]

    if not pdfs:
        print("No PDFs found for Bates stamping.")
        return {"total_pages": 0, "errors": []}

    errors = []

    # 1. Reformat all PDFs to Letter (unless DRY_RUN)
    if DRY_RUN:
        print("\n(DRY RUN) Would reformat all PDFs to US Letter before Bates stamping.")
    else:
        print("\n--- REFORMAT ALL PDFs TO US LETTER ---")
        for pdf in pdfs:
            try:
                reformat_pdf_to_letter_in_place(pdf)
            except Exception as e:
                msg = f"{pdf}: {e}"
                print(f"⚠️  Error reformatting {msg}")
                errors.append(msg)

    # 2. Bates stamp
    total_pages = 0

    for pdf in pdfs:
        try:
            reader = PdfReader(str(pdf))
            total_pages += len(reader.pages)
        except Exception:
            pass

        try:
            apply_bates_to_pdf(pdf)
        except Exception as e:
            msg = f"{pdf}: {e}"
            print(f"⚠️  Failed to Bates-stamp {msg}")
            errors.append(msg)

    if DRY_RUN:
        print("\n(DRY RUN) No Bates labels were actually written.")
    else:
        print("\n✅ All eligible PDFs Bates-stamped.")

    return {"total_pages": total_pages, "errors": errors}


# ---------- Folder range + renaming ----------

def collect_folder_bates_ranges(root: Path):
    """
    Build a mapping: folder_path -> (min_bates, max_bates)
    based on all Bates-labeled files inside that folder (recursively).

    Uses current filenames (call AFTER file renames).
    """
    folder_ranges = {}

    for path in iter_finder_order_files(root):
        if not path.is_file():
            continue
        if BACKUP_FOLDER_NAME in path.parts:
            continue

        m = BATES_NAME_PATTERN.match(path.stem)
        if not m:
            continue

        start = int(m.group("start"))
        end_str = m.group("end")
        end = int(end_str) if end_str else start

        parent = path.parent
        while True:
            if BACKUP_FOLDER_NAME in parent.parts:
                break
            cur = folder_ranges.get(parent)
            if cur is None:
                folder_ranges[parent] = (start, end)
            else:
                folder_ranges[parent] = (min(cur[0], start), max(cur[1], end))

            if parent == root:
                break
            parent = parent.parent

    return folder_ranges


def rename_folders_with_bates(root: Path, folder_ranges):
    """
    Rename folders based on their Bates range.

    Uses:
      - RENAME_FOLDERS (toggle)
      - KEEP_FOLDER_NAME (toggle)
    """
    if not folder_ranges:
        return []

    renames = []

    # Deepest first so child paths remain valid as we rename
    dirs = sorted(
        folder_ranges.keys(),
        key=lambda p: len(p.relative_to(root).parts),
        reverse=True,
    )

    for folder in dirs:
        if folder == root:
            continue
        if BACKUP_FOLDER_NAME in folder.parts:
            continue

        name = folder.name
        if name.startswith(".") or name.startswith("~"):
            continue

        start, end = folder_ranges[folder]
        if start <= 0 or end < start:
            continue

        if start == end:
            base = f"{PREFIX} {start:0{DIGITS}d}"
        else:
            base = f"{PREFIX} {start:0{DIGITS}d}-{end:0{DIGITS}d}"

        if KEEP_FOLDER_NAME:
            new_name = f"{base} - {name}"
        else:
            new_name = base

        if new_name == name:
            continue

        dst = folder.with_name(new_name)
        if dst.exists():
            print(f"⚠️ Folder rename skipped (target exists): {folder} -> {dst}")
            continue

        try:
            folder.rename(dst)
            print(f"📁 Renamed folder: {folder} -> {dst}")
            renames.append((str(folder), str(dst)))
        except Exception as e:
            print(f"⚠️ Failed to rename folder {folder} -> {dst}: {e}")

    return renames


# ---------- Combined final PDF ----------

def create_combined_final_pdf(root: Path):
    """
    Combine all Bates-labeled PDFs in order into a single PDF
    named like: 'CF 0001- CF 0244.pdf' covering the full range.
    """
    folder_ranges = collect_folder_bates_ranges(root)
    if root not in folder_ranges:
        print("ℹ️  No Bates range found for root; skipping combined PDF.")
        return None

    start, end = folder_ranges[root]
    if start <= 0 or end < start:
        print("ℹ️  Invalid Bates range for root; skipping combined PDF.")
        return None

    # Collect PDFs with Bates in filename, sorted by start number
    pdf_infos = []
    for path in iter_finder_order_files(root):
        if not path.is_file():
            continue
        if path.suffix.lower() != PDF_EXT:
            continue
        if BACKUP_FOLDER_NAME in path.parts:
            continue

        m = BATES_NAME_PATTERN.match(path.stem)
        if not m:
            continue
        s = int(m.group("start"))
        pdf_infos.append((s, path))

    if not pdf_infos:
        print("ℹ️  No Bates-labeled PDFs to combine.")
        return None

    pdf_infos.sort(key=lambda t: t[0])

    out_name = f"{PREFIX} {start:0{DIGITS}d}- {PREFIX} {end:0{DIGITS}d}.pdf"
    out_path = root / out_name

    if DRY_RUN:
        print(f"(DRY RUN) Would create combined PDF: {out_path}")
        return str(out_path)

    writer = PdfWriter()
    for _, pdf_path in pdf_infos:
        try:
            reader = PdfReader(str(pdf_path))
            for page in reader.pages:
                writer.add_page(page)
        except Exception as e:
            print(f"⚠️  Skipping {pdf_path} while combining: {e}")

    with open(out_path, "wb") as f:
        writer.write(f)

    print(f"✅ Created combined PDF: {out_path}")
    return str(out_path)


# ---------- Public entrypoint used by GUI/CLI ----------

def run_pipeline(
    root_folder: str,
    prefix: str = "CF",
    digits: int = 4,
    start_counter: int = 1,
    dry_run: bool = True,
    backup_before_bates: bool = True,
    keep_original_name: bool = True,
    rename_folders: bool = False,
    keep_original_folder_name: bool = True,
    number_videos_at_end: bool = True,
    combine_final: bool = False,
    conversion_only: bool = False,
):
    """
    Run full pipeline and return a summary dict:

    {
        "total_files": int,
        "total_pages": int,
        "renamed": [(src, dst), ...],
        "skipped": [str, ...],
        "errors": [str, ...],
    }
    """
    global ROOT_FOLDER, PREFIX, DIGITS, START_COUNTER, DRY_RUN
    global BACKUP_BEFORE_BATES, KEEP_ORIGINAL_NAME, RENAME_FOLDERS
    global KEEP_FOLDER_NAME, NUMBER_VIDEOS_AT_END, COMBINE_FINAL, CONVERSION_ONLY

    ROOT_FOLDER = root_folder
    PREFIX = prefix
    DIGITS = digits
    START_COUNTER = start_counter
    DRY_RUN = dry_run
    BACKUP_BEFORE_BATES = backup_before_bates
    KEEP_ORIGINAL_NAME = keep_original_name
    RENAME_FOLDERS = rename_folders
    KEEP_FOLDER_NAME = keep_original_folder_name
    NUMBER_VIDEOS_AT_END = number_videos_at_end
    COMBINE_FINAL = combine_final
    CONVERSION_ONLY = conversion_only

    root = Path(ROOT_FOLDER)
    if not root.is_dir():
        raise ValueError(f"Root folder not found: {root}")

    print(f"📂 Scanning recursively (Finder-style): {root}")
    print(f"Keep original filename after Bates (files): {KEEP_ORIGINAL_NAME}")
    print(f"Rename folders with Bates ranges: {RENAME_FOLDERS}")
    if RENAME_FOLDERS:
        print(f"Keep original folder name after Bates: {KEEP_FOLDER_NAME}")
    print(f"Number videos at end: {NUMBER_VIDEOS_AT_END}")
    print(f"Create combined final PDF: {COMBINE_FINAL}")
    print(f"Conversion-only mode: {CONVERSION_ONLY}")

    # Backup originals once at the very start (if enabled, non-dry-run)
    if BACKUP_BEFORE_BATES and not DRY_RUN:
        backup_originals(root)

    # === CONVERSION ONLY MODE ===
    if CONVERSION_ONLY:
        renamed_list = []
        error_list = []
        skipped_list = []

        # Run all conversions (images, HTML, TXT, DOCX)
        img_conv, img_err = convert_images_in_tree(root, delete_original=not DRY_RUN)
        html_conv, html_err = convert_htmls_in_tree(root, delete_original=not DRY_RUN)
        txt_conv, txt_err = convert_txts_in_tree(root, delete_original=not DRY_RUN)
        docx_conv, docx_err = convert_docx_in_tree(root)

        renamed_list.extend(img_conv)
        renamed_list.extend(html_conv)
        renamed_list.extend(txt_conv)
        renamed_list.extend(docx_conv)

        error_list.extend(img_err)
        error_list.extend(html_err)
        error_list.extend(txt_err)
        error_list.extend(docx_err)

        # Reformat all PDFs to Letter
        pdfs = [
            p for p in iter_finder_order_files(root)
            if p.is_file() and p.suffix.lower() == PDF_EXT
            and BACKUP_FOLDER_NAME not in p.parts
        ]

        if DRY_RUN:
            print("\n(DRY RUN) Would reformat all PDFs to US Letter (conversion-only mode).")
        else:
            print("\n--- REFORMAT ALL PDFs TO US LETTER (conversion-only mode) ---")
            for pdf in pdfs:
                try:
                    reformat_pdf_to_letter_in_place(pdf)
                except Exception as e:
                    msg = f"{pdf}: {e}"
                    print(f"⚠️  Error reformatting {msg}")
                    error_list.append(msg)

        total_files = len(pdfs)
        total_pages = 0  # Not computed here

        print("\n✅ Conversion-only pipeline complete (no renaming / no Bates).")

        return {
            "total_files": total_files,
            "total_pages": total_pages,
            "renamed": renamed_list,
            "skipped": skipped_list,
            "errors": error_list,
        }

    # === FULL PIPELINE (with renaming / Bates) ===

    # 0. Auto-convert images, HTML, TXT (DOCX handled in plan_items)
    image_conversions, image_errors = convert_images_in_tree(root, delete_original=not DRY_RUN)
    html_conversions, html_errors = convert_htmls_in_tree(root, delete_original=not DRY_RUN)
    txt_conversions, txt_errors = convert_txts_in_tree(root, delete_original=not DRY_RUN)

    # 1. Block unsupported file types (.doc/.eml/.msg)
    blocking = find_blocking_files(root)
    if blocking:
        print("\n❌ Blocked file types detected (.doc/.eml/.msg). Remove or handle these before running:")
        for p in blocking:
            print(f" - {p}")
        return {
            "total_files": 0,
            "total_pages": 0,
            "renamed": image_conversions + html_conversions + txt_conversions,
            "skipped": [str(p) for p in blocking],
            "errors": ["Blocked file types detected. Run aborted."]
                      + image_errors + html_errors + txt_errors,
        }

    # 2. Build logical items
    items = plan_items(root)
    if not items:
        print("No eligible files found to process.")
        return {
            "total_files": 0,
            "total_pages": 0,
            "renamed": image_conversions + html_conversions + txt_conversions,
            "skipped": [],
            "errors": ["No eligible files found to process."]
                      + image_errors + html_errors + txt_errors,
        }

    items = reorder_items_for_videos(items)

    # 3. Build rename operations
    operations, _ = build_renames(items)

    renamed_list = []
    renamed_list.extend(image_conversions)
    renamed_list.extend(html_conversions)
    renamed_list.extend(txt_conversions)
    renamed_list.extend((str(src), str(dst)) for (src, dst) in operations)

    skipped_list = []
    error_list = []
    error_list.extend(image_errors)
    error_list.extend(html_errors)
    error_list.extend(txt_errors)
    total_files = len(items)
    total_pages = 0

    combined_path = None

    if DRY_RUN:
        print("\n🔎 Dry run enabled — no files or folders will be modified.")
        print(f"Planned file renames ({len(operations)}):")
        for src, dst in operations:
            if src != dst:
                print(f"  {src} -> {dst}")
        if RENAME_FOLDERS:
            print("Folder renaming is enabled, but only simulated in dry run.")
        if COMBINE_FINAL:
            print("Combined final PDF option is enabled, but only simulated in dry run.")
    else:
        apply_renames(operations)

        # Optional folder rename based on Bates ranges (uses renamed filenames)
        if RENAME_FOLDERS:
            folder_ranges = collect_folder_bates_ranges(root)
            folder_renames = rename_folders_with_bates(root, folder_ranges)
            renamed_list.extend(folder_renames)

        bates_result = apply_bates_to_all_pdfs(root)
        total_pages = bates_result.get("total_pages", 0)
        error_list.extend(bates_result.get("errors", []))

        if COMBINE_FINAL:
            combined_path = create_combined_final_pdf(root)
            if combined_path:
                renamed_list.append(("COMBINED", combined_path))

    print("\n✅ All steps complete.")

    return {
        "total_files": total_files,
        "total_pages": total_pages,
        "renamed": renamed_list,
        "skipped": skipped_list,
        "errors": error_list,
    }


# ---------- CLI wrapper ----------

def parse_args_or_prompt():
    parser = argparse.ArgumentParser(
        description="Bates rename & stamp pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent(
            """\
            Examples:
              python3 core.py /path/to/folder --dry-run
              python3 core.py /path/to/folder --prefix DEF --digits 5 --start 1001
            """
        ),
    )

    parser.add_argument("root", nargs="?", help="Root folder to process")
    parser.add_argument("--prefix", default=PREFIX, help=f"Bates prefix (default: {PREFIX})")
    parser.add_argument("--digits", type=int, default=DIGITS, help=f"Zero padding (default: {DIGITS})")
    parser.add_argument("--start", type=int, default=START_COUNTER, help=f"Starting number (default: {START_COUNTER})")
    parser.add_argument("--no-backup", action="store_true", help="Disable backup before processing")
    parser.add_argument("--dry-run", action="store_true", help="Preview only (no changes)")
    parser.add_argument(
        "--no-keep-name",
        action="store_true",
        help="Do NOT append original filename after Bates range",
    )
    parser.add_argument(
        "--rename-folders",
        action="store_true",
        help="Rename folders based on Bates ranges of their contents",
    )
    parser.add_argument(
        "--no-folder-keep-name",
        action="store_true",
        help="When renaming folders, do NOT append original folder name",
    )
    parser.add_argument(
        "--videos-inline",
        action="store_true",
        help="Number videos inline (instead of at the end)",
    )
    parser.add_argument(
        "--combine-final",
        action="store_true",
        help="Create a single combined PDF for the full Bates range",
    )
    parser.add_argument(
        "--conversion-only",
        action="store_true",
        help="Conversion-only mode: convert/format only (no renaming, no Bates)",
    )

    args = parser.parse_args()

    if args.root:
        return (
            args.root,
            args.prefix,
            args.digits,
            args.start,
            args.dry_run or DRY_RUN,
            not args.no_backup if not args.dry_run else True,
            not args.no_keep_name,              # keep_original_name
            args.rename_folders,                # rename_folders
            not args.no_folder_keep_name,       # keep_original_folder_name
            not args.videos_inline,             # number_videos_at_end
            args.combine_final,                 # combine_final
            args.conversion_only,               # conversion_only
        )

    # Interactive fallback
    print("\nNo root folder argument provided. Let's set things up interactively.\n")

    while True:
        root = input("Enter root folder path: ").strip().strip('"').strip("'")
        if root and Path(root).is_dir():
            break
        print("Invalid folder. Try again.\n")

    prefix = input(f"Prefix [{PREFIX}]: ").strip() or PREFIX

    digits_str = input(f"Zero padding digits [{DIGITS}]: ").strip()
    digits = int(digits_str) if digits_str.isdigit() else DIGITS

    start_str = input(f"Starting number [{START_COUNTER}]: ").strip()
    start = int(start_str) if start_str.isdigit() else START_COUNTER

    dry_in = input("Dry run only? (y/N): ").strip().lower()
    dry_run = dry_in == "y"

    conv_only_in = input("Conversion-only mode (no renaming / no Bates)? (y/N): ").strip().lower()
    conversion_only = conv_only_in == "y"

    keep_name_in = input("Append original filename after Bates? (Y/n): ").strip().lower()
    keep_original_name = keep_name_in != "n"

    rename_folders_in = input("Rename folders with Bates ranges? (y/N): ").strip().lower()
    rename_folders = rename_folders_in == "y"

    if rename_folders:
        keep_folder_name_in = input(
            "Append original folder name after Bates for folders? (Y/n): "
        ).strip().lower()
        keep_original_folder_name = keep_folder_name_in != "n"
    else:
        keep_original_folder_name = True

    videos_inline_in = input("Number videos inline (instead of at the end)? (y/N): ").strip().lower()
    number_videos_at_end = not (videos_inline_in == "y")

    combine_final_in = input("Create combined final PDF for full Bates range? (y/N): ").strip().lower()
    combine_final = combine_final_in == "y"

    if dry_run:
        backup = True
    else:
        backup_in = input("Backup originals before processing? (Y/n): ").strip().lower()
        backup = backup_in != "n"

    print("\n--- Configuration ---")
    print(f"Root folder: {root}")
    print(f"Prefix: {prefix}")
    print(f"Digits: {digits}")
    print(f"Start #: {start}")
    print(f"Dry run: {dry_run}")
    print(f"Conversion-only mode: {conversion_only}")
    print(f"Backup originals: {backup}")
    print(f"Keep original filename after Bates (files): {keep_original_name}")
    print(f"Rename folders with Bates ranges: {rename_folders}")
    if rename_folders:
        print(f"Keep original folder name after Bates: {keep_original_folder_name}")
    print(f"Number videos at end: {number_videos_at_end}")
    print(f"Create combined final PDF: {combine_final}")
    print("----------------------\n")

    return (
        root,
        prefix,
        digits,
        start,
        dry_run,
        backup,
        keep_original_name,
        rename_folders,
        keep_original_folder_name,
        number_videos_at_end,
        combine_final,
        conversion_only,
    )


if __name__ == "__main__":
    (
        root,
        prefix,
        digits,
        start,
        dry_run,
        backup,
        keep_original_name,
        rename_folders,
        keep_original_folder_name,
        number_videos_at_end,
        combine_final,
        conversion_only,
    ) = parse_args_or_prompt()

    run_pipeline(
        root_folder=root,
        prefix=prefix,
        digits=digits,
        start_counter=start,
        dry_run=dry_run,
        backup_before_bates=backup,
        keep_original_name=keep_original_name,
        rename_folders=rename_folders,
        keep_original_folder_name=keep_original_folder_name,
        number_videos_at_end=number_videos_at_end,
        combine_final=combine_final,
        conversion_only=conversion_only,
    )
