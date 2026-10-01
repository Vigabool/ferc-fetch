#!/usr/bin/env python3
"""
ferc_text.py  (v1.0)
--------------------
Turn FERC eLibrary downloads into AI-ready text bundles. Used by ferc_fetch.py
(--save text / --save all), and runnable on its own against any folder of
downloads, including folders made by Kyle's original download script.
Built on Kyle Olcott's text bundler.

What it does
  1. Extracts text from each file: PDF (page-marked), DOCX, PPTX, XLSX, TXT,
     CSV, RTF. Scanned/image-only files (TIF, image PDFs) are flagged as
     needing OCR and listed in needs_ocr.csv rather than silently dropped.
  2. Caches each document's text under text/ so reruns are fast.
  3. Packs the text, oldest filing first, into bundles of at most
     --max-chars characters (default 3,000,000), each document under a header
     with its accession number, date, type, filer and eLibrary link.
     Documents too large for one bundle are split across bundles in parts.
  4. Writes <name>_Index.txt: one line per filing (including restricted and
     scanned ones) saying where its text is. Give this to the AI with the
     bundles so it knows what exists, not just what it can read.

OCR
  Run OCR (e.g. Adobe Acrobat) on the files in needs_ocr.csv and save the text
  as the .txt path given in its text_path column, or replace the original with
  a searchable PDF of the same name. Rerun and the text is picked up.

Reads PDFs with pypdf, loaded from pypdf.zip beside this file (no install
needed) or from an installed copy. Everything else is standard library.

USAGE
    python ferc_text.py ./Oconto                 # folder from ferc_fetch.py
    python ferc_text.py ./Oconto --name Oconto --max-chars 2000000
"""

import argparse
import csv
import datetime as dt
import io
import logging
import os
import re
import struct
import sys
import zipfile
import xml.etree.ElementTree as ET

try:
    import ferc_package
except ImportError:  # packaging is optional when this file is used alone
    ferc_package = None

# pypdf reads PDFs. An installed copy is used if present; otherwise the
# pypdf.zip shipped beside this file, so nobody has to pip install anything.
# Double-clicking pypdf.zip unpacks it into a "pypdf" folder here, which Python
# would mistake for the library, so each candidate is checked before use.
def _load_pypdf():
    import importlib
    here = os.path.dirname(os.path.abspath(__file__))
    for extra in (None, os.path.join(here, "pypdf.zip"), os.path.join(here, "pypdf")):
        if extra:
            if not os.path.exists(extra):
                continue
            sys.path.insert(0, extra)
        for m in [m for m in sys.modules if m == "pypdf" or m.startswith("pypdf.")]:
            del sys.modules[m]
        importlib.invalidate_caches()
        try:
            mod = importlib.import_module("pypdf")
            if hasattr(mod, "PdfReader"):
                return mod
        except ImportError:
            pass
        if extra:
            sys.path.remove(extra)
    return None


pypdf = _load_pypdf()
HAVE_PYPDF = pypdf is not None
if HAVE_PYPDF:
    logging.getLogger("pypdf").setLevel(logging.ERROR)

PYPDF_HINT = ("Reading PDFs needs pypdf.zip in the same folder as ferc_text.py.\n"
              "Copy it back from the shared folder, or install pypdf with:\n"
              "    python -m pip install pypdf")

MAX_CHARS = 3_000_000
NOTEBOOK_FILE_LIMIT = 300
TEXT_DIR = "text"
BUNDLE_DIR = "bundles"
SKIP_DIRS = {TEXT_DIR, BUNDLE_DIR}
TEXT_COLS = ["text_status", "text_note", "text_chars", "text_path", "bundle_file"]

# Statuses where text exists in the cache.
HAS_TEXT = ("extracted", "partial-ocr", "ocr-supplied")
OCR_TYPES = (".tif", ".tiff", ".jpg", ".jpeg", ".png", ".bmp", ".gif")
# cp1252 punctuation that DOS-era (cp437) text never uses; cp437 accented letters.
CP1252_MARKS = set(b"\x91\x92\x93\x96\x97")
CP437_LETTERS = set(range(0x80, 0xA6))


# ----------------------------------------------------------------------------
# Extraction
# ----------------------------------------------------------------------------

def read_text_safely(raw):
    """Decode text bytes: BOM first, then UTF-8, BOM-less UTF-16, cp1252, latin-1."""
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig")
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        pass
    sample = raw[:4096]
    if sample:
        pairs = max(1, len(sample) // 2)
        for nulls, codec in ((sample[1::2].count(0), "utf-16-le"),
                             (sample[0::2].count(0), "utf-16-be")):
            if nulls / pairs > 0.20:
                try:
                    return raw.decode(codec)
                except UnicodeError:
                    pass
    # Older FERC .TXT files are DOS text (cp437): "Hébert" is byte 0x82, which
    # cp1252 would show as "H‚bert".
    high = {b for b in raw if b >= 0x80}
    if high and not high & CP1252_MARKS and high <= CP437_LETTERS:
        return raw.decode("cp437")
    try:
        return raw.decode("cp1252")
    except UnicodeDecodeError:
        return raw.decode("latin-1")


def clean(text):
    text = text.replace("\x00", "").replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t\f\v]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _letters(s):
    return sum(c.isalnum() for c in s)


def _fragmented(t):
    """True when a page came out roughly one word per line."""
    lines = [l for l in t.split("\n") if l.strip()]
    return len(lines) >= 30 and sum(len(l.split()) <= 1 for l in lines) / len(lines) > 0.6


def extract_pdf(data):
    reader = pypdf.PdfReader(io.BytesIO(data))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception:
            return "", "error", "encrypted PDF"
    pages, blank = [], 0
    for i, page in enumerate(reader.pages, 1):
        try:
            t = page.extract_text() or ""
            if _fragmented(t):
                # Some PDFs place each word separately; layout mode rebuilds lines.
                t = page.extract_text(extraction_mode="layout") or t
                t = "\n".join(re.sub(r"[ \t]{2,}", "  ", l).strip() for l in t.split("\n"))
        except Exception:
            t = ""
        if _letters(t) < 30:
            blank += 1
        pages.append(f"[Page {i}]\n{t.strip()}")
    n = len(pages)
    if not n or blank == n:
        return "", "needs-ocr", f"no text layer on {n} pages"
    text = "\n\n".join(pages)
    if blank / n > 0.2:
        return text, "partial-ocr", f"{blank} of {n} pages have no text layer"
    return text, "extracted", f"{n} pages"


def _xml_text(data, para_tag):
    """Text from Office XML: one line per paragraph-like element."""
    root = ET.fromstring(data)
    lines = []
    for p in root.iter():
        if p.tag.rsplit("}", 1)[-1] != para_tag:
            continue
        parts = []
        for el in p.iter():
            tag = el.tag.rsplit("}", 1)[-1]
            if tag == "t" and el.text:
                parts.append(el.text)
            elif tag == "tab":
                parts.append("\t")
            elif tag in ("br", "cr"):
                parts.append("\n")
        lines.append("".join(parts))
    return "\n".join(lines)


def extract_docx(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    names = ["word/document.xml"] + sorted(
        n for n in z.namelist() if re.match(r"word/(footnotes|endnotes)\.xml$", n))
    return "\n\n".join(_xml_text(z.read(n), "p") for n in names if n in z.namelist())


def extract_pptx(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    slides = sorted((n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                    key=lambda n: int(re.search(r"(\d+)", n.rsplit("/", 1)[1]).group(1)))
    return "\n\n".join(f"[Slide {i}]\n{_xml_text(z.read(n), 'p')}"
                       for i, n in enumerate(slides, 1))


def extract_xlsx(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(ns + "si"):
            shared.append("".join(t.text or "" for t in si.iter(ns + "t")))
    out = []
    sheets = sorted(n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
    for i, name in enumerate(sheets, 1):
        rows = []
        for row in ET.fromstring(z.read(name)).iter(ns + "row"):
            cells = []
            for c in row.iter(ns + "c"):
                v = c.find(ns + "v")
                if c.get("t") == "s" and v is not None:
                    cells.append(shared[int(v.text)])
                elif c.get("t") == "inlineStr":
                    cells.append("".join(t.text or "" for t in c.iter(ns + "t")))
                elif v is not None:
                    cells.append(v.text or "")
            if any(cells):
                rows.append("\t".join(cells))
        out.append(f"[Sheet {i}]\n" + "\n".join(rows))
    return "\n\n".join(out)


def _ole_streams(data):
    """Minimal OLE2 compound-file reader: {stream name: bytes}."""
    if data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        raise ValueError("not an OLE2 file")
    sec = 1 << struct.unpack_from("<H", data, 30)[0]
    mini_sec = 1 << struct.unpack_from("<H", data, 32)[0]
    dir_start, mini_cutoff, minifat_start = struct.unpack_from("<I4xII", data, 48)
    difat_start, n_difat = struct.unpack_from("<II", data, 68)

    def sector(i):
        return data[512 + i * sec: 512 + (i + 1) * sec]

    fat_secs = [x for x in struct.unpack_from("<109I", data, 76) if x < 0xFFFFFFFA]
    d = difat_start
    for _ in range(n_difat):
        if d >= 0xFFFFFFFA:
            break
        vals = struct.unpack(f"<{sec // 4}I", sector(d))
        fat_secs += [x for x in vals[:-1] if x < 0xFFFFFFFA]
        d = vals[-1]
    fat = []
    for f in fat_secs:
        fat += struct.unpack(f"<{sec // 4}I", sector(f))

    def chain(start, table, read):
        out, i, seen = [], start, set()
        while i < 0xFFFFFFFA and i < len(table) and i not in seen:
            seen.add(i)
            out.append(read(i))
            i = table[i]
        return b"".join(out)

    dirs = chain(dir_start, fat, sector)
    entries = []
    for k in range(0, len(dirs), 128):
        e = dirs[k:k + 128]
        nlen = struct.unpack_from("<H", e, 64)[0]
        name = e[:max(0, nlen - 2)].decode("utf-16-le", "replace")
        etype = e[66]
        start, size = struct.unpack_from("<II", e, 116)
        entries.append((name, etype, start, size))
    root = next(e for e in entries if e[1] == 5)
    mini_stream = chain(root[2], fat, sector)
    minifat = []
    if minifat_start < 0xFFFFFFFA:
        mf = chain(minifat_start, fat, sector)
        minifat = list(struct.unpack(f"<{len(mf) // 4}I", mf))

    streams = {}
    for name, etype, start, size in entries:
        if etype != 2:
            continue
        if size < mini_cutoff:
            body = chain(start, minifat,
                         lambda i: mini_stream[i * mini_sec:(i + 1) * mini_sec])
        else:
            body = chain(start, fat, sector)
        streams[name] = body[:size]
    return streams


def extract_doc(data):
    """Word 97-2003 .doc text via the document's piece table."""
    st = _ole_streams(data)
    wd = st["WordDocument"]
    flags = struct.unpack_from("<H", wd, 0x0A)[0]
    if flags & 0x0100:
        raise ValueError("encrypted .doc")
    table = st["1Table" if flags & 0x0200 else "0Table"]
    fc_clx, lcb_clx = struct.unpack_from("<II", wd, 0x01A2)
    clx = table[fc_clx:fc_clx + lcb_clx]
    i = 0
    while i < len(clx) and clx[i] == 0x01:  # skip Prc formatting entries
        i += 3 + struct.unpack_from("<H", clx, i + 1)[0]
    if i >= len(clx) or clx[i] != 0x02:
        raise ValueError("no piece table")
    lcb = struct.unpack_from("<I", clx, i + 1)[0]
    plc = clx[i + 5:i + 5 + lcb]
    n = (len(plc) - 4) // 12
    cps = struct.unpack_from(f"<{n + 1}I", plc, 0)
    parts = []
    for k in range(n):
        fc = struct.unpack_from("<I", plc, 4 * (n + 1) + 8 * k + 2)[0]
        count = cps[k + 1] - cps[k]
        if fc & 0x40000000:
            start = (fc & 0x3FFFFFFF) // 2
            parts.append(wd[start:start + count].decode("cp1252", "replace"))
        else:
            parts.append(wd[fc:fc + 2 * count].decode("utf-16-le", "replace"))
    text = "".join(parts)
    # Word control characters: paragraph/cell marks, field codes, objects.
    text = re.sub(r"\x13[^\x14\x15]*\x15", "", text)  # fields with no result (PAGE etc.)
    text = re.sub(r"\x13[^\x14\x15]*\x14", "", text)  # field instructions, keep results
    text = text.replace("\r", "\n").replace("\x07", "\t").replace("\x0b", "\n").replace("\x0c", "\n")
    return re.sub(r"[\x00-\x08\x0e-\x1f]", "", text)


def extract_rtf(data):
    s = read_text_safely(data)
    s = re.sub(r"\\'([0-9a-fA-F]{2})", lambda m: bytes.fromhex(m.group(1)).decode("cp1252"), s)
    s = re.sub(r"\\(par|line)\b ?", "\n", s)
    s = re.sub(r"\{\\\*[^{}]*\}", "", s)
    s = re.sub(r"\\[a-zA-Z]+-?\d* ?", "", s)
    return s.replace("{", "").replace("}", "")


ZIP_MAX_FILES = 500
ZIP_MAX_BYTES = 1_000_000_000  # stop reading a zip after 1 GB uncompressed


def extract_zip(data, depth):
    """Text of every readable file inside a zip, each under a '[File in zip: ...]' line."""
    z = zipfile.ZipFile(io.BytesIO(data))
    members = [i for i in z.infolist() if not i.is_dir()
               and not i.filename.startswith("__MACOSX/")
               and not os.path.basename(i.filename).startswith(".")]
    parts, counts, total = [], {}, 0
    for i in members[:ZIP_MAX_FILES]:
        total += i.file_size
        if total > ZIP_MAX_BYTES:
            counts["skipped (size limit)"] = counts.get("skipped (size limit)", 0) + 1
            continue
        if i.flag_bits & 0x1:
            st, text = "error", ""  # password-protected member
        else:
            text, st, _ = extract(i.filename, z.read(i), depth + 1)
        counts[st] = counts.get(st, 0) + 1
        if text:
            parts.append(f"[File in zip: {i.filename}]\n{text}")
    note = f"{len(members)} files in zip: " + ", ".join(f"{n} {k}" for k, n in counts.items())
    if parts:
        return "\n\n".join(parts), ("partial-ocr" if counts.get("needs-ocr") else "extracted"), note
    for st in ("needs-ocr", "error", "empty", "unsupported"):
        if counts.get(st):
            return "", st, note
    return "", "empty", note or "empty zip"


def extract(name, data, depth=0):
    """Return (text, status, note). status: extracted, partial-ocr, needs-ocr,
    unsupported, empty, error."""
    ext = os.path.splitext(name.lower())[1]
    try:
        if ext == ".zip":
            if depth >= 2:
                return "", "unsupported", "zip nested too deeply"
            text, status, note = extract_zip(data, depth)
            return clean(text), status, note
        if ext == ".pdf":
            if not HAVE_PYPDF:
                return "", "error", "pypdf not installed"
            text, status, note = extract_pdf(data)
        elif ext in OCR_TYPES:
            return "", "needs-ocr", "scanned image"
        elif ext == ".docx":
            text, status, note = extract_docx(data), "extracted", ""
        elif ext == ".pptx":
            text, status, note = extract_pptx(data), "extracted", ""
        elif ext == ".xlsx":
            text, status, note = extract_xlsx(data), "extracted", ""
        elif ext == ".doc":
            text, status, note = extract_doc(data), "extracted", ""
        elif ext == ".rtf":
            text, status, note = extract_rtf(data), "extracted", ""
        elif ext in (".txt", ".csv", ".xml", ".htm", ".html"):
            text, status, note = read_text_safely(data), "extracted", ""
        else:
            return "", "unsupported", f"{ext or 'no extension'} files are not read"
    except Exception as e:
        return "", "error", f"{type(e).__name__}: {e}"[:200]
    text = clean(text)
    if status == "extracted" and _letters(text) < 30:
        return "", "empty", "no readable text"
    return text, status, note


def text_rel_path(row):
    """Where a row's text lives, relative to the output folder."""
    if row.get("text_path"):
        return row["text_path"]
    rel = row.get("relative_path") or ""
    return os.path.join(TEXT_DIR, rel + ".txt") if rel else ""


def save_text(out, row, name, data):
    """Extract text from file bytes into the cache and record the result on the row.

    Returns True if text was saved.
    """
    text, status, note = extract(name, data)
    row["text_status"], row["text_note"], row["_fresh"] = status, note, True
    if not text:
        row["text_chars"] = 0
        return False
    tpath = os.path.join(out, row["text_path"])
    os.makedirs(os.path.dirname(tpath), exist_ok=True)
    with open(tpath, "w", encoding="utf-8") as f:
        f.write(text)
    row["text_chars"] = len(text)
    return True


def prepare_text(out, rows, log=print, refresh=False):
    """Make sure every row whose file is on disk has cached text (or a status).

    refresh re-reads text from original files even when text is cached, except
    where re-reading finds none (e.g. a scan whose cached text came from OCR).
    """
    done = 0
    for row in rows:
        row["text_path"] = text_rel_path(row)
        if not row["text_path"]:
            continue
        tpath = os.path.join(out, row["text_path"])
        fpath = os.path.join(out, row.get("relative_path") or "")
        cached = os.path.exists(tpath) and os.path.getsize(tpath) > 0
        if refresh and cached and row.get("relative_path") and os.path.isfile(fpath) \
                and not row.get("_fresh"):
            with open(fpath, "rb") as f:
                data = f.read()
            text, status, note = extract(row.get("file_name") or fpath, data)
            if text:
                with open(tpath, "w", encoding="utf-8") as f:
                    f.write(text)
                row.update(text_status=status, text_note=note, text_chars=len(text), _fresh=True)
                done += 1
                continue
        if cached:
            if row.get("text_status") in ("needs-ocr", "partial-ocr", "unsupported",
                                          "empty", "error"):
                # Text appeared since the last run, e.g. OCR output the user saved here.
                row["text_status"], row["text_note"] = "ocr-supplied", ""
            elif row.get("text_status") not in HAS_TEXT:
                row["text_status"] = "extracted"
            row["text_chars"] = row.get("text_chars") or os.path.getsize(tpath)
            continue
        if row.get("_fresh") or not row.get("relative_path") or not os.path.isfile(fpath):
            continue
        with open(fpath, "rb") as f:
            save_text(out, row, row.get("file_name") or fpath, f.read())
        done += 1
        if done % 50 == 0:
            log(f"    ... extracted text from {done} files")
    return done


# ----------------------------------------------------------------------------
# Bundling
# ----------------------------------------------------------------------------

DIV = "=" * 100


def doc_header(row, part=1, parts=1):
    acc = row.get("accession_number") or "unknown"
    part_s = f"   (PART {part} OF {parts})" if parts > 1 else ""
    kind = " - ".join(x for x in (row.get("document_class"), row.get("document_type")) if x)
    fields = [
        ("Filed", row.get("filed_date")),
        ("Category", row.get("category")),
        ("Type", kind),
        ("Dockets", row.get("dockets") or row.get("docket")),
        ("Author/filer", row.get("author")),
        ("Recipient", row.get("recipient")),
        ("Description", row.get("description")),
        ("File", row.get("file_name")),
        ("eLibrary", row.get("elibrary_url")),
        ("Note", row.get("text_note") if row.get("text_status") == "partial-ocr" else ""),
    ]
    body = "\n".join(f"{k}: {v}" for k, v in fields if v)
    return f"{DIV}\nBEGIN DOCUMENT  Accession {acc}{part_s}\n{body}\n{DIV}\n\n"


def doc_footer(row, part=1, parts=1):
    part_s = f" (part {part} of {parts})" if parts > 1 else ""
    return f"\n\n{DIV}\nEND DOCUMENT  Accession {row.get('accession_number') or 'unknown'}{part_s}\n{DIV}\n\n"


def split_text(text, size):
    """Split at page markers or line breaks so each piece is at most `size` chars."""
    pieces = []
    while len(text) > size:
        cut = text.rfind("\n[Page ", 0, size)
        if cut < size // 2:
            cut = text.rfind("\n", 0, size)
        if cut < size // 2:
            cut = size
        pieces.append(text[:cut])
        text = text[cut:].lstrip("\n")
    pieces.append(text)
    return pieces


def sort_key(row):
    return (row.get("filed_date") or "9999", row.get("accession_number") or "",
            row.get("file_name") or "")


def build_bundles(out, rows, name, max_chars=MAX_CHARS, log=print):
    """Write bundles/<name>_Bundle_NNN.txt and <name>_Index.txt; set row['bundle_file']."""
    bdir = os.path.join(out, BUNDLE_DIR)
    os.makedirs(bdir, exist_ok=True)
    prefix = safe_name(name)
    for f in os.listdir(bdir):  # bundles/ is generated: rebuild it from scratch
        if f.endswith(".txt"):
            os.remove(os.path.join(bdir, f))

    # Blocks in chronological order; large documents split into parts.
    room = max_chars - 3000  # header + footer allowance
    blocks = []
    for row in sorted(rows, key=sort_key):
        row["bundle_file"] = ""
        if row.get("text_status") not in HAS_TEXT or not row.get("text_path"):
            continue
        tpath = os.path.join(out, row["text_path"])
        if not os.path.exists(tpath):
            continue
        with open(tpath, encoding="utf-8", errors="replace") as f:
            text = f.read()
        pieces = split_text(text, room)
        for i, piece in enumerate(pieces, 1):
            block = doc_header(row, i, len(pieces)) + piece + doc_footer(row, i, len(pieces))
            blocks.append((row, block))

    # Greedy packing.
    bundles, cur, size = [], [], 0
    for row, block in blocks:
        if cur and size + len(block) > max_chars:
            bundles.append(cur)
            cur, size = [], 0
        cur.append((row, block))
        size += len(block)
    if cur:
        bundles.append(cur)

    total = len(bundles)
    written = []
    for n, items in enumerate(bundles, 1):
        fname = f"{prefix}_Bundle_{n:03d}.txt"
        docs = list(dict.fromkeys(r.get("accession_number") for r, _ in items))
        dates = [r.get("filed_date") for r, _ in items if r.get("filed_date")]
        header = (
            f"PROJECT: {name}\n"
            f"BUNDLE {n:03d} OF {total:03d}: {len(docs)} filings"
            + (f", filed {min(dates)} to {max(dates)}" if dates else "") + "\n"
            f"SOURCE: FERC eLibrary (https://elibrary.ferc.gov/eLibrary/search)\n"
            f"Each document starts with a BEGIN DOCUMENT block giving its accession number,\n"
            f"filing date, type and filer. Cite accession numbers when referring to documents.\n"
            f"Documents are in filing-date order across bundles. The complete list of filings,\n"
            f"including any without text here, is in {prefix}_Index.txt.\n"
            f"{'#' * 100}\n\n"
        )
        with open(os.path.join(bdir, fname), "w", encoding="utf-8") as f:
            f.write(header)
            for row, block in items:
                f.write(block)
                if fname not in row["bundle_file"].split("; "):
                    row["bundle_file"] = "; ".join(x for x in (row["bundle_file"], fname) if x)
        written.append((fname, len(docs), sum(len(b) for _, b in items)))

    index_path = write_index(out, rows, name, prefix)
    for fname, ndocs, chars in written:
        log(f"  {fname}: {ndocs} filings, {chars:,} characters")
    if total + 1 > NOTEBOOK_FILE_LIMIT:
        log(f"  ! {total + 1} files exceeds a Copilot notebook's {NOTEBOOK_FILE_LIMIT}-file "
            f"limit. Rerun with a larger --max-chars.")
    return written, index_path


def write_index(out, rows, name, prefix):
    """One line per filing, so the AI knows what exists even without text."""
    by_acc = {}
    for row in sorted(rows, key=sort_key):
        by_acc.setdefault(row.get("accession_number") or "unknown", []).append(row)

    def where(rs):
        bundles = sorted({b for r in rs for b in (r.get("bundle_file") or "").split("; ") if b})
        if bundles:
            return "text in " + ", ".join(bundles)
        statuses = {r.get("status", "") for r in rs} | {r.get("text_status", "") for r in rs}
        if "restricted" in statuses:
            return "NO TEXT: restricted (" + (rs[0].get("availability") or "non-public") + ")"
        if "covered" in statuses:
            return "text in FERC .TXT version"
        if "needs-ocr" in statuses:
            return "NO TEXT: scanned image, not yet OCR'd"
        if statuses & {"metadata-only", "no-files-listed"}:
            return "NO TEXT: metadata only"
        if statuses & {"unsupported", "empty"}:
            return "NO TEXT: file format not readable"
        return "NO TEXT: not downloaded"

    lines = [
        f"PROJECT: {name}",
        f"INDEX OF FERC eLibrary FILINGS ({len(by_acc)} filings), generated "
        f"{dt.datetime.now():%Y-%m-%d}",
        "Columns: filed date | accession | category | document type | author/filer | "
        "where the text is | description",
        "#" * 100,
    ]
    for acc, rs in by_acc.items():
        r = rs[0]
        lines.append(" | ".join([
            r.get("filed_date") or "", acc, r.get("category") or "",
            r.get("document_type") or "", r.get("author") or "", where(rs),
            re.sub(r"\s+", " ", r.get("description") or "")]))
    path = os.path.join(out, BUNDLE_DIR, f"{prefix}_Index.txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return path


def mark_covered(rows):
    """A scanned TIF whose filing also has a readable .TXT (FERC's own OCR text)
    is not a gap: mark it so it is not listed for OCR."""
    with_text = {r.get("accession_number") for r in rows
                 if r.get("text_status") in HAS_TEXT
                 and (r.get("file_name") or "").lower().endswith(".txt")}
    for r in rows:
        if (r.get("text_status") == "needs-ocr" and r.get("accession_number") in with_text
                and (r.get("file_name") or "").lower().endswith(OCR_TYPES)):
            r["text_status"], r["text_note"] = "covered", "FERC text version in same filing"


def write_needs_ocr(out, rows):
    path = os.path.join(out, "needs_ocr.csv")
    todo = [r for r in rows if r.get("text_status") in ("needs-ocr", "partial-ocr")]
    if not todo:
        if os.path.exists(path):
            os.remove(path)
        return None, 0
    cols = ["accession_number", "filed_date", "description", "relative_path", "text_status",
            "text_note", "text_path"]
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        w.writerows(sorted(todo, key=sort_key))
    return path, len(todo)


def safe_name(s):
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", str(s)).strip("_") or "FERC"


# ----------------------------------------------------------------------------
# Standalone use
# ----------------------------------------------------------------------------

def load_rows(out):
    """Rows from manifest.csv if present, else one row per file found on disk."""
    mpath = os.path.join(out, "manifest.csv")
    if os.path.exists(mpath):
        with open(mpath, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
        for r in rows:  # older manifests (Kyle's original script)
            r.setdefault("dockets", r.get("docket", ""))
        return rows, True
    rows = []
    for root, dirs, files in os.walk(out):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS and not d.startswith(".")]
        for fn in sorted(files):
            if fn.startswith(".") or fn.lower().endswith((".csv", ".json", ".log")):
                continue
            rel = os.path.relpath(os.path.join(root, fn), out)
            m = re.match(r"(\d{8}-\d{4})", fn)
            acc = m.group(1) if m else ""
            rows.append({"accession_number": acc, "file_name": fn, "relative_path": rel,
                         "filed_date": f"{acc[:4]}-{acc[4:6]}-{acc[6:8]}" if acc else "",
                         "status": "downloaded"})
    return rows, False


def rebuild_folder(out, name=None, fmt="docx", bundles=False, max_chars=MAX_CHARS, log=print,
                   refresh=False):
    """Extract text and rebuild the Copilot package (and optional notebook bundles)
    for an existing download folder. No network needed."""
    if not os.path.isdir(out):
        raise SystemExit(f"Folder not found: {out}")
    rows, from_manifest = load_rows(out)
    if not rows:
        raise SystemExit(f"No files found in {out}")
    if not HAVE_PYPDF and any((r.get("file_name") or "").lower().endswith(".pdf") for r in rows):
        raise SystemExit(PYPDF_HINT)
    name = name or os.path.basename(os.path.abspath(out))
    log(f"{len(rows)} files ({'from manifest.csv' if from_manifest else 'found on disk'})")
    n = prepare_text(out, rows, log, refresh)
    mark_covered(rows)
    log(f"Read text from {n} files" + ("" if refresh else " (others already cached or not on disk)"))
    package_and_bundle(out, rows, name, fmt, bundles, max_chars, (), log)
    ocr_path, n_ocr = write_needs_ocr(out, rows)

    if from_manifest:
        mpath = os.path.join(out, "manifest.csv")
        with open(mpath, newline="", encoding="utf-8-sig") as f:
            cols = csv.DictReader(f).fieldnames or []
        cols += [c for c in TEXT_COLS if c not in cols]
        with open(mpath, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
            w.writeheader()
            w.writerows(rows)
    print_text_summary(rows, out, ocr_path, n_ocr, log)


def package_and_bundle(out, rows, name, fmt, bundles, max_chars, focus, log=print):
    if ferc_package is None:
        raise SystemExit("Building the Copilot package needs ferc_package.py beside ferc_text.py.")
    log("Copilot package:")
    ferc_package.build_package(out, rows, name, fmt, focus, log)
    if bundles:
        log("Notebook bundles:")
        build_bundles(out, rows, name, max_chars, log)


def main():
    p = argparse.ArgumentParser(description="Turn FERC downloads into a Copilot knowledge package.")
    p.add_argument("folder", help="Folder of downloads (e.g. made by ferc_fetch.py)")
    p.add_argument("--name", help="Project name used in the package (default: folder name)")
    p.add_argument("--format", choices=["docx", "md"], default="docx",
                   help="File format for the package (default docx)")
    p.add_argument("--notebook-bundles", action="store_true",
                   help="Also write large combined text files for Copilot notebooks (300-file limit)")
    p.add_argument("--max-chars", type=int, default=MAX_CHARS,
                   help=f"Maximum characters per notebook bundle (default {MAX_CHARS:,})")
    p.add_argument("--refresh-text", action="store_true",
                   help="Re-read text from original files instead of reusing cached text")
    args = p.parse_args()
    rebuild_folder(args.folder, args.name, args.format, args.notebook_bundles, args.max_chars,
                   refresh=args.refresh_text)


def print_text_summary(rows, out, ocr_path, n_ocr, log=print):
    from collections import Counter
    c = Counter(r.get("text_status") or "" for r in rows if r.get("text_path"))
    log(f"  text       : {c['extracted'] + c['ocr-supplied']} files read, "
        f"{c['covered']} scans covered by FERC text, "
        f"{c['partial-ocr']} partly scanned, {c['needs-ocr']} need OCR, "
        f"{c['unsupported'] + c['empty']} unreadable, {c['error']} errors")
    log(f"Copilot package: {os.path.join(out, 'copilot')}")
    if ocr_path:
        log(f"Needs OCR ({n_ocr} files): {ocr_path}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(130)
