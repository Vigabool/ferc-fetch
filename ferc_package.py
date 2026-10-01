#!/usr/bin/env python3
"""
ferc_package.py  (v1.0)
-----------------------
Build a Copilot knowledge package from a ferc_fetch.py download folder.

Copilot (agents, notebooks, SharePoint grounding) never reads a whole record.
It splits files into passages, retrieves a handful per question, and answers
from those. This package is shaped for that:

  Knowledge/
    00 Project Overview      who, what, when: facts compiled from the filing metadata
    01 Timeline              every issuance and application, by date
    02 Catalog/<decade>      every filing, one self-contained line each
    Documents/<year>/        one file per filing, named by date, accession and
                             description, with a context line repeated through
                             the text so any passage says which filing it is from
  Agent Instructions.txt     paste into the agent's instructions

Point a Copilot agent at the Knowledge folder (upload it to SharePoint first).
The overview, timeline and catalog answer whole-record questions ("list the
orders", "was a license issued?") that passage retrieval cannot.

Text cleanup: page numbers, repeated page headers/footers, end-of-document
certificates of service, and duplicate filings are removed.

Output is Word (.docx, default) or Markdown (.md). Standard library only.
"""

import collections
import datetime as dt
import hashlib
import os
import re
import shutil
import zipfile
from xml.sax.saxutils import escape

PACKAGE_DIR = "copilot"
KNOWLEDGE_DIR = "Knowledge"
HAS_TEXT = ("extracted", "partial-ocr", "ocr-supplied")
CTX_EVERY = 1500  # characters between repeated context lines

DECADE_SPLIT = 600  # catalog lines per file before splitting a decade by year


# ----------------------------------------------------------------------------
# Document model: a list of blocks rendered to .docx or .md
# ----------------------------------------------------------------------------

class Doc:
    def __init__(self, title):
        self.title = title
        self.blocks = [("title", title)]

    def h1(self, t): self.blocks.append(("h1", t))
    def h2(self, t): self.blocks.append(("h2", t))
    def p(self, t): self.blocks.append(("p", t))
    def li(self, t): self.blocks.append(("li", t))
    def ctx(self, t): self.blocks.append(("ctx", t))
    def fields(self, pairs): self.blocks.append(("fields", [(k, v) for k, v in pairs if v]))

    def save(self, path_no_ext, fmt):
        path = f"{path_no_ext}.{fmt}"
        os.makedirs(os.path.dirname(path), exist_ok=True)
        (write_docx if fmt == "docx" else write_md)(self, path)
        return path


def write_md(doc, path):
    out, prev = [], None
    for kind, v in doc.blocks:
        if prev == "li" and kind != "li":
            out.append("")
        prev = kind
        if kind == "title":
            out.append(f"# {v}\n")
        elif kind == "h1":
            out.append(f"## {v}\n")
        elif kind == "h2":
            out.append(f"### {v}\n")
        elif kind == "fields":
            out.append("\n".join(f"- **{k}:** {val}" for k, val in v) + "\n")
        elif kind == "li":
            out.append(f"- {v}")
        elif kind == "ctx":
            out.append(f"\n> {v}\n")
        else:
            out.append(v.replace("\n", "  \n") + "\n")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(out) + "\n")


_XML_BAD = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _run(text, bold=False):
    parts = _XML_BAD.sub("", text).split("\n")
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    body = "<w:br/>".join(f'<w:t xml:space="preserve">{escape(x)}</w:t>' for x in parts)
    return f"<w:r>{rpr}{body}</w:r>"


def _para(text, style=None, bold_prefix=None):
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    runs = (_run(bold_prefix, True) if bold_prefix else "") + _run(text)
    return f"<w:p>{ppr}{runs}</w:p>"


STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {W}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
<w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>
<w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
<w:pPr><w:keepNext/><w:spacing w:before="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>
<w:pPr><w:keepNext/><w:spacing w:before="200"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/>
<w:pPr><w:spacing w:after="40"/><w:ind w:left="360" w:hanging="360"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Context"><w:name w:val="Context"/><w:basedOn w:val="Normal"/>
<w:pPr><w:spacing w:before="120" w:after="60"/></w:pPr><w:rPr><w:i/><w:color w:val="666666"/><w:sz w:val="18"/></w:rPr></w:style>
</w:styles>"""


def write_docx(doc, path):
    body = []
    for kind, v in doc.blocks:
        if kind == "title":
            body.append(_para(v, "Title"))
        elif kind == "h1":
            body.append(_para(v, "Heading1"))
        elif kind == "h2":
            body.append(_para(v, "Heading2"))
        elif kind == "fields":
            body += [_para(str(val), None, f"{k}: ") for k, val in v]
        elif kind == "li":
            body.append(_para("• " + v, "ListBullet"))
        elif kind == "ctx":
            body.append(_para(v, "Context"))
        else:
            body.append(_para(v))
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
                + "".join(body) + "<w:sectPr/></w:body></w:document>")
    now = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:title>{escape(_XML_BAD.sub("", doc.title))}</dc:title><dc:creator>FERC eLibrary (ferc_fetch)</dc:creator>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created></cp:coreProperties>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                   '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                   '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                   '</Types>')
        z.writestr("_rels/.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                   '</Relationships>')
        z.writestr("word/_rels/document.xml.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                   '</Relationships>')
        z.writestr("word/document.xml", document)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("docProps/core.xml", core)


# ----------------------------------------------------------------------------
# Text cleanup
# ----------------------------------------------------------------------------

PAGE_NUM = re.compile(
    r"^(?:page\s+)?[-–(]?\s*(?:\d{1,4}|[ivxlcdm]{1,7}|[A-Z]?-?\d{1,3}-\d{1,4})\s*[-–)]?"
    r"(?:\s+of\s+\d{1,4})?$", re.I)
SERVICE = re.compile(r"^\s*(CERTIFICATE\s+OF\s+SERVICE|SERVICE\s+LIST)\s*$", re.I | re.M)


def split_pages(text):
    """[(page number or None, page text)] from '[Page N]' markers."""
    parts = re.split(r"^\[Page (\d+)\]\n", text, flags=re.M)
    if len(parts) == 1:
        return [(None, text)]
    pages = [(None, parts[0])] if parts[0].strip() else []
    pages += [(int(parts[i]), parts[i + 1]) for i in range(1, len(parts) - 1, 2)]
    return pages


def strip_page_furniture(pages):
    """Drop page numbers and header/footer lines repeated across pages."""
    norm = lambda s: re.sub(r"\d+", "#", s.strip().lower())
    edge_count = collections.Counter()
    split = []
    for n, t in pages:
        lines = t.split("\n")
        split.append((n, lines))
        nonblank = [l for l in lines if l.strip()]
        for l in set(norm(x) for x in nonblank[:2] + nonblank[-2:]):
            edge_count[l] += 1
    repeated = ({l for l, c in edge_count.items() if c >= max(4, len(pages) * 0.5) and len(l) > 3}
                if len(pages) >= 4 else set())

    cleaned = []
    for n, lines in split:
        idx = [i for i, l in enumerate(lines) if l.strip()]
        drop = set()
        for i in idx[:3] + idx[-3:]:
            s = lines[i].strip()
            if PAGE_NUM.match(s) or norm(s) in repeated:
                drop.add(i)
        cleaned.append((n, "\n".join(l for i, l in enumerate(lines) if i not in drop).strip()))
    return cleaned


def cut_service_list(text):
    """Remove a certificate of service / service list at the end of a document."""
    ms = list(SERVICE.finditer(text))
    if ms and ms[-1].start() > len(text) * 0.6:
        return text[:ms[-1].start()].rstrip(), True
    return text, False


def chunks(text, size=CTX_EVERY):
    """Split at paragraph breaks into pieces of roughly `size` characters."""
    out, cur = [], ""
    for para in re.split(r"\n{2,}", text):
        if cur and len(cur) + len(para) > size:
            out.append(cur)
            cur = ""
        cur = f"{cur}\n\n{para}" if cur else para
    if cur:
        out.append(cur)
    return out


# ----------------------------------------------------------------------------
# Building blocks
# ----------------------------------------------------------------------------

def base_docket(d):
    m = re.match(r"^([A-Z]+\d*-\d+)(?:-\d{3})?$", d.strip().upper())
    return m.group(1) if m else d.strip()


def project_label(dockets, focus):
    """'P-2523' for the package's own docket, else the filing's, plus others."""
    bases = list(dict.fromkeys(base_docket(d) for d in re.split(r";\s*", dockets or "") if d))
    if not bases:
        return ""
    first = next((b for b in bases if b in focus), bases[0])
    others = [b for b in bases if b != first]
    if len(others) > 2:
        return f"{first} et al. ({len(bases)} dockets)"
    return ", ".join([first] + others)


def short(s, n):
    s = re.sub(r"\s+", " ", s or "").strip()
    return s if len(s) <= n else s[:n - 1].rsplit(" ", 1)[0] + "…"


def file_safe(s, n=80):
    s = re.sub(r'[<>:"/\\|?*#%&{}~\x00-\x1f]', " ", s or "")
    s = re.sub(r"\s+", " ", s).strip(" .")
    return short(s, n).rstrip("… .") or "filing"


def kind(r):
    return " - ".join(x for x in (r.get("document_class"), r.get("document_type")) if x)


def filer(r):
    """First author's organization; for individuals, the person's name."""
    first = (r.get("author") or "").split(";")[0].strip()
    org = re.sub(r"\s*\([^)]*\)", "", first).strip()
    if re.search(r"individual|no affiliation", org, re.I):
        person = re.search(r"\(([^)]*)\)", first)
        return f"{person.group(1)} (individual)" if person else "Individual"
    return org


def org_key(name):
    """'Great River Hydro, LLC' and 'GREAT RIVER HYDRO LLC' count as one filer."""
    k = re.sub(r"[^a-z0-9 ]", " ", name.lower())
    k = re.sub(r"\b(llc|inc|co|corp|corporation|company|lp|llp|ltd|the)\b", " ", k)
    return re.sub(r"\s+", " ", k).strip()


def count_filers(rows):
    """[(display name, count)] grouping spelling variants; individuals excluded."""
    groups = collections.defaultdict(collections.Counter)
    for r in rows:
        f = filer(r)
        if f and not f.endswith("(individual)") and f != "Individual":
            groups[org_key(f)][f] += 1
    out = [(c.most_common(1)[0][0], sum(c.values())) for c in groups.values()]
    return sorted(out, key=lambda x: -x[1])


def first_type(r):
    return (r.get("document_type") or r.get("document_class") or "").split(";")[0].strip()


def plural(n, word):
    return f"{n} {word}" + ("" if n == 1 else "s")


def group_filings(rows):
    by = collections.OrderedDict()
    for r in sorted(rows, key=lambda r: (r.get("filed_date") or "9999",
                                         r.get("accession_number") or "")):
        by.setdefault(r.get("accession_number") or "unknown", []).append(r)
    return by


def no_text_reason(rs):
    st = {r.get("status", "") for r in rs} | {r.get("text_status", "") for r in rs}
    if "restricted" in st:
        return "restricted (" + (rs[0].get("availability") or "non-public") + "); not available"
    if "needs-ocr" in st:
        return "scanned image; text not yet OCR'd"
    if st & {"metadata-only", "no-files-listed"}:
        return "metadata only; files not downloaded"
    if st & {"unsupported", "empty"}:
        return "file format could not be read"
    return "not downloaded"


# ----------------------------------------------------------------------------
# Package
# ----------------------------------------------------------------------------

def build_package(out, rows, name, fmt="docx", focus=(), log=print):
    """Write <out>/copilot/Knowledge/... and Agent Instructions.txt."""
    focus = {base_docket(f) for f in focus}
    pdir = os.path.join(out, PACKAGE_DIR)
    kdir = os.path.join(pdir, KNOWLEDGE_DIR)
    shutil.rmtree(kdir, ignore_errors=True)  # generated: rebuild from scratch
    os.makedirs(kdir, exist_ok=True)

    filings = group_filings(rows)
    if not focus:
        c = collections.Counter(base_docket(d) for rs in filings.values()
                                for d in re.split(r";\s*", rs[0].get("dockets") or "") if d)
        focus = {c.most_common(1)[0][0]} if c else set()

    seen_hash, stats = {}, collections.Counter()
    info = {}  # accession -> (doc file name, has_text, reason)
    for acc, rs in filings.items():
        r0 = rs[0]
        label = project_label(r0.get("dockets"), focus)
        desc = r0.get("description") or kind(r0) or "Filing"
        year = (r0.get("filed_date") or "0000")[:4]
        fname = file_safe(f"{r0.get('filed_date') or ''} {acc} {desc}", 110)

        texts = []
        for r in rs:
            tp = os.path.join(out, r.get("text_path") or "")
            if r.get("text_status") in HAS_TEXT and r.get("text_path") and os.path.isfile(tp):
                with open(tp, encoding="utf-8", errors="replace") as f:
                    texts.append((r, f.read()))

        doc = Doc(desc)
        doc.fields([
            ("Accession number", acc),
            ("Filed", r0.get("filed_date")),
            ("Document date", r0.get("document_date") if r0.get("document_date") != r0.get("filed_date") else ""),
            ("Project / dockets", r0.get("dockets")),
            ("Category", r0.get("category")),
            ("Document type", kind(r0)),
            ("Author / filer", r0.get("author")),
            ("Recipient", r0.get("recipient")),
            ("Library", r0.get("library")),
            ("eLibrary link", r0.get("elibrary_url")),
        ])

        reason = ""
        if not texts:
            reason = no_text_reason(rs)
            doc.p(f"Text not available: {reason}. The description above is from the "
                  f"eLibrary record.")
            stats["stub"] += 1
        else:
            ctx_base = f"{label} · {acc} · {r0.get('filed_date') or ''} · {short(desc, 70)}"
            for r, text in texts:
                if len(texts) > 1:
                    doc.h1(f"File: {r.get('file_name')}")
                # The same attachment is often filed more than once (re-filings,
                # a zip holding the PDF next to it): keep the text once.
                digest = (hashlib.sha1(re.sub(r"\W+", "", text.lower()).encode()).hexdigest()
                          if len(text) > 300 else None)
                if digest in seen_hash:
                    where = seen_hash[digest]
                    doc.p(f"Same text as {where[1]} in accession {where[0]}"
                          + (", above." if where[0] == acc else ", filed earlier; see that document."))
                    stats["duplicate"] += 1
                    continue
                if digest:
                    seen_hash[digest] = (acc, r.get("file_name"))
                text, cut = cut_service_list(text)
                stats["service lists removed"] += cut
                pages = strip_page_furniture(split_pages(text))
                for n, page in pages:
                    if not page:
                        continue
                    for i, piece in enumerate(chunks(page)):
                        doc.ctx(f"[{ctx_base}" + (f" · p. {n}" if n else "") + "]")
                        doc.p(piece)
                if cut:
                    doc.p("[Certificate of service / service list omitted.]")
            stats["documents"] += 1
        path = doc.save(os.path.join(kdir, "Documents", year, fname), fmt)
        info[acc] = (os.path.basename(path), bool(texts), reason)

    write_overview(kdir, name, filings, info, focus, fmt)
    write_timeline(kdir, name, filings, info, fmt)
    n_cat = write_catalog(kdir, name, filings, info, fmt)
    write_instructions(pdir, name, focus)

    log(f"  {stats['documents']} documents with text, {stats['stub']} without text "
        f"(listed with metadata), {stats['duplicate']} duplicate files folded, "
        f"{stats['service lists removed']} service lists removed")
    log(f"  overview, timeline, {n_cat} catalog file(s)")
    return kdir


def filing_line(acc, rs, info, with_text=True):
    r = rs[0]
    parts = [r.get("filed_date") or "", acc, r.get("category") or "", first_type(r),
             ("filed by " if r.get("category") == "Submittal" else "from ") + (filer(r) or "unknown"),
             short(r.get("description"), 300)]
    if with_text:
        _, has, reason = info[acc]
        parts.append("full text in Documents" if has else f"no text: {reason}")
    return " · ".join(p for p in parts if p)


def write_overview(kdir, name, filings, info, focus, fmt):
    rows = [r for rs in filings.values() for r in rs]
    firsts = [rs[0] for rs in filings.values()]
    dates = sorted(r.get("filed_date") for r in firsts if r.get("filed_date"))
    subs = [r for r in firsts if r.get("category") == "Submittal"]
    filers = count_filers(subs)
    individuals = sum(1 for r in subs if filer(r).endswith("(individual)") or filer(r) == "Individual")
    names = collections.Counter()
    for r in firsts:
        for m in re.finditer(r"((?:[A-Z][\w.'&-]*\s){1,4})(?:Hydroelectric|Hydro|Hydropower|Pumped Storage)?\s?Project",
                             r.get("description") or ""):
            n = m.group(1).strip().split("'s ")[-1]
            if n and n.split()[0] not in ("The", "For", "Of", "Under", "Re", "And", "Proposed"):
                names[n] += 1
    with_text = sum(1 for v in info.values() if v[1])

    doc = Doc(f"{name}: project overview")
    doc.p("Facts compiled automatically from FERC eLibrary filing metadata. Counts "
          "cover the filings retrieved for this package, which may be a subset of the "
          "full record. See the timeline and catalog for individual filings.")
    doc.h1("Record")
    doc.fields([
        ("Dockets", ", ".join(sorted(focus)) or "various"),
        ("Project names most often used in filings",
         "; ".join(f"{n} ({c})" for n, c in names.most_common(3))),
        ("Most frequent filer (likely licensee or applicant)",
         f"{filers[0][0]} ({plural(filers[0][1], 'filing')})" if filers else ""),
        ("Filings by individuals (members of the public)", str(individuals) if individuals else ""),
        ("Filings in this package", f"{len(filings)} ({sum(1 for r in firsts if r.get('category') == 'Issuance')} "
                                    f"FERC issuances, {len(subs)} submittals)"),
        ("Date range", f"{dates[0]} to {dates[-1]}" if dates else ""),
        ("Text available", f"{with_text} of {len(filings)} filings"),
    ])

    doc.h1("Filings by decade")
    dec = collections.Counter((d[:3] + "0s") for d in dates)
    for d, c in sorted(dec.items()):
        doc.li(f"{d}: {plural(c, 'filing')}")

    if filers[1:]:
        doc.h1("Most active filers other than the licensee")
        for f, c in filers[1:16]:
            doc.li(f"{f}: {plural(c, 'filing')}")

    doc.h1("Most common document types")
    for t, c in collections.Counter(first_type(r) for r in firsts if first_type(r)).most_common(12):
        doc.li(f"{t}: {c}")

    doc.h1("FERC orders")
    orders = [(a, rs) for a, rs in filings.items() if rs[0].get("category") == "Issuance"
              and "Order" in (rs[0].get("document_class") or "") + (rs[0].get("document_type") or "")]
    for a, rs in orders:
        doc.li(filing_line(a, rs, info))
    if not orders:
        doc.p("No orders among the filings in this package.")

    gaps = collections.Counter(v[2].split(";")[0] for v in info.values() if not v[1])
    if gaps:
        doc.h1("Filings without text")
        for g_, c in gaps.most_common():
            doc.li(f"{plural(c, 'filing')}: {g_}")
    doc.save(os.path.join(kdir, "00 Project Overview"), fmt)


MILESTONE = re.compile(r"applicat|notice of intent|pre-application|rehearing|settlement|"
                       r"study plan|amend|transfer|surrender|exemption|petition", re.I)


def write_timeline(kdir, name, filings, info, fmt):
    doc = Doc(f"{name}: timeline")
    doc.p("Every FERC issuance plus applications, petitions, rehearing requests and "
          "settlements, oldest first. One line per filing: date, accession, category, "
          "type, author, description, and whether its full text is in the Documents folder.")
    year = None
    for acc, rs in filings.items():
        r = rs[0]
        if r.get("category") != "Issuance" and not MILESTONE.search(
                f"{r.get('document_class')} {r.get('document_type')}"):
            continue
        y = (r.get("filed_date") or "Undated")[:4]
        if y != year:
            doc.h1(y)
            year = y
        doc.li(filing_line(acc, rs, info))
    doc.save(os.path.join(kdir, "01 Timeline"), fmt)


def write_catalog(kdir, name, filings, info, fmt):
    groups = collections.OrderedDict()
    for acc, rs in filings.items():
        y = (rs[0].get("filed_date") or "0000")[:4]
        groups.setdefault(y[:3] + "0s", []).append((y, acc, rs))
    n = 0
    for decade, items in groups.items():
        # Split very busy decades into one file per year.
        parts = ([(decade, items)] if len(items) <= DECADE_SPLIT else
                 [(y, [i for i in items if i[0] == y]) for y in dict.fromkeys(i[0] for i in items)])
        for label, its in parts:
            doc = Doc(f"{name}: filing catalog, {label}")
            doc.p("Every filing retrieved for this period, one line each: date, accession, "
                  "category, type, author, description, and whether its full text is in "
                  "the Documents folder.")
            year = None
            for y, acc, rs in its:
                if y != year:
                    doc.h1(y)
                    year = y
                doc.li(filing_line(acc, rs, info))
            doc.save(os.path.join(kdir, "02 Catalog", f"Catalog {label}"), fmt)
            n += 1
    return n


INSTRUCTIONS = """\
You are a research assistant for the FERC record of {name} ({dockets}). Your
knowledge is a package built from FERC eLibrary:

- "00 Project Overview": facts about the record (filers, dates, orders, gaps).
- "01 Timeline": every FERC issuance and application, by date.
- "02 Catalog": every filing in the package, one line each.
- "Documents": the text of each filing, one file per filing, named by filing
  date, accession number and description.

How to answer:
1. For questions about what exists, what happened when, or lists of filings,
   use the overview, timeline and catalog first. They cover every filing,
   including ones whose text is not available.
2. For questions about what a filing says, use the Documents files. Each passage
   starts with a bracketed line giving project, accession number, filing date,
   description and page.
3. Cite the accession number (format YYYYMMDD-NNNN) and filing date for every
   fact you state from a filing, and the page when given.
4. Prefer the most recent order or filing when filings conflict, and say that
   the position changed and when.
5. Never conclude that a filing or event does not exist because you did not find
   it in the documents. Check the catalog. If the catalog lists it without text
   (restricted, scanned or not downloaded), say so and give its accession
   number so the user can open it in eLibrary.
6. If the package does not answer the question, say what you searched and what
   is missing instead of guessing.
"""


def write_instructions(pdir, name, focus):
    with open(os.path.join(pdir, "Agent Instructions.txt"), "w", encoding="utf-8") as f:
        f.write(INSTRUCTIONS.format(name=name, dockets=", ".join(sorted(focus)) or "FERC dockets"))
