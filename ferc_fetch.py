#!/usr/bin/env python3
"""
ferc_fetch.py  (v2.3)
---------------------
Pull filings out of FERC eLibrary into a local folder, with a metadata manifest,
for downstream AI / Copilot notebook use. Built on Kyle Olcott's original
eLibrary download script (v1.2).

You choose exactly what to pull. Every filter on eLibrary's Advanced Search is
available here, in any combination:

    what        docket(s) / sub-docket, keywords (description or full text),
                accession number(s)
    when        date range on filed, document, or posted date
    kind        category (Issuance / Submittal), library (Hydro, Gas, ...),
                document class / type ("Environmental Assessment", "Order", ...)
    who         filer / author / recipient affiliation ("Consumers Energy")
    access      public only (default) or everything (restricted docs listed,
                never downloaded)
    how much    max results, file types, max file size
    output      manifest only, + original files, or + a Copilot knowledge package

Uses the JSON endpoints the eLibrary web UI itself calls:
    POST https://elibrary.ferc.gov/eLibrarywebapi/api/Search/AdvancedSearch
    POST https://elibrary.ferc.gov/eLibrarywebapi/api/File/DownloadP8File
    GET  https://elibrary.ferc.gov/eLibrarywebapi/api/Search/GetClassTypes

v2.3 changes:
  - Two stages. Stage 1 downloads originals and metadata (always kept).
    Stage 2 builds the Copilot package from the folder; it runs after a
    download by default (--save package) or alone with --rebuild. Improvements
    to text extraction reach existing folders with --rebuild --refresh-text,
    no re-download. '--save text' and '--save all' now mean '--save package'.

v2.2 changes:
  - --save text/all builds a Copilot knowledge package (ferc_package.py):
    one file per filing with self-identifying passages, plus a project
    overview, timeline and catalog. Notebook bundles are now optional
    (--notebook-bundles). --rebuild regenerates it with no download.

v2.1 changes:
  - --save chooses the output: manifest | files | text | all. 'text' keeps
    originals only where text could not be read; 'all' keeps every original.

v2.0 changes:
  - New filters: --text, --full-text, --accession, --date-type, --library,
    --doc-type, --affiliation/--role, --max-results, --sort. Docket is optional.
  - Guided mode: run with no arguments (or -i) and it asks what you want.
  - --preview shows the match count and first results without downloading.
  - Saved queries: --save-query recipe.json / --query recipe.json.
  - FIX: date filter. v1.2 sent fromDate/toDate, which the API silently ignores,
    so --from/--to returned every filing. Now uses startDate/endDate.
  - FIX: public-only filter. v1.2 sent availability "public", which the API
    silently ignores, so CEII filings came back and failed as 401s. Now ["P"].
  - FIX: manifest columns class_type, library, availability were always blank
    (wrong field names). Now filled in, plus document class/type, author,
    recipient, all docket numbers, file size, and an eLibrary link.
  - Skips oversized files using the listed size, before downloading.
  - A filing on several requested dockets is downloaded once, not per docket.
  - Reruns merge into the existing manifest instead of overwriting it.
  - Uses macOS system certificates if Python's own are not installed.

NOTE ON RESTRICTED MATERIAL
  This script performs no authentication and cannot retrieve CEII, privileged, or
  otherwise non-public documents. Access to those requires a determination from
  FERC and the applicable non-disclosure agreement. restricted_manifest.csv is
  provided to identify such documents, not to circumvent the restriction.

Standard library only. Python 3.9+. Building the package also needs ferc_text.py,
ferc_package.py and pypdf.zip beside this file.

USAGE
    python ferc_fetch.py                                  # guided mode
    python ferc_fetch.py P-2523 P-2689 --out ./Oconto     # whole record (as v1.2)
    python ferc_fetch.py P-2523 --doc-type order --text license --preview
    python ferc_fetch.py CP25-10 --from 2025-01-01 --text "environmental assessment"
    python ferc_fetch.py --library Hydro --affiliation "Consumers Energy" \\
        --from 2024-01-01 --metadata-only --out ./CE_2024
    python ferc_fetch.py --text "fish passage" --full-text --library Hydro \\
        --from 2023-01-01 --max-results 200
    python ferc_fetch.py --accession 20260916-3016 20250828-5123
    python ferc_fetch.py P-2523 --out ./Oconto                 # download + package
    python ferc_fetch.py P-2523 --save files --out ./Oconto    # stage 1 only
    python ferc_fetch.py --rebuild --out ./Oconto              # stage 2 only
    python ferc_fetch.py --list-doc-types environmental
Rerunning the same command resumes: files already on disk are skipped.
"""

import argparse
import csv
import datetime as dt
import html
import json
import os
import re
import ssl
import sys
import time
import urllib.error
import urllib.request

try:
    import ferc_text  # text extraction + bundling; keep ferc_text.py beside this file
except ImportError:
    ferc_text = None

API = "https://elibrary.ferc.gov/eLibrarywebapi/api"
ELIB_LINK = "https://elibrary.ferc.gov/eLibrary/filelist?accession_number={}"
PER_PAGE = 100
UA = "Merjent-FERC-Research/2.0 (project document retrieval)"

KEEP_EXT = (".pdf", ".doc", ".docx", ".xls", ".xlsx", ".txt", ".rtf",
            ".zip", ".tif", ".tiff", ".csv", ".ppt", ".pptx")

# HTTP codes meaning "this content exists but you are not entitled to it".
RESTRICTED_CODES = (401, 403)

# Option values as the API expects them (verified against the live service).
CATEGORIES = {"issuance": "Issuance", "submittal": "Submittal"}
LIBRARIES = {"hydro": "Hydro", "gas": "Gas", "electric": "Electric", "oil": "Oil",
             "rulemaking": "Rulemaking", "general": "General"}
DATE_TYPES = {"filed": "filed_date", "document": "issued_date", "posted": "posted_date"}
ROLES = {"any": "", "author": "AUTHOR", "recipient": "RECIPIENT", "agent": "AGENT"}
SORTS = {"newest": "reversefileddate", "oldest": "fileddate"}
AVAIL_NAMES = {"P": "Public", "C": "CEII", "N": "Privileged", "S": "Protected"}
SAVES = {"manifest": "manifest only (no downloads)",
         "files": "manifest + original files",
         "package": "manifest + original files + Copilot package"}
OLD_SAVES = {"text": "package", "all": "package"}  # names used before v2.3

# Every query setting, with its default. This is also the saved-query format.
QUERY_DEFAULTS = {
    "dockets": [],
    "text": "",
    "full_text": False,
    "accessions": [],
    "date_from": None,
    "date_to": None,
    "date_type": "filed",
    "categories": [],
    "libraries": [],
    "doc_types": [],
    "affiliations": [],
    "role": "any",
    "availability": "public",
    "sort": "newest",
    "max_results": None,
    "file_types": [],
    "max_mb": 250.0,
    "save": "package",
    "package_format": "docx",
    "notebook_bundles": False,
    "bundle_chars": 3_000_000,
    "bundle_name": None,
    "out": "ferc_docs",
}


# ----------------------------------------------------------------------------
# HTTP
# ----------------------------------------------------------------------------

class FatalHTTP(Exception):
    """A client error that will not succeed on retry."""


class RestrictedContent(Exception):
    """Server declined the document (401/403). Expected for CEII/privileged material."""

    def __init__(self, code, path, detail=""):
        self.code = code
        self.detail = detail
        super().__init__(f"HTTP {code} restricted on {path}")


def _ssl_context():
    """Default TLS context, falling back to certifi or the macOS system bundle.

    python.org installs on macOS ship without CA certificates until the user runs
    'Install Certificates.command', which makes every HTTPS request fail.
    """
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        pass
    ctx = ssl.create_default_context()
    if ctx.cert_store_stats().get("x509_ca", 0) == 0:
        for bundle in ("/etc/ssl/cert.pem", "/usr/local/etc/openssl/cert.pem"):
            if os.path.exists(bundle):
                ctx.load_verify_locations(bundle)
                break
    return ctx


SSL_CTX = _ssl_context()


def _request(path, body=None, raw=False, timeout=180, retries=5):
    """GET (body None) or POST JSON to the eLibrary web API with status-aware retry/backoff."""
    data = json.dumps(body).encode("utf-8") if body is not None else None
    last = None

    for attempt in range(retries):
        req = urllib.request.Request(
            f"{API}/{path}",
            data=data,
            headers={
                "Content-Type": "application/json",
                "Accept": "application/json, */*",
                "User-Agent": UA,
            },
            method="POST" if data is not None else "GET",
        )
        try:
            with urllib.request.urlopen(req, timeout=timeout, context=SSL_CTX) as r:
                return r.read() if raw else json.load(r)

        except urllib.error.HTTPError as e:
            code = e.code
            try:
                detail = e.read(400).decode("utf-8", "replace").strip()
            except Exception:
                detail = ""

            # Not entitled to this document. Expected; surface distinctly.
            if code in RESTRICTED_CODES:
                raise RestrictedContent(code, path, detail)

            last = FatalHTTP(f"HTTP {code} on {path}: {detail[:200]}")

            # Throttled: honor Retry-After, else escalating cool-down.
            if code in (429, 503):
                hdr = e.headers.get("Retry-After") if e.headers else None
                wait = int(hdr) if (hdr or "").isdigit() else min(60 * (attempt + 1), 300)
                print(f"    ! HTTP {code} throttled - sleeping {wait}s", file=sys.stderr)
                time.sleep(wait)
                continue

            # Other 4xx: retrying cannot help.
            if 400 <= code < 500:
                print(f"    ! HTTP {code} on {path} (not retrying) {detail[:120]}", file=sys.stderr)
                raise last

            wait = 2 ** attempt
            print(f"    ! HTTP {code} on {path}, retry in {wait}s", file=sys.stderr)
            time.sleep(wait)

        except urllib.error.URLError as e:
            if isinstance(e.reason, ssl.SSLCertVerificationError):
                raise SystemExit(
                    "SSL certificate check failed. On a Mac with python.org Python, run\n"
                    "  /Applications/Python*/Install\\ Certificates.command\n"
                    "or: pip install certifi") from e
            last = e
            wait = 2 ** attempt
            print(f"    ! {type(e).__name__}: {e} - retry in {wait}s", file=sys.stderr)
            time.sleep(wait)

        except (TimeoutError, OSError) as e:
            last = e
            wait = 2 ** attempt
            print(f"    ! {type(e).__name__}: {e} - retry in {wait}s", file=sys.stderr)
            time.sleep(wait)

    raise last


def download_file(file_id):
    return _request("File/DownloadP8File", {"fileidLst": [file_id]}, raw=True)


_CLASS_TYPES = None


def class_types():
    """eLibrary's list of document class/type pairs (cached for the run)."""
    global _CLASS_TYPES
    if _CLASS_TYPES is None:
        _CLASS_TYPES = _request("Search/GetClassTypes")
    return _CLASS_TYPES


# ----------------------------------------------------------------------------
# Query building
# ----------------------------------------------------------------------------

def split_docket(s):
    """'P-2523-007' -> ('P-2523', ['007']); 'CP25-10' -> ('CP25-10', [])."""
    s = s.strip().upper()
    m = re.match(r"^([A-Z]+\d*-\d+)-(\d{3})$", s)
    return (m.group(1), [m.group(2)]) if m else (s, [])


def resolve_doc_types(terms, categories=()):
    """Turn user terms into eLibrary {documentClass, documentType} pairs.

    Each term may be 'Class|Type' (exact), a class name (all its types), or any
    word or phrase, matched against class and type names. With categories given,
    only types in those categories are kept.
    """
    cats = {CATEGORIES[c.lower()] for c in categories}
    pool = [x for x in class_types() if not cats or x["Category"] in cats]
    pairs, seen = [], set()
    for term in terms:
        t = term.strip().lower()
        if "|" in t:
            c, _, ty = (x.strip() for x in t.partition("|"))
            hits = [x for x in pool
                    if x["Class"].lower() == c and x["Type"].lower() == ty]
        else:
            hits = [x for x in pool if x["Class"].lower() == t] or \
                   [x for x in pool
                    if t in x["Class"].lower() or t in x["Type"].lower()]
        if not hits:
            within = f" among {'/'.join(sorted(cats))} types" if cats else ""
            raise SystemExit(f"No eLibrary document type matches '{term}'{within}. "
                             f"Try: --list-doc-types {term.split()[0]}")
        for x in hits:
            key = (x["Class"], x["Type"])
            if key not in seen:
                seen.add(key)
                pairs.append({"documentClass": x["Class"], "documentType": x["Type"]})
    return pairs


def search_body(q, page, docket=None, accession=None, doc_pairs=None):
    date_searches = []
    if q["date_from"] or q["date_to"]:
        date_searches.append({
            "dateType": DATE_TYPES[q["date_type"]],
            "startDate": q["date_from"] or "1904-01-01",
            "endDate": q["date_to"] or dt.date.today().isoformat(),
        })

    dockets = []
    if docket:
        base, subs = split_docket(docket)
        dockets.append({"docketNumber": base, "subDocketNumbers": subs})

    return {
        "searchText": q["text"] or "*",
        "searchFullText": bool(q["full_text"] and q["text"]),
        "searchDescription": True,
        "dateSearches": date_searches,
        "allDates": not date_searches,
        "availability": ["P"] if q["availability"] == "public" else None,
        "affiliations": [{"afType": ROLES[q["role"]], "affiliation": a.upper(),
                          "lastName": None, "firstInitial": None, "middleInitial": None}
                         for a in q["affiliations"]],
        "categories": [CATEGORIES[c.lower()] for c in q["categories"]],
        "libraries": [LIBRARIES[l.lower()] for l in q["libraries"]],
        "accessionNumber": accession,
        "eFiling": False,
        "docketSearches": dockets,
        "resultsPerPage": PER_PAGE,
        "curPage": page,
        "classTypes": doc_pairs or [],
        "sortBy": SORTS[q["sort"]],
        "groupBy": "NONE",
        "idolResultID": "",
    }


def search_plan(q):
    """One search per accession, else one per docket, else a single open search."""
    if q["accessions"]:
        return [("accession " + a, None, a) for a in q["accessions"]]
    if q["dockets"]:
        return [(d.upper(), d, None) for d in q["dockets"]]
    return [("search", None, None)]


def iter_hits(q, docket, accession, doc_pairs, pause, limit=None, quiet=False):
    """Yield search hits page by page. Sets iter_hits.total for the caller."""
    page, total, n = 1, None, 0
    while True:
        data = _request("Search/AdvancedSearch",
                        search_body(q, page, docket, accession, doc_pairs))
        if total is None:
            total = int(g(data, "totalHits", default=0) or 0)
            iter_hits.total = total
            if not quiet:
                shown = f" (taking first {limit})" if limit and limit < total else ""
                print(f"  {total:,} filings matched{shown}")
        hits = g(data, "searchHits", default=[]) or []
        if not hits:
            return
        for h in hits:
            if limit and n >= limit:
                return
            n += 1
            yield h
        if page * PER_PAGE >= total:
            return
        page += 1
        time.sleep(pause)


iter_hits.total = 0


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------

def safe(name, maxlen=120):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name or ""))
    name = re.sub(r"\s+", " ", name).strip(" .")
    return name[:maxlen] or "unnamed"


def g(d, *keys, default=None):
    """First present key. eLibrary misspells accessionNumber as 'acesssionNumber'."""
    for k in keys:
        if isinstance(d, dict) and d.get(k) not in (None, ""):
            return d[k]
    return default


def iso(mdy):
    """'09/16/2026' -> '2026-09-16' so the manifest sorts correctly."""
    m = re.match(r"^(\d{2})/(\d{2})/(\d{4})", mdy or "")
    return f"{m.group(3)}-{m.group(1)}-{m.group(2)}" if m else (mdy or "")[:10]


def parties(h, af_type):
    out = []
    for a in g(h, "affiliations", default=[]) or []:
        if a.get("afType") == af_type:
            person = " ".join(x for x in (a.get("firstInitial"), a.get("lastName")) if x)
            org = html.unescape(a.get("affiliation") or "").strip()
            out.append(f"{org} ({person})" if org and person else org or person)
    return "; ".join(dict.fromkeys(out))


def base_meta(h):
    acc = g(h, "acesssionNumber", "accessionNumber", default="unknown")
    cts = g(h, "classTypes", default=[]) or []
    return {
        "accession_number": acc,
        "elibrary_url": ELIB_LINK.format(acc),
        "filed_date": iso(g(h, "filedDate")),
        "document_date": iso(g(h, "issuedDate", "documentDate")),
        "posted_date": iso(g(h, "postedDate")),
        "category": g(h, "category", default=""),
        "document_class": "; ".join(dict.fromkeys(c.get("documentClass", "") for c in cts)),
        "document_type": "; ".join(dict.fromkeys(c.get("documentType", "") for c in cts)),
        "library": "; ".join(g(h, "libraries", default=[]) or []),
        "availability": AVAIL_NAMES.get(g(h, "availCode", default=""), g(h, "availCode", default="")),
        "dockets": "; ".join(g(h, "docketNumbers", default=[]) or []),
        "author": parties(h, "AUTHOR"),
        "recipient": parties(h, "RECIPIENT"),
        "description": html.unescape(g(h, "description", default="") or "").strip(),
    }


def folder_for(h, requested_docket):
    """Requested docket if given, else the filing's first docket (sub-docket dropped)."""
    if requested_docket:
        return safe(split_docket(requested_docket)[0])
    ds = g(h, "docketNumbers", default=[]) or []
    return safe(split_docket(ds[0])[0]) if ds else "_no_docket"


def parse_cooldown(s):
    """'25:60' -> (25, 60): sleep 60s after every 25 downloads."""
    if not s:
        return (0, 0)
    n, _, secs = s.partition(":")
    return (int(n), int(secs or 60))


def describe(q):
    """One-line plain-English summary of a query."""
    parts = []
    if q["accessions"]:
        parts.append("accession " + ", ".join(q["accessions"]))
    if q["dockets"]:
        parts.append("docket " + ", ".join(q["dockets"]))
    if q["text"]:
        parts.append(f'"{q["text"]}" in ' + ("full text" if q["full_text"] else "description"))
    if q["date_from"] or q["date_to"]:
        parts.append(f'{q["date_type"]} {q["date_from"] or "start"} to {q["date_to"] or "today"}')
    if q["categories"]:
        parts.append("/".join(q["categories"]))
    if q["libraries"]:
        parts.append("library " + "/".join(q["libraries"]))
    if q["doc_types"]:
        parts.append("type " + ", ".join(q["doc_types"]))
    if q["affiliations"]:
        role = "" if q["role"] == "any" else q["role"] + " "
        parts.append(role + "affiliation " + ", ".join(q["affiliations"]))
    parts.append("public only" if q["availability"] == "public" else "all availability")
    if q["max_results"]:
        parts.append(f'max {q["max_results"]}')
    parts.append("save " + SAVES[q["save"]])
    return "; ".join(parts)


def validate(q):
    for key in ("date_from", "date_to"):
        if q[key] and not re.match(r"^\d{4}-\d{2}-\d{2}$", q[key]):
            raise SystemExit(f"Dates must be YYYY-MM-DD (got {q[key]!r}).")
    checks = (("categories", CATEGORIES), ("libraries", LIBRARIES))
    for key, allowed in checks:
        for v in q[key]:
            if v.lower() not in allowed:
                raise SystemExit(f"Unknown {key[:-1]} '{v}'. Choose from: "
                                 + ", ".join(allowed.values()))
    for key, allowed in (("date_type", DATE_TYPES), ("role", ROLES), ("sort", SORTS),
                         ("availability", {"public": 1, "all": 1}), ("save", SAVES)):
        if q[key] not in allowed:
            raise SystemExit(f"Unknown {key} '{q[key]}'. Choose from: " + ", ".join(allowed))


# ----------------------------------------------------------------------------
# Manifests (merged across reruns)
# ----------------------------------------------------------------------------

COLS = ["accession_number", "elibrary_url", "filed_date", "document_date", "posted_date",
        "category", "document_class", "document_type", "library", "availability",
        "dockets", "author", "recipient", "description", "file_name", "file_type",
        "listed_bytes", "file_id", "relative_path", "status", "http_code", "bytes",
        "text_status", "text_note", "text_chars", "text_path", "bundle_file"]

RCOLS = ["accession_number", "elibrary_url", "filed_date", "category", "document_class",
         "document_type", "dockets", "author", "description", "file_name", "availability",
         "http_code"]


def _row_key(r):
    return (r.get("accession_number", ""), r.get("file_id") or r.get("file_name", ""))


def merge_rows(out, rows):
    """This run's rows on top of the existing manifest's, oldest filing first."""
    merged = {}
    mpath = os.path.join(out, "manifest.csv")
    if os.path.exists(mpath):
        with open(mpath, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                r.setdefault("dockets", r.get("docket", ""))  # v1.2 manifests
                merged[_row_key(r)] = r
    for r in rows:
        old = merged.get(_row_key(r))
        if old and not r.get("text_status") and old.get("text_status"):
            # Text cached by an earlier run: keep what that run learned about it.
            for k in ("text_status", "text_note", "text_chars", "text_path"):
                r[k] = old.get(k, "")
        merged[_row_key(r)] = r
    return sorted(merged.values(), key=lambda r: (r.get("filed_date", ""),
                                                  r.get("accession_number", "")))


def write_manifests(out, final, raw_meta, q):
    mpath = os.path.join(out, "manifest.csv")
    with open(mpath, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=COLS, extrasaction="ignore")
        w.writeheader()
        w.writerows(final)

    restricted = [r for r in final if r.get("status") == "restricted"]
    rpath = os.path.join(out, "restricted_manifest.csv")
    if restricted:
        with open(rpath, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=RCOLS, extrasaction="ignore")
            w.writeheader()
            w.writerows(restricted)

    jpath = os.path.join(out, "manifest.json")
    records = {}
    if os.path.exists(jpath):
        try:
            with open(jpath, encoding="utf-8") as f:
                for rec in json.load(f).get("records", []):
                    records[g(rec, "acesssionNumber", "accessionNumber", default=id(rec))] = rec
        except (ValueError, OSError):
            pass
    for rec in raw_meta:
        records[g(rec, "acesssionNumber", "accessionNumber", default=id(rec))] = rec
    with open(jpath, "w", encoding="utf-8") as f:
        json.dump({"generated": dt.datetime.now().isoformat(timespec="seconds"),
                   "query": q, "records": list(records.values())}, f, indent=2)

    with open(os.path.join(out, "query.json"), "w", encoding="utf-8") as f:
        json.dump(q, f, indent=2)

    return mpath, (rpath if restricted else None), len(final)


# ----------------------------------------------------------------------------
# Guided mode
# ----------------------------------------------------------------------------

def ask(prompt, default=""):
    shown = f" [{default}]" if default not in ("", None) else ""
    try:
        v = input(f"{prompt}{shown}: ").strip()
    except EOFError:
        v = ""
    return v or (default if default is not None else "")


def ask_list(prompt):
    v = ask(prompt)
    return [x.strip() for x in re.split(r"[,;]", v) if x.strip()]


def yes(prompt, default=False):
    v = ask(prompt + (" (Y/n)" if default else " (y/N)")).lower()
    return default if not v else v.startswith("y")


def pick(prompt, options, default):
    print(f"{prompt}: " + "  ".join(f"[{i}] {o}" for i, o in enumerate(options, 1)))
    v = ask("  choose", str(options.index(default) + 1))
    try:
        return options[int(v) - 1]
    except (ValueError, IndexError):
        return default


def guided(q):
    print("\nFERC eLibrary fetch - guided mode. Press Enter to skip any question.\n")
    q["dockets"] = ask_list("Docket number(s), e.g. P-2523, CP25-10 (blank = any docket)")
    if not q["dockets"]:
        q["accessions"] = ask_list("Specific accession number(s) (blank = none)")
    if not q["accessions"]:
        q["text"] = ask("Keywords (blank = everything)")
        if q["text"]:
            q["full_text"] = yes("Search inside document text, not just descriptions?")
        q["date_from"] = ask("From date YYYY-MM-DD (blank = earliest)") or None
        q["date_to"] = ask("To date YYYY-MM-DD (blank = today)") or None
        if q["date_from"] or q["date_to"]:
            q["date_type"] = pick("Which date", list(DATE_TYPES), "filed")
        cat = pick("Category", ["both", "Issuance", "Submittal"], "both")
        q["categories"] = [] if cat == "both" else [cat]
        q["libraries"] = ask_list("Library: " + ", ".join(LIBRARIES.values()) + " (blank = all)")
        while True:
            terms = ask_list("Document type keywords, e.g. license, environmental assessment, "
                             "comment ('?' to browse, blank = all)")
            if terms[:1] == ["?"]:
                list_doc_types(ask("  filter list by word (blank = show all)"))
                continue
            q["doc_types"] = terms
            break
        q["affiliations"] = ask_list("Filed by / affiliation, e.g. Consumers Energy (blank = anyone)")
        if q["affiliations"]:
            q["role"] = pick("Their role", list(ROLES), "any")
    q["availability"] = "public" if yes("Public documents only?", True) else "all"
    mr = ask("Maximum number of filings (blank = no limit)")
    q["max_results"] = int(mr) if mr.isdigit() else None
    print("What to save:")
    for i, (k, v) in enumerate(SAVES.items(), 1):
        print(f"  [{i}] {v}" + ("  <- for Copilot agents" if k == "package" else ""))
    v = ask("  choose", "3")
    q["save"] = list(SAVES)[int(v) - 1] if v in ("1", "2", "3") else "package"
    slug = safe("_".join(q["dockets"] or q["accessions"][:1] or [q["text"] or "search"]), 40)
    q["out"] = ask("Output folder", os.path.join(".", f"ferc_{slug.replace(' ', '_')}"))
    return q


def list_doc_types(word=""):
    word = (word or "").lower()
    rows = sorted({(x["Category"], x["Class"], x["Type"]) for x in class_types()
                   if not word or word in x["Class"].lower() or word in x["Type"].lower()})
    w = max([len(r[1]) for r in rows] + [5])
    print(f"\n  {'CATEGORY':<10} {'CLASS':<{w}}  TYPE")
    for cat, cls, typ in rows:
        print(f"  {cat:<10} {cls:<{w}}  {typ}")
    print(f"\n  {len(rows)} types. Use any word from CLASS or TYPE with --doc-type, "
          f"or exact 'Class|Type'.\n")


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------

def build_parser():
    p = argparse.ArgumentParser(
        description="Pull filings from FERC eLibrary. Run with no arguments for guided mode.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="All filters combine (AND). Lists accept several values: --library Hydro Gas")
    S = argparse.SUPPRESS  # unset flags must not override a --query file

    w = p.add_argument_group("what")
    w.add_argument("dockets", nargs="*", default=S,
                   help="Docket numbers, e.g. P-2523 CP25-10. Add -NNN for one sub-docket (P-2523-007)")
    w.add_argument("--text", default=S, help="Keywords to search for")
    w.add_argument("--full-text", action="store_true", default=S,
                   help="Search --text inside document contents, not just descriptions")
    w.add_argument("--accession", dest="accessions", nargs="+", default=S,
                   help="Specific accession numbers, e.g. 20260916-3016")

    f = p.add_argument_group("filters")
    f.add_argument("--from", dest="date_from", default=S, help="Start date YYYY-MM-DD")
    f.add_argument("--to", dest="date_to", default=S, help="End date YYYY-MM-DD")
    f.add_argument("--date-type", choices=list(DATE_TYPES), default=S,
                   help="Which date --from/--to apply to (default filed)")
    f.add_argument("--category", dest="categories", nargs="+", default=S,
                   metavar="{Issuance,Submittal}")
    f.add_argument("--library", dest="libraries", nargs="+", default=S,
                   metavar="{" + ",".join(LIBRARIES.values()) + "}")
    f.add_argument("--doc-type", dest="doc_types", nargs="+", default=S,
                   help="Document class/type words, e.g. 'environmental assessment' license. "
                        "See --list-doc-types")
    f.add_argument("--affiliation", dest="affiliations", nargs="+", default=S,
                   help="Organization name (partial OK), e.g. 'Consumers Energy'")
    f.add_argument("--role", choices=list(ROLES), default=S,
                   help="Affiliation role to match (default any). Filers are 'author'")
    f.add_argument("--availability", choices=["public", "all"], default=S,
                   help="'all' also lists CEII/privileged filings (never downloaded). Default public")
    f.add_argument("--include-nonpublic", dest="availability", action="store_const",
                   const="all", default=S, help="Same as --availability all (v1.2 flag)")

    o = p.add_argument_group("output")
    o.add_argument("--out", default=S, help="Output directory (default ferc_docs)")
    o.add_argument("--max-results", type=int, default=S, help="Stop after this many filings per search")
    o.add_argument("--sort", choices=list(SORTS), default=S, help="Order filings are taken in (default newest)")
    o.add_argument("--types", dest="file_types", nargs="+", default=S,
                   help="Only these file extensions, e.g. --types pdf docx")
    o.add_argument("--max-mb", type=float, default=S, help="Skip files larger than this (default 250)")
    o.add_argument("--save", choices=list(SAVES) + list(OLD_SAVES), default=S,
                   metavar="{manifest,files,package}",
                   help="manifest = spreadsheet only; files = + original files; "
                        "package = + Copilot knowledge package built from them (default)")
    o.add_argument("--metadata-only", "--dry-run", dest="save", action="store_const",
                   const="manifest", default=S, help="Same as --save manifest")
    o.add_argument("--name", dest="bundle_name", default=S,
                   help="Project name used in the Copilot package (default: output folder name)")
    o.add_argument("--format", dest="package_format", choices=["docx", "md"], default=S,
                   help="Copilot package file format (default docx)")
    o.add_argument("--notebook-bundles", action="store_true", default=S,
                   help="Also write large combined text files for Copilot notebooks (300-file limit)")
    o.add_argument("--bundle-chars", type=int, default=S,
                   help="Maximum characters per notebook bundle (default 3,000,000)")
    o.add_argument("--rebuild", "--package-only", "--bundle-only", dest="rebuild",
                   action="store_true",
                   help="Stage 2 only: build the Copilot package from an existing --out folder, "
                        "no download. Works on folders from Kyle's original script too")
    o.add_argument("--refresh-text", action="store_true",
                   help="Re-read text from the original files instead of reusing cached text "
                        "(after an update to text extraction)")

    r = p.add_argument_group("run")
    r.add_argument("-i", "--interactive", action="store_true", help="Guided mode")
    r.add_argument("--preview", nargs="?", const=15, type=int, metavar="N",
                   help="Show the match count and first N filings (default 15), then stop")
    r.add_argument("--query", metavar="FILE", help="Load settings from a saved query (.json)")
    r.add_argument("--save-query", metavar="FILE", help="Save these settings to a query file")
    r.add_argument("--list-doc-types", nargs="?", const="", metavar="WORD",
                   help="List eLibrary document classes/types (optionally filtered) and exit")
    r.add_argument("-y", "--yes", action="store_true", help="Do not ask before large broad searches")
    r.add_argument("--pause", type=float, default=2.0, help="Seconds between requests (default 2.0)")
    r.add_argument("--cooldown", default="25:60",
                   help="'N:SECS' - pause SECS after every N downloads. '0:0' disables.")
    r.add_argument("--max-consecutive-failures", type=int, default=8,
                   help="Abort a search after this many real failures in a row "
                        "(restricted 401/403 responses do not count)")
    return p


def preview(q, doc_pairs, n):
    print(f"\nQuery: {describe(q)}")
    for label, docket, acc in search_plan(q):
        print(f"\n=== {label} ===")
        cap = q["max_results"]
        rows = list(iter_hits(q, docket, acc, doc_pairs, 0, limit=min(n, cap or n), quiet=True))
        total = iter_hits.total
        print(f"  {total:,} filings matched" + (f" (run takes first {cap})" if cap and cap < total else ""))
        for h in rows:
            m = base_meta(h)
            print(f"  {m['filed_date']}  {m['accession_number']:<15} {m['category'][:9]:<9} "
                  f"{m['document_type'][:28]:<28} {m['description'][:70]}")
        if cap:
            total = min(total, cap)
        if total > len(rows):
            print(f"  ... and {total - len(rows):,} more")
        time.sleep(1)


def main():
    p = build_parser()
    args = vars(p.parse_args())

    if args["list_doc_types"] is not None:
        list_doc_types(args["list_doc_types"])
        return

    q = dict(QUERY_DEFAULTS)
    if args.get("query"):
        with open(args["query"], encoding="utf-8") as f:
            saved = json.load(f)
        if saved.get("metadata_only"):
            saved["save"] = "manifest"
        q.update({k: v for k, v in saved.items() if k in QUERY_DEFAULTS})
    q.update({k: v for k, v in args.items() if k in QUERY_DEFAULTS})
    q["save"] = OLD_SAVES.get(q["save"], q["save"])

    if args["rebuild"]:
        require_text_support()
        ferc_text.rebuild_folder(q["out"], q["bundle_name"], q["package_format"],
                                 q["notebook_bundles"], q["bundle_chars"],
                                 refresh=args["refresh_text"])
        return

    query_given = any(k in args for k in QUERY_DEFAULTS) or args.get("query")
    if args["interactive"] or not query_given:
        q = guided(q)
        if not args["preview"]:
            args["preview"] = 10
            args["_confirm_after_preview"] = True

    q["dockets"] = [d.strip().upper() for d in q["dockets"]]
    validate(q)
    if q["save"] == "package":
        require_text_support()
    doc_pairs = resolve_doc_types(q["doc_types"], q["categories"]) if q["doc_types"] else []
    if doc_pairs:
        print(f"Document types matched ({len(doc_pairs)}): "
              + "; ".join(f"{d['documentClass']} | {d['documentType']}" for d in doc_pairs[:12])
              + (" ..." if len(doc_pairs) > 12 else ""))

    if args["save_query"]:
        with open(args["save_query"], "w", encoding="utf-8") as f:
            json.dump(q, f, indent=2)
        print(f"Saved query to {args['save_query']}")

    if args["preview"]:
        preview(q, doc_pairs, args["preview"])
        if not args.get("_confirm_after_preview"):
            print("\n(Preview only. Remove --preview to run it.)")
            return
        print()
        if not yes(f"Run this and save to {q['out']}?", True):
            if yes("Save this query to reuse later?"):
                path = ask("File name", "ferc_query.json")
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(q, f, indent=2)
                print(f"Saved. Run it with: python {os.path.basename(sys.argv[0])} --query {path}")
            return

    run(q, doc_pairs, args)


def require_text_support():
    """Fail before any downloading if text bundling cannot work."""
    if ferc_text is None or ferc_text.ferc_package is None:
        raise SystemExit("The Copilot package needs ferc_text.py and ferc_package.py "
                         "in the same folder as ferc_fetch.py.")
    if not ferc_text.HAVE_PYPDF:
        raise SystemExit(ferc_text.PYPDF_HINT)


def run(q, doc_pairs, args):
    exts = tuple("." + e.lower().lstrip(".") for e in q["file_types"]) if q["file_types"] else KEEP_EXT
    os.makedirs(q["out"], exist_ok=True)
    print(f"\nQuery: {describe(q)}\nOutput: {os.path.abspath(q['out'])}")
    rows, raw_meta, errors = [], [], []
    counts = {"downloaded": 0, "skipped": 0, "restricted": 0, "failed": 0}
    seen_acc = set()

    interrupted = False
    try:
        for label, docket, acc_q in search_plan(q):
            if not fetch_search(q, label, docket, acc_q, doc_pairs, args, exts, rows,
                                raw_meta, errors, seen_acc, counts):
                return
    except KeyboardInterrupt:
        interrupted = True
        print("\nStopped. Writing manifest for what was fetched; rerun to resume.",
              file=sys.stderr)

    finish(q, rows, raw_meta, errors, seen_acc, counts)
    if interrupted:
        sys.exit(130)


def fetch_search(q, label, docket, acc_q, doc_pairs, args, exts, rows, raw_meta, errors,
                 seen_acc, counts):
    """Run one search and download its files. Returns False if the user cancelled."""
    cool_every, cool_secs = parse_cooldown(args["cooldown"])
    pause = args["pause"]
    out = q["out"]
    broad = not q["dockets"] and not q["accessions"]
    print(f"\n=== {label} ===")
    try:
        gen = iter_hits(q, docket, acc_q, doc_pairs, pause, limit=q["max_results"])
        first = next(gen, None)
        total = iter_hits.total
        if broad and first and total > 500 and not args["yes"] and sys.stdin.isatty() \
                and not q["max_results"]:
            if not yes(f"  This matches {total:,} filings across all dockets. Continue?"):
                return False
        hits = ([first] if first else []) + list(gen)
        # eLibrary's search index occasionally drops a few hits from a result set.
        # A second pass catches them; it costs one request per 100 filings.
        if hits and not q["max_results"]:
            have = {g(h, "acesssionNumber", "accessionNumber") for h in hits}
            extra = [h for h in iter_hits(q, docket, acc_q, doc_pairs, pause, quiet=True)
                     if g(h, "acesssionNumber", "accessionNumber") not in have]
            if extra:
                print(f"  + {len(extra)} more found on a second search pass")
                hits += extra
    except Exception as e:
        errors.append(f"{label}: search failed: {e}")
        print(f"  ! search failed: {e}", file=sys.stderr)
        return True

    consecutive = 0
    aborted = False

    for h in hits:
        base = base_meta(h)
        acc = base["accession_number"]
        if acc in seen_acc:
            continue  # already handled under an earlier docket in this run
        seen_acc.add(acc)
        raw_meta.append(h)
        folder = folder_for(h, docket)
        ddir = os.path.join(out, folder)
        avail_code = g(h, "availCode", default="P")

        files = g(h, "transmittals", default=[]) or []
        if not files:
            rows.append(dict(base, status="no-files-listed"))
            continue

        for t in files:
            fid = g(t, "fileId", "fileID", "id")
            fname = g(t, "fileName", "name", default=f"{acc}.pdf")
            if not fname.lower().endswith(KEEP_EXT):
                fname += "." + (g(t, "fileType", default="pdf") or "pdf").lower()

            fn = safe(f"{acc}__{fname}")
            path = os.path.join(ddir, fn)
            listed = g(t, "fileSize", default="")
            row = dict(base, file_id=fid, file_name=fname,
                       file_type=g(t, "fileType", default=""), listed_bytes=listed,
                       relative_path=os.path.join(folder, fn))

            if not fname.lower().endswith(exts):
                row["status"] = "skipped-file-type"
                row["relative_path"] = ""
            elif avail_code != "P":
                # Known non-public: list it, do not ask the server for it.
                row["status"] = "restricted"
                row["relative_path"] = ""
                counts["restricted"] += 1
                print(f"    - restricted ({row['availability']}): {fname}")
            elif q["save"] == "manifest":
                if os.path.exists(path) and os.path.getsize(path) > 0:
                    row["status"] = "already-present"
                    row["bytes"] = os.path.getsize(path)
                else:
                    row["status"] = "metadata-only"
                    row["relative_path"] = ""
            elif os.path.exists(path) and os.path.getsize(path) > 0:
                row["status"] = "already-present"
                row["bytes"] = os.path.getsize(path)
                counts["skipped"] += 1
            elif listed and int(listed) > q["max_mb"] * 1024 * 1024:
                row["status"] = "too-large-skipped"
                row["relative_path"] = ""
                counts["skipped"] += 1
            elif not fid:
                row["status"] = "no-file-id"
                row["relative_path"] = ""
                counts["failed"] += 1
            else:
                try:
                    blob = download_file(fid)
                    if not blob or (len(blob) < 400 and blob.lstrip().startswith(b'{"Message"')):
                        row["status"] = "empty-or-restricted"
                        row["relative_path"] = ""
                        counts["failed"] += 1
                    else:
                        os.makedirs(ddir, exist_ok=True)
                        with open(path, "wb") as fh:
                            fh.write(blob)
                        row["status"] = "downloaded"
                        row["bytes"] = len(blob)
                        counts["downloaded"] += 1
                        consecutive = 0
                        print(f"    + {folder}/{fn} ({len(blob):,} bytes)")
                        if cool_every and counts["downloaded"] % cool_every == 0:
                            print(f"    ... cooldown {cool_secs}s after {counts['downloaded']} files")
                            time.sleep(cool_secs)
                    time.sleep(pause)

                except RestrictedContent as e:
                    # Expected for CEII / privileged material. Record and continue.
                    # Deliberately does NOT touch `consecutive`.
                    row["status"] = "restricted"
                    row["http_code"] = e.code
                    row["relative_path"] = ""
                    counts["restricted"] += 1
                    print(f"    - restricted (HTTP {e.code}): {fname}")
                    time.sleep(pause)

                except Exception as e:
                    row["status"] = f"error: {e}"
                    row["relative_path"] = ""
                    errors.append(f"{acc} / {fname}: {e}")
                    counts["failed"] += 1
                    consecutive += 1
                    if consecutive >= args["max_consecutive_failures"]:
                        msg = (f"  ! {consecutive} consecutive failures - stopping {label}. "
                               f"Last error: {e}")
                        print(msg, file=sys.stderr)
                        errors.append(msg)
                        rows.append(row)
                        aborted = True
                        break

            rows.append(row)

        if aborted:
            break
    return True


def finish(q, rows, raw_meta, errors, seen_acc, counts):
    out = q["out"]
    final = merge_rows(out, rows)
    ocr = None
    if q["save"] == "package":
        # Stage 2: build text and the Copilot package from the files on disk.
        print("\nBuilding Copilot package from the downloaded files:")
        ferc_text.prepare_text(out, final)
        ferc_text.mark_covered(final)
        ferc_text.package_and_bundle(
            out, final, q["bundle_name"] or os.path.basename(os.path.abspath(out)),
            q["package_format"], q["notebook_bundles"], q["bundle_chars"], q["dockets"])
        ocr = ferc_text.write_needs_ocr(out, final)
    mpath, rpath, n_rows = write_manifests(out, final, raw_meta, q)

    if errors:
        with open(os.path.join(out, "errors.log"), "a", encoding="utf-8") as f:
            f.write(f"--- {dt.datetime.now().isoformat(timespec='seconds')}\n")
            f.write("\n".join(errors) + "\n")

    # ---- summary ----
    print("\n" + "-" * 58)
    print(f"  filings    : {len(seen_acc)}")
    print(f"  downloaded : {counts['downloaded']}")
    print(f"  skipped    : {counts['skipped']}   (already on disk or too large)")
    print(f"  restricted : {counts['restricted']}   (not retrievable without a FERC determination)")
    print(f"  failed     : {counts['failed']}")
    if ocr:
        ferc_text.print_text_summary(final, out, *ocr, log=lambda m: print(m) if m.startswith("  ") else None)
    print("-" * 58)
    print(f"Manifest: {mpath}  ({n_rows} file rows, including any from earlier runs)")
    if ocr:
        print(f"Copilot package: {os.path.join(out, 'copilot')}  "
              f"(upload the Knowledge folder; paste Agent Instructions.txt)")
        if ocr[0]:
            print(f"Needs OCR ({ocr[1]} files): {ocr[0]}")
    if rpath:
        print(f"Restricted list: {rpath}")
    if errors:
        print(f"Errors: {os.path.join(out, 'errors.log')}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nStopped. Rerun the same command to resume.", file=sys.stderr)
        sys.exit(130)
