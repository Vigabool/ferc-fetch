# FERC eLibrary fetch

`ferc_fetch.py` downloads filings from [FERC eLibrary](https://elibrary.ferc.gov/eLibrary/search) to a local folder. It writes a manifest spreadsheet and, optionally, a knowledge package built for a Copilot agent. It builds on the eLibrary download script and text bundler Kyle Olcott wrote, which proved the approach on real projects.

The tool is four files that must stay together in one folder. There's nothing to install beyond Python:

- `ferc_fetch.py`: the tool
- `ferc_text.py`: reads text from PDF, Word, and other files
- `ferc_package.py`: builds the Copilot knowledge package
- `pypdf.zip`: the open-source [pypdf](https://pypi.org/project/pypdf/) PDF reader (BSD license), loaded directly from the zip. There's no need to unzip it, but it still works if you do.

## Getting started on your Windows laptop

You run the tool by typing commands into a terminal, a text window where you type a command and press Enter. You don't need any programming knowledge: each step below is one thing to click or type.



### 1. Put the tool in its own folder

1. Open **File Explorer** and go to **This PC > Local Disk (C:) > Users >** your user name.
2. Right-click an empty area, choose **New > Folder**, and name it `FERC`.
3. Download the tool from GitHub and copy its files into `FERC`:
   1. Go to https://github.com/Vigabool/ferc-fetch and sign in to GitHub. The repository is private, so you need a GitHub account that has been given access; ask Nick Vigdahl. (The tool will move to Merjent's GitHub later; this guide will be updated with the new address.)
   2. Click the green **Code** button, then **Download ZIP**.
   3. Open your **Downloads** folder, right-click `ferc-fetch-main.zip`, choose **Extract All**, and click **Extract**.
   4. Open the extracted `ferc-fetch-main` folder (there may be a second `ferc-fetch-main` folder inside it). Select the four files `ferc_fetch.py`, `ferc_text.py`, `ferc_package.py`, and `pypdf.zip`, and copy them into your `FERC` folder. Copy `pypdf.zip` as it is; don't extract it.
   5. To update later, repeat these steps and replace the four files. Your downloaded projects in `FERC` aren't affected.

Use this folder rather than your Desktop, Documents, or a OneDrive or SharePoint folder. Those usually sync to the cloud, and syncing thousands of downloaded files slows the run and your laptop. When the package is ready, copy only its `Knowledge` folder to OneDrive or SharePoint.



### 2. Open a terminal in that folder

1. Open the `FERC` folder in File Explorer.
2. Right-click an empty area of the folder and choose **Open in Terminal**.
   - If you don't see it, click the address bar at the top of the File Explorer window (where it shows the folder path), type `cmd`, and press Enter.

A window opens with a blinking cursor after a line like `PS C:\Users\jsmith\FERC>` or `C:\Users\jsmith\FERC>`. That line is the prompt. It shows which folder the terminal is working in, and it should end in `FERC`.



### 3. Check that Python is installed

Type this and press Enter:

```
python --version
```

- **You see `Python 3.9` or higher** (for example `Python 3.12.4`): you're ready.
- **You see "not recognized," or the Microsoft Store opens:** try `py --version` instead. If that works, type `py` wherever this guide says `python`. If neither works, Python isn't installed. Ask Nick Vigdahl, since installing it may need IT.



### 4. Run the tool

Type this and press Enter:

```
python ferc_fetch.py
```

This starts guided mode, which asks you questions one at a time. Type an answer and press Enter, or just press Enter to skip a question. It shows what matches and asks before downloading anything.

When it finishes, open your results folder (for example `Oconto`) inside `FERC` in File Explorer, or type `explorer Oconto` in the terminal.

### Terminal tips

- **Paste** a command with **Ctrl+V** or a right-click. Commands in this guide can be copied and pasted as they are.
- **Rerun or edit** a previous command by pressing the **Up arrow**, changing what you need, and pressing Enter.
- **Stop a run** with **Ctrl+C**. What's been downloaded is kept; run the same command again to pick up where it stopped.
- **Don't close the window** while a run is going, because that stops it.
- **Keep the laptop awake for long runs.** A whole project can take hours. Plug in, go to **Settings > System > Power & battery > Screen and sleep**, and set sleep when plugged in to **Never** for the run, or leave the lid open. Set it back afterward.
- **How commands are written.** Options start with two dashes (`--out`). Put anything with spaces in quotes: `--name "Bellows Falls"`. Each command is one line.

### If something goes wrong

| Message | What to do |
|---|---|
| `'python' is not recognized...` | See step 3. |
| `can't open file ... ferc_fetch.py` | The terminal isn't in the `FERC` folder. Close it and repeat step 2. |
| `The Copilot package needs ferc_text.py and ferc_package.py...` or `Reading PDFs needs pypdf.zip...` | One of the four files is missing from the `FERC` folder. Copy it back. |
| `HTTP 429` or `HTTP 503`, "throttled" | FERC is asking the tool to slow down. It waits and continues on its own. |
| A run stops partway (network drop, laptop slept) | Run the same command again. It skips everything already downloaded. |

## How it works: two stages

1. **Download.** Searches eLibrary and saves the original files plus a manifest. This is the slow part: hours for a large record, at a pace FERC's servers tolerate.
2. **Package.** Reads the text from the files on disk and builds the Copilot knowledge package. This takes minutes and needs no network.

By default one command runs both. You can also run them separately:

- `--save files` downloads only.
- `--rebuild --out <folder>` packages an existing folder without downloading anything.

Keeping the stages apart means improvements to text reading, OCR, or the package format reach a folder you already downloaded, with no re-download: run `--rebuild --refresh-text`.

## What to save

Choose with `--save`, or answer the question in guided mode:

| `--save` | You get | Use it for |
|---|---|---|
| `manifest` | The spreadsheet only | Seeing what's in a record before committing to a download |
| `files` | The spreadsheet + original files | Stage 1 only |
| `package` (default) | The spreadsheet + original files + Copilot knowledge package | Both stages |

## Easiest: guided mode

```
python ferc_fetch.py
```

It asks what you want (docket, keywords, dates, document types, who filed, and so on). Press Enter to skip any question. It then shows how many filings match and the first few, and asks before downloading anything. You can also save your answers to reuse later.

## Command examples

| I want... | Command |
|---|---|
| A project's whole public record, packaged for Copilot | `python ferc_fetch.py P-2523 --name "Oconto Falls" --out Oconto` |
| Several projects in one folder | `python ferc_fetch.py P-2523 P-2689 --out Oconto` |
| Only license orders | `python ferc_fetch.py P-2523 --doc-type order --text license` |
| Only FERC issuances since 2020 | `python ferc_fetch.py P-2523 --category Issuance --from 2020-01-01` |
| Environmental assessments on a pipeline docket | `python ferc_fetch.py CP25-10 --doc-type "environmental assessment"` |
| Everything a company filed this year, any docket | `python ferc_fetch.py --affiliation "Consumers Energy" --role author --from 2026-01-01` |
| Hydro filings that mention a topic anywhere in the text | `python ferc_fetch.py --text "fish passage" --full-text --library Hydro --from 2025-01-01` |
| Specific documents | `python ferc_fetch.py --accession 20220727-3025 19971125-0691` |
| Download only, package later | add `--save files` |
| Package a folder that's already downloaded | `python ferc_fetch.py --rebuild --out Oconto` |
| Just the spreadsheet, no downloads | add `--save manifest` |
| Check a query before running it | add `--preview` |
| Only PDFs, nothing over 50 MB | add `--types pdf --max-mb 50` |
| The most recent 100 matching filings | add `--max-results 100` |

All filters combine: each one you add narrows the results. Run `python ferc_fetch.py --help` for the full list.

### Finding document types

eLibrary groups documents into about 260 class/type pairs. To see the ones that match a word:

```
python ferc_fetch.py --list-doc-types environmental
```

`--doc-type` accepts any word or phrase from those lists. Watch the naming: license orders are filed as class "Order/Opinion", not under a "license" type. To get them, combine `--doc-type order` with `--text license`. The script prints which types it matched so you can check.

### Saving a query

```
python ferc_fetch.py P-2523 --doc-type order --text license --save-query license_orders.json --preview
python ferc_fetch.py --query license_orders.json --out Oconto
```

Every output folder also gets a `query.json` file recording the query that produced it.

## What you get

```
Oconto/
  P-2523/                     downloaded files, named <accession>__<file name>
  manifest.csv                one row per file: accession, eLibrary link, dates, category,
                              document class/type, author/filer, recipient, all dockets,
                              description, file name/size, download status
  restricted_manifest.csv     CEII/privileged filings that can't be downloaded (only
                              written with --availability all)
  manifest.json               full raw metadata from eLibrary
  query.json                  the query that produced this folder

  from the package stage:
  copilot/Knowledge/          copy this folder to OneDrive or SharePoint for the agent
  copilot/Agent Instructions.txt   paste into the agent's instructions
  text/                       each document's raw text (a cache; rebuilds are fast)
  needs_ocr.csv               scans with no readable text (only written if any)
```

## The Copilot knowledge package

Copilot never reads a whole record. It splits files into short passages, finds the few that best match each question, and answers from those. The package is built around that.

```
Knowledge/
  00 Project Overview         licensee, other active parties, date range, every FERC order, gaps
  01 Timeline                 every FERC issuance and application, by date, one line each
  02 Catalog/Catalog 1990s    every filing, one line each (busy decades are split by year)
  Documents/1998/1998-11-13 19981117-0254 Order on rehearing of ... .docx
```

- **One file per filing,** named by date, accession number, and description. Copilot's search weighs file names heavily, and its citations show them.
- **Passages that identify themselves.** A line like `[P-2523 · 19981117-0254 · 1998-11-13 · Order on rehearing… · p. 4]` repeats roughly every 1,500 characters, so any passage Copilot retrieves says which filing and page it came from.
- **Whole-record questions answered up front.** Copilot can't list every order by searching passages. The overview, timeline, and catalog cover every filing, including restricted and scanned ones, in a few short files.
- **Less noise.** Page numbers, repeated page headers and footers, and certificates of service are removed. When a file repeats an earlier one word for word (re-filed attachments, or a zip holding the same PDF), it's replaced with a pointer to the earlier copy.
- **File format.** Word (.docx) is the default because SharePoint and Copilot handle it everywhere. Use `--format md` for Markdown.

### Setting up an agent

1. Copy the `Knowledge` folder into a OneDrive folder (just for you) or a SharePoint site (for a project team).
2. Create an agent from that folder: use "Create an agent" in SharePoint, or the agent builder in the Copilot app with the OneDrive or SharePoint folder as its knowledge source.
3. Paste `Agent Instructions.txt` into the agent's instructions. The instructions tell it to use the overview and catalog first, to cite accession numbers, and never to assume a filing doesn't exist just because search didn't find it.
4. Share the agent with the project team.

When new filings come in, rerun the same command (it downloads only what's new), then replace the `Knowledge` folder in OneDrive or SharePoint.

### Text extraction

- **What it reads.** PDF, Word (both .docx and the older .doc), PowerPoint (.pptx), Excel (.xlsx), TXT, CSV, and RTF. Older WordPerfect, .xls, and .ppt files aren't read.
- **ZIP files.** Files inside .zip attachments are read too, including zips inside zips.
- **Scans.** Many older filings are scanned TIFs, but FERC often includes a .TXT version of the same filing. When it does, the script uses that text and leaves the scan off the OCR list. Scans with no text version are listed in `needs_ocr.csv`. They still appear in the overview, timeline, and catalog, marked as having no text yet.
- **Running OCR.** Run OCR on the files in `needs_ocr.csv` (Adobe Acrobat works). Then either save the text to the path in the `text_path` column, or replace the original with a searchable PDF of the same name. Then rebuild:

  ```
  python ferc_fetch.py --rebuild --out Oconto
  ```

- **After an update to text reading,** run `--rebuild --refresh-text` to re-read every original. Text you supplied from OCR is kept.
- **Existing folders.** `--rebuild` also works on folders made with Kyle's original download script. You can also run `python ferc_text.py <folder>` on any folder of downloaded files, with or without a manifest.
- **Copilot notebooks.** A notebook holds at most 300 files, so a large project's package won't fit. Add `--notebook-bundles` to also write a few large combined text files for notebooks.

Rerunning a command resumes where it stopped: files already on disk are skipped. A rerun with a different query into the same folder adds to the manifest instead of replacing it. For example, `--from` with last month's date picks up only new filings. Pressing Ctrl-C stops the run and still saves the manifest.

## Notes

- **Restricted material.** By default only public filings are searched. With `--availability all`, CEII and privileged filings appear in the manifest but are never downloaded. Access requires a FERC determination under 18 CFR 388.113.
- **Pace.** The script waits 2 seconds between requests and takes a 60-second break every 25 downloads, so a large project can take hours. Leave it running overnight.

## Fixed from Kyle's original script

- `--from`/`--to` had no effect: the API ignored the date format and returned every filing. Fixed.
- "Public only" had no effect either, so CEII filings came back and failed as 401 errors. Fixed.
- The manifest's class/type, library, and availability columns were always blank. They are now filled in, along with author/filer, recipient, and an eLibrary link.

## Backlog

Ideas raised while building this, not yet done. Roughly in priority order.

### Prove it works
- **Copilot scorecard.** Have someone who knows a project well write 15 to 20 questions with known answers, mixing detail ("What does Article 403 require?"), list ("Which agencies commented on the DEIS?"), and change-over-time ("How did the fish passage position change?") questions. Score the current notebook approach against an agent built on the knowledge package.
- **Check in our Microsoft 365 tenant.** Confirm current SharePoint and Copilot limits (file size, file count, supported formats), how agents are created and shared, and that the generated Word files open cleanly in Word and are indexed as expected.

### Text coverage
- **OCR for scanned filings.** About 8% of files in the Bellows Falls test had no text layer, many of them agency and stakeholder letters. Options:
  1. Windows' built-in OCR, called from Python through PowerShell. Nothing to install; needs testing on a Merjent Windows laptop, and may be blocked where IT restricts PowerShell.
  2. Tesseract, a free OCR engine. IT must approve and install it.
  3. Microsoft's OCR in SharePoint (SharePoint Premium/Syntex, pay-as-you-go). No laptop software, but an admin must turn it on.
- **Encrypted PDFs.** Some signed letters are encrypted against editing (AES), and pypdf needs a compiled library to read them. A small pure-Python decryptor would cover them without an install.
- **Older formats.** WordPerfect (.wpd), .xls, and .ppt files are not read. FERC usually provides another format alongside WordPerfect files.
- **Near-duplicates.** Exact duplicates are folded, but "track changes" and "clean" versions of the same exhibit are both kept in full.

### Better answers
- **AI-written summaries.** A 150 to 300 word summary per filing and a written project brief, generated once while the package is built. This is the most promising fix for "how did positions evolve" questions. It needs an AI service such as the Claude API. There's a place for it in each document file.
- **A shared library and cross-project agent.** One SharePoint library with a folder per project and an agent per project, plus a central agent for questions across projects. Copilot only retrieves files a user has permission to see.

### Ease of use
- **No-install interface.** Options considered:
  - A web page plus a small cloud relay. FERC's server blocks the browser's pre-check before downloads, so a page can't call eLibrary directly.
  - A single Windows .exe. Needs IT approval as an unsigned program.
  - Revisit if the pilot shows demand.
- **Python availability.** Users need Python 3.9 or newer. Whether people can install it themselves, or need IT, decides how far the pilot can spread.

