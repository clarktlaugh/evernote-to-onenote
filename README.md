# enex2onenote

Migrate Evernote notebooks (`.enex` exports) into OneNote. Claude reads each
note, including the text inside attached PDFs and, when there is no text, the
scanned pages or photos. It then proposes a descriptive **title**, **tags**
and a short **summary**. You can review and edit those before anything is
uploaded.

```
Scan_0012.pdf  ->  Comcast invoice - March 2021          #bills #invoice #comcast
Untitled       ->  Buttermilk pancakes recipe - Grandma Ruth   #recipe #family
```

Run it on one Evernote notebook export at a time, or on a folder of them
(one subfolder per OneNote notebook).

## How it works

1. **`analyze`** parses the `.enex` file and asks Claude about each note. The
   results go to `migration-state/<folder>/<file name>/review.csv`, named
   after the folder the `.enex` file is in and its file name.
2. **Review (optional).** Open `review.csv` in Excel and edit `final_title`
   and `final_tags` (separated by `;`). Put `y` in `skip` for notes you don't
   want to migrate.
3. **`upload`** creates one OneNote page per note in the notebook and section
   you choose:
   - Formatting, images and attachments are kept. PDFs are attached, and their
     pages are also shown as printouts on the page.
   - Each page keeps the Evernote creation date.
   - The top of each page shows the tags as `#hashtags` (searchable in
     OneNote), the original Evernote title and the source URL.

Everything is resumable. Re-running a command skips notes that are already
done and retries notes that failed, so you can stop at any time. Tags that
Claude assigned in earlier notebooks are passed to later runs, which keeps
your tag vocabulary consistent across notebooks.

## Setup

Requires Python 3.10+.

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -e ".[com]"          # Windows: includes the OneNote desktop app support
pip install -e .                 # macOS/Linux: upload with --target graph
```

### 1. Choose how to reach Claude

| Backend | Use it when | Setup |
| --- | --- | --- |
| `claude-code` | You have a Claude **Pro/Max subscription** | Install [Claude Code](https://claude.com/claude-code) and run `claude` once to log in. Analysis counts against your plan's usage limits. |
| `api` | You have Claude **API** credits (billed separately from the subscription) | Create a key at [platform.claude.com](https://platform.claude.com) and `export ANTHROPIC_API_KEY=...` |
| `none` | Keep Evernote titles, no AI | — |

The default, `--backend auto`, uses `api` when `ANTHROPIC_API_KEY` is set and
`claude-code` otherwise.

> **A claude.ai subscription does not include API access.** If you only have a
> subscription, use `claude-code`.

With the `api` backend, the default model is `claude-opus-5-5` at low effort.
Change them with `--model` and `--effort`. Only the first 20,000 characters of
a note and its attachments are sent (`--max-chars`); when content is cut off,
the prompt tells Claude so.

### 2. Choose how pages get into OneNote

| Target | Runs on | Setup | Notes |
| --- | --- | --- | --- |
| `--target com` (default) | **Windows** with the OneNote desktop app (Microsoft 365 / OneNote 2016) | `pip install -e ".[com]"` | No Azure app, no sign-in, no size limits. Works offline. |
| `--target graph` | Any OS | Azure app registration + Microsoft sign-in (below) | 4 MB per-request limit; large files go to OneDrive. |

#### Option A: OneNote desktop app (COM, the default)

```powershell
pip install -e ".[com]"
enex2onenote notebooks          # checks that COM works and lists your notebooks
```

- **Create the notebook in OneNote first.** `--notebook` must name a notebook
  that's open in OneNote. Create new ones with **File → New → OneDrive**, so
  that OneNote syncs them to OneDrive itself.
  - **Why the tool won't create notebooks.** COM can only create a notebook
    from a folder of files. OneNote then treats it as a local notebook, or
    the OneDrive app ends up syncing its files alongside OneNote.
  - **If the notebook isn't open,** the tool stops and lists the notebooks
    that are.
  - **Sections** are created as needed.
- **Attachments.** Files are embedded as file attachments. With PyMuPDF
  installed (part of `.[com]`), each PDF's pages are also shown as images on
  the page. The first 30 pages are shown; change this with
  `--printout-max-pages`.
- **Troubleshooting.** If COM fails to start, open OneNote once, then delete
  `%LOCALAPPDATA%\Temp\gen_py` and try again.
- **"Library not registered".** Some Click-to-Run Office installs leave
  OneNote's COM type library registration incomplete: it's registered for
  32-bit programs only, next to an empty version 1.0 entry. When that
  happens, no automation call works.
  - **Detect it.** `enex2onenote notebooks` detects this and
    prints PowerShell commands to fix and undo it.
  - **What the fix does.** It adds entries under
    `HKCU\Software\Classes\TypeLib` that point at the library OneNote already
    ships. It applies only to your Windows account and needs no admin
    rights.
  - **After applying it.** Fully restart OneNote. It can keep running in the
    background after you close its window.

#### Option B: Microsoft Graph (macOS/Linux, or without the desktop app)

Pass `--target graph` to `upload`, `migrate` and `notebooks`. Graph needs a
one-time app registration:

1. Go to <https://entra.microsoft.com> → **App registrations** → **New registration**.
   - Name: `enex2onenote`
   - Supported account types: choose **"Personal Microsoft accounts only"**
     for a personal OneNote, or include work/school accounts if you need them.
   - Redirect URI: platform **Public client/native**, value
     `https://login.microsoftonline.com/common/oauth2/nativeclient`
2. Under **Authentication**, set **Allow public client flows** to **Yes**.
3. Under **API permissions**, add the Microsoft Graph *delegated* permissions
   `Notes.ReadWrite` and `Files.ReadWrite`.
4. Copy the **Application (client) ID**:

```bash
export ENEX2ONENOTE_CLIENT_ID=<application-client-id>
export ENEX2ONENOTE_TENANT=consumers   # personal accounts; "organizations" or "common" for work accounts
enex2onenote login                     # opens a device-code sign-in; the token is cached in ~/.enex2onenote
```

## Usage

Export a notebook from Evernote (right-click the notebook → **Export
notebook** → `.enex`). Then:

```bash
# Try it on the first 10 notes
enex2onenote analyze "Household Bills.enex" --limit 10

# Analyze everything, then edit migration-state/<folder>/Household Bills/review.csv
enex2onenote analyze "Household Bills.enex"

# See the pages locally without uploading (OneNote XML files in
# migration-state/.../preview; with --target graph, the HTML sent to Graph)
enex2onenote upload "Household Bills.enex" --dry-run

# Upload to OneNote notebook "Evernote Import", section "Household Bills"
enex2onenote upload "Household Bills.enex" --notebook "Evernote Import"                  # desktop app
enex2onenote upload "Household Bills.enex" --notebook "Evernote Import" --target graph   # Graph

# Or do both steps in one go, without the review step
enex2onenote migrate "Recipes.enex" --notebook "Evernote Import" --section "Recipes"
```

### A folder of exports: one folder per notebook

Put your exports in one folder of notebook folders:

```
Exports/
  Finances/              -> OneNote notebook "Finances"
    Receipts.enex        -> section "Receipts", one page per note
    Taxes.enex           -> section "Taxes"
  Family/                -> OneNote notebook "Family"
    Recipes.enex
```

Then pass the folder instead of a file:

```bash
enex2onenote analyze Exports            # analyzes every .enex file
# ...review migration-state/<Notebook>/<Section>/review.csv files...
enex2onenote upload Exports             # uploads all of them

enex2onenote upload Exports/Finances    # just one notebook folder
```

- **Create the notebooks first.** Create each notebook in OneNote (File →
  New → OneDrive) with the same name as its folder. Before anything is
  analyzed or uploaded, `upload` and `migrate` check that every notebook
  exists. They list all the missing ones at once.
- **One level only.** The tool stops instead of skipping files if:
  - an `.enex` file sits in a folder inside a notebook folder;
  - loose `.enex` files sit next to notebook folders;
  - two files in a notebook would get the same section name. Section names
    are limited to 50 characters, and characters OneNote can't use (such as
    `& # %`) become `-`.
- **Naming comes from the folders.** `--notebook` and `--section` can't be
  used with a folder.
- **`--limit N` applies per file.** It takes the first N notes of each file.
- **Summary.** Each file gets its own summary, followed by totals for the
  whole run.

### Merging several exports into one notebook

For single files, `--notebook` picks the OneNote notebook. With the desktop app, create it in
OneNote first; with `--target graph`, it's created on the first upload. Each
export becomes its own section unless you pass `--section`:

```bash
# One notebook, one section per export ("Bills 2019", "Bills 2020")
enex2onenote upload "Bills 2019.enex" --notebook "Finances"
enex2onenote upload "Bills 2020.enex" --notebook "Finances"

# One notebook, all pages in a single section
enex2onenote upload "Bills 2019.enex" --notebook "Finances" --section "Bills"
enex2onenote upload "Bills 2020.enex" --notebook "Finances" --section "Bills"
```

Progress is tracked per `.enex` file, in `migration-state/<folder>/<file name>/`.
This means:

- **Same file name in different folders is fine.** If you keep one folder
  per notebook, `Finances/Receipts.enex` and `Household/Receipts.enex` each
  get their own progress and `review.csv`.
- **Two folders with the same name are refused.** If they're in different
  places (`A/Finances/Receipts.enex` and `B/Finances/Receipts.enex`), the
  tool stops instead of mixing their progress. Rename one, or pass a
  different `--state-dir`.
- **Moving or renaming the export folder is fine.** Its progress comes
  along.
- **Uploaded notes stay uploaded.** Once a note is uploaded, it isn't
  uploaded again from that export, even
  if you change `--notebook` or `--section` later.
- A note that appears in two exports is uploaded once from each.

Useful options:

- `--instructions "Start every title with the document date as YYYY-MM-DD"`:
  adds your own title or tag conventions to the prompt (`--instructions-file`
  reads them from a file).
- `--keep-good-titles`: keeps the Evernote title when Claude judges it
  already accurate.
- `--reanalyze`: runs Claude again on notes that already have results.
  Edits you made in `review.csv` for those notes are replaced.
- `--no-pdf-printouts`, `--no-summary`: plainer pages.
- `--workers N`: analyzes N notes in parallel (default 4).

## Limits and behavior

- **Notes whose analysis failed.** These notes are not uploaded under their
  old Evernote title. Re-run `analyze`, or type a `final_title` for them in
  `review.csv`.
- **Size limits (Graph only).** Microsoft Graph caps each request at 4 MB and 5
  attachments, so the tool splits notes with many attachments into one create
  request and several follow-up updates. Attachments larger than about 3.4 MB
  can't go through the OneNote API. The tool uploads those to OneDrive under
  `Evernote Attachments/<notebook>/<section>/` and links them from the page.
- **Tags.** OneNote has no free-form tags like Evernote. Tags become
  `#hashtags` in the page header, so OneNote search (`#taxes`) finds them.
- **Full sections (Graph only).** When a section reaches OneNote's page limit, the upload
  continues in "Section (2)".
- **Content that isn't migrated.** Encrypted Evernote text is replaced with a
  notice. Word or Excel attachments are uploaded but not read by Claude.

## Development

```bash
pip install -e ".[dev]"
pytest
```

The tests in `tests/test_onenote_live.py` use the real OneNote desktop app
(Windows, `pip install -e ".[com,dev]"`). They are skipped unless you opt
in. Each test creates a throwaway local notebook in a temp folder, then
closes it and deletes its files. Your other notebooks are not touched.

```powershell
$env:ENEX2ONENOTE_LIVE_ONENOTE = "1"
pytest -m onenote
```
