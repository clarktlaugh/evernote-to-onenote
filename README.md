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

Run it once per Evernote notebook, whenever you're ready to move that notebook.

## How it works

1. **`analyze`** parses the `.enex` file and asks Claude about each note. The
   results go to `migration-state/<notebook>/review.csv`.
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
pip install -e .
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
| `--target com` | **Windows** with the OneNote desktop app (Microsoft 365 / OneNote 2016) | `pip install -e ".[com]"` | No Azure app, no sign-in, no size limits. Works offline. |
| `--target graph` (default) | Any OS | Azure app registration + Microsoft sign-in (below) | 4 MB per-request limit; large files go to OneDrive. |

#### Option A: OneNote desktop app (COM)

```powershell
pip install -e ".[com]"
enex2onenote notebooks --target com          # checks that COM works and lists your notebooks
```

- **Notebook location.** If `--notebook` names a notebook that's already open
  in OneNote, pages go there. If it isn't open, the tool creates a **local**
  notebook in OneNote's default notebook folder (or in `--notebook-folder`).
  To have a new notebook sync across devices, create it on OneDrive in
  OneNote first (**File → New → OneDrive**), then pass its name.
- **Attachments.** Files are embedded as file attachments. With PyMuPDF
  installed (part of `.[com]`), each PDF's pages are also shown as images on
  the page. The first 30 pages are shown; change this with
  `--printout-max-pages`.
- **Troubleshooting.** If COM fails to start, open OneNote once, then delete
  `%LOCALAPPDATA%\Temp\gen_py` and try again.

#### Option B: Microsoft Graph (register an app, one time)

OneNote is updated through Microsoft Graph, which needs an app registration:

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

# Analyze everything, then edit migration-state/Household Bills/review.csv
enex2onenote analyze "Household Bills.enex"

# See the pages locally without uploading (files in migration-state/.../preview;
# add --target com to preview the OneNote XML instead of Graph's HTML)
enex2onenote upload "Household Bills.enex" --dry-run

# Upload to OneNote notebook "Evernote Import", section "Household Bills"
enex2onenote upload "Household Bills.enex" --notebook "Evernote Import" --target com   # desktop app
enex2onenote upload "Household Bills.enex" --notebook "Evernote Import"                # Graph

# Or do both steps in one go, without the review step
enex2onenote migrate "Recipes.enex" --notebook "Evernote Import" --section "Recipes"
```

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
