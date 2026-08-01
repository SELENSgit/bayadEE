# SELENS Fines Lookup

A static site (GitHub Pages) that lets students search their Student ID to
see any outstanding event fines. Data is pulled automatically from the
SELENS Excel sheet on SharePoint every 30 minutes via a GitHub Action —
no server, no Power Automate, no Apps Script.

## How it works

1. A GitHub Action runs on a schedule (`.github/workflows/update-data.yml`).
2. It downloads the Excel file directly from a public SharePoint link
   (`scripts/build-data.mjs`) and converts it into `data/fines.json`.
3. If the data changed, the Action commits the updated JSON back to the repo.
4. `index.html` is a static page that fetches `data/fines.json` and lets a
   student search their own record by Student ID.

## One-time setup

### 1. Make the Excel file link public
In the Excel file on SharePoint: **Share → change the link to "Anyone with
the link" → Can view** (not just "People at MSU-Gensan" — GitHub's servers
need to reach it without logging in). Copy that link, then append
`&download=1` (or `?download=1` if there's no `?` already) so it serves the
raw file instead of the Excel Online viewer.

Test it in an incognito browser window first — it should immediately
download an `.xlsx` file with no login prompt. If it prompts for login,
the sharing setting isn't public yet.

### 2. Create the GitHub repo
Push this folder to a new GitHub repo (public or private both work for
Pages, but public is required for the free tier if you're on a personal
account without GitHub Pro).

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<your-username>/<repo-name>.git
git push -u origin main
```

### 3. Add the Excel URL as a secret
In the repo: **Settings → Secrets and variables → Actions → New repository
secret**
- Name: `EXCEL_URL`
- Value: the public sharing link with `&download=1` appended

Keeping it as a secret (rather than hardcoding it in the script) avoids
exposing the exact SharePoint file ID publicly, and makes it easy to
rotate later if you regenerate the sharing link.

### 4. Allow the Action to push commits
**Settings → Actions → General → Workflow permissions →** select
**"Read and write permissions"**, then Save. Without this, the Action can
download the data but can't commit `data/fines.json` back to the repo.

### 5. Enable GitHub Pages
**Settings → Pages → Source: Deploy from a branch → Branch: `main` /
root.** Save. Your site will be live at
`https://<your-username>.github.io/<repo-name>/` within a minute or two.

### 6. Run it once manually
**Actions tab → "Update fines data" → Run workflow.** This does the first
sync so the site isn't showing an empty placeholder. After that it runs
automatically every 30 minutes.

## Adjusting the refresh interval
Edit the `cron` line in `.github/workflows/update-data.yml`. For example,
`*/15 * * * *` for every 15 minutes, or `0 */6 * * *` for every 6 hours.
Don't go below every 5 minutes — GitHub throttles very frequent schedules
and it also eats into your free Actions minutes faster than needed for
data that doesn't change that often.

## If the Action fails
Check the Actions tab → the failed run → expand "Fetch latest Excel data
and rebuild JSON". The script deliberately fails loudly with a clear
message if:
- `EXCEL_URL` secret isn't set, or
- the downloaded file isn't a real `.xlsx` (almost always means the
  SharePoint link reverted to private/login-required — someone may have
  re-tightened the sharing settings).

## Logo
`index.html` references `org-logo.png` in the header. Add your logo file
to the repo root with that exact filename (case-sensitive) — if it's
missing, the browser will just show a small broken-image icon in the
header instead of breaking the rest of the page.

## Column mapping
The build script expects these column headers in the first row of the
first sheet (case-insensitive, exact wording otherwise):

```
TXN ID | Student ID | Full Name | Email Address | Event / Offense |
Fine Amount | Date | Status | Date Paid
```

If you rename or reorder columns, update the key mapping in
`scripts/build-data.mjs`.
