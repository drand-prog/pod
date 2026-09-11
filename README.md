# The Melting Pod — Analytics Dashboard (project repo)

This repository is set up to be handed to **Claude Code on the web** (claude.ai/code),
which will build an interactive dashboard from the consolidated analytics workbook.

## What's here
- `CLAUDE.md` — the full project brief and context (Claude Code reads this automatically).
- `PROMPT.md` — a starter task to paste into the Claude Code session.
- `data/consolidated/` — the 12-tab analytics workbook.
- `data/raw/` — the original CSV exports.
- `data/screenshots/` — platform screenshots that had no CSV export.

## How to use it with Claude Code on the web
1. Create an **empty GitHub repository** (private is fine).
2. Upload this folder's contents to it — on the repo page use **Add file → Upload files** and
   drag everything in (no git command line needed).
3. Go to **claude.ai/code**, sign in, and connect GitHub when prompted (a paid Claude plan is required).
4. Select this repository, then paste the contents of `PROMPT.md` as the task.
5. Claude Code builds the dashboard on a branch and opens a pull request for review.

## Viewing the dashboard
- Simplest: after merging, download `index.html` and open it in a browser.
- Hosted: enable **GitHub Pages** (Settings → Pages → deploy from branch) to get a shareable URL.

## Refreshing each month
Drop the new platform exports into `data/raw/` (and any new screenshots into `data/screenshots/`),
then start a new Claude Code session with a prompt like: "Refresh the dashboard from the latest
files in data/ and update the screenshot-derived figures." See `CLAUDE.md` for the caveats to preserve.
