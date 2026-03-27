# DEADLIFT.JP (deadlift-jp)

This repo is a small, mostly-static toolset:

- `counseling-form.html`: Coaching intake form (submits to Formspree).
- `coach-dashboard.html`: Coach dashboard that parses the Formspree email body, calls Anthropic Messages API from the browser, renders the program, and can request an Excel file from a local server.
- `generate_excel_server.py`: Python script that builds an `.xlsx` training program (uses `openpyxl`).

## What to optimize for

- Keep the UI clean and responsive (mobile-first).
- Prefer **minimal, localized edits** over rewrites (these files are large single-page documents).
- Do not introduce build tooling unless explicitly requested (no bundlers/frameworks by default).

## Security / privacy constraints (important)

- **Never hardcode API keys**. `coach-dashboard.html` reads the Anthropic key from an `<input>` and saves it in `localStorage` (`dljp_api_key`). Keep it that way.
- Avoid adding any logging that could leak personal data from the counseling form (names, email, injury details, etc.).
- Do not add new third-party scripts/CDNs unless necessary.

## Editing guidelines

- Keep Japanese UI text consistent and natural.
- Preserve the existing visual style tokens (CSS variables) and typography choices.
- When changing parsing rules, keep backward compatibility with previously received Formspree email formats where possible.

## How to run / test locally

### Static pages

- Open `counseling-form.html` or `coach-dashboard.html` directly in a browser.

### Excel generation

`coach-dashboard.html` expects a local HTTP server at `http://localhost:8765` for Excel generation.

If you change the Python side, verify:

- The response still returns JSON with `{ "xlsx": "<base64>" }`.
- The produced workbook opens cleanly in Excel/Numbers/Google Sheets.

## Code style

- HTML/CSS/JS: vanilla; keep functions small and readable; avoid dependencies.
- Python: keep it script-friendly (simple entrypoint), avoid adding heavy frameworks.

