# Contego MSSP Report Generator (HTML)

Read-only tool that generates branded **HTML** MSSP reports from the Acronis
Cyber Protect Cloud API (backup + security posture), one per customer.

> **This repository contains code only** — no credentials and no client data.
> Each run reads its own API credentials at runtime and writes reports to
> `automation/output/` (gitignored). The pipeline issues **READ-ONLY** calls
> (GET requests + the OAuth token + audit search).

## Run in Google Colab (HTML only)

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/oness24/acronis_report_tool/blob/main/Generate-Reports.ipynb)

1. Click the badge above (opens the notebook in Colab — no GitHub sign-in needed; this repo is public).
2. Add Colab **Secrets** (🔑 icon, left sidebar), toggling "Notebook access" on each:
   - `US_CLOUD_ID`, `US_CLOUD_SECRET` — us-cloud API client id/secret
   - `BR02_ID`, `BR02_SECRET` — br02 API client id/secret
3. **Runtime → Run all.** A `.zip` of the HTML reports downloads at the end.

The report PDF is skipped on Colab (no headless browser there by default). The prepare cell (step 4) installs both `python-pptx` and `chromium-browser`, so each client also gets a Gamma-style `.deck.pdf` (landscape slides, self-contained) in addition to the `.pptx`. This path produces **HTML + PPTX + deck.pdf**.

## Run locally (Windows / PowerShell — also produces PDF)

1. Copy `automation/config/secrets.example.json` → `automation/config/secrets.json` and fill in the credentials.
2. ```
   pwsh -File automation/Run-MonthlyReports.ps1 -Month 2026-05
   ```
   Omit `-Month` to use the previous month. Reports land in `automation/output/<month>/reports/`.

## What it does

- Auto-discovers every customer tenant under each configured data-center root.
- Collects usage, backup activity, alerts, agents and (where available) M365 figures.
- Renders a branded, self-contained HTML report per client (cover, KPI dashboard,
  SVG charts, SLA, action plan, alert glossary) — no JavaScript, offline-safe.

Section visibility is driven by each client's active services; add per-tenant
contract data to `automation/config/contracts.json` to gate sections by contract
instead of by live usage.
