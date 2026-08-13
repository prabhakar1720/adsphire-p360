# Adsphire P360 Hosted Team Dashboard

A private team version of the P360 dashboard.

## What changed

- Team login with one shared password.
- Separate admin password.
- AppsFlyer V2 tokens are stored on the backend by PRT ID.
- Tokens are not included in Excel and are never sent to the browser.
- Team members can upload Excel rows containing App ID, App Name, PID, PRT, Reporting Timezone and an optional Billable Event Name.
- Team members can save new app rows directly from the dashboard and edit existing saved rows.
- Manual rows and admin token mappings are stored in `/data`.
- Dashboard interface is English-only.

## Environment variables

- `TEAM_PASSWORD`: shared password for the 3–4 dashboard users.
- `ADMIN_PASSWORD`: password for `/admin`.
- `SECRET_KEY`: Flask session secret. Render can generate it.
- `DATA_DIR`: use `/data` on Render.
- `AFFISE_BASE_URL`: your Affise admin/API base URL, for example `https://offers-saltoro.affise.com`.
- `AFFISE_API_KEY`: Affise API key with permission to read custom statistics.

## Local run

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
export TEAM_PASSWORD='your-team-password'
export ADMIN_PASSWORD='your-admin-password'
export SECRET_KEY='replace-with-a-long-random-string'
export AFFISE_BASE_URL='https://offers-saltoro.affise.com'
export AFFISE_API_KEY='your-affise-api-key'
python app.py
```

Open `http://127.0.0.1:8787`.

## Render

Use `render.yaml` as a Blueprint, or create a Python Web Service manually.
The included Blueprint uses a Starter service and a 1 GB persistent disk mounted at `/data`.

## Persistent Excel configuration

The latest uploaded Excel workbook is saved under the Render persistent disk at `/data` and automatically loads for every authenticated team member. Upload a new workbook only when replacing the configuration. The dashboard also provides **Saved Excel** download and **Remove Saved Excel** controls. PRT tokens and manually added rows remain separate and are not removed when the saved Excel is cleared.

## Global filters update
The report table now includes filters for PID, PRT, Account, App, App ID, and Day. These filters apply to all grouping tabs, KPI cards, publisher reporting, and current-view CSV exports.

## Single-app run selector

Use **App to check** before clicking **Run Summary**:

- **All configured apps** preserves the existing full-dashboard behavior.
- Selecting one app limits the AppsFlyer API requests and report calculation to that App ID and its configured PID/PRT scopes.
- The selector is rebuilt automatically whenever the saved Excel or manual app configuration changes.

## Billable events

Add the exact, case-sensitive AppsFlyer **Billable Event Name** to each App + PID + PRT row. During a run, the dashboard requests the accepted raw in-app-event report, keeps only rows matching the configured scope and event name, and shows the count in the KPI, report tables and CSV exports.

The dashboard does not request or count the blocked in-app-event report. Existing blocked-install fraud reporting is unchanged. To add an event name to an existing manually saved row, open **Manage Saved Apps**, click **Edit**, enter the event name and click **Update App Row**.

## AppsFlyer impressions

Every run also reads impressions from AppsFlyer's aggregate `partners_by_date_report`. Only impressions belonging to configured App + PID + PRT scopes are included. Impression counts appear in the KPI cards, grouped report table and current-view CSV export. The response is cached and reused if install fallback or the Affise click comparison needs the same aggregate report.

Impressions are not added to the publisher-level fraud report because the aggregate report does not provide the raw `af_ad` or `af_adset` publisher fields.

## Affise vs AppsFlyer clicks

Select one app before running the summary. Add that app's comma-separated **Affise Offer IDs** in the Excel Apps sheet or its manually saved row. The dashboard sums raw Affise clicks for those offers and compares them with AppsFlyer aggregate clicks matching all configured PIDs for the selected app.

Offer mapping is intentional: Affise click statistics support filtering by date, offer, and affiliate, but not reliably by advertiser. The Affise API key stays on the backend and is never sent to the browser.
