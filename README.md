# Adsphire P360 Hosted Team Dashboard

A private team version of the P360 dashboard.

## What changed

- Team login with one shared password.
- Separate admin password.
- AppsFlyer V2 tokens are stored on the backend by PRT ID.
- Tokens are not included in Excel and are never sent to the browser.
- Team members can upload Excel rows containing App ID, App Name, PID, PRT and Reporting Timezone.
- Team members can also save new app rows directly from the dashboard.
- Manual rows and admin token mappings are stored in `/data`.
- Dashboard interface is English-only.
- A separate **Target Dashboard** at `/targets` tracks daily campaign pacing for clicks, impressions, installs and billable events.
- Admin can store both **Agency** and **Ad Network** AppsFlyer data accounts and assign each campaign to the correct credential.
- Install source selection is **non-organic raw report → install postback report → aggregate report**.
- Billable event source selection is **non-organic in-app event report → in-app event postback report → aggregate report**.
- Fallback happens when report access is unavailable. A successful report containing zero matching rows remains a valid zero result.
- Blocked-event postback reports are never requested, and fraud/blocked rows are excluded if a source returns such fields.
- The new daily tracker has Overall, Day-wise and PID-level views and intentionally excludes Site ID. The existing P360 fraud dashboard keeps its original Site ID view and remains restricted to the PIDs in its own saved P360 configuration.
- The daily tracker uses the orange-and-white P360 visual style and provides global PID, PRT, account, app, App ID, campaign, day and status filters. Filters apply to KPI cards, the table and CSV export.
- AppsFlyer responses are cached centrally under `/data/af_report_cache`: aggregate data for 1 hour on current-day ranges and 24 hours for past ranges; raw/postback data for 15 minutes on current-day ranges and 6 hours for past ranges. Page reloads and filters reuse this shared cache, automatic refresh is disabled, and stale saved data is used if AppsFlyer responds with a rate-limit/server error.
- Opening or reloading `/targets` makes no AppsFlyer report calls. Reports load only when a user clicks **Refresh Performance**; changing filters, breakdown tabs and exporting CSV uses the already-loaded browser data.

## Environment variables

- `TEAM_PASSWORD`: shared password for the 3–4 dashboard users.
- `ADMIN_PASSWORD`: password for `/admin`.
- `SECRET_KEY`: Flask session secret. Render can generate it.
- `DATA_DIR`: use `/data` on Render.

## Local run

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
export TEAM_PASSWORD='your-team-password'
export ADMIN_PASSWORD='your-admin-password'
export SECRET_KEY='replace-with-a-long-random-string'
python app.py
```

Open `http://127.0.0.1:8787`.

## Configure campaign targets

1. Open `/admin` and add an **AppsFlyer Data Account**. Choose Agency or Ad Network and paste its V2 token.
2. Add each campaign under **Campaign Daily Targets** with its App ID, PID, data account, billing KPI, exact billable event name and daily target.
   Campaign targets can be paused and enabled again without deleting their setup.
3. Open `/targets` and choose the delivery date. The page refreshes automatically every 15 minutes and supports CSV export.

The daily tracker does not reuse the P360 dashboard's five-PID scope. For every configured AppsFlyer data account and App ID, its aggregate request is unfiltered and every non-organic PID returned by AppsFlyer is shown. PIDs without a saved target appear as **No Target** with their aggregate impressions, clicks and installs; add them in Admin when billable-event pacing is required.

For Angel One, Upstox and other ad-network-only offers, select the Ad Network credential. If neither raw nor postback access is available, installs and events are taken from `partners_by_date_report`. When AppsFlyer returns only an app-level aggregate without PID, the row is labelled **AGGREGATE APP TOTAL** and is used only when it can be assigned unambiguously.

Use a currently valid V2 token. AppsFlyer revoked tokens generated before March 10, 2026, so older credentials must be regenerated before this dashboard can pull data.

## Render

Use `render.yaml` as a Blueprint, or create a Python Web Service manually.
The included Blueprint uses a Starter service and a 1 GB persistent disk mounted at `/data`.

## Persistent Excel configuration

The latest uploaded Excel workbook is saved under the Render persistent disk at `/data` and automatically loads for every authenticated team member. Upload a new workbook only when replacing the configuration. The dashboard also provides **Saved Excel** download and **Remove Saved Excel** controls. PRT tokens and manually added rows remain separate and are not removed when the saved Excel is cleared.

## Global filters update
The report table now includes filters for PID, PRT, Account, App, App ID, and Day. These filters apply to all grouping tabs, KPI cards, fraud reasons, and current-view CSV exports.

## Single-app run selector

Use **App to check** before clicking **Run Summary**:

- **All configured apps** preserves the existing full-dashboard behavior.
- Selecting one app limits the AppsFlyer API requests and report calculation to that App ID and its configured PID/PRT scopes.
- The selector is rebuilt automatically whenever the saved Excel or manual app configuration changes.
