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
