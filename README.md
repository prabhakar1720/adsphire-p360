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
