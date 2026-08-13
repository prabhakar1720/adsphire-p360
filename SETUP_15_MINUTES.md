# 15-Minute Setup — Adsphire P360 Hosted Team Dashboard

## What you need

- A GitHub account
- A Render account
- A payment method for the Render Starter service and 1 GB persistent disk
- Your shared team password
- A separate admin password
- AppsFlyer V2 tokens for each PRT

## 1. Create a private GitHub repository — 2 minutes

1. Open GitHub.
2. Click **New repository**.
3. Name it `adsphire-p360`.
4. Select **Private**.
5. Create the repository.
6. Click **Add file → Upload files**.
7. Upload every file and folder from this package.
8. Click **Commit changes**.

Never upload real AppsFlyer tokens to GitHub. Tokens are added later from the Admin page.

## 2. Deploy with Render Blueprint — 5 minutes

1. Open Render and connect your GitHub account.
2. Click **New → Blueprint**.
3. Select the private `adsphire-p360` repository.
4. Render detects `render.yaml`.
5. Enter the requested secret values:
   - `TEAM_PASSWORD`: shared password for the 3–4 team members
   - `ADMIN_PASSWORD`: separate password only for the administrator
6. Confirm the Blueprint and start deployment.
7. Wait for status **Live**.

The Blueprint creates:

- One Python web service in Singapore
- One 1 GB persistent disk mounted at `/data`
- A generated Flask `SECRET_KEY`

## 3. Add AppsFlyer tokens — 3 minutes

1. Open the Render URL.
2. Sign in with `TEAM_PASSWORD`.
3. Click **Admin**.
4. Enter `ADMIN_PASSWORD`.
5. Add every PRT and its AppsFlyer V2 token.

Example:

- PRT: `adsphirein749`
- Token: the V2 token for that AppsFlyer account

If several PRT IDs use the same token, save the token separately against each PRT ID.

## 4. Add app rows — 3 minutes

### Excel method

1. Download **Excel Template** from the dashboard.
2. Fill the Apps sheet:
   - Enabled
   - App Name
   - App ID
   - PRT ID
   - Media Source / PID
   - Reporting Timezone
   - Notes
3. Upload the Excel file.

### Manual method

Use **Add an App Row Manually**. These rows are saved on the server and shared with all logged-in users.

## 5. Test — 2 minutes

1. Select **Yesterday** first because post-attribution data for Today can still change.
2. Click **Run Summary**.
3. Search one app you know, such as Upstox.
4. Confirm attributed installs, blocked installs and fraud rate.
5. Test Site ID view with the App and PID drill-down filters.
6. Download Current View CSV.

## URLs

- Dashboard: `https://YOUR-SERVICE.onrender.com/`
- Admin: `https://YOUR-SERVICE.onrender.com/admin`
- Health check: `https://YOUR-SERVICE.onrender.com/health`

## Important

The persistent disk is required. Without it, saved PRT tokens and manually added app rows can disappear after a restart or redeploy.

## Saved Excel behavior

After the first Excel upload, the workbook is stored on the persistent disk and automatically loaded for all users. Use **Upload or Replace Excel Configuration** only when updating the app list. **Remove Saved Excel** deletes only the shared workbook configuration; PRT tokens and manually saved rows remain.
# Daily Target Dashboard

After deployment, open `/admin` and add every AppsFlyer Agency or Ad Network credential under **AppsFlyer Data Accounts**. Then add the campaigns and daily targets below it. Team members can open `/targets` from the main P360 dashboard.

The target dashboard loads clicks and impressions from AppsFlyer aggregate data. Installs and billable events use the access-aware fallback shown in the dashboard: non-organic report, then postback report, then aggregate data.
