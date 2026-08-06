# Update the existing hosted dashboard

1. In GitHub, open the `adsphire-p360` repository.
2. Replace `app.py` and `dashboard.html` with the files from this package. You may upload the complete package and overwrite matching files.
3. Commit the changes to the `main` branch.
4. Render should redeploy automatically. If it does not, open the Blueprint and click **Manual sync** once.
5. After the service becomes Live, sign in and upload the Excel configuration one final time.
6. The workbook is then saved at `/data/saved_excel_config.xlsx` and auto-loads for all team members and incognito sessions.

Existing PRT tokens and manually saved app rows are not deleted by this update because they already live on the same `/data` persistent disk.
