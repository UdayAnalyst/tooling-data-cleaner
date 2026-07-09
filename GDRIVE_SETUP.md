# Setting up auto-fetch from Google Drive

This lets the app pull CSV/Excel files straight from a folder in Google
Drive, so you don't have to upload them by hand each time — just drop files
into that folder and open the app. Without this setup the app still works
exactly as before — you'll just see the upload button.

This reuses the same Google Cloud service account already set up for the PO
$ registry (see `PO_REGISTRY_SETUP.md`) — no separate app registration or
admin approval needed, since sharing a Drive folder is self-serve, the same
way you already shared the PO Registry Sheet.

## 1. Confirm the Google Drive API is enabled

If you followed `PO_REGISTRY_SETUP.md`, this is already done (it has you
enable both **Google Sheets API** and **Google Drive API** on the same
project). If not: https://console.cloud.google.com/ > your project > APIs &
Services > Library > search **Google Drive API** > Enable.

## 2. Create the Drive folder

Create (or pick) a folder in Google Drive to drop source files into, e.g.
"Tooling Uploads".

## 3. Share the folder with the service account

Open the folder's **Share** dialog and add your service account's email as
a **Viewer** (read-only is enough — the app never writes to this folder).
The email is in your `gcp_service_account` secrets block as `client_email`
— for this project that's:

```
sheets-api-sa@insurance-insights-youtube.iam.gserviceaccount.com
```

## 4. Get the folder ID

Open the folder in Drive and copy the ID out of the URL:

```
https://drive.google.com/drive/folders/THIS_PART_IS_THE_FOLDER_ID
```

## 5. Add the folder ID to Streamlit secrets

**Locally** (for testing with `streamlit run main.py`): add this line to
`.streamlit/secrets.toml` in this project folder (this file is gitignored —
never commit it), alongside the existing `gcp_service_account` block:

```toml
gdrive_folder_id = "THE_FOLDER_ID_FROM_STEP_4"
```

**On Streamlit Community Cloud**: open your app > **Settings > Secrets**,
and add the same line there instead of using a local file.

## 6. Verify

Restart the app. If it's working, the upload button is replaced by a
caption saying files are fetched automatically from Google Drive, plus a
"Refresh from Drive" button — and Step 2's table populates without you
uploading anything. Drop a new file into the Drive folder and click
**Refresh from Drive** to confirm it picks it up.
