# Setting up the PO $ registry (Google Sheet)

This lets the app remember `Total PO $` values by `Okay PN` across days, so you
only enter a value once. Without this setup the app still works exactly as
before — it just won't pre-fill anything, and you'll see a warning banner.

## 1. Create the Google Sheet

Create a new blank Google Sheet (any name). Copy its ID out of the URL:

```
https://docs.google.com/spreadsheets/d/THIS_PART_IS_THE_ID/edit
```

The app will create a "PO Registry" tab inside it automatically the first
time it runs — you don't need to add any columns yourself.

## 2. Create a Google Cloud service account

1. Go to https://console.cloud.google.com/ and create a project (or reuse one).
2. Enable **Google Sheets API** and **Google Drive API** for that project
   (APIs & Services > Library > search each > Enable).
3. Go to **APIs & Services > Credentials > Create Credentials > Service account**.
   Give it any name, no roles needed, click through to finish.
4. Open the new service account > **Keys** tab > **Add Key > Create new key
   > JSON**. This downloads a `.json` file — keep it private, it's a credential.

## 3. Share the Sheet with the service account

Open the downloaded JSON file and copy the `client_email` value (looks like
`something@your-project.iam.gserviceaccount.com`). In your Google Sheet, click
**Share** and add that email as an **Editor**.

## 4. Add the credentials to Streamlit

**Locally** (for testing with `streamlit run main.py`): create
`.streamlit/secrets.toml` in this project folder (this file is gitignored —
never commit it) with:

```toml
po_registry_sheet_id = "THE_SHEET_ID_FROM_STEP_1"

[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "...@....iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
```

Copy each field's value straight out of the downloaded JSON file into the
matching key above.

**On Streamlit Community Cloud**: open your app > **Settings > Secrets**, and
paste the same TOML content there instead of using a local file.

## 5. Verify

Restart the app and upload a file. If it's working, the Step 2 caption will
say values are being pre-filled from the saved registry instead of showing
the yellow "PO $ registry isn't connected yet" warning.
