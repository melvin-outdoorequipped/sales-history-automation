import os
import io
import json
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# --- CONFIGURATION ---
SOURCE_FOLDER_ID = "104ARgCkdAfalJQLO1_vIwAVmGwR-Yqul"
DEST_FOLDER_ID = "1HvxOOvh8kwMO1qHS1Vg5mI1sB2VlAA3g"
SCOPES = ['https://www.googleapis.com/auth/drive']

DRY_RUN = False

def get_drive_service():
    print("Authenticating with Google Drive...")
    
    # --- GITHUB SECRET AUTHENTICATION ---
    creds_json = os.environ.get("GDRIVE_CREDENTIALS")
    if not creds_json:
        raise ValueError("❌ Missing GDRIVE_CREDENTIALS environment variable. If running locally, set it or use the local file version.")
    
    creds_dict = json.loads(creds_json)
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

def process_latest_file(service):
    print(f"Checking Source Folder... (Dry Run: {DRY_RUN})")

    # 1. Get the latest .xlsb file
    results = service.files().list(
        q=f"'{SOURCE_FOLDER_ID}' in parents and name contains '.xlsb' and trashed=false",
        orderBy="modifiedTime desc",
        pageSize=1,
        fields="files(id, name, modifiedTime)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()

    files = results.get('files', [])
    if not files:
        print("❌ No .xlsb files found.")
        return

    latest_file = files[0]
    file_id = latest_file['id']
    print(f"✅ Found latest file: {latest_file['name']}")

    # 2. Download the file
    print("Downloading file...")
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()

    fh.seek(0)

    # 3. Parse with pandas
    print("Parsing Excel data...")
    try:
        raw_df = pd.read_excel(fh, sheet_name=0, engine='calamine', header=None, nrows=20)

        header_row_idx = None
        for i in range(len(raw_df)):
            row_vals = [str(val).strip().lower() for val in raw_df.iloc[i].values]
            if "date" in row_vals and ("brand" in row_vals or "sku" in row_vals):
                header_row_idx = i
                break

        if header_row_idx is None:
            print("⚠️ Could not find the header row. Defaulting to row 0.")
            header_row_idx = 0

        fh.seek(0)
        df = pd.read_excel(fh, sheet_name=0, engine='calamine', header=header_row_idx)
        print(f"✅ Found headers on row {header_row_idx + 1}.")

    except Exception as e:
        print(f"❌ Error reading Excel: {e}")
        return

    # 4. Filter Columns
    required_columns = [
        "Date", "BRAND", "amazon-order-id", "fulfillment-channel",
        "sku", "asin", "quantity", "item-price", "Fulfillment"
    ]

    actual_cols = {str(c).strip().lower(): c for c in df.columns}
    filtered_df = pd.DataFrame()

    for req_col in required_columns:
        if req_col in df.columns:
            filtered_df[req_col] = df[req_col]
        elif req_col.lower() in actual_cols:
            filtered_df[req_col] = df[actual_cols[req_col.lower()]]
        else:
            print(f"⚠️ Warning: Column '{req_col}' not found. Creating empty column.")
            filtered_df[req_col] = ""

    print(f"✅ Filtered data! Shape: {filtered_df.shape[0]} rows, {filtered_df.shape[1]} columns.")

    # 5. Convert to CSV
    csv_buffer = io.StringIO()
    filtered_df.to_csv(csv_buffer, index=False)
    csv_bytes = csv_buffer.getvalue().encode('utf-8')
    new_name = latest_file['name'].replace('.xlsb', '.csv')

    # 6. Check Destination Folder
    print("Checking Destination Folder for existing CSVs...")
    dest_results = service.files().list(
        q=f"'{DEST_FOLDER_ID}' in parents and mimeType='text/csv' and trashed=false",
        orderBy="modifiedTime desc",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    old_files = dest_results.get('files', [])

    # 7. Execute
    if DRY_RUN:
        print("\n--- DRY RUN RESULTS ---")
        print(f"Would have UPLOADED: {new_name}")
        for f in old_files:
            print(f"Would have replaced/trashed existing CSV: {f['name']}")
    else:
        print("\n--- LIVE RUN ---")
        media = MediaIoBaseUpload(io.BytesIO(csv_bytes), mimetype='text/csv', resumable=True)

        if old_files:
            # Overwrite the newest existing CSV (content + name).
            keep = old_files[0]
            service.files().update(
                fileId=keep['id'],
                body={'name': new_name},
                media_body=media,
                supportsAllDrives=True
            ).execute()
            print(f"✅ Updated existing file in place: {new_name}")

            # Try to trash any extra leftover CSVs; skip quietly if not permitted
            for extra in old_files[1:]:
                try:
                    service.files().update(
                        fileId=extra['id'],
                        body={'trashed': True},
                        supportsAllDrives=True
                    ).execute()
                    print(f"🗑️ Trashed extra CSV: {extra['name']}")
                except Exception as e:
                    print(f"⚠️ Could not trash {extra['name']}: {e}")
        else:
            service.files().create(
                body={'name': new_name, 'parents': [DEST_FOLDER_ID]},
                media_body=media,
                fields='id',
                supportsAllDrives=True
            ).execute()
            print(f"✅ Successfully uploaded: {new_name}")

if __name__ == "__main__":
    service = get_drive_service()
    process_latest_file(service)
