import os
import io
import re
import json
import time
import httplib2
import pandas as pd
import base64
import pickle
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# --- CONFIGURATION ---

# 1. AMAZON AUTOMATION (Original)
AMAZON_SOURCE_ID = "104ARgCkdAfalJQLO1_vIwAVmGwR-Yqul"
AMAZON_DEST_ID = "1HvxOOvh8kwMO1qHS1Vg5mI1sB2VlAA3g"

# 2. PRICING AUTOMATION (New)
PRICING_SOURCE_ID = "1OLCSaeQTiBUKapL3_y-phslrmPDjULAy"
PRICING_DEST_ID = "1t1Rh5K3eLMh54xesHLNlIGIGmxEc36os"

SCOPES = ['https://www.googleapis.com/auth/drive']

# 🛑 SET TO FALSE FOR LIVE UPLOADS
DRY_RUN = False 


def get_drive_service():
    print("Authenticating with Google Drive via OAuth...")
    creds = None
    
    token_b64 = os.environ.get("GOOGLE_TOKEN_PICKLE_B64")
    
    if token_b64:
        print("Using OAuth token from Environment Variable...")
        creds = pickle.loads(base64.b64decode(token_b64))
        if not creds.valid:
            if creds.expired and creds.refresh_token:
                print("Token expired, refreshing...")
                creds.refresh(Request())
    else:
        if os.path.exists('token.pickle'):
            print("Using local token.pickle file...")
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
            if not creds.valid:
                if creds.expired and creds.refresh_token:
                    print("Token expired, refreshing...")
                    creds.refresh(Request())
        else:
            raise ValueError("❌ No OAuth token found! Please run generate_token.py first.")

    return build('drive', 'v3', credentials=creds)


def execute_resumable(request, max_attempts=6):
    response = None
    attempt = 0
    while response is None:
        try:
            status, response = request.next_chunk(num_retries=5)
            if status:
                print(f"   Upload progress: {int(status.progress() * 100)}%")
        except (OSError, httplib2.HttpLib2Error) as e:
            attempt += 1
            if attempt > max_attempts:
                raise
            wait = min(2 ** attempt, 30)
            print(f"⚠️ Connection problem ({type(e).__name__}); retry {attempt}/{max_attempts} in {wait}s...")
            time.sleep(wait)
    return response


def get_latest_file(service, folder_id, is_pricing=False):
    if is_pricing:
        query = f"'{folder_id}' in parents and (name contains '.xlsb' or name contains '.xlsx' or mimeType='application/vnd.google-apps.spreadsheet') and trashed=false"
    else:
        query = f"'{folder_id}' in parents and name contains '.xlsb' and trashed=false"
        
    results = service.files().list(
        q=query,
        orderBy="modifiedTime desc",
        pageSize=1,
        fields="files(id, name, modifiedTime, mimeType)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    
    files = results.get('files', [])
    return files[0] if files else None


def handle_upload(service, dest_folder_id, new_name, csv_bytes):
    print("Checking Destination Folder for existing files...")
    
    # Catch absolutely everything in the folder
    dest_results = service.files().list(
        q=f"'{dest_folder_id}' in parents and trashed=false",
        orderBy="modifiedTime desc",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    old_files = dest_results.get('files', [])

    if DRY_RUN:
        print(f"[DRY RUN] Would have CREATED a brand new file: {new_name}")
        for f in old_files:
            print(f"[DRY RUN] Would have TRASHED old file: {f['name']}")
        return

    media = MediaIoBaseUpload(
        io.BytesIO(csv_bytes),
        mimetype='text/csv',
        chunksize=8 * 1024 * 1024,
        resumable=True
    )

    print(f"Uploading NEW file: '{new_name}'...")
    execute_resumable(service.files().create(
        body={'name': new_name, 'parents': [dest_folder_id]},
        media_body=media,
        fields='id',
        supportsAllDrives=True
    ))
    print(f"✅ Successfully uploaded: {new_name}")

    # 🧹 CLEANUP: Trash all old files
    if old_files:
        for old_file in old_files:
            try:
                service.files().update(
                    fileId=old_file['id'],
                    body={'trashed': True},
                    supportsAllDrives=True
                ).execute()
                print(f"🗑️ Moved to Trash: {old_file['name']}")
            except Exception as e:
                print(f"⚠️ FAILED to trash {old_file['name']}: {e}")


# ==========================================
# AUTOMATION 1: AMAZON ORDER DATA
# ==========================================
def process_amazon_automation(service):
    print("\n" + "="*50)
    print("🛒 RUNNING AUTOMATION 1: AMAZON ORDER DATA")
    print("="*50)
    
    latest_file = get_latest_file(service, AMAZON_SOURCE_ID, is_pricing=False)
    if not latest_file:
        print("❌ No .xlsb files found in Amazon Source folder.")
        return

    file_id = latest_file['id']
    print(f"✅ Found latest file: {latest_file['name']}")

    print("Downloading file...")
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    fh.seek(0)

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
            header_row_idx = 0

        fh.seek(0)
        df = pd.read_excel(fh, sheet_name=0, engine='calamine', header=header_row_idx)
        print(f"✅ Found headers on row {header_row_idx + 1}.")
    except Exception as e:
        print(f"❌ Error reading Excel: {e}")
        return

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
            filtered_df[req_col] = ""

    print(f"✅ Filtered data! Shape: {filtered_df.shape[0]} rows, {filtered_df.shape[1]} columns.")

    csv_buffer = io.StringIO()
    filtered_df.to_csv(csv_buffer, index=False)
    csv_bytes = csv_buffer.getvalue().encode('utf-8')

    # AMAZON NAMING
    new_name = re.sub(r'\.xlsb$', ' - Extracted.csv', latest_file['name'], flags=re.IGNORECASE)
    print(f"📁 Target file name: {new_name}")

    handle_upload(service, AMAZON_DEST_ID, new_name, csv_bytes)


# ==========================================
# AUTOMATION 2: PRICING TEAM DATA
# ==========================================
def process_pricing_automation(service):
    print("\n" + "="*50)
    print("🏷️ RUNNING AUTOMATION 2: PRICING TEAM DATA")
    print("="*50)
    
    latest_file = get_latest_file(service, PRICING_SOURCE_ID, is_pricing=True)
    if not latest_file:
        print("❌ No spreadsheet files found in Pricing Source folder.")
        return

    file_id = latest_file['id']
    original_name = latest_file['name']
    mime_type = latest_file['mimeType']
    print(f"✅ Found latest file: {original_name}")

    print("Downloading file...")
    fh = io.BytesIO()
    
    if mime_type == 'application/vnd.google-apps.spreadsheet':
        print("   (Google Sheet detected. Exporting to .xlsx format...)")
        request = service.files().export_media(
            fileId=file_id, 
            mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        request = service.files().get_media(fileId=file_id, supportsAllDrives=True)

    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    fh.seek(0)

    print("Parsing Excel data...")
    try:
        df = pd.read_excel(fh, sheet_name=0, engine='calamine')
        df.columns = [str(c).strip().upper() for c in df.columns]
        print(f"✅ Successfully read file with {len(df)} rows.")
    except Exception as e:
        print(f"❌ Error reading Excel: {e}")
        return

    required_columns = [
        "SKU", "ITEM_ID", "COST", "MIN_PRICE", "MAX_PRICE", 
        "CURRENT_PRICE", "MANUAL_PRICE", "STRATEGY_ID", 
        "BUYBOX_PRICE", "BUYBOX_WINNER"
    ]
    filtered_df = pd.DataFrame()

    for req_col in required_columns:
        if req_col in df.columns:
            filtered_df[req_col] = df[req_col]
        else:
            filtered_df[req_col] = ""

    filtered_df = filtered_df.fillna("")
    print(f"✅ Filtered data! Shape: {filtered_df.shape[0]} rows, {filtered_df.shape[1]} columns.")

    csv_buffer = io.StringIO()
    filtered_df.to_csv(csv_buffer, index=False)
    csv_bytes = csv_buffer.getvalue().encode('utf-8')

    # 🟢 PRICING NAMING: Uses exact source file name + _Trimmed.csv
    if mime_type == 'application/vnd.google-apps.spreadsheet':
        base_name = original_name
    else:
        base_name = re.sub(r'\.(xlsb|xlsx)$', '', original_name, flags=re.IGNORECASE)
        
    new_name = f"{base_name}_Trimmed.csv"
    print(f"📁 Target file name: {new_name}")

    handle_upload(service, PRICING_DEST_ID, new_name, csv_bytes)


# ==========================================
# MAIN EXECUTION
# ==========================================
SCRIPT_VERSION = "v18-source-name-pricing"

if __name__ == "__main__":
    print(f"Script version: {SCRIPT_VERSION}")
    service = get_drive_service()
    
    process_amazon_automation(service)
    process_pricing_automation(service)
    
    print("\n✅ All automations finished successfully!")
