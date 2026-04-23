import csv
import os
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta, timezone

# Use absolute path for Streamlit Cloud compatibility
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "usage_log.csv")

def is_gsheet_configured():
    """Checks if Google Sheets connection is configured in st.secrets."""
    try:
        # Standard Streamlit GSheets connection format
        return (hasattr(st, "secrets") and 
                "connections" in st.secrets and 
                "gsheets" in st.secrets["connections"])
    except Exception:
        return False

def log_event(user, event_type, details):
    """Logs an event to the appropriate destination (CSV or GSheets)."""
    vn_tz = timezone(timedelta(hours=7))
    timestamp = datetime.now(vn_tz).strftime("%Y-%m-%d %H:%M:%S")
    
    if is_gsheet_configured():
        try:
            # Import GSheetsConnection only when needed
            from streamlit_gsheets import GSheetsConnection
            conn = st.connection("gsheets", type=GSheetsConnection)
            
            # Read existing data to append (GSheets connection doesn't have native append yet)
            # Use ttl=0 to always get the latest data
            df = conn.read(ttl=0)
            
            new_row = pd.DataFrame([{
                "Timestamp": timestamp,
                "User": user,
                "Event Type": event_type,
                "Details": details
            }])
            
            updated_df = pd.concat([df, new_row], ignore_index=True)
            conn.update(data=updated_df)
            return # Success
        except Exception as e:
            print(f"Error logging to GSheet: {e}")
            # Fallback to local CSV if GSheet fails

    # Default: Log to local CSV
    file_exists = os.path.isfile(LOG_FILE)
    try:
        with open(LOG_FILE, mode='a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["Timestamp", "User", "Event Type", "Details"])
            writer.writerow([timestamp, user, event_type, details])
    except Exception as e:
        print(f"Error logging event to CSV: {e}")

def get_logs():
    """Reads all logs from the appropriate source."""
    if is_gsheet_configured():
        try:
            from streamlit_gsheets import GSheetsConnection
            conn = st.connection("gsheets", type=GSheetsConnection)
            df = conn.read(ttl=0)
            if not df.empty:
                return df.to_dict('records')[::-1]
        except Exception as e:
            print(f"Error reading from GSheet: {e}")

    # Fallback to local CSV
    if not os.path.exists(LOG_FILE):
        return []
    
    try:
        logs = []
        with open(LOG_FILE, mode='r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                logs.append(row)
        return logs[::-1]
    except Exception:
        return []
