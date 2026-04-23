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
        # Check for both standard and common variations
        if hasattr(st, "secrets"):
            if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
                return True
            if "gsheets" in st.secrets:
                return True
    except Exception:
        pass
    return False

def get_logging_mode():
    """Returns a string indicating where logs are currently being saved."""
    if is_gsheet_configured():
        return "Google Sheets (Cloud)"
    return "CSV (Local)"

def log_event(user, event_type, details):
    """Logs an event to the appropriate destination (CSV or GSheets)."""
    vn_tz = timezone(timedelta(hours=7))
    timestamp = datetime.now(vn_tz).strftime("%Y-%m-%d %H:%M:%S")
    
    if is_gsheet_configured():
        try:
            print(f"DEBUG: Attempting to log to Google Sheets for user {user}...")
            # Import GSheetsConnection only when needed
            from streamlit_gsheets import GSheetsConnection
            
            # Use specific connection key if available
            conn_key = "gsheets" if "connections" in st.secrets else "gsheets"
            
            # Robustly find the spreadsheet URL from secrets
            ss_url = None
            if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
                ss_url = st.secrets["connections"]["gsheets"].get("spreadsheet") or st.secrets["connections"]["gsheets"].get("url")
            elif "gsheets" in st.secrets:
                ss_url = st.secrets["gsheets"].get("spreadsheet") or st.secrets["gsheets"].get("url")
            
            conn = st.connection(conn_key, type=GSheetsConnection)
            
            # Read existing data with 0 TTL to ensure freshness
            # Pass the URL explicitly if we found it to be extra safe
            df = conn.read(spreadsheet=ss_url, ttl=0)
            
            new_row = pd.DataFrame([{
                "Timestamp": timestamp,
                "User": user,
                "Event Type": event_type,
                "Details": details
            }])
            
            # Ensure columns match
            if df.empty:
                updated_df = new_row
            else:
                updated_df = pd.concat([df, new_row], ignore_index=True)
            
            conn.update(spreadsheet=ss_url, data=updated_df)
            print("DEBUG: Google Sheets log update successful.")
            st.session_state["last_log_status"] = "Success: Logged to Google Sheets"
            return # Success
        except Exception as e:
            error_msg = f"CRITICAL ERROR: Failed to log to GSheet: {str(e)}"
            print(error_msg)
            st.session_state["last_log_status"] = f"Error: {str(e)}"
            # Final fallback just in case, but now we have a status
            log_to_csv_fallback(timestamp, user, event_type, details)
            return

    # Default: Log to local CSV (for localhost)
    log_to_csv_fallback(timestamp, user, event_type, details)
    st.session_state["last_log_status"] = "Success: Logged to local CSV"

def log_to_csv_fallback(timestamp, user, event_type, details):
    """Fallback helper for CSV logging."""
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
            conn_key = "gsheets" if "connections" in st.secrets else "gsheets"
            
            # Find the URL
            ss_url = None
            if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
                ss_url = st.secrets["connections"]["gsheets"].get("spreadsheet") or st.secrets["connections"]["gsheets"].get("url")
            elif "gsheets" in st.secrets:
                ss_url = st.secrets["gsheets"].get("spreadsheet") or st.secrets["gsheets"].get("url")
                
            conn = st.connection(conn_key, type=GSheetsConnection)
            df = conn.read(spreadsheet=ss_url, ttl=0)
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
