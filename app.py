import streamlit as st
import pandas as pd
import io
import json
import os
import platform
from extractor import extract_table_data
from processor import compare_dataframes
from usage_logger import log_event, get_logs, get_logging_mode
from datetime import datetime, timedelta, timezone

# Page Configuration
st.set_page_config(
    page_title="Word Number Comparison",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USERS_FILE = os.path.join(BASE_DIR, "users.json")
CSS_FILE = os.path.join(BASE_DIR, "style.css")

# Load CSS
if os.path.exists(CSS_FILE):
    with open(CSS_FILE, "r", encoding="utf-8") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# Authentication Logic
def load_users():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)["users"]
    return []

def check_credentials(username, password):
    users = load_users()
    for user in users:
        if user["username"] == username and user["password"] == password:
            return True, user.get("role", "user")
    return False, None

def save_user(username, password, role, auto_fill=False):
    users = load_users()
    if any(u["username"] == username for u in users):
        return False, f"Username '{username}' already exists."
    
    users.append({
        "username": username,
        "password": password,
        "role": role,
        "auto_fill": auto_fill
    })
    
    try:
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump({"users": users}, f, indent=4)
        return True, f"User '{username}' added successfully."
    except Exception as e:
        return False, f"Error saving user: {e}"

def remove_user(username):
    users = load_users()
    updated_users = [u for u in users if u["username"] != username]
    try:
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump({"users": updated_users}, f, indent=4)
        return True, f"User '{username}' deleted successfully."
    except Exception as e:
        return False, f"Error deleting user: {e}"

def update_user_data(old_username, new_username, new_password, new_role, auto_fill):
    users = load_users()
    if old_username != new_username:
        if any(u["username"] == new_username for u in users):
            return False, f"Username '{new_username}' already exists."
    
    for user in users:
        if user["username"] == old_username:
            user["username"] = new_username
            user["password"] = new_password
            user["role"] = new_role
            user["auto_fill"] = auto_fill
            break
            
    try:
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump({"users": users}, f, indent=4)
        return True, f"User '{new_username}' updated successfully."
    except Exception as e:
        return False, f"Error updating user: {e}"

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'username' not in st.session_state:
    st.session_state.username = ""
if 'role' not in st.session_state:
    st.session_state.role = "user"

def handle_autofill():
    users = load_users()
    selected_user = next((u for u in users if u["username"] == st.session_state.login_user), None)
    if selected_user and selected_user.get("auto_fill"):
        st.session_state.login_password = selected_user["password"]
    else:
        st.session_state.login_password = ""

def login_screen():
    st.markdown('<div class="login-container">', unsafe_allow_html=True)
    st.markdown('<div class="login-header">🔐 User Login</div>', unsafe_allow_html=True)
    
    users = load_users()
    usernames = [u["username"] for u in users]
    
    if "login_user" not in st.session_state:
        st.session_state.login_user = "user" if "user" in usernames else (usernames[0] if usernames else "")
        handle_autofill()

    st.selectbox("Select Username", usernames, key="login_user", on_change=handle_autofill)
    
    with st.form("login_form"):
        password = st.text_input("Password", key="login_password", type="password")
        submit = st.form_submit_button("Sign In", use_container_width=True)
        
        if submit:
            is_valid, role = check_credentials(st.session_state.login_user, password)
            if is_valid:
                st.session_state.authenticated = True
                st.session_state.username = st.session_state.login_user
                st.session_state.role = role
                log_event(st.session_state.username, "Login", "Successfully signed in")
                st.rerun()
            else:
                st.error("Invalid username or password")
    
    st.markdown('</div>', unsafe_allow_html=True)

if not st.session_state.authenticated:
    login_screen()
    st.stop()

# --- GLOBAL SIDEBAR (Authenticated Only) ---
active_view = "compare"
with st.sidebar:
    st.markdown(f"👤 **Logged in as:** `{st.session_state.get('username', '')}`")
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()
    st.divider()

    # Admin Navigation
    if st.session_state.get('role') == 'admin':
        st.markdown("### 🧭 Navigation")
        nav_choice = st.radio(
            "Select View",
            ["📊 Number Comparison", "🛡️ Admin Panel"],
            index=0,
            label_visibility="collapsed"
        )
        active_view = "admin" if nav_choice == "🛡️ Admin Panel" else "compare"
        st.divider()

    st.caption(f"v2.1 ({platform.system()} Edition)")

# App UI Header
st.markdown('<div class="app-logo">📊</div>', unsafe_allow_html=True)
st.markdown('<div class="main-header">Word Number Comparison</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Data Comparison & Auditing Utility</div>', unsafe_allow_html=True)

# --- VIEW 1: NUMBER COMPARISON (Direct display, no redundant tab bar) ---
if active_view == "compare":
    # 1. Compact Horizontal Settings Bar at the Top
    st.markdown('<div class="settings-panel">', unsafe_allow_html=True)
    st.markdown("### ⚙️ Settings")
    
    set_col1, set_col2, set_col3, set_col4 = st.columns([1.5, 1, 1, 1.5], gap="medium")
    
    with set_col1:
        st.markdown('<div class="settings-label">1. Number Formats</div>', unsafe_allow_html=True)
        s_col1, s_col2 = st.columns(2)
        with s_col1:
            st.caption("Doc 1")
            format_1 = st.radio("F1", ["Vietnam", "US"], index=0, key="f1_fmt", label_visibility="collapsed")
        with s_col2:
            st.caption("Doc 2")
            format_2 = st.radio("F2", ["Vietnam", "US"], index=1, key="f2_fmt", label_visibility="collapsed")
    
    with set_col2:
        st.markdown('<div class="settings-label">2. Target</div>', unsafe_allow_html=True)
        extract_mode = st.radio("Target", ["Number", "Text"], index=0, horizontal=True, label_visibility="collapsed")
    
    with set_col3:
        st.markdown('<div class="settings-label">3. Filter</div>', unsafe_allow_html=True)
        view_mode = st.radio("Filter", ["Mismatches", "All Results"], index=0, label_visibility="collapsed")

    with set_col4:
        st.info("💡 Header check suggested first.")
    st.markdown('</div>', unsafe_allow_html=True)

    # 2. Upload Area
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        st.markdown('<div class="app-card">', unsafe_allow_html=True)
        st.info("📂 **Upload Document 1**")
        file1 = st.file_uploader("Select first .docx file", type=["docx"], key="doc1")
        st.markdown('</div>', unsafe_allow_html=True)

    with col_c2:
        st.markdown('<div class="app-card">', unsafe_allow_html=True)
        st.info("📂 **Upload Document 2**")
        file2 = st.file_uploader("Select second .docx file", type=["docx"], key="doc2")
        st.markdown('</div>', unsafe_allow_html=True)

    if file1 and file2:
        if st.button("🚀 Run Comparison", use_container_width=True):
            with st.spinner("Extracting and comparing data..."):
                try:
                    df1 = extract_table_data(file1, format_1, mode=extract_mode)
                    df2 = extract_table_data(file2, format_2, mode=extract_mode)
                    
                    if df1.empty or df2.empty:
                        st.warning(f"One of the documents contains no {extract_mode.lower()} table data.")
                    else:
                        merged_df, _, mismatch_count = compare_dataframes(df1, df2, mode=extract_mode)
                        log_event(st.session_state.username, "Comparison", f"Files: {file1.name}, {file2.name} | Mode: {extract_mode}")
                        
                        st.divider()
                        m_col1, m_col2, m_col3 = st.columns(3)
                        m_col1.metric("Rows in Doc 1", len(df1))
                        m_col2.metric("Rows in Doc 2", len(df2))
                        delta_color = "normal" if mismatch_count == 0 else "inverse"
                        m_col3.metric("Mismatches", mismatch_count, delta=mismatch_count, delta_color=delta_color)
                        
                        if mismatch_count == 0:
                            st.success("Verification successful! All values match perfectly.")
                        else:
                            st.error(f"Found {mismatch_count} discrepancies. Please review the table below.")
                            st.markdown("""
                                <div class="word-hint">
                                    <b>💡 Tip: Quickly Find Tables in MS Word</b><br>
                                    1. Press <b>Ctrl + G</b> to open 'Go To' dialog.<br>
                                    2. Select <b>Table</b> in the list.<br>
                                    3. Enter the <b>Table Number</b> from the list below.<br>
                                    4. Click <b>Go To</b> to jump directly to the table.
                                </div>
                            """, unsafe_allow_html=True)
                        
                        display_df = merged_df.copy()
                        if extract_mode == 'Number':
                            mismatch_mask = display_df['Diff'].abs() > 1e-6
                        else:
                            mismatch_mask = display_df['Text 1'].fillna("") != display_df['Text 2'].fillna("")
                        
                        if view_mode == "Mismatches":
                            display_df = display_df[mismatch_mask]
                        
                        if extract_mode == 'Text':
                            display_df = display_df.drop(columns=['Value 1', 'Value 2', 'Diff'])
                        
                        def highlight_diff(row):
                            if extract_mode == 'Number':
                                is_diff = abs(row['Diff']) > 1e-6
                            else:
                                is_diff = str(row['Text 1']) != str(row['Text 2'])
                            return ['background-color: #fef2f2' if is_diff else '' for _ in row]

                        styled_df = display_df.style.apply(highlight_diff, axis=1)
                        if extract_mode == 'Number':
                            styled_df = styled_df.format({'Value 1': "{:,.2f}", 'Value 2': "{:,.2f}", 'Diff': "{:,.2f}"})
                        
                        st.dataframe(styled_df, use_container_width=True, height=600)
                        
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            merged_df.to_excel(writer, index=False, sheet_name='Comparison')
                        st.download_button(
                            label="📥 Download Comparison Report (Excel)",
                            data=output.getvalue(),
                            file_name="Comparison_Report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                except Exception as e:
                    st.error(f"Error during analysis: {str(e)}")
    else:
        st.info("Please upload Word documents to start.")

# --- VIEW 2: ADMIN PANEL (Admin Only) ---
elif active_view == "admin" and st.session_state.get('role') == 'admin':
    admin_tabs = st.tabs(["📊 Usage Analytics", "👥 User Control"])
    
    # --- SUB-TAB: ANALYTICS & LOGS ---
    with admin_tabs[0]:
        mode = get_logging_mode()
        mode_color = "#2e7d32" if "Google Sheets" in mode else "#1976d2"
        
        st.markdown(f"""
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
                <h3 style="margin:0;">📊 System Usage & Activity</h3>
                <div style="display: flex; gap: 10px;">
                    <span style="background-color: {mode_color}; color: white; padding: 4px 12px; border-radius: 12px; font-size: 0.8rem; font-weight: bold;">
                        Backend: {mode}
                    </span>
                </div>
            </div>
        """, unsafe_allow_html=True)

        # Show last error if any
        if "last_log_status" in st.session_state:
            status = st.session_state["last_log_status"]
            if "Error" in status:
                st.error(f"⚠️ **Logging Issue Detected:** {status}")
                st.info("""
                    **How to fix Google Sheets logging:**
                    1. Open your Google Sheet.
                    2. Click **Share** (top right).
                    3. Add the **Client Email** from your Service Account JSON as an **Editor**.
                    4. Ensure the `spreadsheet` URL in Streamlit Secrets is correct.
                """)
            elif "Google Sheets" in status:
                st.success("✅ System is successfully writing to Google Sheets.")
        
        logs = get_logs()
        if logs:
            log_df = pd.DataFrame(logs)
            vn_tz = timezone(timedelta(hours=7))
            
            # 1. Dashboard Metrics
            m1, m2, m3, m4 = st.columns(4)
            total_events = len(log_df)
            unique_users = log_df['User'].nunique()
            most_active = log_df['User'].value_counts().idxmax() if not log_df.empty else "N/A"
            
            try:
                cutoff = datetime.now(vn_tz) - timedelta(hours=24)
                recent_24h = sum(
                    1 for t in log_df['Timestamp']
                    if datetime.strptime(str(t).strip(), "%Y-%m-%d %H:%M:%S").replace(tzinfo=vn_tz) >= cutoff
                )
            except Exception:
                recent_24h = total_events
            
            m1.metric("Total Events", total_events)
            m2.metric("Unique Users", unique_users)
            m3.metric("Top User", most_active)
            m4.metric("Last 24h", recent_24h)
            
            st.divider()
            
            # 2. Filters
            st.markdown("#### 🔍 Filter Logs")
            f_col1, f_col2, f_col3 = st.columns([2, 1, 1])
            
            with f_col1:
                search_query = st.text_input("Search details or users...", placeholder="e.g. login, user1, comparison")
            with f_col2:
                event_filter = st.selectbox("Event Type", ["All"] + list(log_df['Event Type'].unique()))
            with f_col3:
                row_limit = st.selectbox("Display Limit", [100, 500, 1000, "All"], index=1)

            # Applying filters
            filtered_df = log_df.copy()
            if search_query:
                filtered_df = filtered_df[
                    filtered_df['Details'].str.contains(search_query, case=False, na=False) | 
                    filtered_df['User'].str.contains(search_query, case=False, na=False)
                ]
            if event_filter != "All":
                filtered_df = filtered_df[filtered_df['Event Type'] == event_filter]
            
            if row_limit != "All":
                filtered_df = filtered_df.head(int(row_limit))

            # 3. Log Display
            st.markdown(f"Showing **{len(filtered_df)}** entries")
            st.dataframe(filtered_df, use_container_width=True, height=600)
            
            # 4. Actions
            output_log = io.BytesIO()
            with pd.ExcelWriter(output_log, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, index=False, sheet_name='FilteredLogs')
            
            st.download_button(
                label="📥 Export Current View to Excel",
                data=output_log.getvalue(),
                file_name=f"System_Log_Export_{datetime.now(vn_tz).strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.info("No system logs available yet.")

    # --- SUB-TAB: USER MANAGEMENT ---
    with admin_tabs[1]:
        st.markdown("### 👥 User Account Management")
        
        uc_col1, uc_col2 = st.columns([1, 2], gap="large")
        
        with uc_col1:
            st.markdown('<div class="app-card">', unsafe_allow_html=True)
            st.markdown("#### ➕ Create New Account")
            with st.form("add_user_form_v2", clear_on_submit=True):
                new_user = st.text_input("Username")
                new_pass = st.text_input("Password", type="password")
                new_role = st.selectbox("Role", ["user", "admin"])
                new_auto_fill = st.checkbox("Enable Auto-fill", value=False)
                if st.form_submit_button("Register User", use_container_width=True):
                    if new_user and new_pass:
                        success_add, msg_add = save_user(new_user, new_pass, new_role, new_auto_fill)
                        if success_add: st.success(msg_add)
                        else: st.error(msg_add)
            st.markdown('</div>', unsafe_allow_html=True)

        with uc_col2:
            st.markdown("#### 📋 Existing Users")
            current_users_list = load_users()
            
            for idx, user_data in enumerate(current_users_list):
                is_admin_account = (user_data["username"] == "admin")
                with st.expander(f"👤 {user_data['username']} — Role: {user_data['role'].upper()}", expanded=False):
                    e_col1, e_col2 = st.columns(2)
                    with e_col1:
                        edit_name = st.text_input("Username", value=user_data["username"], disabled=is_admin_account, key=f"v2_un_{idx}")
                        edit_pass = st.text_input("Password", value=user_data["password"], type="password", key=f"v2_pw_{idx}")
                    with e_col2:
                        edit_role = st.selectbox("Role", ["admin", "user"], index=0 if user_data["role"]=="admin" else 1, key=f"v2_rl_{idx}")
                        edit_auto = st.checkbox("Auto-fill", value=user_data.get("auto_fill", False), key=f"v2_af_{idx}")
                    
                    st.divider()
                    ec1, ec2 = st.columns(2)
                    with ec1:
                        if st.button("💾 Save Changes", key=f"v2_upd_{idx}", use_container_width=True):
                            success_upd, msg_upd = update_user_data(user_data["username"], edit_name, edit_pass, edit_role, edit_auto)
                            if success_upd: st.rerun()
                            else: st.error(msg_upd)
                    with ec2:
                        if is_admin_account:
                            st.button("🚫 System Admin", disabled=True, key=f"v2_del_dis_{idx}", use_container_width=True)
                        else:
                            if st.button("🗑️ Delete User", key=f"v2_del_{idx}", use_container_width=True):
                                remove_user(user_data["username"])
                                st.rerun()

# Footer
st.divider()
st.caption("Word Number Comparison | Streamlit Application")
