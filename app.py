import streamlit as pd_stream
import pandas as pd
from datetime import datetime, timedelta
import os
import time

# Set up the web page styling and title
pd_stream.set_page_config(page_title="IT Helpdesk System", layout="wide")

EXCEL_FILE = "it_tickets_database.xlsx"
# ✨ SCREENSHOT STORAGE CONFIGURATION (OPTION A)
SCREENSHOT_DIR = "ticket_screenshots"
if not os.path.exists(SCREENSHOT_DIR):
    os.makedirs(SCREENSHOT_DIR)

# Helper function to read data from Excel safely (handles all sheets)
def load_data(sheet_name="Tickets"):
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
            # Safe-guard: Ensure the Screenshot column exists if missing in older files
            if sheet_name in ["Tickets", "Archived_Tickets"] and not df.empty and "Screenshot Path" not in df.columns:
                df["Screenshot Path"] = "—"
            return df
        except Exception:
            if sheet_name == "Inventory":
                return create_default_inventory()
            elif sheet_name == "IT_Status":
                return pd.DataFrame([{"Current Status": "🟢 Online"}])
            elif sheet_name == "Departments":
                return pd.DataFrame({"Department Name": ["Administration", "Finance", "HR", "Operations", "Sales/Marketing", "Other"]})
            elif sheet_name == "Archived_Tickets":
                return create_default_tickets()
            return create_default_tickets()
    else:
        if sheet_name == "Inventory":
            return create_default_inventory()
        elif sheet_name == "IT_Status":
            return pd.DataFrame([{"Current Status": "🟢 Online"}])
        elif sheet_name == "Departments":
            return pd.DataFrame({"Department Name": ["Administration", "Finance", "HR", "Operations", "Sales/Marketing", "Other"]})
        elif sheet_name == "Archived_Tickets":
            return create_default_tickets()
        return create_default_tickets()

def create_default_tickets():
    columns = ["Ticket ID", "Date Created", "Date Resolved", "Employee Name", "Department", "Category", "Urgency", "Subject", "Description", "Status", "IT Notes", "Screenshot Path"]
    return pd.DataFrame(columns=columns)

def create_default_inventory():
    items = ["System Unit", "Monitor", "UPS", "Keyboard", "Mouse", "Headset", "Wi-Fi Dongle", "Power Cable", "VGA Cable", "HDMI Cable"]
    df_inv = pd.DataFrame({
        "Equipment Item": items,
        "Available Stocks": [0] * len(items),
        "Defective": [0] * len(items),
        "Currently Deployed": [0] * len(items),
        "Needed to be Disposed": [0] * len(items),
        "Total Stock": [0] * len(items)
    })
    return df_inv

# Clean helper function to save data safely without losing tabs
def save_data(df_to_save, sheet_name="Tickets"):
    items = ["System Unit", "Monitor", "UPS", "Keyboard", "Mouse", "Headset", "Wi-Fi Dongle", "Power Cable", "VGA Cable", "HDMI Cable"]
    default_inv = pd.DataFrame({"Equipment Item": items, "Available Stocks": [0]*10, "Defective": [0]*10, "Currently Deployed": [0]*10, "Needed to be Disposed": [0]*10, "Total Stock": [0]*10})
    default_users = pd.DataFrame(columns=["Username", "Password", "Role", "Full Name", "Department"])
    default_status = pd.DataFrame([{"Current Status": "🟢 Online"}])
    default_dept = pd.DataFrame({"Department Name": ["Administration", "Finance", "HR", "Operations", "Sales/Marketing", "Other"]})
    
    df_tickets = load_data(sheet_name="Tickets") if sheet_name != "Tickets" else df_to_save
    df_archive = load_data(sheet_name="Archived_Tickets") if sheet_name != "Archived_Tickets" else df_to_save
    
    if os.path.exists(EXCEL_FILE):
        try: df_inv = pd.read_excel(EXCEL_FILE, sheet_name="Inventory") if sheet_name != "Inventory" else df_to_save
        except Exception: df_inv = default_inv
            
        try: df_users = pd.read_excel(EXCEL_FILE, sheet_name="Users") if sheet_name != "Users" else df_to_save
        except Exception: df_users = default_users

        try: df_status = pd.read_excel(EXCEL_FILE, sheet_name="IT_Status") if sheet_name != "IT_Status" else df_to_save
        except Exception: df_status = default_status
            
        try: df_dept = pd.read_excel(EXCEL_FILE, sheet_name="Departments") if sheet_name != "Departments" else df_to_save
        except Exception: df_dept = default_dept
    else:
        df_inv = df_to_save if sheet_name == "Inventory" else default_inv
        df_users = df_to_save if sheet_name == "Users" else default_users
        df_status = df_to_save if sheet_name == "IT_Status" else default_status
        df_dept = df_to_save if sheet_name == "Departments" else default_dept

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_tickets.to_excel(writer, sheet_name="Tickets", index=False)
        df_archive.to_excel(writer, sheet_name="Archived_Tickets", index=False)
        df_inv.to_excel(writer, sheet_name="Inventory", index=False)
        df_users.to_excel(writer, sheet_name="Users", index=False)
        df_status.to_excel(writer, sheet_name="IT_Status", index=False)
        df_dept.to_excel(writer, sheet_name="Departments", index=False)

# --- AUTOMATIC MAINTENANCE ENGINE (15-Day Archive / 30-Day Delete) ---
def run_database_maintenance():
    df_active = load_data(sheet_name="Tickets")
    df_archive = load_data(sheet_name="Archived_Tickets")
    
    now = datetime.now()
    active_changed = False
    archive_changed = False
    
    if not df_active.empty and "Date Resolved" in df_active.columns:
        resolved_mask = (df_active["Status"] == "Resolved") & (df_active["Date Resolved"] != "—") & (df_active["Date Resolved"].notna())
        indices_to_archive = []
        for idx, row in df_active[resolved_mask].iterrows():
            try:
                res_date = datetime.strptime(str(row["Date Resolved"]).strip(), "%Y-%m-%d %H:%M")
                if (now - res_date) >= timedelta(days=15):
                    indices_to_archive.append(idx)
            except Exception:
                pass
        
        if indices_to_archive:
            rows_to_move = df_active.loc[indices_to_archive]
            df_archive = pd.concat([df_archive, rows_to_move], ignore_index=True)
            df_active = df_active.drop(indices_to_archive).reset_index(drop=True)
            active_changed = True
            archive_changed = True

    if not df_archive.empty and "Date Resolved" in df_archive.columns:
        indices_to_delete = []
        for idx, row in df_archive.iterrows():
            try:
                if str(row["Date Resolved"]) != "—" and pd.notna(row["Date Resolved"]):
                    res_date = datetime.strptime(str(row["Date Resolved"]).strip(), "%Y-%m-%d %H:%M")
                    if (now - res_date) >= timedelta(days=30):
                        indices_to_delete.append(idx)
            except Exception:
                pass
                
        if indices_to_delete:
            df_archive = df_archive.drop(indices_to_delete).reset_index(drop=True)
            archive_changed = True
            
    if active_changed or archive_changed:
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            df_active.to_excel(writer, sheet_name="Tickets", index=False)
            df_archive.to_excel(writer, sheet_name="Archived_Tickets", index=False)
            for sheet in ["Inventory", "Users", "IT_Status", "Departments"]:
                load_data(sheet_name=sheet).to_excel(writer, sheet_name=sheet, index=False)

run_database_maintenance()

# --- USER DATABASE ---
def load_users():
    credentials_dict = {}
    if os.path.exists(EXCEL_FILE):
        try:
            df_users = pd.read_excel(EXCEL_FILE, sheet_name="Users")
            for _, row in df_users.iterrows():
                username = str(row["Username"]).strip().lower()
                credentials_dict[username] = {
                    "password": str(row["Password"]),
                    "role": str(row["Role"]),
                    "name": str(row["Full Name"]),
                    "department": str(row.get("Department", "IT Department")) 
                }
        except Exception:
            pass
    return credentials_dict

# --- INITIALIZE TRACKING MEMORY (Session States) ---
if "logged_in" not in pd_stream.session_state:
    pd_stream.session_state["logged_in"] = False
if "user_role" not in pd_stream.session_state:
    pd_stream.session_state["user_role"] = None
if "logged_in_user" not in pd_stream.session_state:
    pd_stream.session_state["logged_in_user"] = None
if "logged_in_username" not in pd_stream.session_state:
    pd_stream.session_state["logged_in_username"] = None
if "user_dept" not in pd_stream.session_state:
    pd_stream.session_state["user_dept"] = "Other"
if "it_status" not in pd_stream.session_state:
    pd_stream.session_state["it_status"] = "🟢 Online"
# ✨ TRACKING EDITING STATE FOR TICKET MODIFICATIONS
if "is_editing_ticket" not in pd_stream.session_state:
    pd_stream.session_state["is_editing_ticket"] = False
if "last_selected_ticket" not in pd_stream.session_state:
    pd_stream.session_state["last_selected_ticket"] = None

if os.path.exists(EXCEL_FILE):
    try:
        df_status = pd.read_excel(EXCEL_FILE, sheet_name="IT_Status")
        if not df_status.empty:
            pd_stream.session_state["it_status"] = str(df_status.iloc[0]["Current Status"])
    except Exception:
        pass

# --- LOGIN SCREEN INTERFACE ---
if not pd_stream.session_state["logged_in"]:
    pd_stream.sidebar.title("🔒 System Login")
    input_user = pd_stream.sidebar.text_input("Username").strip().lower()
    input_pass = pd_stream.sidebar.text_input("Password", type="password")
    
    if pd_stream.sidebar.button("Login"):
        current_credentials = load_users()
        if input_user in current_credentials and current_credentials[input_user]["password"] == input_pass:
            pd_stream.session_state["logged_in"] = True
            pd_stream.session_state["user_role"] = current_credentials[input_user]["role"]
            pd_stream.session_state["logged_in_user"] = current_credentials[input_user]["name"]
            pd_stream.session_state["logged_in_username"] = input_user
            pd_stream.session_state["user_dept"] = current_credentials[input_user]["department"]
            pd_stream.rerun()
        else:
            pd_stream.sidebar.error("❌ Invalid Username or Password")
    pd_stream.stop()

# --- IF LOGGED IN: SETUP COMPANION SIDEBAR ---
user_role = pd_stream.session_state["user_role"]
current_name = pd_stream.session_state["logged_in_user"]
active_username = pd_stream.session_state["logged_in_username"]

pd_stream.sidebar.title(f"👤 Welcome, {current_name}!")

if pd_stream.sidebar.button("🔄 Sync System Data", use_container_width=True):
    run_database_maintenance()
    pd_stream.toast("Data synchronized with Excel server!", icon="📥")
    time.sleep(0.5)
    pd_stream.rerun()

if user_role == "IT Department":
    pd_stream.sidebar.markdown("---")
    pd_stream.sidebar.subheader("📢 Availability Status")
    status_options = ["🟢 Online", "🟡 Away (Field Work)", "🔴 Offline (Closed)"]
    current_idx = status_options.index(pd_stream.session_state["it_status"]) if pd_stream.session_state["it_status"] in status_options else 0
    chosen_status = pd_stream.sidebar.selectbox("Set Your Live Status:", status_options, index=current_idx)
    
    if chosen_status != pd_stream.session_state["it_status"]:
        pd_stream.session_state["it_status"] = chosen_status
        df_status_update = pd.DataFrame([{"Current Status": chosen_status}])
        save_data(df_status_update, sheet_name="IT_Status")
        pd_stream.toast(f"Status saved as {chosen_status}!", icon="💾")

# --- SIDEBAR SECURITY / PASSWORD MANAGEMENT PANEL ---
pd_stream.sidebar.markdown("---")
with pd_stream.sidebar.expander("🔑 Security Settings", expanded=False):
    with pd_stream.form(key="sidebar_password_form", clear_on_submit=True):
        current_pwd_input = pd_stream.text_input("Current Password", type="password")
        new_pwd_input = pd_stream.text_input("New Password", type="password")
        confirm_pwd_input = pd_stream.text_input("Confirm New Password", type="password")
        
        if pd_stream.form_submit_button("Update Password", use_container_width=True):
            current_roster_data = load_users()
            if not current_pwd_input or not new_pwd_input or not confirm_pwd_input:
                pd_stream.error("All password fields are required.")
            elif current_roster_data[active_username]["password"] != current_pwd_input:
                pd_stream.error("Incorrect current password.")
            elif new_pwd_input != confirm_pwd_input:
                pd_stream.error("New passwords do not match.")
            elif len(new_pwd_input) < 4:
                pd_stream.error("Password must be at least 4 characters.")
            else:
                if os.path.exists(EXCEL_FILE):
                    try:
                        df_u_temp = pd.read_excel(EXCEL_FILE, sheet_name="Users")
                        df_u_temp['Username'] = df_u_temp['Username'].astype(str).str.strip().str.lower()
                        match_idx = df_u_temp[df_u_temp["Username"] == active_username].index
                        if not match_idx.empty:
                            df_u_temp.at[match_idx[0], "Password"] = str(new_pwd_input)
                            save_data(df_u_temp, sheet_name="Users")
                            pd_stream.toast("Password updated successfully!", icon="🔒")
                            time.sleep(0.6)
                            pd_stream.rerun()
                        else:
                            pd_stream.error("System configuration context mismatch.")
                    except Exception as e:
                        pd_stream.error(f"Error accessing database server: {e}")

pd_stream.sidebar.markdown("---")
if pd_stream.sidebar.button("Log Out", use_container_width=True):
    pd_stream.session_state["logged_in"] = False
    pd_stream.session_state["user_role"] = None
    pd_stream.session_state["logged_in_user"] = None
    pd_stream.session_state["logged_in_username"] = None
    pd_stream.session_state["user_dept"] = "Other"
    pd_stream.rerun()
    
# ---------------------------------------------------------
# PAGE 1: EMPLOYEE PORTAL (Ticket Submission & History)
# ---------------------------------------------------------
if user_role == "Employee Portal":
    current_dept = pd_stream.session_state.get("user_dept", "Other")
    it_presence = pd_stream.session_state["it_status"]
    
    if "Online" in it_presence:
        pd_stream.success(f"💻 **IT Support Status:** {it_presence} (We are active and ready to help!)")
    elif "Away" in it_presence:
        pd_stream.warning(f"⏳ **IT Support Status:** {it_presence} (Expect slight delays in response)")
    else:
        pd_stream.error(f"🔴 **IT Support Status:** {it_presence} (IT Office is currently closed)")

    pd_stream.title("📋 Submit a Ticket / Assistance Request")
    pd_stream.write(f"Logged in as: **{current_name}** | Department: **{current_dept}**")
    
    with pd_stream.form(key="ticket_form", clear_on_submit=True):
        col1, col2 = pd_stream.columns(2)
        with col1:
            pd_stream.text_input("Your Name", value=current_name, disabled=True)
            pd_stream.text_input("Department", value=current_dept, disabled=True)
        with col2:
            category = pd_stream.selectbox("What do you need help with?", ["Hardware (PC, Monitor, Printer)", "Software / Access Issues", "Wi-Fi & Internet", "General Inquiry"])
            urgency = pd_stream.select_slider("Urgency Level", options=["Low", "Medium", "High"])
            
        subject = pd_stream.text_input("Subject / Short Summary *")
        description = pd_stream.text_area("Detailed Description of the problem *")
        
        # ✨ COMPONENT ADDED: EMPLOYEE SCREENSHOT FILE UPLOAD CONTAINER
        uploaded_screenshot = pd_stream.file_uploader("Upload an Error Screenshot (Optional)", type=["png", "jpg", "jpeg"])
        
        submit_button = pd_stream.form_submit_button(label="Submit Ticket")
        
        if submit_button:
            if not subject or not description:
                pd_stream.error("⚠️ Please fill out all required fields (*)")
            else:
                df = load_data()
                ticket_id = f"IT-{datetime.now().strftime('%m%d-%H%M%S')}"
                date_str = datetime.now().strftime("%Y-%m-%d %H:%M")
                
                # Handling File IO Extraction & Saving locally under Option A framework
                saved_img_path = "—"
                if uploaded_screenshot is not None:
                    file_ext = os.path.splitext(uploaded_screenshot.name)[1]
                    unique_filename = f"{ticket_id}{file_ext}"
                    full_destination_path = os.path.join(SCREENSHOT_DIR, unique_filename)
                    
                    with open(full_destination_path, "wb") as file_stream:
                        file_stream.write(uploaded_screenshot.getbuffer())
                    saved_img_path = full_destination_path
                
                new_ticket = {
                    "Ticket ID": ticket_id, "Date Created": date_str, "Date Resolved": "—", "Employee Name": current_name, 
                    "Department": current_dept, "Category": category, "Urgency": urgency, 
                    "Subject": subject, "Description": description, "Status": "Pending", "IT Notes": "",
                    "Screenshot Path": saved_img_path
                }
                
                for col in df.columns:
                    if col not in new_ticket:
                        new_ticket[col] = ""
                        
                df = pd.concat([df, pd.DataFrame([new_ticket])], ignore_index=True)
                save_data(df)
                pd_stream.toast(f"🎉 Ticket submitted! ID: {ticket_id}", icon="🚀")
                time.sleep(0.6)
                pd_stream.rerun()

    pd_stream.divider()
    pd_stream.subheader("📂 Your Ticket History")
    df = load_data()
    user_tickets = df[df["Employee Name"] == current_name]
    
    if not user_tickets.empty:
        pd_stream.write(f"Showing updates for tickets submitted by **{current_name}**:")
        display_cols = ["Ticket ID", "Date Created", "Category", "Status", "IT Notes"]
        if "Date Resolved" in user_tickets.columns:
            display_cols.insert(2, "Date Resolved")
        pd_stream.dataframe(user_tickets[display_cols], use_container_width=True, hide_index=True)
    else:
        pd_stream.info("No active or past tickets found under your account.")

# ---------------------------------------------------------
# PAGE 2: IT DEPARTMENT DASHBOARD (Management Panel)
# ---------------------------------------------------------
elif user_role == "IT Department":
    pd_stream.title("🖥️ IT Central Command Dashboard")
    df_metrics = load_data(sheet_name="Tickets")
    
    with pd_stream.expander("📊 Operational Metrics & Data Visualizations", expanded=True):
        if df_metrics.empty:
            pd_stream.info("Analytics engine waiting for ticketing logs to map metrics charts...")
        else:
            m_col1, m_col2 = pd_stream.columns(2)
            with m_col1:
                pd_stream.markdown("##### 🏢 Ticket Frequency By Corporate Department")
                if "Department" in df_metrics.columns and not df_metrics["Department"].dropna().empty:
                    dept_counts = df_metrics["Department"].value_counts()
                    pd_stream.bar_chart(dept_counts)
                else:
                    pd_stream.info("No department logs logged to process distribution charts.")
            with m_col2:
                pd_stream.markdown("##### ⚠️ Ticket Distribution By Urgency Level")
                if "Urgency" in df_metrics.columns and not df_metrics["Urgency"].dropna().empty:
                    urgency_counts = df_metrics["Urgency"].value_counts()
                    pd_stream.bar_chart(urgency_counts)
                else:
                    pd_stream.info("No urgency metadata records logged to analyze breakdown distributions.")
                
    pd_stream.write("") 
    tab1, tab2, tab3 = pd_stream.tabs(["🎟️ Ticket Management", "📦 Inventory Tracking", "👥 User Access Control"])

    # TAB 1: TICKET MANAGEMENT
    with tab1:
        df_tickets = load_data(sheet_name="Tickets")
        
        if df_tickets.empty:
            pd_stream.info("No tickets submitted yet! Everything is quiet... for now. ☕")
        else:
            with pd_stream.expander("🎟️ Active Helpdesk Tickets Queue", expanded=True):
                pending_count = len(df_tickets[df_tickets["Status"] == "Pending"])
                resolved_count = len(df_tickets[df_tickets["Status"] == "Resolved"])
                
                c1, c2, c3 = pd_stream.columns(3)
                c1.metric("Total Received", len(df_tickets))
                c2.metric("Pending Attention", pending_count, delta=pending_count, delta_color="inverse")
                c3.metric("Resolved Tickets", resolved_count)
                
                pd_stream.markdown("---")
                f_col1, f_col2 = pd_stream.columns(2)
                with f_col1:
                    status_filter = pd_stream.multiselect("Filter by Status:", ["Pending", "In Progress", "Resolved"], default=[])
                with f_col2:
                    urgency_filter = pd_stream.multiselect("Filter by Urgency:", ["Low", "Medium", "High"], default=[])
                
                filtered_df = df_tickets.copy()
                if status_filter:
                    filtered_df = filtered_df[filtered_df["Status"].isin(status_filter)]
                if urgency_filter:
                    filtered_df = filtered_df[filtered_df["Urgency"].isin(urgency_filter)]
                
                # Dropping path string formatting raw values on the macro viewer to keep output matrix tidy
                clean_view_df = filtered_df.copy()
                if "Screenshot Path" in clean_view_df.columns:
                    clean_view_df = clean_view_df.drop(columns=["Screenshot Path"])
                pd_stream.dataframe(clean_view_df, use_container_width=True, hide_index=True)
            
            # 🔄 RESPOND & UPDATE ENGINE WITH STATE CONTROL
            with pd_stream.expander("✍️ Respond & Update Ticket Status", expanded=True):
                target_id = pd_stream.selectbox("Select Ticket ID to Modify:", df_tickets["Ticket ID"].unique())
                
                # Reset editing state if user switches to a completely different ticket ID
                if pd_stream.session_state["last_selected_ticket"] != target_id:
                    pd_stream.session_state["is_editing_ticket"] = False
                    pd_stream.session_state["last_selected_ticket"] = target_id
                
                selected_row = df_tickets[df_tickets["Ticket ID"] == target_id].iloc[0]
                current_status_val = selected_row["Status"]
                current_notes_val = str(selected_row["IT Notes"]) if pd.notna(selected_row["IT Notes"]) else ""
                
                # Fetching screenshot file route details safely
                screenshot_file_link = str(selected_row.get("Screenshot Path", "—"))
                
                status_options = ["Pending", "In Progress", "Resolved"]
                status_idx = status_options.index(current_status_val) if current_status_val in status_options else 0
                
                # Dynamic visual indicators letting IT know if fields are locked or unlocked
                if not pd_stream.session_state["is_editing_ticket"]:
                    pd_stream.info(f"🔒 Fields are locked. Click **'✏️ Edit Ticket Details'** below to modify Ticket **{target_id}**.")
                else:
                    pd_stream.success(f"🔓 Fields unlocked. Modifying Ticket **{target_id}**... Click **'💾 Save Updates'** when finished.")

                # Inputs render disabled or enabled based on editing state
                status_update = pd_stream.selectbox("Set New Status:", status_options, index=status_idx, disabled=not pd_stream.session_state["is_editing_ticket"])
                notes_update = pd_stream.text_area("IT Actions Taken / Internal Notes:", value=current_notes_val, disabled=not pd_stream.session_state["is_editing_ticket"])
                
                # ✨ COMPONENT ADDED: LIVE SCREENSHOT VIEWER INTERFACE FOR IT OPERATIONS
                pd_stream.markdown("##### 🖼️ Attached Screenshot File")
                if screenshot_file_link != "—" and os.path.exists(screenshot_file_link):
                    pd_stream.image(screenshot_file_link, caption=f"Screenshot for Ticket {target_id}", use_container_width=False, width=600)
                else:
                    pd_stream.text("ℹ️ (No screenshot attached for this ticket)")
                pd_stream.markdown("---")

                # Stateful Action Buttons
                if not pd_stream.session_state["is_editing_ticket"]:
                    if pd_stream.button("✏️ Edit Ticket Details", use_container_width=True):
                        pd_stream.session_state["is_editing_ticket"] = True
                        pd_stream.rerun()
                else:
                    col_btn1, col_btn2 = pd_stream.columns(2)
                    with col_btn1:
                        if pd_stream.button("💾 Save Updates", type="primary", use_container_width=True):
                            idx = df_tickets[df_tickets["Ticket ID"] == target_id].index[0]
                            
                            df_tickets["Status"] = df_tickets["Status"].astype(str)
                            df_tickets["IT Notes"] = df_tickets["IT Notes"].astype(str)
                            if "Date Resolved" not in df_tickets.columns:
                                df_tickets["Date Resolved"] = "—"
                            df_tickets["Date Resolved"] = df_tickets["Date Resolved"].astype(str)
                            
                            if status_update == "Resolved":
                                df_tickets.at[idx, "Date Resolved"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                            else:
                                df_tickets.at[idx, "Date Resolved"] = "—"
                            
                            df_tickets.at[idx, "Status"] = str(status_update)
                            df_tickets.at[idx, "IT Notes"] = str(notes_update)
                            
                            save_data(df_tickets, sheet_name="Tickets")
                            
                            # Lock fields down and trigger success flow
                            pd_stream.session_state["is_editing_ticket"] = False
                            pd_stream.toast(f"Ticket {target_id} updated successfully!", icon="✅")
                            time.sleep(0.6)
                            pd_stream.rerun()
                    with col_btn2:
                        if pd_stream.button("❌ Cancel Changes", use_container_width=True):
                            pd_stream.session_state["is_editing_ticket"] = False
                            pd_stream.rerun()

    # TAB 2: INVENTORY TRACKING
    with tab2:
        with pd_stream.expander("📦 Corporate IT Hardware Inventory Registry", expanded=True):
            df_inv = load_data(sheet_name="Inventory")
            pd_stream.dataframe(df_inv, use_container_width=True, hide_index=True)
        
        with pd_stream.expander("🔄 Adjust Asset Volumes & Stock Counts", expanded=True):
            # ✨ FIX PART 1: Kept the item dropdown outside the form so updates trigger a clean component data read
            target_item = pd_stream.selectbox("Select Equipment Profile to Adjust:", df_inv["Equipment Item"].unique())
            
            # ✨ FIX PART 2: Read current row values from the Excel dataframe based on selection
            inv_row = df_inv[df_inv["Equipment Item"] == target_item].iloc[0]
            curr_avail = int(inv_row.get("Available Stocks", 0))
            curr_defective = int(inv_row.get("Defective", 0))
            curr_deployed = int(inv_row.get("Currently Deployed", 0))
            curr_disposed = int(inv_row.get("Needed to be Disposed", 0))
            
            # ✨ FIX PART 3: The form now serves exclusively to lock number values and safely execute updates on click
            with pd_stream.form(key="inventory_mod_form", clear_on_submit=False):
                col_inv1, col_inv2 = pd_stream.columns(2)
                with col_inv1:
                    # Dynamically pass the existing counts as the default value parameter
                    avail = pd_stream.number_input("Available Stock (In Storage)", min_value=0, value=curr_avail)
                    defective = pd_stream.number_input("Defective / Damaged Units", min_value=0, value=curr_defective)
                with col_inv2:
                    deployed = pd_stream.number_input("Currently Deployed / In Use", min_value=0, value=curr_deployed)
                    disposed = pd_stream.number_input("Marked for Disposal / Scrap", min_value=0, value=curr_disposed)
                    
                if pd_stream.form_submit_button("Commit Stock Adjustments", use_container_width=True):
                    inv_idx = df_inv[df_inv["Equipment Item"] == target_item].index[0]
                    df_inv.at[inv_idx, "Available Stocks"] = avail
                    df_inv.at[inv_idx, "Defective"] = defective
                    df_inv.at[inv_idx, "Currently Deployed"] = deployed
                    df_inv.at[inv_idx, "Needed to be Disposed"] = disposed
                    df_inv.at[inv_idx, "Total Stock"] = avail + defective + deployed + disposed
                    
                    save_data(df_inv, sheet_name="Inventory")
                    pd_stream.toast(f"✅ Stock successfully modified for {target_item}!", icon="📦")
                    time.sleep(0.6)
                    pd_stream.rerun()

    # TAB 3: USER ACCESS CONTROL & DYNAMIC DEPARTMENTS
    with tab3:
        df_dept_list = load_data(sheet_name="Departments")
        
        with pd_stream.expander("🏢 Corporate Department Management Engine", expanded=False):
            d_col1, d_col2 = pd_stream.columns([3, 2])
            with d_col1:
                pd_stream.dataframe(df_dept_list, use_container_width=True, hide_index=True)
            with d_col2:
                new_dept_input = pd_stream.text_input("New Department Name:")
                if pd_stream.button("➕ Add Department", use_container_width=True):
                    if new_dept_input and new_dept_input not in df_dept_list["Department Name"].values:
                        new_row = pd.DataFrame({"Department Name": [new_dept_input]})
                        df_dept_list = pd.concat([df_dept_list, new_row], ignore_index=True)
                        save_data(df_dept_list, sheet_name="Departments")
                        pd_stream.toast(f"Added {new_dept_input}!", icon="🏢")
                        time.sleep(0.5)
                        pd_stream.rerun()
                
                remove_dept_select = pd_stream.selectbox("Select Department to Remove:", df_dept_list["Department Name"].unique())
                if pd_stream.button("❌ Remove Department", use_container_width=True):
                    df_dept_list = df_dept_list[df_dept_list["Department Name"] != remove_dept_select]
                    save_data(df_dept_list, sheet_name="Departments")
                    pd_stream.toast(f"Removed {remove_dept_select}!", icon="🗑️")
                    time.sleep(0.5)
                    pd_stream.rerun()

        current_roster = load_users()
        roster_list = []
        for uname, details in current_roster.items():
            roster_list.append({
                "Username": uname, 
                "Full Name / Identifier": details["name"], 
                "Assigned System Role": details["role"],
                "Department": details.get("department", "IT Department")
            })
            
        with pd_stream.expander("📋 Registered System Users Directory", expanded=True):
            pd_stream.dataframe(pd.DataFrame(roster_list), use_container_width=True, hide_index=True)
        
        with pd_stream.expander("❌ Remove User Account Profile (Offboarding)", expanded=False):
            removable_users = [uname for uname in current_roster.keys() if current_roster[uname]['name'] != current_name]
            
            if not removable_users:
                pd_stream.info("No other user profiles are currently registered to remove.")
            else:
                with pd_stream.form(key="user_removal_form", clear_on_submit=True):
                    user_to_remove = pd_stream.selectbox(
                        "Select Username to Terminate / Delete:", 
                        options=removable_users,
                        format_func=lambda x: f"{x} ({current_roster[x]['name']} - {current_roster[x]['department']})"
                    )
                    confirm_delete = pd_stream.checkbox("⚠️ I confirm that I want to permanently delete this user account profile.")
                    
                    if pd_stream.form_submit_button("Terminate Account"):
                        if not confirm_delete:
                            pd_stream.error("Please check the confirmation box before deleting an account.")
                        else:
                            if os.path.exists(EXCEL_FILE):
                                try:
                                    df_u = pd.read_excel(EXCEL_FILE, sheet_name="Users")
                                    df_u['Username'] = df_u['Username'].astype(str).str.strip().str.lower()
                                    df_u = df_u[df_u["Username"] != user_to_remove]
                                    
                                    try: df_t = pd.read_excel(EXCEL_FILE, sheet_name="Tickets")
                                    except Exception: df_t = create_default_tickets()
                                    try: df_i = pd.read_excel(EXCEL_FILE, sheet_name="Inventory")
                                    except Exception: df_i = create_default_inventory()
                                    try: df_s = pd.read_excel(EXCEL_FILE, sheet_name="IT_Status")
                                    except Exception: df_s = pd.DataFrame([{"Current Status": "🟢 Online"}])
                                    
                                    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                                        df_t.to_excel(writer, sheet_name="Tickets", index=False)
                                        df_i.to_excel(writer, sheet_name="Inventory", index=False)
                                        df_u.to_excel(writer, sheet_name="Users", index=False)
                                        df_s.to_excel(writer, sheet_name="IT_Status", index=False)
                                        df_dept_list.to_excel(writer, sheet_name="Departments", index=False)
                                        
                                    pd_stream.toast(f"Wiped account '{user_to_remove}'", icon="🔒")
                                    time.sleep(0.6)
                                    pd_stream.rerun()
                                except Exception as e:
                                    pd_stream.error(f"Error accessing records: {e}")

        with pd_stream.expander("➕ Provision New Account Profile", expanded=False):
            with pd_stream.form(key="user_provision_form", clear_on_submit=True):
                reg_name = pd_stream.text_input("Full Name (e.g., 'K. Ramos' or 'Tech Specialist') *")
                reg_user = pd_stream.text_input("Desired Username (Lowercase, no spaces) *").strip().lower()
                reg_pass = pd_stream.text_input("Assign Access Password *")
                reg_role = pd_stream.selectbox("Assigned Interface Clearance Level:", ["Employee Portal", "IT Department"])
                reg_dept = pd_stream.selectbox("Assign User's Department:", df_dept_list["Department Name"].unique())
                
                if pd_stream.form_submit_button("Provision Account"):
                    if not reg_name or not reg_user or not reg_pass:
                        pd_stream.error("⚠️ All fields are mandatory.")
                    elif reg_user in current_roster:
                        pd_stream.error(f"❌ Account identity profile '{reg_user}' already exists.")
                    else:
                        if os.path.exists(EXCEL_FILE):
                            try: df_u = pd.read_excel(EXCEL_FILE, sheet_name="Users")
                            except Exception: df_u = pd.DataFrame(columns=["Username", "Password", "Role", "Full Name", "Department"])
                        else:
                            df_u = pd.DataFrame(columns=["Username", "Password", "Role", "Full Name", "Department"])
                            
                        new_profile = {
                            "Username": reg_user, "Password": reg_pass, "Role": reg_role, "Full Name": reg_name, "Department": reg_dept
                        }
                        df_u = pd.concat([df_u, pd.DataFrame([new_profile])], ignore_index=True)
                        
                        if os.path.exists(EXCEL_FILE):
                            try: df_t = pd.read_excel(EXCEL_FILE, sheet_name="Tickets")
                            except Exception: df_t = create_default_tickets()
                            try: df_i = pd.read_excel(EXCEL_FILE, sheet_name="Inventory")
                            except Exception: df_i = create_default_inventory()
                            try: df_s = pd.read_excel(EXCEL_FILE, sheet_name="IT_Status")
                            except Exception: df_s = pd.DataFrame([{"Current Status": "🟢 Online"}])
                        else:
                            df_i = create_default_inventory()
                            df_t = create_default_tickets()
                            df_s = pd.DataFrame([{"Current Status": "🟢 Online"}])
                            
                        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                            df_t.to_excel(writer, sheet_name="Tickets", index=False)
                            df_i.to_excel(writer, sheet_name="Inventory", index=False)
                            df_u.to_excel(writer, sheet_name="Users", index=False)
                            df_s.to_excel(writer, sheet_name="IT_Status", index=False)
                            df_dept_list.to_excel(writer, sheet_name="Departments", index=False)
                            
                        pd_stream.toast(f"👤 Provisioned account {reg_user}!", icon="🎉")
                        time.sleep(0.6)
                        pd_stream.rerun()
