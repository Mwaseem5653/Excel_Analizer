import streamlit as st
import json
import os
import datetime

# Define the path for the users JSON file
USERS_FILE = "users.json"
EYECON_LOGS_FILE = "eyecon_logs.json"

def load_users():
    """Loads users from the JSON file. If the file doesn't exist, creates it with a default admin."""
    if not os.path.exists(USERS_FILE):
        default_users = {
            "jerrynaeem19@gmail.com": {
                "password": "Pakistan1122@",
                "tokens": 100, # Default tokens for admin
                "eyecon_tokens": 0,
                "services": [
                    "Application Extractor",
                    "Excel Analyzer",
                    "PTA Services",
                    "CDR Format",
                    "Admin",
                    "Vehicle and Mobile"
                ]
            },
            
        }
        with open(USERS_FILE, 'w') as f:
            json.dump(default_users, f, indent=4)
        return default_users
    else:
        with open(USERS_FILE, 'r') as f:
            users = json.load(f)
            # Migration: Ensure all users have a 'tokens', 'eyecon_tokens', 'eyecon_consumed' and 'eyecon_pool' field
            is_modified = False
            for email in users:
                if "tokens" not in users[email]:
                    users[email]["tokens"] = 0
                    is_modified = True
                if "eyecon_tokens" not in users[email]:
                    users[email]["eyecon_tokens"] = 0
                    is_modified = True
                if "eyecon_consumed" not in users[email]:
                    users[email]["eyecon_consumed"] = 0
                    is_modified = True
                if "eyecon_pool" not in users[email]:
                    users[email]["eyecon_pool"] = 0
                    is_modified = True
            
            if is_modified:
                with open(USERS_FILE, 'w') as f:
                    json.dump(users, f, indent=4)
            return users

def save_users():
    """Saves the current USERS dictionary to the JSON file."""
    with open(USERS_FILE, 'w') as f:
        json.dump(USERS, f, indent=4)

# Load users at the start of the application
USERS = load_users()

def get_eyecon_logs():
    """Loads eyecon logs from JSON."""
    if not os.path.exists(EYECON_LOGS_FILE):
        return []
    try:
        with open(EYECON_LOGS_FILE, 'r') as f:
            return json.load(f)
    except:
        return []

def log_eyecon_transaction(admin_email, target_user, amount):
    """Logs the issuance of eyecon tokens."""
    logs = get_eyecon_logs()
    entry = {
        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "admin": admin_email,
        "user": target_user,
        "amount": amount
    }
    logs.append(entry)
    with open(EYECON_LOGS_FILE, 'w') as f:
        json.dump(logs, f, indent=4)

def login():
    """Displays the login form."""
    st.title("Login")
    email = st.text_input("Email")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if email in USERS and USERS[email]["password"] == password:
            st.session_state["logged_in"] = True
            st.session_state["email"] = email
            st.success("Logged in successfully!")
            st.rerun()
        else:
            st.error("Invalid email or password")

def is_logged_in():
    """Checks if the user is logged in."""
    return st.session_state.get("logged_in", False)

def get_user_services():
    """Returns the list of services for the logged-in user."""
    if is_logged_in():
        email = st.session_state["email"]
        return USERS.get(email, {}).get("services", [])
    return []

def get_tokens(email=None):
    """Returns the token balance for a user (current user if email is None)."""
    if email is None:
        if is_logged_in():
            email = st.session_state["email"]
        else:
            return 0
    return USERS.get(email, {}).get("tokens", 0)

def get_eyecon_tokens(email=None):
    """Returns the eyecon token balance for a user."""
    if email is None:
        if is_logged_in():
            email = st.session_state["email"]
        else:
            return 0
    return USERS.get(email, {}).get("eyecon_tokens", 0)

def add_tokens(email, amount):
    """Adds tokens to a user's balance."""
    if email in USERS:
        USERS[email]["tokens"] = USERS[email].get("tokens", 0) + amount
        save_users()
        return True
    return False

def get_eyecon_pool(email=None):
    """Returns the eyecon pool balance for a user (usually admin)."""
    if email is None:
        if is_logged_in():
            email = st.session_state["email"]
        else:
            return 0
    return USERS.get(email, {}).get("eyecon_pool", 0)

def add_eyecon_tokens(target_email, amount, admin_email="System"):
    """Transfers eyecon tokens from Admin's pool to Target's balance and logs it."""
    
    # 1. Check if Admin has enough in Pool
    admin_pool = USERS.get(admin_email, {}).get("eyecon_pool", 0)
    
    # If admin is "System" (e.g. initial setup), we might skip check, but here admin is a user.
    # We assume the logged-in user (admin_email) is performing the action.
    
    if admin_email in USERS:
        if admin_pool < amount:
            return "Insufficient Pool Balance"
        
        # Deduct from Pool
        USERS[admin_email]["eyecon_pool"] = admin_pool - amount
    else:
        # Fallback if system/unknown admin
        pass

    # 2. Add to Target
    if target_email in USERS:
        USERS[target_email]["eyecon_tokens"] = USERS[target_email].get("eyecon_tokens", 0) + amount
        save_users()
        log_eyecon_transaction(admin_email, target_email, amount)
        return "Success"
    
    return "User not found"

def deduct_tokens(amount, email=None):
    """Deducts tokens from a user's balance if sufficient funds exist."""
    if email is None:
        if is_logged_in():
            email = st.session_state["email"]
        else:
            return False
            
    current_tokens = USERS.get(email, {}).get("tokens", 0)
    if current_tokens >= amount:
        USERS[email]["tokens"] = current_tokens - amount
        save_users()
        return True
    return False

def deduct_eyecon_tokens(amount, email=None):
    """Deducts eyecon tokens from a user's balance."""
    if email is None:
        if is_logged_in():
            email = st.session_state["email"]
        else:
            return False
            
    current = USERS.get(email, {}).get("eyecon_tokens", 0)
    if current >= amount:
        USERS[email]["eyecon_tokens"] = current - amount
        USERS[email]["eyecon_consumed"] = USERS[email].get("eyecon_consumed", 0) + amount
        save_users()
        return True
    return False

def logout():
    """Logs the user out and clears the session."""
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

def admin_section():
    """Displays the admin panel for user management."""
    st.title("Admin Section")
    
    # --- Eyecon Pool Stats ---
    admin_email = st.session_state.get("email")
    pool_balance = get_eyecon_pool(admin_email)
    
    # Calculate Total Issued from Logs or Current Balances
    # Let's use current balances for "Active Issued"
    total_active_issued = sum(u.get("eyecon_tokens", 0) for u in USERS.values())
    
    st.markdown("### 🏦 Eyecon Token Stats")
    col_stat1, col_stat2 = st.columns(2)
    with col_stat1:
        st.metric("🎱 Admin Pool (Remaining)", pool_balance)
    with col_stat2:
        st.metric("📤 Total Active Issued", total_active_issued)
    st.write("---")

    st.subheader("Manage Users and Permissions")

    col1, col2 = st.columns(2)

    # --- Token Management ---
    with col1:
        st.write("### 💰 Issue General Tokens")
        with st.form("issue_tokens_form"):
            user_to_issue = st.selectbox("Select User", list(USERS.keys()))
            tokens_to_add = st.number_input("Tokens Amount", min_value=1, value=100, step=10)
            
            issue_submitted = st.form_submit_button("Issue Tokens")
            if issue_submitted:
                if add_tokens(user_to_issue, tokens_to_add):
                    st.success(f"Successfully added {tokens_to_add} tokens to {user_to_issue}")
                    st.rerun()
                else:
                    st.error("Failed to add tokens.")
    
    # --- Eyecon Token Management ---
    with col2:
        st.write("### 👁️ Issue Eyecon Tokens")
        # Subheading removed as requested
        with st.form("issue_eyecon_tokens_form"):
            e_user_to_issue = st.selectbox("Select User (Eyecon)", list(USERS.keys()))
            e_tokens_to_add = st.number_input("Eyecon Tokens Amount", min_value=1, value=100, step=10)
            
            e_issue_submitted = st.form_submit_button("Issue Eyecon Tokens")
            if e_issue_submitted:
                result = add_eyecon_tokens(e_user_to_issue, e_tokens_to_add, admin_email)
                if result == "Success":
                    st.success(f"Successfully issued {e_tokens_to_add} Eyecon tokens to {e_user_to_issue}")
                    st.rerun()
                elif result == "Insufficient Pool Balance":
                    st.error(f"❌ Insufficient Pool Balance! You have {pool_balance}, trying to issue {e_tokens_to_add}.")
                else:
                    st.error("Failed to add Eyecon tokens.")

    st.write("---")

    # Display current users
    st.write("### User Management")
    if st.button("👥 View/Refresh All Users"):
        st.session_state.show_users = not st.session_state.get('show_users', False)

    if st.session_state.get('show_users'):
        for email, data in USERS.items():
            st.write(f"**Email:** {email}")
            st.write(f"**Tokens:** {data.get('tokens', 0)} | **Eyecon Tokens:** {data.get('eyecon_tokens', 0)}")
            st.write(f"**Eyecon Consumed:** {data.get('eyecon_consumed', 0)}")
            st.write(f"**Services:** {', '.join(data['services'])}")
            st.write("---")

    # Eyecon Logs
    st.write("### 📜 Eyecon Token Logs")
    logs = get_eyecon_logs()
    
    if logs:
        # 1. Summarize by User
        summary = {}
        for entry in logs:
            user = entry.get("user")
            amount = entry.get("amount", 0)
            if user:
                summary[user] = summary.get(user, 0) + amount
        
        # Display Summary
        st.write("#### Issued Tokens Summary (Per User)")
        if summary:
            st.table([{"User": u, "Total Issued": t} for u, t in summary.items()])
        
        # Display Detailed Logs
        st.write("#### Detailed Logs")
        st.dataframe(logs)

        # Download Button
        csv_data = json.dumps(logs, indent=4) # Using JSON for now as it's simple list of dicts
        # Or better, convert to CSV string manually or use pandas if available (but let's stick to simple json/text for no deps issue here, actually user asked for file)
        
        # Let's make a CSV string
        header = ["timestamp", "admin", "user", "amount"]
        csv_rows = [",".join(header)]
        for entry in logs:
            row = [
                str(entry.get("timestamp", "")),
                str(entry.get("admin", "")),
                str(entry.get("user", "")),
                str(entry.get("amount", ""))
            ]
            csv_rows.append(",".join(row))
        csv_string = "\n".join(csv_rows)

        st.download_button(
            label="📥 Download Logs (CSV)",
            data=csv_string,
            file_name="eyecon_token_logs.csv",
            mime="text/csv"
        )

    else:
        st.info("No logs found.")
    
    st.write("---")

    # Add a new user
    st.write("### Add a new user")
    with st.form("add_user_form"):
        new_email = st.text_input("New User Email")
        new_password = st.text_input("New User Password", type="password")
        
        all_services = [
            "Application Extractor", "Excel Analyzer", "PTA Services", 
            "CDR Format", "Vehicle and Mobile", "Admin", "Settings / Future Tools",
            "Eyecon Info"
        ]
        selected_services = st.multiselect("Select services for the new user", all_services)
        initial_tokens = st.number_input("Initial Tokens", min_value=0, value=0)
        initial_eyecon_tokens = st.number_input("Initial Eyecon Tokens", min_value=0, value=0)
        
        submitted = st.form_submit_button("Add User")
        if submitted:
            if new_email and new_password and selected_services:
                USERS[new_email] = {
                    "password": new_password,
                    "tokens": initial_tokens,
                    "eyecon_tokens": initial_eyecon_tokens,
                    "services": selected_services
                }
                save_users() # Save changes to the file
                st.success(f"User {new_email} added successfully!")
                st.rerun()
            else:
                st.error("Please fill in all the fields.")

    # Remove a user
    st.write("### Remove a user")
    users_to_remove = [email for email in USERS if email != "jerrynaeem19@gmail.com"]
    if users_to_remove:
        user_to_remove = st.selectbox("Select user to remove", users_to_remove)
        if st.button("Remove User"):
            if user_to_remove in USERS:
                del USERS[user_to_remove]
                save_users() # Save changes to the file
                st.success(f"User {user_to_remove} removed successfully!")
                st.rerun()
            else:
                st.error("User not found.")
    else:
        st.info("No other users to remove.")