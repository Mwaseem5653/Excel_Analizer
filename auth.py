import streamlit as st
import json
import os

# Define the path for the users JSON file
USERS_FILE = "users.json"

def load_users():
    """Loads users from the JSON file. If the file doesn't exist, creates it with a default admin."""
    if not os.path.exists(USERS_FILE):
        default_users = {
            "jerrynaeem19@gmail.com": {
                "password": "Pakistan1122@",
                "tokens": 100, # Default tokens for admin
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
            # Migration: Ensure all users have a 'tokens' field
            is_modified = False
            for email in users:
                if "tokens" not in users[email]:
                    users[email]["tokens"] = 0
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

def add_tokens(email, amount):
    """Adds tokens to a user's balance."""
    if email in USERS:
        USERS[email]["tokens"] = USERS[email].get("tokens", 0) + amount
        save_users()
        return True
    return False

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

def logout():
    """Logs the user out and clears the session."""
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

def admin_section():
    """Displays the admin panel for user management."""
    st.title("Admin Section")
    st.subheader("Manage Users and Permissions")

    # --- Token Management ---
    st.write("### 💰 Issue Tokens")
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
    
    st.write("---")

    # Display current users
    st.write("### Current Users")
    for email, data in USERS.items():
        st.write(f"**Email:** {email}")
        st.write(f"**Tokens:** {data.get('tokens', 0)}")
        st.write(f"**Services:** {', '.join(data['services'])}")
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
        
        submitted = st.form_submit_button("Add User")
        if submitted:
            if new_email and new_password and selected_services:
                USERS[new_email] = {
                    "password": new_password,
                    "tokens": initial_tokens,
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