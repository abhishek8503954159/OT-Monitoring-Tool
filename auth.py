"""
Simple Azure AD SSO Authentication with Excel user verification.
Uses OAuth2/OIDC with tenant_id, client_id, client_secret.
"""

import os
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from typing import Optional, Dict, Any

# Load environment variables from .env file
from dotenv import load_dotenv
load_dotenv()

try:
    import msal
    MSAL_AVAILABLE = True
except ImportError:
    MSAL_AVAILABLE = False

# Azure AD Configuration - loaded from .env file
TENANT_ID = os.getenv("AZURE_TENANT_ID", "YOUR_TENANT_ID")
CLIENT_ID = os.getenv("AZURE_CLIENT_ID", "YOUR_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET", "YOUR_CLIENT_SECRET")
REDIRECT_URI = os.getenv("REDIRECT_URI", "http://localhost:8501")

# Users Excel file for authorization
USERS_FILE = os.getenv("USERS_FILE", "Users_List.xlsx")

# Development mode - set to True to bypass Azure AD and login with email directly
DEV_MODE = os.getenv("DEV_MODE", "false").lower() == "true"

# Azure AD endpoints
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["User.Read"]


class UserVerification:
    """Verify users against Excel file."""
    
    def __init__(self, users_file: str = None):
        self.users_file = users_file or USERS_FILE
        self._cache = None
        self._cache_time = None
        self._email_col = None
        self._name_col = None
    
    def _load_users(self) -> pd.DataFrame:
        """Load users with 5-minute cache."""
        now = datetime.now()
        if self._cache is None or self._cache_time is None or (now - self._cache_time).seconds > 300:
            if os.path.exists(self.users_file):
                # Try reading with sheet name 'Users', fallback to first sheet
                try:
                    self._cache = pd.read_excel(self.users_file, sheet_name='Users')
                except:
                    self._cache = pd.read_excel(self.users_file, sheet_name=0)
                self._cache.columns = self._cache.columns.str.strip()
                
                # Find email column (flexible matching)
                self._email_col = self._find_column(['Email', 'email', 'E-mail', 'User Email', 
                                                      'Username', 'username', 'User', 'UserName',
                                                      'Mail', 'mail', 'EmailAddress', 'email_address'])
                
                # Find name column
                self._name_col = self._find_column(['Name', 'name', 'Full Name', 'FullName', 
                                                     'Employee Name', 'Display Name', 'DisplayName'])
            else:
                self._cache = pd.DataFrame()
            self._cache_time = now
        return self._cache
    
    def _find_column(self, possible_names: list) -> Optional[str]:
        """Find column by checking possible names."""
        if self._cache is None:
            return None
        for col in self._cache.columns:
            col_lower = col.lower().strip()
            for name in possible_names:
                if col_lower == name.lower():
                    return col
        return None
    
    def verify_user(self, email: str) -> Optional[Dict[str, Any]]:
        """Check if user email/username exists in Users_List.xlsx."""
        df = self._load_users()
        if df.empty or self._email_col is None:
            return None
        
        email_lower = email.lower().strip()
        # Also try matching just the username part (before @)
        username_part = email_lower.split('@')[0] if '@' in email_lower else email_lower
        
        # Check for exact email match or username match
        df_lower = df[self._email_col].astype(str).str.lower().str.strip()
        mask = (df_lower == email_lower) | (df_lower == username_part)
        
        # Also check if stored value is username and matches
        if not mask.any():
            mask = df_lower.apply(lambda x: x.split('@')[0] if '@' in x else x) == username_part
        
        if mask.any():
            user = df[mask].iloc[0]
            
            # Check if active (if column exists)
            is_active_col = self._find_column(['Is_Active', 'IsActive', 'Active', 'Status'])
            if is_active_col:
                is_active = user.get(is_active_col, True)
                if not (pd.isna(is_active) or is_active == True or str(is_active).lower() in ['true', 'yes', 'active', '1']):
                    return None
            
            # Get user details
            name = user.get(self._name_col, 'Unknown') if self._name_col else 'Unknown'
            
            emp_id_col = self._find_column(['Employee_ID', 'EmployeeID', 'Emp_ID', 'EmpID', 'ID', 'Personnel Number'])
            dept_col = self._find_column(['Department', 'Dept', 'Area', 'Team'])
            role_col = self._find_column(['Role', 'role', 'Access', 'Permission'])
            
            return {
                'email': email,
                'employee_id': str(user.get(emp_id_col, '')) if emp_id_col else '',
                'name': str(name) if pd.notna(name) else 'Unknown',
                'role': str(user.get(role_col, 'viewer')) if role_col else 'viewer',
                'department': str(user.get(dept_col, 'Unknown')) if dept_col else 'Unknown'
            }
        return None


def get_msal_app():
    """Create MSAL confidential client application."""
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )


def get_auth_url():
    """Get Azure AD login URL."""
    app = get_msal_app()
    return app.get_authorization_request_url(
        scopes=SCOPE,
        redirect_uri=REDIRECT_URI
    )


def get_token_from_code(auth_code: str) -> Optional[Dict]:
    """Exchange authorization code for access token."""
    app = get_msal_app()
    result = app.acquire_token_by_authorization_code(
        auth_code,
        scopes=SCOPE,
        redirect_uri=REDIRECT_URI
    )
    return result if "access_token" in result else None


def get_user_info(access_token: str) -> Optional[Dict]:
    """Get user info from Microsoft Graph API."""
    import requests
    
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get("https://graph.microsoft.com/v1.0/me", headers=headers)
    
    if response.status_code == 200:
        return response.json()
    return None


def init_session_state():
    """Initialize authentication session state."""
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'user_info' not in st.session_state:
        st.session_state.user_info = None
    if 'login_time' not in st.session_state:
        st.session_state.login_time = None


def check_auth_callback():
    """Check for OAuth callback with authorization code."""
    query_params = st.query_params
    auth_code = query_params.get("code")
    
    if auth_code and not st.session_state.authenticated:
        try:
            token_result = get_token_from_code(auth_code)
            if token_result:
                user_data = get_user_info(token_result["access_token"])
                if user_data:
                    email = user_data.get("mail") or user_data.get("userPrincipalName", "")
                    
                    # Verify against Excel
                    verifier = UserVerification()
                    verified = verifier.verify_user(email)
                    
                    if verified:
                        st.session_state.authenticated = True
                        st.session_state.user_info = {
                            **verified,
                            'name': user_data.get("displayName", verified['name'])
                        }
                        st.session_state.login_time = datetime.now().isoformat()
                        st.query_params.clear()
                        st.rerun()
                    else:
                        st.error(f"Access denied. User '{email}' not found in authorized users list.")
                        st.query_params.clear()
        except Exception as e:
            st.error(f"Authentication error: {e}")
            st.query_params.clear()


def logout():
    """Clear session."""
    st.session_state.authenticated = False
    st.session_state.user_info = None
    st.session_state.login_time = None


def show_login_page():
    """Display login page."""
    st.markdown("""
    <style>
    .login-box {
        max-width: 400px;
        margin: 80px auto;
        padding: 40px;
        background: #f8f9fa;
        border-radius: 10px;
        text-align: center;
    }
    </style>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("## 🔐 OT Monitoring Dashboard")
        st.markdown("---")
        
        if DEV_MODE:
            st.warning("⚠️ Development Mode - SSO Bypassed")
            email = st.text_input("Email", placeholder="Enter your email")
            
            if st.button("Login", type="primary", use_container_width=True):
                if email:
                    verifier = UserVerification()
                    user = verifier.verify_user(email)
                    if user:
                        st.session_state.authenticated = True
                        st.session_state.user_info = user
                        st.session_state.login_time = datetime.now().isoformat()
                        st.rerun()
                    else:
                        st.error("User not authorized. Check users.xlsx")
                else:
                    st.warning("Enter your email")
        else:
            if not MSAL_AVAILABLE:
                st.error("MSAL library not installed. Run: pip install msal")
                return
            
            if TENANT_ID == "YOUR_TENANT_ID":
                st.error("Azure AD not configured. Set AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET")
                return
            
            st.markdown("### Sign in with Microsoft")
            
            try:
                auth_url = get_auth_url()
                st.link_button("🔑 Login with Microsoft", auth_url, use_container_width=True)
            except Exception as e:
                st.error(f"Error generating login URL: {e}")
        
        st.markdown("---")
        st.caption("Contact IT support if you need access.")


def show_user_sidebar():
    """Show logged-in user info in sidebar."""
    if st.session_state.authenticated and st.session_state.user_info:
        user = st.session_state.user_info
        with st.sidebar:
            st.markdown("---")
            st.markdown(f"**👤 {user['name']}**")
            st.caption(f"📧 {user['email']}")
            st.caption(f"🏢 {user['department']}")
            st.caption(f"Role: {user['role']}")
            if st.button("Logout", use_container_width=True):
                logout()
                st.rerun()


def require_auth():
    """Main authentication check. Call at start of app."""
    init_session_state()
    
    if not DEV_MODE and MSAL_AVAILABLE:
        check_auth_callback()
    
    if not st.session_state.authenticated:
        show_login_page()
        st.stop()
    
    show_user_sidebar()
