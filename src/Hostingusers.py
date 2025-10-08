import os
import logging
import requests
import pyodbc
from azure.identity import ClientSecretCredential
from datetime import datetime, timedelta

# -----------------------------
# CONFIGURATION VIA ENVIRONMENT VARIABLES
# -----------------------------
TENANT_ID = os.getenv("GRAPH_TENANT_ID", "<tenant_id>")
CLIENT_ID = os.getenv("GRAPH_CLIENT_ID", "<client_id>")
CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET", "<client_secret>")
SCOPES = ["https://graph.microsoft.com/.default"]

SQL_SERVER = os.getenv("SQL_SERVER", "<sql_server>")
SQL_DATABASE = os.getenv("SQL_DATABASE", "<sql_database>")
SQL_USER = os.getenv("SQL_USER", "<sql_user>")
SQL_PASSWORD = os.getenv("SQL_PASSWORD", "<sql_password>")
SQL_PORT = int(os.getenv("SQL_PORT", 1433))

TEAMS_MEETING_PLANS = [
    "MCOSTANDARD", "MCOEV", "TEAMS1", "ENTERPRISEPACK", "ENTERPRISEPREMIUM",
    "ENTERPRISEWITHSCAL", "STANDARDPACK", "STANDARDWOFFPACK", "BUSINESS_PREMIUM",
    "M365_BUSINESS_BASIC", "M365_BUSINESS_STD", "M365_E3", "M365_E5",
    "SPE_E3", "SPE_E5", "DEVELOPERPACK"
]

# -----------------------------
# HELPER FUNCTIONS
# -----------------------------
def get_graph_token():
    try:
        credential = ClientSecretCredential(
            tenant_id=TENANT_ID,
            client_id=CLIENT_ID,
            client_secret=CLIENT_SECRET
        )
        return credential.get_token(*SCOPES).token
    except Exception as e:
        logging.error(f"Failed to get Graph token: {e}")
        return None

def can_host_meetings(user_id: str, token: str) -> bool:
    """Check if user has a Teams license that allows hosting meetings."""
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/licenseDetails"
    headers = {"Authorization": f"Bearer {token}"}
    try:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        licenses = resp.json().get("value", [])
        for lic in licenses:
            for plan in lic.get("servicePlans", []):
                plan_name = plan.get("servicePlanName", "")
                status = plan.get("provisioningStatus", "")
                if plan_name in TEAMS_MEETING_PLANS and status == "Success":
                    return True
    except Exception as e:
        logging.warning(f"Failed licenseDetails for {user_id}: {e}")
    return False

def connect_sql():
    conn_str = (
        f"DRIVER={{ODBC Driver 18 for SQL Server}};"
        f"SERVER={SQL_SERVER},{SQL_PORT};"
        f"DATABASE={SQL_DATABASE};"
        f"UID={SQL_USER};PWD={SQL_PASSWORD};"
        f"Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
    )
    return pyodbc.connect(conn_str)

def fetch_users(token: str) -> list:
    url = "https://graph.microsoft.com/v1.0/users?$top=999&$select=id,mail,userPrincipalName"
    headers = {"Authorization": f"Bearer {token}"}
    try:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        return resp.json().get("value", [])
    except Exception as e:
        logging.error(f"Failed to fetch users: {e}")
        return []

def insert_user_into_sql(cursor, user_id, email, can_host):
    now = datetime.utcnow()
    sub_expiry = now + timedelta(hours=70)
    cursor.execute("""
        INSERT INTO TeamsHostingUsers
        (UserId, Email, CanHostMeetings, LastValidatedAt, SubscriptionExpiresAt)
        VALUES (?, ?, ?, ?, ?)
    """, user_id, email, int(can_host), now, sub_expiry)

# -----------------------------
# MAIN LOGIC
# -----------------------------
def main():
    logging.info("Starting Teams Hosting Users SQL update...")
    
    token = get_graph_token()
    if not token:
        logging.error("Unable to obtain Microsoft Graph token. Exiting.")
        return

    users = fetch_users(token)
    if not users:
        logging.info("No users fetched from Microsoft Graph.")
        return

    conn = connect_sql()
    cursor = conn.cursor()

    for u in users:
        email = (u.get("mail") or u.get("userPrincipalName") or "").lower()
        user_id = u.get("id")
        if not email.endswith("@mobilelive.ca") or not user_id:
            continue

        can_host = can_host_meetings(user_id, token)
        insert_user_into_sql(cursor, user_id, email, can_host)
        logging.info(f"Inserted {email} into SQL (CanHostMeetings={can_host})")

    conn.commit()
    conn.close()
    logging.info("✅ SQL population completed successfully.")

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    main()
