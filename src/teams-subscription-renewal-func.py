import logging
import azure.functions as func
import requests
import json
import pymssql
from azure.identity import ClientSecretCredential
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

# -----------------------------
# CONFIGURATION (PLACEHOLDERS)
# -----------------------------
TENANT_ID = "<YOUR_TENANT_ID>"
CLIENT_ID = "<YOUR_CLIENT_ID>"
CLIENT_SECRET = "<YOUR_CLIENT_SECRET>"

FUNCTION_URL = "<YOUR_FUNCTION_APP_URL>"  # Endpoint for Graph API notifications
SCOPES = ["https://graph.microsoft.com/.default"]

SQL_SERVER = "<SQL_SERVER_NAME>"
SQL_DATABASE = "<SQL_DATABASE_NAME>"
SQL_USER = "<SQL_USERNAME>"
SQL_PASSWORD = "<SQL_PASSWORD>"
SQL_PORT = 1433

# -----------------------------
# FUNCTION APP INIT
# -----------------------------
app = func.FunctionApp()

# -----------------------------
# HELPERS
# -----------------------------
def get_graph_token():
    """Get Microsoft Graph API access token."""
    try:
        credential = ClientSecretCredential(
            tenant_id=TENANT_ID,
            client_id=CLIENT_ID,
            client_secret=CLIENT_SECRET
        )
        token = credential.get_token(*SCOPES)
        return token.token
    except Exception as e:
        logging.error(f"Failed to get Graph token: {e}")
        return None

def get_teams_users_from_sql():
    """Fetch users authorized to host meetings from SQL."""
    try:
        conn = pymssql.connect(
            server=SQL_SERVER,
            user=SQL_USER,
            password=SQL_PASSWORD,
            database=SQL_DATABASE,
            port=SQL_PORT
        )
        cursor = conn.cursor(as_dict=True)
        cursor.execute("""
            SELECT UserId, Email
            FROM TeamsHostingUsers
            WHERE CanHostMeetings = 1
        """)
        rows = cursor.fetchall()
        cursor.close()
        conn.close()
        return rows
    except Exception as e:
        logging.error(f"Failed to fetch Teams users from SQL: {e}")
        return []

def update_last_validated_for_all(new_validated_at, new_expiry):
    """Update subscription timestamps for all users in SQL once."""
    try:
        conn = pymssql.connect(
            server=SQL_SERVER,
            user=SQL_USER,
            password=SQL_PASSWORD,
            database=SQL_DATABASE,
            port=SQL_PORT
        )
        cursor = conn.cursor()
        sql = """
        UPDATE TeamsHostingUsers
        SET LastValidatedAt = %s,
            SubscriptionExpiresAt = %s
        WHERE CanHostMeetings = 1
        """
        cursor.execute(sql, (new_validated_at, new_expiry))
        conn.commit()
        cursor.close()
        conn.close()
        logging.info("✅ Updated subscription timestamps for all users")
    except Exception as e:
        logging.error(f"Error updating subscriptions for all users: {e}")

def renew_subscription_for_user(token, user):
    """Renew Graph subscription for a single user."""
    sub_url = "https://graph.microsoft.com/v1.0/subscriptions"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    now = datetime.utcnow()
    new_expiry = now + timedelta(hours=70)

    body = {
        "changeType": "created,updated,deleted",
        "notificationUrl": FUNCTION_URL,
        "resource": f"/users/{user['UserId']}/events",
        "expirationDateTime": new_expiry.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "clientState": "SecretClientValue"
    }

    try:
        resp = requests.post(sub_url, headers=headers, data=json.dumps(body))
        if resp.status_code in [200, 201]:
            logging.info(f"✅ Subscription renewed for {user['Email']} ({user['UserId']})")
            return True
        else:
            logging.warning(f"⚠️ Cannot renew subscription for {user['Email']}: {resp.status_code} {resp.text}")
            return False
    except Exception as e:
        logging.error(f"Error renewing subscription for {user['Email']}: {e}")
        return False

# -----------------------------
# TIMER TRIGGER FUNCTION
# -----------------------------
@app.function_name(name="TimerTriggerRenewSubscriptions")
@app.schedule(schedule="0 */70 * * * *", arg_name="myTimer", run_on_startup=True, use_monitor=False)
def timer_trigger_renew_subscriptions(myTimer: func.TimerRequest) -> None:
    if myTimer.past_due:
        logging.warning("The timer is past due!")

    logging.info("Subscription renewal timer started.")

    token = get_graph_token()
    if not token:
        logging.error("Could not acquire Graph token, aborting subscription renewal.")
        return

    users = get_teams_users_from_sql()
    if not users:
        logging.info("No Teams users found with CanHostMeetings=True.")
        return

    # ------------- Parallel Subscription Renewal -------------
    max_workers = 10  # adjust based on expected load
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(renew_subscription_for_user, token, u) for u in users]
        for _ in as_completed(futures):
            pass  # wait for all to complete

    # ------------- Update SQL once for all users -------------
    now = datetime.utcnow()
    new_expiry = now + timedelta(hours=70)
    update_last_validated_for_all(now, new_expiry)

    logging.info("Subscription renewal timer completed.")
