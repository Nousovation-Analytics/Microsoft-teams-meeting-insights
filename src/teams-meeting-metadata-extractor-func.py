import azure.functions as func
import logging
import json
from datetime import datetime, timezone
import requests
from azure.identity import ClientSecretCredential
import uuid
from urllib.parse import quote

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

# --------- PLACEHOLDERS / ENVIRONMENT VARIABLES ---------
TENANT_ID = os.getenv("GRAPH_TENANT_ID", "<tenant_id>")
CLIENT_ID = os.getenv("GRAPH_CLIENT_ID", "<client_id>")
CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET", "<client_secret>")
SCOPES = ["https://graph.microsoft.com/.default"]

# SQL placeholders
SQL_SERVER = os.getenv("SQL_SERVER", "<sql_server>")
SQL_DATABASE = os.getenv("SQL_DATABASE", "<sql_database>")
SQL_USER = os.getenv("SQL_USER", "<sql_user>")
SQL_PASSWORD = os.getenv("SQL_PASSWORD", "<sql_password>")
SQL_PORT = int(os.getenv("SQL_PORT", 1433))
SQL_TABLE = os.getenv("SQL_TABLE", "TeamsMeetings")

# --------- GRAPH TOKEN ---------
def get_graph_token():
    """Fetch Microsoft Graph access token using client credentials."""
    try:
        credential = ClientSecretCredential(
            tenant_id=TENANT_ID,
            client_id=CLIENT_ID,
            client_secret=CLIENT_SECRET
        )
        token = credential.get_token(*SCOPES)
        return token.token
    except Exception as e:
        logging.error(f"Error fetching Graph token: {e}")
        return None

# --------- HELPER FUNCTIONS ---------
def get_event_details(user_id, event_id, access_token):
    """Fetch event details from Graph API."""
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/events/{event_id}"
    headers = {"Authorization": f"Bearer {access_token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json()

def get_official_meeting_id_by_join_url(user_id, join_url, access_token):
    """Get official Teams meeting ID via onlineMeetings filter."""
    if not all([user_id, join_url, access_token]):
        return None
    encoded_join_url = quote(join_url, safe='')
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/onlineMeetings?$filter=joinWebUrl eq '{encoded_join_url}'"
    headers = {"Authorization": f"Bearer {access_token}"}
    try:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        meetings = resp.json().get("value", [])
        return meetings[0].get("id") if meetings else None
    except Exception as e:
        logging.error(f"Failed to fetch official meeting ID: {e}")
        return None

def normalize_datetime(dt_str):
    """Normalize ISO datetime to UTC-aware datetime object."""
    if not dt_str:
        return None
    dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
    return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)

def safe_get(dct, keys):
    """Safe nested dictionary access."""
    for key in keys:
        if not isinstance(dct, dict):
            return None
        dct = dct.get(key)
        if dct is None:
            return None
    return dct

def get_user_object_id_by_email(email, access_token):
    """Fetch Azure AD object ID for a given email."""
    url = f"https://graph.microsoft.com/v1.0/users/{email}"
    headers = {"Authorization": f"Bearer {access_token}"}
    try:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        return resp.json().get("id")
    except Exception as e:
        logging.warning(f"Failed to get object ID for {email}: {e}")
        return None

def upsert_meeting_sql(meeting_id, organizer_email, subject, start_time, end_time, join_url,
                      transcript_status, transcript_url, adls_path, last_checked, meeting_series_id,
                      meeting_completion_status, notes, status, organizer_object_id):
    """Placeholder for SQL upsert. Replace with actual DB logic in production."""
    logging.info(f"Upsert meeting {meeting_id}: {subject}, organizer {organizer_email}, status {status}")
    # Actual SQL connection code using pymssql can be inserted here

# --------- MAIN HTTP TRIGGER ---------
@app.route(route="http_trigger_webhooks", methods=["GET", "POST"])
def http_trigger_webhooks(req: func.HttpRequest) -> func.HttpResponse:
    """Handles incoming Teams webhook notifications."""
    validation_token = req.params.get("validationToken")
    if validation_token:
        return func.HttpResponse(validation_token, status_code=200)

    try:
        body = req.get_json()
        access_token = get_graph_token()
        if not access_token:
            return func.HttpResponse("Failed to obtain Graph token", status_code=500)

        processed_meeting_ids = set()  # Deduplication in batch

        for notification in body.get("value", []):
            resource_data = notification.get("resourceData")
            if not resource_data:
                continue

            event_id = resource_data.get("id")
            resource_parts = notification.get("resource", "").split("/")
            user_id = resource_parts[1] if len(resource_parts) > 1 else None

            if not event_id or not user_id:
                continue

            event_details = get_event_details(user_id, event_id, access_token)
            if not event_details:
                continue

            join_url = safe_get(event_details, ["onlineMeetingUrl"]) or safe_get(event_details, ["onlineMeeting", "joinUrl"])
            if not join_url:
                continue

            meeting_id = get_official_meeting_id_by_join_url(user_id, join_url, access_token)
            if not meeting_id or meeting_id in processed_meeting_ids:
                continue
            processed_meeting_ids.add(meeting_id)

            subject = safe_get(event_details, ["subject"]) or "No subject"
            organizer_email = safe_get(event_details, ["organizer", "emailAddress", "address"]) or "unknown"
            organizer_object_id = safe_get(event_details, ["organizer", "emailAddress", "id"]) or get_user_object_id_by_email(organizer_email, access_token) or str(uuid.uuid4())

            start_time = normalize_datetime(safe_get(event_details, ["start", "dateTime"]))
            end_time = normalize_datetime(safe_get(event_details, ["end", "dateTime"]))
            series_id = safe_get(event_details, ["seriesMasterId"])
            now_utc = datetime.now(timezone.utc)

            transcript_status = "MEETING_ID_FETCHED"
            transcript_url = None
            meeting_completion_status = ""
            notes = "Meeting metadata saved"
            status = "MEETING_ID_FETCHED"

            upsert_meeting_sql(
                meeting_id, organizer_email, subject, start_time, end_time, join_url,
                transcript_status, transcript_url, None, now_utc, series_id,
                meeting_completion_status, notes, status, organizer_object_id
            )

    except Exception as ex:
        logging.error(f"Function exception: {ex}")
        return func.HttpResponse("Error", status_code=500)

    return func.HttpResponse("Received", status_code=202)
