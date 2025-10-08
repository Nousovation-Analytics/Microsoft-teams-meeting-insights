import azure.functions as func
import logging
from datetime import datetime, timezone
import time
import pymssql
import requests
import uuid
from azure.identity import ClientSecretCredential
from azure.storage.filedatalake import DataLakeServiceClient
import re

# -----------------------------
# CONFIGURATION VIA ENVIRONMENT VARIABLES
# -----------------------------
ADLS_ACCOUNT_NAME = os.getenv("ADLS_ACCOUNT_NAME", "<adls_account_name>")
ADLS_ACCOUNT_KEY = os.getenv("ADLS_ACCOUNT_KEY", "<adls_account_key>")
ADLS_FILE_SYSTEM_NAME = os.getenv("ADLS_FILE_SYSTEM_NAME", "<filesystem_name>")

TENANT_ID = os.getenv("GRAPH_TENANT_ID", "<tenant_id>")
CLIENT_ID = os.getenv("GRAPH_CLIENT_ID", "<client_id>")
CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET", "<client_secret>")
SCOPES = ["https://graph.microsoft.com/.default"]

SQL_SERVER = os.getenv("SQL_SERVER", "<sql_server>")
SQL_DATABASE = os.getenv("SQL_DATABASE", "<sql_database>")
SQL_USER = os.getenv("SQL_USER", "<sql_user>")
SQL_PASSWORD = os.getenv("SQL_PASSWORD", "<sql_password>")
SQL_PORT = int(os.getenv("SQL_PORT", 1433))
MAX_RETRIES = int(os.getenv("MAX_RETRIES", 96))
RETRY_INTERVAL_SECONDS = int(os.getenv("RETRY_INTERVAL_SECONDS", 2))

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
        token = credential.get_token(*SCOPES)
        return token.token
    except Exception as e:
        logging.error(f"Failed to get Graph token: {e}")
        return None

def get_pending_meetings():
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
            SELECT TeamsMeetingId, OrganizerEmail, OrganizerObjectId, Subject, StartTime, Status, TranscriptStatus
            FROM TeamsMeetings
            WHERE (Status LIKE 'TRANSCRIPT_RUN_%' OR Status = 'MEETING_ID_FETCHED')
            AND IsParent = 1
        """)
        rows = cursor.fetchall()
        cursor.close()
        conn.close()
        return rows
    except Exception as e:
        logging.error(f"Failed to fetch pending meetings from SQL: {e}")
        return []

def determine_next_status(current_status):
    if current_status == "MEETING_ID_FETCHED":
        return "TRANSCRIPT_RUN_1"
    if current_status and current_status.startswith("TRANSCRIPT_RUN_"):
        try:
            run_num = int(current_status.split("_")[-1])
            if run_num < MAX_RETRIES:
                return f"TRANSCRIPT_RUN_{run_num + 1}"
            else:
                return "TRANSCRIPT_FAILED"
        except Exception:
            return "TRANSCRIPT_FAILED"
    return None

def parse_iso_datetime(dt_str):
    if dt_str is None:
        return None
    fixed = re.sub(r'(\.\d{6})\d+', r'\1', dt_str)
    try:
        return datetime.fromisoformat(fixed.replace("Z", "+00:00"))
    except Exception as e:
        logging.error(f"Cannot parse date {dt_str} (fixed: {fixed}): {e}")
        return None

def fetch_transcript_list(access_token, organizer_object_id, meeting_id):
    url = f"https://graph.microsoft.com/v1.0/users/{organizer_object_id}/onlineMeetings/{meeting_id}/transcripts"
    headers = {"Authorization": f"Bearer {access_token}"}
    try:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        return resp.json().get("value", [])
    except Exception as e:
        logging.error(f"Error fetching transcript list for meeting {meeting_id}: {e}")
        return []

def fetch_transcript_content(access_token, content_url):
    headers = {"Authorization": f"Bearer {access_token}"}
    try:
        resp = requests.get(content_url, headers=headers)
        resp.raise_for_status()
        return resp.text
    except Exception as e:
        logging.error(f"Error fetching transcript content: {e}")
        return None

def sanitize_filename_component(s: str) -> str:
    return re.sub(r'[\/*?:"<>|]', "_", s).strip().replace(" ", "_")

def save_transcript_to_adls(transcript_text, subject, organizer_email, meeting_start_time):
    if not transcript_text:
        return None
    try:
        service_client = DataLakeServiceClient(
            account_url=f"https://{ADLS_ACCOUNT_NAME}.dfs.core.windows.net",
            credential=ADLS_ACCOUNT_KEY
        )
        filesystem_client = service_client.get_file_system_client(ADLS_FILE_SYSTEM_NAME)
        safe_subject = sanitize_filename_component(subject or "UnknownSubject")
        safe_organizer = sanitize_filename_component(organizer_email or "UnknownOrganizer")
        if isinstance(meeting_start_time, str):
            meeting_start_time = parse_iso_datetime(meeting_start_time)
        date_str = meeting_start_time.strftime("%Y-%m-%d") if meeting_start_time else "UnknownDate"
        folder_path = f"{safe_organizer}/{safe_subject}/{date_str}"
        filename = f"{safe_subject}_transcript.txt"
        full_path = f"{folder_path}/{filename}"
        file_client = filesystem_client.get_file_client(full_path)
        file_client.upload_data(transcript_text.encode("utf-8"), overwrite=True)
        logging.info(f"Saved transcript to ADLS path: {full_path}")
        return full_path
    except Exception as e:
        logging.error(f"Failed to save transcript to ADLS: {e}")
        return None

def transcript_already_saved(transcript_url):
    try:
        conn = pymssql.connect(
            server=SQL_SERVER,
            user=SQL_USER,
            password=SQL_PASSWORD,
            database=SQL_DATABASE,
            port=SQL_PORT
        )
        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM TeamsMeetings WHERE TranscriptUrl = %s", (transcript_url,))
        exists = cursor.fetchone() is not None
        cursor.close()
        conn.close()
        return exists
    except Exception as e:
        logging.error(f"Error checking transcript existence for URL {transcript_url}: {e}")
        return False

# -----------------------------
# TIMER FUNCTION
# -----------------------------
app = func.FunctionApp()

@app.function_name(name="TimerTriggerTranscripts")
@app.schedule(schedule="0 */15 * * * *", arg_name="timer", run_on_startup=False, use_monitor=True)
def timer_trigger_transcripts(timer: func.TimerRequest) -> None:
    logging.info(f"Transcript timer trigger started at {datetime.now(timezone.utc).isoformat()}")
    access_token = get_graph_token()
    if not access_token:
        logging.error("Could not acquire Microsoft Graph token, aborting transcript fetch.")
        return

    meetings = get_pending_meetings()
    if not meetings:
        logging.info("No meetings pending transcript fetch at this time.")
        return

    for meeting in meetings:
        meeting_id = meeting["TeamsMeetingId"]
        organizer_object_id = meeting["OrganizerObjectId"]
        subject = meeting["Subject"]
        organizer_email = meeting["OrganizerEmail"]
        current_status = meeting.get("Status")

        logging.info(f"Processing meeting {meeting_id} with status {current_status}")

        transcript_list = fetch_transcript_list(access_token, organizer_object_id, meeting_id)
        transcripts_fetched = False
        for transcript_info in transcript_list:
            transcript_url = transcript_info.get("transcriptContentUrl")
            if transcript_already_saved(transcript_url):
                continue
            meeting_start = transcript_info.get("createdDateTime")
            transcript_content = fetch_transcript_content(access_token, transcript_url)
            if transcript_content:
                adls_path = save_transcript_to_adls(transcript_content, subject, organizer_email, meeting_start)
                transcripts_fetched = True
        next_status = determine_next_status(current_status) or "TRANSCRIPT_FAILED"
        logging.info(f"Meeting {meeting_id} processing completed. Next status: {next_status}")
        time.sleep(RETRY_INTERVAL_SECONDS)

    logging.info(f"Transcript timer trigger completed at {datetime.now(timezone.utc).isoformat()}")
