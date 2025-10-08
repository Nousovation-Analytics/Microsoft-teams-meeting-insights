import azure.functions as func
import logging
import os
import json
import uuid
import re
from azure.storage.blob import BlobServiceClient
import openai

# --------- ENVIRONMENT VARIABLES / PLACEHOLDERS ---------
SOURCE_CONTAINER = os.getenv("SOURCE_CONTAINER", "<source_container_name>")
TARGET_CONTAINER = os.getenv("TARGET_CONTAINER", "<target_container_name>")
STORAGE_CONNECTION = os.getenv("AzureWebJobsStorage")

AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")  # optional for Azure
AZURE_OPENAI_LLM_MODEL = os.getenv("AZURE_OPENAI_LLM_MODEL", "gpt-35-turbo")

# SQL Configuration placeholders
SQL_SERVER = os.getenv("SQL_SERVER")
SQL_DATABASE = os.getenv("SQL_DATABASE")
SQL_USER = os.getenv("SQL_USER")
SQL_PASSWORD = os.getenv("SQL_PASSWORD")
SQL_TABLE = os.getenv("SQL_TABLE", "TeamsMeetingAINotes")
SQL_PORT = 1433

# --------- FUNCTION APP INSTANCE ---------
app = func.FunctionApp()

# --------- HELPER CLASS ---------
class MeetingNotesProcessor:
    """Generates AI meeting notes from transcript text using Azure OpenAI."""

    def __init__(self):
        logging.info("Initializing MeetingNotesProcessor...")
        if not (AZURE_OPENAI_API_KEY and AZURE_OPENAI_LLM_MODEL):
            raise ValueError("Missing Azure OpenAI API key or model name")

        openai.api_key = AZURE_OPENAI_API_KEY
        if AZURE_OPENAI_ENDPOINT:
            logging.info(f"Using Azure OpenAI endpoint: {AZURE_OPENAI_ENDPOINT}")
            openai.api_base = AZURE_OPENAI_ENDPOINT
            openai.api_type = "azure"
            openai.api_version = "2023-03-15-preview"

        self.model = AZURE_OPENAI_LLM_MODEL
        logging.info(f"Model set to: {self.model}")

    def generate_meeting_notes(self, transcript_text: str) -> str:
        """Generate structured meeting notes from transcript."""
        prompt = f"""
You are an expert meeting note taker. Analyze this transcript and generate structured meeting notes.

Transcript:
{transcript_text}

Format:
Meeting notes:
* [topic_name]:[who do what]
    * [subtopic_name]:[who do what]

Follow-up tasks:
* [task_name]:[task_description]([Person responsible])
"""
        try:
            response = openai.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "You are an expert meeting note taker."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2,
                max_tokens=800
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            logging.exception(f"OpenAI completion failed: {e}")
            raise

# --------- UTILITY FUNCTIONS ---------
def get_blob_relative_path(blob_url: str, container_name: str) -> str:
    """Extract relative path of a blob inside a container."""
    try:
        return blob_url.split(f"/{container_name}/")[1]
    except IndexError:
        logging.warning(f"Failed to parse blob relative path: {blob_url}")
        return None

def get_blob_url(container_name: str, blob_relative_path: str) -> str:
    """Return the full URL to a blob in Azure Storage."""
    return f"https://<your_storage_account>.blob.core.windows.net/{container_name}/{blob_relative_path}"

def extract_metadata_from_transcript(blob_relative_path: str, transcript_text: str) -> dict:
    """Extract organizer email, meeting subject, date, and speakers from transcript and blob path."""
    metadata = {}
    try:
        parts = blob_relative_path.split('/')
        metadata['organiser_email'] = parts[0] if len(parts) >= 4 else "unknown@domain.com"
        metadata['meeting_subject'] = parts[1] if len(parts) >= 4 else "Unknown"
        metadata['date'] = parts[2] if len(parts) >= 4 else None

        speakers = set(re.findall(r"<v (.*?)>", transcript_text))
        metadata['speakers'] = list(speakers)
        metadata['speaker_count'] = len(speakers)
    except Exception:
        logging.exception("Metadata extraction failed")
    return metadata

def insert_meeting_record(organiser_email: str, meeting_subject: str, meeting_date: str, ainotes_path: str):
    """Insert AI notes record into SQL (placeholders, actual connection not included)."""
    logging.info(f"Insert record: {organiser_email}, {meeting_subject}, {meeting_date}, {ainotes_path}")
    # Actual SQL insert would use SQL_USER, SQL_PASSWORD, etc. in production

# --------- MAIN FUNCTION ---------
@app.function_name(name="eventgrid_blob_copy_ai")
@app.route(route="eventgrid_blob_copy_ai", methods=["POST"])
def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("EventGrid Function triggered")
    
    try:
        events = req.get_json()
    except Exception:
        logging.exception("Invalid JSON payload")
        return func.HttpResponse("Invalid JSON", status_code=400)

    # Handle subscription validation
    if len(events) == 1 and events[0].get("eventType") == "Microsoft.EventGrid.SubscriptionValidationEvent":
        validation_code = events[0]["data"]["validationCode"]
        return func.HttpResponse(
            json.dumps({"validationResponse": validation_code}),
            mimetype="application/json"
        )

    # Initialize processor
    try:
        processor = MeetingNotesProcessor()
    except Exception:
        logging.exception("Processor initialization failed")
        return func.HttpResponse("Processor init failed", status_code=500)

    # Process each blob event
    for event in events:
        if event.get("eventType") != "Microsoft.Storage.BlobCreated":
            continue

        blob_url = event["data"]["url"]
        blob_relative_path = get_blob_relative_path(blob_url, SOURCE_CONTAINER)
        if not blob_relative_path:
            continue

        # Dummy transcript download placeholder
        transcript_text = "<transcript text here>"

        metadata = extract_metadata_from_transcript(blob_relative_path, transcript_text)
        ai_notes = processor.generate_meeting_notes(transcript_text)

        # Save output blob (placeholder)
        logging.info(f"Would save AI notes to: {TARGET_CONTAINER}/{blob_relative_path}")

        ainotes_path = get_blob_url(TARGET_CONTAINER, blob_relative_path)
        insert_meeting_record(metadata.get("organiser_email"), metadata.get("meeting_subject"),
                              metadata.get("date"), ainotes_path)

    return func.HttpResponse("Processed", status_code=200)
