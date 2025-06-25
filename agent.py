import os
import re
import base64
import pandas as pd
import hashlib
from datetime import datetime, timedelta
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from langchain.tools import tool
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain.agents import AgentExecutor, create_tool_calling_agent
from langchain_core.prompts import ChatPromptTemplate
from langchain_community.agent_toolkits import GmailToolkit
from langchain_community.tools.gmail.utils import (
    build_resource_service,
    get_gmail_credentials,
)
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Ensure the OPENAI_API_KEY environment variable is set
api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    raise ValueError("The OPENAI_API_KEY environment variable is not set")


SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
BASE_DIR = os.getcwd()  # gets the current working directory
SAVE_PATH = os.path.join(BASE_DIR, 'download')
LOG_FILE = os.path.join(BASE_DIR, "email_download_log.xlsx")

def get_gmail_service():
    """Initialize Gmail service with credentials."""
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    return build('gmail', 'v1', credentials=creds)


def initialize_log_file():
    """Initialize the Excel log file with required columns if it doesn't exist."""
    if not os.path.exists(LOG_FILE):
        # Create initial DataFrame with required columns - one row per file
        df = pd.DataFrame(columns=[
            # Document Identity
            'subject',
            'email_id',
            'thread_id',
            'sender',
            
            # Processing Timeline
            'first_inbox_msg',
            'last_check_date',
            'download_date',
            'duplicate_check_date',
            
            # File Management
            'count_download',
            'list_name_count',
            'attachment_names',
            'file_paths',
            'original_filenames',
            'res_path',           # New field: Result file path location
            
            # Data Integrity
            'message_hash',
            'file_hashes',
            'unique_file_ids',
            
            # Processing Status
            'process_status',
            'classification',
            'duplicate_status',
            'markdown',
            'json',
            'res_status'         # New field: Result processing status
        ])
        df.to_excel(LOG_FILE, index=False)
        print(f"Created new structured log file: {LOG_FILE}")
    return LOG_FILE


def load_log_data():
    """Load existing log data from Excel file."""
    try:
        if os.path.exists(LOG_FILE):
            df = pd.read_excel(LOG_FILE)
            # Add new columns if they don't exist (for backwards compatibility)
            required_columns = [
                'subject', 'email_id', 'thread_id', 'sender',
                'first_inbox_msg', 'last_check_date', 'download_date', 'duplicate_check_date',
                'count_download', 'list_name_count', 'attachment_names', 'file_paths', 
                'original_filenames', 'res_path', 'message_hash', 'file_hashes', 
                'unique_file_ids', 'process_status', 'classification', 'duplicate_status',
                'markdown', 'json', 'res_status'
            ]
            
            for col in required_columns:
                if col not in df.columns:
                    df[col] = ''
            
            return df
        else:
            initialize_log_file()
            return pd.DataFrame(columns=[
                'subject', 'email_id', 'thread_id', 'sender',
                'first_inbox_msg', 'last_check_date', 'download_date', 'duplicate_check_date',
                'count_download', 'list_name_count', 'attachment_names', 'file_paths', 
                'original_filenames', 'res_path', 'message_hash', 'file_hashes', 
                'unique_file_ids', 'process_status', 'classification', 'duplicate_status',
                'markdown', 'json', 'res_status'
            ])
    except Exception as e:
        print(f"Error loading log data: {e}")
        return pd.DataFrame()



def save_log_data(df):
    """Save log data to Excel file."""
    try:
        os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
        df.to_excel(LOG_FILE, index=False)
        print(f"Log data saved to {LOG_FILE}")
    except Exception as e:
        print(f"Error saving log data: {e}")



def generate_message_hash(msg_data):
    """Generate unique hash for message to detect duplicates."""
    headers = msg_data['payload'].get('headers', [])
    subject = next((h['value'] for h in headers if h['name'] == 'Subject'), '')
    message_id = msg_data.get('id', '')
    thread_id = msg_data.get('threadId', '')
    
    hash_string = f"{message_id}_{thread_id}_{subject}"
    return hashlib.md5(hash_string.encode()).hexdigest()



def generate_file_hash(file_data):
    """Generate hash of file content to detect identical files."""
    return hashlib.sha256(file_data).hexdigest()


def generate_unique_file_id(filename, file_hash, email_id):
    """Generate unique identifier for a file based on name, content and email."""
    return f"{filename}_{email_id}_{file_hash[:16]}"


def is_file_already_downloaded(log_df, filename, file_data, email_id, thread_id=None):
    """
    Check if file has already been downloaded using multiple criteria.
    Returns: (bool, str) - (is_duplicate, reason)
    """
    if log_df.empty:
        return False, "No previous downloads"

    file_hash = generate_file_hash(file_data)
    unique_file_id = generate_unique_file_id(filename, file_hash, email_id)

    # Check for exact content match (identical files)
    exact_matches = log_df[log_df['file_hashes'].str.contains(file_hash, na=False)]
    if not exact_matches.empty:
        match = exact_matches.iloc[0]
        return True, (
            f"Identical file content already exists (downloaded on {match['download_date']} "
            f"from {match['sender']}, subject: {match['subject']})"
        )

    # Check for same file in the same email thread
    if thread_id:
        thread_matches = log_df[
            (log_df['thread_id'] == thread_id) & 
            (log_df['attachment_names'].str.contains(re.escape(filename), na=False))
        ]
        if not thread_matches.empty:
            match = thread_matches.iloc[0]
            return True, (
                f"Same filename already downloaded in this email thread "
                f"(original download: {match['download_date']}, subject: {match['subject']})"
            )

    # Check for similar filenames with different content
    similar_names = log_df[log_df['attachment_names'].str.contains(re.escape(filename), na=False)]
    if not similar_names.empty:
        # If same filename exists but with different content, add a note
        print(f"Note: Found file with same name but different content: {filename}")

    # Check if this exact combination of file and email has been processed
    unique_matches = log_df[log_df['unique_file_ids'].str.contains(unique_file_id, na=False)]
    if not unique_matches.empty:
        match = unique_matches.iloc[0]
        return True, (
            f"This exact file from this email has already been processed "
            f"(original download: {match['download_date']})"
        )

    return False, "File is new"


@tool 
def monitor_gmail_for_new_attachments_with_logging() -> str:
    """Monitor Gmail for new emails with document attachments in the last 24 hours with enhanced duplicate prevention."""
    try:
        # Initialize logging
        initialize_log_file()
        log_df = load_log_data()
        
        service = get_gmail_service()
        
        # Search for recent emails with attachments (last 24 hours)
        yesterday = datetime.now() - timedelta(days=1)
        query = f'has:attachment after:{yesterday.strftime("%Y/%m/%d")}'
        
        results = service.users().messages().list(userId='me', q=query).execute()
        messages = results.get('messages', [])

        if not messages:
            return "📭 No new emails with attachments found in the last 24 hours."

        downloaded = []
        skipped = []
        processed_emails = []
        skip_details = []
        
        for msg in messages:
            try:
                msg_data = service.users().messages().get(userId='me', id=msg['id']).execute()
                
                # Extract email information
                headers = msg_data['payload'].get('headers', [])
                subject = next((h['value'] for h in headers if h['name'] == 'Subject'), 'No Subject')
                sender = next((h['value'] for h in headers if h['name'] == 'From'), 'Unknown Sender')
                date_header = next((h['value'] for h in headers if h['name'] == 'Date'), '')
                thread_id = msg_data.get('threadId', '')
                
                # Generate message hash
                message_hash = generate_message_hash(msg_data)
                
                # Process attachments
                parts = msg_data['payload'].get('parts', [])
                if not parts and msg_data['payload'].get('filename'):
                    parts = [msg_data['payload']]
                
                document_attachments = []
                for part in parts:
                    filename = part.get('filename', '')
                    if filename and any(filename.lower().endswith(ext) for ext in ['.pdf', '.docx', '.xlsx', '.doc', '.ppt', '.pptx', '.txt']):
                        document_attachments.append(filename)
                
                if not document_attachments:
                    continue
                
                # Download new attachments with enhanced checking - create one row per file
                current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                thread_exists = not log_df[log_df['thread_id'] == thread_id].empty
                
                for part in parts:
                    filename = part.get('filename', '')
                    if filename in document_attachments:
                        try:
                            # Get file data for duplicate checking
                            if 'data' in part.get('body', {}):
                                file_data = base64.urlsafe_b64decode(part['body']['data'])
                            elif 'attachmentId' in part.get('body', {}):
                                attachment = service.users().messages().attachments().get(
                                    userId='me', messageId=msg['id'], id=part['body']['attachmentId']
                                ).execute()
                                file_data = base64.urlsafe_b64decode(attachment['data'])
                            else:
                                continue

                            # Enhanced duplicate check
                            is_duplicate, reason = is_file_already_downloaded(log_df, filename, file_data, msg['id'], thread_id)
                            
                            if is_duplicate:
                                skipped.append(filename)
                                skip_details.append(f"  • {filename}: {reason}")
                                continue

                            os.makedirs(SAVE_PATH, exist_ok=True)
                            
                            # Handle duplicate filenames
                            base_name, ext = os.path.splitext(filename)
                            counter = 1
                            final_filename = filename
                            while os.path.exists(os.path.join(SAVE_PATH, final_filename)):
                                final_filename = f"{base_name}_{counter}{ext}"
                                counter += 1
                            
                            filepath = os.path.join(SAVE_PATH, final_filename)
                            with open(filepath, 'wb') as f:
                                f.write(file_data)
                            
                            # Generate file metadata
                            file_hash = generate_file_hash(file_data)
                            unique_file_id = generate_unique_file_id(final_filename, file_hash, msg['id'])
                            
                            # Create one row per file
                            new_row = {
                                'subject': subject,
                                'email_id': msg['id'],
                                'thread_id': thread_id,
                                'sender': sender,
                                'first_inbox_msg': date_header if not thread_exists else log_df[log_df['thread_id'] == thread_id]['first_inbox_msg'].iloc[0] if thread_exists else date_header,
                                'last_check_date': current_time,
                                'download_date': current_time,
                                'duplicate_check_date': current_time,
                                'count_download': 1,  # Each row represents one file
                                'list_name_count': final_filename,
                                'attachment_names': final_filename,
                                'file_paths': filepath,
                                'original_filenames': filename,
                                'res_path': '',  # Will be filled during processing
                                'message_hash': message_hash,
                                'file_hashes': file_hash,
                                'unique_file_ids': unique_file_id,
                                'process_status': 'downloaded',
                                'classification': '',
                                'duplicate_status': 'unique',
                                'markdown': '',
                                'json': '',
                                'res_status': ''
                            }
                            
                            log_df = pd.concat([log_df, pd.DataFrame([new_row])], ignore_index=True)
                            downloaded.append(final_filename)
                            
                        except Exception as e:
                            print(f"Error downloading {filename}: {e}")
                            continue
                
                if any(filename in downloaded for filename in document_attachments):
                    processed_emails.append(f"📧 {sender}: {subject}")
                
            except Exception as e:
                print(f"Error processing message: {e}")
                continue
        
        # Save updated log
        save_log_data(log_df)
        
        # Generate detailed summary
        summary = []
        summary.append("🔍 Gmail Monitoring Complete")
        summary.append("=" * 40)
        
        if downloaded:
            summary.append(f"✅ Downloaded {len(downloaded)} new attachments:")
            for file in downloaded:
                summary.append(f"   • {file}")
        
        if skipped:
            summary.append(f"\n⏭️ Skipped {len(skipped)} duplicate attachments:")
            summary.extend(skip_details)
        
        if processed_emails:
            summary.append(f"\n📨 Processed emails:")
            for email in processed_emails:
                summary.append(f"   {email}")
        
        summary.append(f"\n📊 Log file: {LOG_FILE}")
        summary.append(f"💾 Save path: {SAVE_PATH}")
        
        return '\n'.join(summary) if downloaded or skipped else "📭 No new document attachments found in recent emails."
        
    except Exception as e:
        return f"❌ Error monitoring emails: {str(e)}"



@tool
def view_download_log() -> str:
    """View the current download log statistics with duplicate prevention details."""
    try:
        if not os.path.exists(LOG_FILE):
            return "📄 No log file found. No downloads recorded yet."
        
        log_df = pd.read_excel(LOG_FILE)
        
        if log_df.empty:
            return "📄 Log file is empty. No downloads recorded yet."
        
        # Generate statistics for one-row-per-file structure
        total_files = len(log_df)  # Each row is one file
        unique_emails = log_df['email_id'].nunique() if 'email_id' in log_df.columns else 0
        unique_senders = log_df['sender'].nunique() if 'sender' in log_df.columns else 0
        unique_threads = log_df['thread_id'].nunique() if 'thread_id' in log_df.columns else 0
        
        # File hash statistics (if available)
        unique_files_by_content = 0
        if 'file_hashes' in log_df.columns:
            all_hashes = []
            for _, row in log_df.iterrows():
                if pd.notna(row['file_hashes']) and row['file_hashes']:
                    hashes = str(row['file_hashes']).split(', ')
                    all_hashes.extend(hashes)
            unique_files_by_content = len(set(all_hashes)) if all_hashes else 0
        
        # Recent downloads (last 7 days)
        if 'download_date' in log_df.columns:
            log_df['download_date'] = pd.to_datetime(log_df['download_date'], errors='coerce')
            recent = log_df[log_df['download_date'] > (datetime.now() - timedelta(days=7))]
            recent_count = len(recent)
        else:
            recent_count = 0
        
        # File type breakdown
        file_types = {}
        if 'attachment_names' in log_df.columns:
            for _, row in log_df.iterrows():
                if pd.notna(row['attachment_names']):
                    filename = str(row['attachment_names'])
                    ext = os.path.splitext(filename)[1].lower()
                    file_types[ext] = file_types.get(ext, 0) + 1
        
        summary = []
        summary.append("📊 Enhanced Download Log Summary (One Row Per File)")
        summary.append("=" * 50)
        summary.append(f"📎 Total files downloaded: {total_files}")
        summary.append(f"📧 Unique emails processed: {unique_emails}")
        summary.append(f"🔗 Unique email threads: {unique_threads}")
        summary.append(f"📄 Unique files by content: {unique_files_by_content}")
        summary.append(f"👥 Unique senders: {unique_senders}")
        summary.append(f"🕐 Recent downloads (7 days): {recent_count}")
        summary.append(f"📁 Log file location: {LOG_FILE}")
        
        if file_types:
            summary.append("\n📋 File type breakdown:")
            for ext, count in sorted(file_types.items()):
                summary.append(f"   • {ext}: {count} files")
        
        if not log_df.empty and len(log_df) > 0:
            summary.append("\n📋 Recent file entries:")
            for _, row in log_df.tail(5).iterrows():
                subject = row.get('subject', 'N/A')[:40] + '...' if len(str(row.get('subject', ''))) > 40 else row.get('subject', 'N/A')
                sender = row.get('sender', 'N/A')
                filename = row.get('attachment_names', 'N/A')
                download_date = row.get('download_date', 'N/A')
                summary.append(f"   • {filename}")
                summary.append(f"     From: {sender} | Subject: {subject}")
                summary.append(f"     Downloaded: {download_date}")
        
        return '\n'.join(summary)
        
    except Exception as e:
        return f"❌ Error reading log: {str(e)}"


@tool
def clear_duplicate_entries() -> str:
    """Clean up duplicate entries in the log file based on file content."""
    try:
        if not os.path.exists(LOG_FILE):
            return "📄 No log file found."
        
        log_df = pd.read_excel(LOG_FILE)
        
        if log_df.empty:
            return "📄 Log file is empty."
        
        original_count = len(log_df)
        
        # Remove duplicates based on unique_file_ids if available
        if 'unique_file_ids' in log_df.columns:
            # Keep the first occurrence of each unique file
            log_df = log_df.drop_duplicates(subset=['unique_file_ids'], keep='first')
        elif 'file_hashes' in log_df.columns:
            # Fallback to file content hash
            log_df = log_df.drop_duplicates(subset=['file_hashes'], keep='first')
        else:
            # Fallback to message_hash + attachment_names
            log_df = log_df.drop_duplicates(subset=['message_hash', 'attachment_names'], keep='first')
        
        cleaned_count = len(log_df)
        removed_count = original_count - cleaned_count
        
        if removed_count > 0:
            save_log_data(log_df)
            return f"✅ Cleaned log file: Removed {removed_count} duplicate entries. {cleaned_count} entries remain."
        else:
            return "✅ No duplicate entries found in log file."
        
    except Exception as e:
        return f"❌ Error cleaning log: {str(e)}"




# ✅ Set your API key
llm = ChatGoogleGenerativeAI(
    model="gemini-1.5-flash",
    temperature=0,
    google_api_key=api_key  # ✅ <-- Replace with your actual key
)



# Combine Gmail toolkit tools with your custom tools 

credentials = get_gmail_credentials(
        token_file="token.json",
        scopes=SCOPES,
        client_secrets_file="credentials.json"
    )

api_resource = build_resource_service(credentials=credentials)
toolkit = GmailToolkit(api_resource=api_resource)
all_tools = toolkit.get_tools() + [
        monitor_gmail_for_new_attachments_with_logging,
        view_download_log,
        clear_duplicate_entries
    ]

prompt = ChatPromptTemplate.from_messages([
        ("system", """You are an advanced Gmail assistant with comprehensive logging and duplicate prevention capabilities. You can:
        
        🔍 **Email Operations:**
        - Search and read emails
        - Send emails
        - Monitor inbox for new messages
        
        📎 **Enhanced Attachment Management:**
        - Download document attachments (PDF, DOCX, XLSX, DOC, PPT, PPTX, TXT)
        - **ONE ROW PER FILE STRUCTURE**: Each downloaded file gets its own row in the log
        - **ENHANCED DUPLICATE PREVENTION**: 
          * Check file content hashes to prevent downloading identical files
          * Prevent downloading same files from email threads
          * Track unique file identifiers with email context
        - Maintain detailed logs in Excel format with structured columns
        
        📊 **Advanced Logging Features:**
        - **Document Identity**: subject, email_id, thread_id, sender
        - **Processing Timeline**: first_inbox_msg, last_check_date, download_date, duplicate_check_date
        - **File Management**: count_download, list_name_count, attachment_names, file_paths, original_filenames, res_path
        - **Data Integrity**: message_hash, file_hashes, unique_file_ids
        - **Processing Status**: process_status, classification, duplicate_status, markdown, json, res_status
        - Generate comprehensive statistics including unique file counts and file type breakdowns
        - Provide detailed reasons for skipping duplicate files
        
        🛡️ **Multi-Level Duplicate Prevention:**
        1. **Content-based**: Identical file content (SHA-256 hash)
        2. **Thread-based**: Same filename from same email thread
        3. **Unique ID**: Combination of filename, email_id and content hash
        
        🧹 **Maintenance Tools:**
        - Clean duplicate entries from log file
        - View detailed statistics with duplicate analysis and file type breakdown
        
        Always provide clear, detailed feedback about actions taken, files downloaded, duplicates skipped with specific reasons.
        """),
        ("user", "{input}"),
        ("placeholder", "{agent_scratchpad}")
    ])


# Create the agent with all tools
agent = create_tool_calling_agent(llm, all_tools, prompt)

# Build the AgentExecutor
agent_executor = AgentExecutor(
    agent=agent,
    tools=all_tools,
    verbose=True,
    handle_parsing_errors=True
)

def run_agent():
    try:
            response = agent_executor.invoke({
                "input": """Monitor my Gmail inbox for new emails with document attachments. 
                Download any PDF, DOCX, XLSX files, but avoid downloading duplicates using 
                enhanced content-based detection. Show me detailed information about any 
                files that were skipped and why."""
            })
            print("-------------------------------------------")
            print("Response:", response["output"])
            print("-------------------------------------------")
    except Exception as e:
            print(f"Error: {e}")