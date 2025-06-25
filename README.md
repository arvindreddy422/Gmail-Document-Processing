# Gmail Document Processing System

A comprehensive automated system for downloading PDF attachments from Gmail, extracting their content using AI, and converting the extracted data into structured JSON format.

## 🏗️ System Architecture

The system follows a modular pipeline architecture with the following main components:

```
📧 Gmail → 📥 Download → 📄 PDF Extraction → 📝 Markdown → 🗂️ JSON
```

## 📋 Main Flow

The entire process is orchestrated by `doc.py`, which serves as the main entry point and coordinates all processing steps:

1. **Email Monitoring** (`agent.py`) - Downloads PDF attachments from Gmail
2. **PDF Processing** (`extract_pdf_md.py`) - Converts PDFs to markdown using AI
3. **JSON Conversion** (`json_from_md.py`) - Transforms markdown to structured JSON

## 🔧 Components Overview

### 1. Main Controller (`doc.py`)

- **Purpose**: Orchestrates the entire processing pipeline
- **Function**: `main()` - Executes all three processing stages sequentially
- **Flow**: `run_agent()` → `process_input_folder()` → `process_md_folder()`

### 2. Gmail Agent (`agent.py`)

- **Purpose**: Monitors Gmail for new PDF attachments and downloads them
- **Key Features**:
  - Gmail API integration with OAuth2 authentication
  - Duplicate detection using file hashes and message IDs
  - Comprehensive logging to Excel file (`email_download_log.xlsx`)
  - Intelligent file naming and organization
- **Output**: PDF files saved to `download/` directory
- **Logging**: Tracks all downloads with metadata including:
  - Email details (subject, sender, thread ID)
  - File information (names, paths, hashes)
  - Processing status and timestamps

### 3. PDF to Markdown Extractor (`extract_pdf_md.py`)

- **Purpose**: Converts PDF documents to structured markdown using AI
- **Process**:
  1. Converts PDF pages to images using PyMuPDF
  2. Processes images with Google Gemini AI
  3. Extracts structured data in markdown format
  4. Handles forms, tables, checkboxes, and radio buttons
- **Input**: PDF files from `download/` directory
- **Output**: Markdown files in `output/{filename}/output_all_pages.md`
- **Features**:
  - AI-powered text extraction from images
  - Form field detection (checkboxes, radio buttons)
  - Table structure preservation
  - Automatic cleanup of temporary images

### 4. Markdown to JSON Converter (`json_from_md.py`)

- **Purpose**: Converts markdown content to structured JSON based on predefined schemas
- **Process**:
  1. Reads markdown files from `output/` directory
  2. Analyzes content structure and identifies document type
  3. Applies appropriate schema for data extraction
  4. Generates structured JSON output
- **Input**: Markdown files from `output/` directory
- **Output**: JSON files in `results/` directory
- **Features**:
  - Schema-based data extraction
  - Validation and error handling
  - Support for complex data structures (tables, arrays)
  - Comprehensive logging and status tracking

## 📁 Directory Structure

```
gmail/
├── doc.py                    # Main controller
├── agent.py                  # Gmail monitoring and download
├── extract_pdf_md.py         # PDF to markdown conversion
├── json_from_md.py           # Markdown to JSON conversion
├── download/                 # Downloaded PDF files
│   ├── A-0031_Qt_2025.pdf
│   ├── B - 083 05.05.2025.pdf
│   └── Digital_Depiction_-_NDA_Signed_1.pdf
├── output/                   # Extracted markdown files
│   ├── A-0031_Qt_2025/
│   │   └── output_all_pages.md
│   ├── B - 083 05.05.2025/
│   │   └── output_all_pages.md
│   └── Digital_Depiction_-_NDA_Signed_1/
│       └── output_all_pages.md
├── results/                  # Final JSON output
│   ├── A-0031_Qt_2025_20250625_120845.json
│   ├── B - 083 05.05.2025_20250625_120847.json
│   └── Digital_Depiction_-_NDA_Signed_1_20250625_120849.json
├── images/                   # Temporary image files (auto-cleaned)
├── email_download_log.xlsx   # Processing log and status tracking
└── token.json               # Gmail API authentication token
```

## 🚀 Getting Started

### Prerequisites

1. **Python Dependencies**:

   ```bash
   pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib
   pip install langchain langchain-google-genai langchain-community
   pip install pandas openpyxl PyMuPDF pdf2image Pillow
   pip install python-dotenv cryptography
   ```

2. **API Keys**:

   - Google Gemini API key (set as `GEMINI_API_KEY` environment variable)
   - Gmail API credentials (stored in `token.json`)

3. **Gmail API Setup**:
   - Enable Gmail API in Google Cloud Console
   - Create OAuth2 credentials
   - Generate `token.json` for authentication

### Environment Variables

Create a `.env` file with:

```
GEMINI_API_KEY=your_gemini_api_key_here
```

### Running the System

Execute the main controller:

```bash
python doc.py
```

This will automatically:

1. Check Gmail for new PDF attachments
2. Download new files to the `download/` directory
3. Convert PDFs to markdown using AI
4. Transform markdown to structured JSON
5. Update the processing log

## 📊 Logging and Monitoring

The system maintains comprehensive logs in `email_download_log.xlsx` with the following information:

- **Document Identity**: Subject, email ID, thread ID, sender
- **Processing Timeline**: First inbox message, download date, processing dates
- **File Management**: Attachment names, file paths, original filenames
- **Data Integrity**: Message hashes, file hashes, unique file IDs
- **Processing Status**: Download status, markdown conversion, JSON generation

## 🔄 Processing Flow Details

### Stage 1: Email Monitoring

- Monitors Gmail inbox for new messages with PDF attachments
- Implements duplicate detection to avoid reprocessing
- Downloads files with organized naming convention
- Updates Excel log with download metadata

### Stage 2: PDF Processing

- Converts each PDF page to high-quality images
- Uses Google Gemini AI to extract structured data
- Handles complex form elements (checkboxes, radio buttons, tables)
- Generates clean markdown output
- Automatically cleans up temporary image files

### Stage 3: JSON Conversion

- Analyzes markdown content to determine document type
- Applies appropriate schema for data extraction
- Converts form data, tables, and structured content to JSON
- Validates output and handles errors gracefully
- Saves structured data with timestamps

## 🛠️ Customization

### Adding New Document Types

1. Define schema in `json_from_md.py`
2. Add document type detection logic
3. Update field definitions and validation rules

### Modifying AI Prompts

- Edit system instructions in `extract_pdf_md.py` for PDF processing
- Update extraction prompts in `json_from_md.py` for JSON conversion

### Changing Output Formats

- Modify the JSON schema definitions
- Update the markdown formatting rules
- Adjust file naming conventions

## 🔍 Troubleshooting

### Common Issues

1. **Gmail API Authentication**: Ensure `token.json` is valid and has proper permissions
2. **PDF Processing Errors**: Check if PDFs are password-protected or corrupted
3. **AI Extraction Issues**: Verify Gemini API key and quota limits
4. **File Permission Errors**: Ensure write permissions for output directories

### Debug Mode

- Check `email_download_log.xlsx` for detailed processing status
- Review console output for error messages
- Verify file paths and directory structure

## 📈 Performance Considerations

- **Batch Processing**: System processes files sequentially to avoid API rate limits
- **Duplicate Prevention**: Comprehensive hash-based duplicate detection
- **Resource Management**: Automatic cleanup of temporary files
- **Error Recovery**: Graceful handling of processing failures

## 🤝 Contributing

When modifying the system:

1. Update the Excel log schema if adding new fields
2. Maintain backward compatibility with existing data
3. Test with various PDF formats and document types
4. Update documentation for any new features

## 📄 License

This project is designed for internal document processing workflows. Ensure compliance with data privacy and security requirements when processing sensitive documents.
