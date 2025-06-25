from langchain_google_genai import GoogleGenerativeAI
import os
from dotenv import load_dotenv
import json
import logging
import pandas as pd
from datetime import datetime


load_dotenv()

# Excel log file path
LOG_FILE = "email_download_log.xlsx"

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# LangChain configuration
llm = GoogleGenerativeAI(
    model="gemini-2.0-flash",
    google_api_key=os.getenv('GEMINI_API_KEY'),
    temperature=0.7,
    # top_p=0.98,
    # top_k=20,
    max_output_tokens=8192
)

# Helper functions
def read_markdown_file(file_path):
    """Read and return the content of a markdown file"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        logger.error(f"Error reading markdown file {file_path}: {str(e)}")
        raise

def validate_markdown_structure(content):
    """Validate if the markdown content has expected structure"""
    # Basic validation - check if content is not empty and contains some structure
    if not content or len(content.strip()) == 0:
        return False
    
    # Check for common markdown elements
    has_headers = '#' in content
    has_content = len(content.strip()) > 50  # Minimum content length
    
    return has_headers or has_content

def extract_markdown_metadata(content):
    """
    Extract key metadata from markdown content for debugging and validation.
    
    Args:
        content (str): Markdown content
        
    Returns:
        dict: Metadata about the content
    """
    if not content:
        return {}
    
    metadata = {
        'total_length': len(content),
        'sections': [],
        'checkboxes': {'checked': 0, 'unchecked': 0},
        'radio_buttons': {'selected': 0, 'unselected': 0},
        'tables': 0,
        'headers': []
    }
    
    lines = content.split('\n')
    
    for line in lines:
        line = line.strip()
        
        # Count headers
        if line.startswith('#'):
            metadata['headers'].append(line)
            if line.startswith('## '):
                section_name = line[3:].strip()
                metadata['sections'].append(section_name)
        
        # Count checkboxes
        if '[x]' in line:
            metadata['checkboxes']['checked'] += 1
        if '[ ]' in line:
            metadata['checkboxes']['unchecked'] += 1
        
        # Count radio buttons
        if '(•)' in line:
            metadata['radio_buttons']['selected'] += 1
        if '( )' in line:
            metadata['radio_buttons']['unselected'] += 1
        
        # Count tables
        if '|' in line and '---' not in line:
            metadata['tables'] += 1
    
    return metadata

def format_field_definitions(field_definitions):
    """Format field definitions for the prompt"""
    field_def_text = "\n"
    for i, field in enumerate(field_definitions, 1):
        name = field.get('name', '').lower()
        description = field.get('description', '')
        field_def_text += f"        {i}. {name}: {description}\n"
        field_def_text += "           If not marked, return null\n\n"
    return field_def_text.strip()

def extract_json_from_response(response_text):
    """Extract JSON from response text with improved error handling"""
    try:
        # If already a dict, return as is
        if isinstance(response_text, dict):
            return response_text

        # Convert response to string if needed
        if not isinstance(response_text, str):
            response_text = str(response_text)

        # Clean up the text
        text = response_text.strip()
        
        try:
            # Try direct JSON parsing first
            return json.loads(text)
        except:
            # Find the first { and last }
            start = text.find('{')
            end = text.rfind('}')
            
            if start != -1 and end != -1:
                # Extract just the JSON part
                json_text = text[start:end+1]
                
                # Clean up common issues
                json_text = json_text.replace('\n', ' ')
                json_text = json_text.replace('\r', ' ')
                json_text = ' '.join(json_text.split())  # Normalize whitespace
                json_text = json_text.replace("'", '"')  # Replace single quotes
                
                try:
                    return json.loads(json_text)
                except json.JSONDecodeError as e:
                    logger.error(f"Failed to parse JSON: {e}")
                    logger.error(f"Attempted to parse: {json_text}")
                    return {}
            
            logger.error("No JSON object found in response")
            logger.error(f"Response text: {text}")
            return {}

    except Exception as e:
        logger.error(f"Error in extract_json_from_response: {str(e)}")
        logger.error(f"Response text: {response_text}")
        return {}

def fix_table_data_format(data):
    """
    Post-process extracted data to ensure table data is properly formatted as arrays of objects.
    
    Args:
        data (dict): The extracted JSON data
        
    Returns:
        dict: Data with corrected table format
    """
    if not isinstance(data, dict):
        return data
    
    fixed_data = data.copy()
    
    for key, value in data.items():
        # Check if this field contains table data (common table field names)
        is_table_field = any(table_indicator in key.lower() for table_indicator in [
            'formulation', 'table', 'data', 'list', 'array', 'rows', 'materials'
        ])
        
        if is_table_field and isinstance(value, dict):
            # Convert dict to array of objects
            try:
                # Check if the dict has numeric keys (like "1", "2", "3")
                numeric_keys = [k for k in value.keys() if str(k).isdigit()]
                
                if numeric_keys:
                    # This is likely table data with row numbers as keys
                    # Convert to array format
                    table_array = []
                    max_row = max(int(k) for k in numeric_keys)
                    
                    logger.info(f"Converting table data for field '{key}' from dict with {len(numeric_keys)} rows")
                    
                    for i in range(1, max_row + 1):
                        row_key = str(i)
                        if row_key in value:
                            row_data = value[row_key]
                            if isinstance(row_data, dict):
                                # Check if all values are null (empty row)
                                if all(v is None for v in row_data.values()):
                                    logger.debug(f"Skipping empty row {i} for field '{key}'")
                                    continue
                                table_array.append(row_data)
                            else:
                                # If row_data is not a dict, create a simple object
                                table_array.append({"data": row_data})
                        else:
                            # Skip missing row numbers
                            logger.debug(f"Row {i} not found for field '{key}'")
                    
                    fixed_data[key] = table_array
                    logger.info(f"Fixed table data format for field '{key}': converted {len(numeric_keys)} dict entries to {len(table_array)} array items")
                
                elif len(value) > 0 and all(isinstance(v, dict) for v in value.values()):
                    # If all values are dicts, convert to array
                    fixed_data[key] = list(value.values())
                    logger.info(f"Fixed table data format for field '{key}': converted dict values to array")
                
            except Exception as e:
                logger.warning(f"Error fixing table data format for field '{key}': {str(e)}")
                # Keep original value if fixing fails
        
        elif is_table_field and isinstance(value, list):
            # Ensure list contains objects
            if value and not all(isinstance(item, dict) for item in value):
                # Convert list items to objects if they're not already
                fixed_data[key] = [{"data": item} if not isinstance(item, dict) else item for item in value]
                logger.info(f"Fixed table data format for field '{key}': ensured list contains objects")
    
    return fixed_data

def load_log_data():
    """Load existing log data from Excel file."""
    try:
        if os.path.exists(LOG_FILE):
            df = pd.read_excel(LOG_FILE)
            # Add json column if it doesn't exist (for backwards compatibility)
            if 'json' not in df.columns:
                df['json'] = ''
            return df
        else:
            return pd.DataFrame(columns=['json'])
    except Exception as e:
        logger.error(f"Error loading log data: {e}")
        return pd.DataFrame()

def save_log_data(df):
    """Save log data to Excel file."""
    try:
        df.to_excel(LOG_FILE, index=False)
        logger.info(f"Log data saved to {LOG_FILE}")
    except Exception as e:
        logger.error(f"Error saving log data: {e}")

def is_markdown_already_processed(folder_name):
    """
    Check if markdown file for a given folder has already been processed.
    
    Args:
        folder_name (str): Name of the folder containing the markdown file
        
    Returns:
        bool: True if already processed, False otherwise
    """
    try:
        log_df = load_log_data()
        if log_df.empty:
            return False
        
        # Look for rows where the folder name appears in file_paths and json column is 'completed'
        for _, row in log_df.iterrows():
            file_paths = str(row.get('file_paths', ''))
            json_status = str(row.get('json', ''))
            
            # Check if folder name is in file paths and json is completed
            if folder_name in file_paths and json_status.lower() == 'completed':
                logger.info(f"Markdown for folder '{folder_name}' already processed (found in log)")
                return True
        
        return False
    except Exception as e:
        logger.error(f"Error checking if markdown already processed: {e}")
        return False

def mark_json_as_completed(folder_name):
    """
    Mark the JSON column as completed for the row containing the folder name.
    
    Args:
        folder_name (str): Name of the folder containing the markdown file
    """
    try:
        log_df = load_log_data()
        if log_df.empty:
            logger.warning("No log data found to update")
            return
        
        # Find rows where the folder name appears in file_paths
        updated = False
        for idx, row in log_df.iterrows():
            file_paths = str(row.get('file_paths', ''))
            
            if folder_name in file_paths:
                log_df.at[idx, 'json'] = 'completed'
                updated = True
                logger.info(f"Marked JSON as completed for folder '{folder_name}' in log")
        
        if updated:
            save_log_data(log_df)
        else:
            logger.warning(f"No matching row found for folder '{folder_name}' in log")
            
    except Exception as e:
        logger.error(f"Error marking JSON as completed: {e}")

# Schema definitions
client_schema = {
    "id": "new-id-001",
    "name": "client site information",
    "description": "Client and site-related details",
    "fields": [
        {
            "description": "",
            "name": "Name of the client",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "City",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "Street",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "Pin code",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "State",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "Building No",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "Email",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "Website",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "Telephone",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "Site Information",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "Cooling water available",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "Place of machine installation",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "Gearbox make preferred",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "End Application details",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "Bench mark available",
            "required": False,
            "type": "text"
        }
    ],
    "created_at": "2025-06-19T12:15:00.000000",
    "updated_at": "2025-06-19T12:15:00.000000"
}

invoice_schema = {
    "id": "35a299fa-d30a-452c-8fa5-ee188e353525",
    "name": "flexo tech quotation",
    "description": "test",
    "fields": [
        {
            "description": "",
            "name": "Payment type",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "Grand Total",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "Dated",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "gstin",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "Kind Attn",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "Enquiry No",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "quotation number",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "To Address",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "validity",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "Sales Co-ordinator number",
            "required": False,
            "type": "text"
        },
        {
            "description": "",
            "name": "Sales Co-ordinator",
            "required": False,
            "type": "text"
        }
    ],
    "created_at": "2025-06-11T04:58:01.533380",
    "updated_at": "2025-06-19T12:00:00.000000"
}

quatation_schema = {
    "id": "general_quotation_template",
    "name": "General Quotation Data Extraction",
    "description": "Universal template for extracting data from quotations, estimates, and purchase orders",
    "version": "1.0",
    "fields": [
        {
            "name": "document_source",
            "description": "Original document filename or reference",
            "required": False,
            "type": "text"
        },
        {
            "name": "quotation_number",
            "description": "Unique quotation/estimate/invoice number",
            "required": False,
            "type": "text"
        },
        {
            "name": "quotation_date",
            "description": "Date when quotation was issued",
            "required": False,
            "type": "date"
        },
        {
            "name": "supplier_name",
            "description": "Name of the supplier/vendor company",
            "required": False,
            "type": "text"
        },
        {
            "name": "supplier_gstin",
            "description": "Supplier's GST identification number",
            "required": False,
            "type": "text"
        },
        {
            "name": "supplier_address",
            "description": "Complete address of supplier",
            "required": False,
            "type": "text"
        },
        {
            "name": "supplier_contact",
            "description": "Phone number, email or contact person details",
            "required": False,
            "type": "text"
        },
        {
            "name": "customer_name",
            "description": "Name of the customer/buyer company",
            "required": False,
            "type": "text"
        },
        {
            "name": "customer_gstin",
            "description": "Customer's GST identification number",
            "required": False,
            "type": "text"
        },
        {
            "name": "goods_description",
            "description": "Detailed description of items/services quoted",
            "required": False,
            "type": "text"
        },
        {
            "name": "item_details",
            "description": "Array of individual items with quantities and rates",
            "required": False,
            "type": "array"
        },
        {
            "name": "subtotal_amount",
            "description": "Amount before taxes",
            "required": False,
            "type": "currency"
        },
        {
            "name": "tax_details",
            "description": "Breakdown of taxes (CGST, SGST, IGST, etc.)",
            "required": False,
            "type": "object"
        },
        {
            "name": "total_amount",
            "description": "Final total amount including all taxes and charges",
            "required": False,
            "type": "currency"
        },
        {
            "name": "currency",
            "description": "Currency code (INR, USD, etc.)",
            "required": False,
            "type": "text",
            "default": "INR"
        },
        {
            "name": "validity",
            "description": "Validity period of the quotation",
            "required": False,
            "type": "text"
        },
        {
            "name": "delivery",
            "description": "Delivery terms and timeframe",
            "required": False,
            "type": "text"
        },
        {
            "name": "payment_terms",
            "description": "Payment terms and conditions",
            "required": False,
            "type": "text"
        },
        {
            "name": "freight_terms",
            "description": "Freight/shipping terms",
            "required": False,
            "type": "text"
        },
        {
            "name": "warranty",
            "description": "Warranty terms if applicable",
            "required": False,
            "type": "text"
        },
        {
            "name": "special_conditions",
            "description": "Any special terms and conditions",
            "required": False,
            "type": "text"
        },
        {
            "name": "hsn_codes",
            "description": "HSN/SAC codes for items",
            "required": False,
            "type": "array"
        },
        {
            "name": "po_reference",
            "description": "Related purchase order reference if any",
            "required": False,
            "type": "text"
        },
        {
            "name": "enquiry_reference",
            "description": "Original enquiry reference number",
            "required": False,
            "type": "text"
        },
        {
            "name": "document_type",
            "description": "Type of document (quotation, estimate, proforma invoice, etc.)",
            "required": False,
            "type": "text"
        },
        {
            "name": "status",
            "description": "Current status of the quotation",
            "required": False,
            "type": "text"
        },
        {
            "name": "extracted_date",
            "description": "Date when data was extracted",
            "required": False,
            "type": "datetime"
        },
        {
            "name": "extraction_notes",
            "description": "Any notes or issues during extraction",
            "required": False,
            "type": "text"
        }
    ],
    "created_at": "2025-06-23T00:00:00.000000",
    "updated_at": "2025-06-23T00:00:00.000000"
}

# Create a PromptTemplate for extraction
EXTRACTION_PROMPT = """You are an expert data analyzer assistant. Your task is to analyse data from markdown document  and return it in JSON format.

extract what are the field name asked and the respected selected options.
the selected options are below the field name with the [x] or (•) which selected to find it easily 
Be concious with the actual table and text options , some times the normal option is also getting in table format with each character as a column

FIELD DEFINITIONS:
{field_definitions}
  

DOCUMENT CONTENT (Markdown Format):
{pdf_text}

INSTRUCTIONS:
1. Extract data for each field defined above from the markdown content
  - we have lable or question and their corresponding answer in checkbox or radio button may be in fill up the field
2. The content is in markdown format with the following conventions:
   - Headers: ## Section Name
   - Checkboxes: [x] for checked, [ ] for unchecked
   - Radio buttons: (•) for selected, ( ) for unselected
   - Tables: | Column1 | Column2 | Column3 |
   - Bold text: **Label:** value
3. Return ONLY a JSON object - no other text or formatting
4. Use null for missing or empty values
5. Use exact field names from the definitions
6. For checkboxes and radio buttons, extract the selected/checked values
7. For tables, extract relevant data as arrays of objects where each object represents a row

RESPONSE FORMAT:
Return only a JSON object like this (using actual field names from definitions):
{{
    "actual_field_name1": "value1",
    "actual_field_name2": null,
    "actual_field_name3": ["value1", "value2"],
    "checkbox_field": "selected_option",
    "table_data": [
        {{"column1": "value1", "column2": "value2"}},
        {{"column1": "value3", "column2": "value4"}}
    ]
}}

"""


def process_md_folder():
    output_dir = "output"
    results_dir = "results"
    
    # Create results directory if it doesn't exist
    if not os.path.exists(results_dir):
        os.makedirs(results_dir)
        print(f"Created results directory: {results_dir}")
    
    # Check if output directory exists
    if not os.path.exists(output_dir):
        print(f"Error: Output directory '{output_dir}' not found.")
        return
    
    # Get all subdirectories in the output directory
    subdirs = [d for d in os.listdir(output_dir) if os.path.isdir(os.path.join(output_dir, d))]
    
    if not subdirs:
        print("No subdirectories found in the output directory.")
        return
    
    # Process each subdirectory
    for subdir in subdirs:
        subdir_path = os.path.join(output_dir, subdir)
        print(f"\nProcessing subdirectory: {subdir}")
        
        # Check if this markdown has already been processed
        if is_markdown_already_processed(subdir):
            print(f"Markdown for folder '{subdir}' already processed. Skipping...")
            continue
        
        # Look for markdown files in the subdirectory
        md_files = [f for f in os.listdir(subdir_path) if f.lower().endswith('.md')]
        
        if not md_files:
            print(f"No markdown files found in {subdir}")
            continue
        
        # Process each markdown file
        for md_file in md_files:
            md_path = os.path.join(subdir_path, md_file)
            print(f"Processing markdown file: {md_file}")
            
            try:
                # Process the markdown file using process_single_md
                result = process_single_md(md_path)
                
                # Check if processing was successful
                if result.get('status') == 'success':
                    # Generate unique filename for the JSON result
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    json_filename = f"{subdir}_{timestamp}.json"
                    json_path = os.path.join(results_dir, json_filename)
                    
                    # Save the result to JSON file
                    with open(json_path, 'w', encoding='utf-8') as f:
                        json.dump(result, f, indent=2, ensure_ascii=False)
                    
                    print(f"Saved result to: {json_path}")
                    
                    # Mark JSON as completed in the Excel log
                    mark_json_as_completed(subdir)
                    print(f"Marked JSON processing as completed for folder '{subdir}'")
                else:
                    print(f"Processing failed for {md_file}: {result.get('error', 'Unknown error')}")
                
            except Exception as e:
                print(f"Error processing {md_file}: {str(e)}")
                continue





def process_single_md(file_path):
    """Process a single markdown file using LangChain with Google Gemini"""
    try:
        # Debug: Print current working directory
        cwd = os.getcwd()
        print(f"Current working directory: {cwd}")
        
        # Get the absolute path to the markdown file
        md_path = os.path.abspath(os.path.normpath(file_path))
        print(f"Input markdown path: {md_path}")
        
        # Verify markdown file exists
        if not os.path.isfile(md_path):
            error_msg = f"Markdown file not found: {md_path}"
            print(error_msg)
            logger.error(error_msg)
            return {
                'status': 'error',
                'data': {},
                'file_path': file_path,
                'error': 'Markdown file not found'
            }
        
        # Read the markdown content
        try:
            print(f"Attempting to read file: {md_path}")
            pdf_text = read_markdown_file(md_path)
            
            # Validate the markdown structure
            if not validate_markdown_structure(pdf_text):
                error_msg = "Markdown file does not contain expected structure"
                print(error_msg)
                logger.warning(error_msg)
                # Continue processing anyway, but log the warning
            
            # Extract and log metadata for debugging
            metadata = extract_markdown_metadata(pdf_text)
            print(f"Markdown metadata: {metadata}")
            logger.info(f"Markdown metadata: {metadata}")
            
        except Exception as e:
            error_msg = f"Error reading file {md_path}: {str(e)}"
            print(error_msg)
            logger.error(error_msg)
            return {
                'status': 'error',
                'data': {},
                'file_path': file_path,
                'error': f'Error reading markdown file: {str(e)}'
            }
        
        # Choose schema based on content analysis
        if "BEACON INDUSTRIES" in pdf_text:
            selected_schema = quatation_schema
            print(f"Using quatation_schema for markdown containing 'BEACON INDUSTRIES'")
        elif "Flexo Tech Products" in pdf_text:
            selected_schema = invoice_schema
            print(f"Using invoice_schema for markdown containing 'Flexo Tech Products'")
        else:
            selected_schema = client_schema
            print(f"Using client_schema for markdown without specific identifiers")
        
        # Use the selected schema's fields for field definitions
        schema_field_definitions = selected_schema.get('fields', [])
        
        # Format field definitions using the selected schema
        field_def_text = format_field_definitions(schema_field_definitions)
        
        # Create the prompt with the content
        prompt = EXTRACTION_PROMPT.format(
            field_definitions=field_def_text,
            pdf_text=pdf_text
        )
        
        # Get response from LLM
        try:
            response = llm.invoke(prompt)
            logger.info(f"Raw LLM response: {response}")
        except Exception as e:
            error_msg = f"Error getting LLM response: {str(e)}"
            print(error_msg)
            logger.error(error_msg)
            return {
                'status': 'error',
                'data': {field.get('name', ''): None for field in schema_field_definitions},
                'file_path': file_path,
                'error': error_msg
            }
        
        # Extract JSON from response
        extraction_json_data = extract_json_from_response(response)
        
        if not extraction_json_data:
            error_msg = "Failed to extract valid JSON from LLM response"
            logger.error(error_msg)
            return {
                'status': 'error',
                'data': {field.get('name', ''): None for field in schema_field_definitions},
                'file_path': file_path,
                'error': error_msg
            }
        
        # Fix table data format
        fixed_data = fix_table_data_format(extraction_json_data)
        
        logger.info(f"Successfully processed markdown: {file_path}")
        logger.info(f"Extracted data: {fixed_data}")

        return {
            'status': 'success',
            'data': fixed_data,
            'file_path': file_path
        }
        
    except Exception as e:
        error_msg = f"Error processing markdown {file_path}: {str(e)}"
        print(error_msg)
        logger.error(error_msg)
        return {
            'status': 'error',
            'data': {},
            'file_path': file_path,
            'error': str(e)
        }


# if __name__ == "__main__":
#     print("Starting markdown to JSON processing...")
#     print("Checking for previously processed files...")
#     process_md_folder()
#     print("Processing complete!")


