import os
import fitz  # PyMuPDF
from pdf2image import convert_from_path
import pandas as pd
import google.generativeai as genai
from PIL import Image
from dotenv import load_dotenv

load_dotenv()

gemini_api_key = os.getenv('GEMINI_API_KEY')

# Excel log file path
LOG_FILE = "email_download_log.xlsx"

def load_log_data():
    """Load existing log data from Excel file."""
    try:
        if os.path.exists(LOG_FILE):
            df = pd.read_excel(LOG_FILE)
            # Add markdown column if it doesn't exist (for backwards compatibility)
            if 'markdown' not in df.columns:
                df['markdown'] = ''
            return df
        else:
            return pd.DataFrame(columns=['markdown'])
    except Exception as e:
        print(f"Error loading log data: {e}")
        return pd.DataFrame()

def save_log_data(df):
    """Save log data to Excel file."""
    try:
        df.to_excel(LOG_FILE, index=False)
        print(f"Log data saved to {LOG_FILE}")
    except Exception as e:
        print(f"Error saving log data: {e}")

def is_pdf_already_processed(pdf_filename):
    """
    Check if a PDF file has already been processed (markdown column = 'completed').
    
    Args:
        pdf_filename (str): Name of the PDF file (without extension)
        
    Returns:
        bool: True if already processed, False otherwise
    """
    try:
        log_df = load_log_data()
        if log_df.empty:
            return False
        
        # Check if any row has this PDF filename in file_paths and markdown = 'completed'
        for _, row in log_df.iterrows():
            if pd.notna(row['file_paths']) and pd.notna(row['markdown']):
                file_paths = str(row['file_paths']).split(', ')
                if any(pdf_filename in path for path in file_paths) and row['markdown'] == 'completed':
                    return True
        return False
    except Exception as e:
        print(f"Error checking if PDF is already processed: {e}")
        return False

def mark_pdf_as_completed(pdf_filename):
    """
    Mark a PDF file as completed in the Excel log file.
    
    Args:
        pdf_filename (str): Name of the PDF file (without extension)
    """
    try:
        log_df = load_log_data()
        if log_df.empty:
            print("No log data found to update.")
            return
        
        # Find rows that contain this PDF filename in file_paths
        updated = False
        for idx, row in log_df.iterrows():
            if pd.notna(row['file_paths']):
                file_paths = str(row['file_paths']).split(', ')
                if any(pdf_filename in path for path in file_paths):
                    log_df.at[idx, 'markdown'] = 'completed'
                    updated = True
                    print(f"Marked {pdf_filename} as completed in log file.")
        
        if updated:
            save_log_data(log_df)
        else:
            print(f"Could not find {pdf_filename} in log file to mark as completed.")
            
    except Exception as e:
        print(f"Error marking PDF as completed: {e}")

def cleanup_images():
    """
    Remove all PNG images from the output directory after text extraction is complete.
    """
    output_dir = "images"
    if not os.path.exists(output_dir):
        return
    
    # Get all PNG files
    image_files = [f for f in os.listdir(output_dir) if f.lower().endswith('.png')]
    
    # Remove each image file
    for image_file in image_files:
        try:
            file_path = os.path.join(output_dir, image_file)
            os.remove(file_path)
            print(f"Removed: {image_file}")
        except Exception as e:
            print(f"Error removing {image_file}: {str(e)}")
    
    print("Cleanup completed. All images have been removed.")


def process_input_folder():
    input_dir = "download"
    
    # Check if input directory exists
    if not os.path.exists(input_dir):
        print(f"Error: Input directory '{input_dir}' not found.")
        return
    
    # Clean up any existing images at the start
    cleanup_images()
    
    # Get all PDF files from input directory
    pdf_files = [f for f in os.listdir(input_dir) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print("No PDF files found in the input directory.")
        return
    
    # Process each PDF file
    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_dir, pdf_file)
        pdf_filename = os.path.splitext(pdf_file)[0]  # Get filename without extension
        
        # Check if PDF is already processed
        if is_pdf_already_processed(pdf_filename):
            print(f"Skipping {pdf_file} - already processed (markdown column = 'completed')")
            continue
        
        print(f"\nProcessing {pdf_file}...")
        result = convert_pdf_to_images(pdf_path)
        
        # If processing was successful, mark as completed
        if result:
            mark_pdf_as_completed(pdf_filename)
        
        # Clean up images after each PDF is processed
        cleanup_images()

def process_images_with_gemini(pdf_filename):
    """
    Process images with Gemini AI and create a markdown report.
    
    Args:
        pdf_filename (str): Name of the original PDF file (without extension)
        
    Returns:
        bool: True if processing was successful, False otherwise
    """
    # Configure Gemini
    genai.configure(api_key=gemini_api_key)
    
    # Configure Gemini model
    gemini_config = {
        "temperature": 1,
        "top_p": 0.98,
        "top_k": 20,
        "max_output_tokens": 8192,
    }
    
    # Initialize chat session
    model = genai.GenerativeModel(
        model_name='gemini-2.0-flash',
        generation_config=gemini_config,
        system_instruction="You are an expert data extraction assistant specialized in processing manufacturing industry quotation and enquiry forms from image. The forms contain tables, checkboxes, radio buttons, input fields, text fields, and filled-in data. Task: Extract all data and return it as a clean Markdown-formatted.for radio options use the (•) Yes  ( ) No, for checkboxes use the [x] for checked and [ ] for unchecked. You should not add any information that is not present in the image.Only the data that is presene in the image no extra information"
    )
    
    chat = model.start_chat(history=[])
    
    # Add initial guidelines to the chat
    initial_guidelines = """
You are an expert AI assistant specialized in extracting data from images of manufacturing industry quotation and enquiry forms, and converting it into a clean, well-formatted Markdown file.
Input: You will receive an image of a form. This form may contain:
Tables (with or without clearly defined borders)
Checkboxes (checked or unchecked)
Radio buttons (selected or unselected)
Input fields (filled or blank)
Text fields (containing single-line or multi-line text)
Handwritten text (if present, try to interpret it and flag if uncertain)
Varying fonts and font sizes
Noise, shadows, or slight distortions
Task:
Text Extraction: Accurately extract all text elements from the image, using OCR or other appropriate methods. Pay close attention to detail to ensure that no information is missed. Correct any OCR errors, especially in technical terms, part numbers, or industry-specific vocabulary.
Data Interpretation: Analyze the extracted text to identify the different form elements (tables, checkboxes, radio buttons, fields, etc.) and their relationships. Understand the form's structure to correctly interpret the data.
Markdown Conversion: Convert the extracted data into a clean, well-formatted Markdown file. Follow these specific formatting 
! important follow this - Do not put ```markdown and ``` at starting and ending of markdown data 
guidelines:
Headings: Use appropriate Markdown headings ( ##, ###, etc.) to structure the document and clearly identify different sections of the form (e.g., "Company Details", "Product Specifications", "Contact Information"). Maintain the hierarchy of the form in the headings.
Lists: Use Markdown lists (-, *, 1., etc.) to represent lists of items.
Checkboxes and Radio Buttons: Represent checkboxes and radio buttons using the following format:
Checked: [x]
Unchecked: [ ]
Radio buttons: 
- Use (•) for selected option
- Use ( ) for unselected option
Enclose these in bullet points within a section describing the options. For example:
Generated markdown
**Material Options:**
- [x] Steel
- [ ] Aluminum
- [ ] Plastic
Use code with caution.

Tables: Convert tables into Markdown tables using | to separate columns and --- to create the header row separator. Ensure that the table is properly aligned and readable. If the table has no visible borders, infer the structure from the data.
Generated markdown
| Header 1 | Header 2 | Header 3 |
|---|---|---|
| Data 1 | Data 2 | Data 3 |
| Data 4 | Data 5 | Data 6 |
Use code with caution.

Input Fields and Text Fields: Represent input fields and text fields with the field label followed by the extracted text (if any) or a placeholder (e.g., ______) if the field is blank. If the label is missing, infer it from the surrounding context. For multi-line text fields, preserve the line breaks in the Markdown output.
Generated markdown
Name: John Doe
Email: john.doe@example.com
Comments:
This is a multi-line comment.
It spans several lines.


Important Notes: Display important notes or disclaimers in a distinct way, using blockquotes or bold text.
Generated markdown
> **Note:** All prices are in USD and do not include shipping.


Uncertainty Flagging: If you are uncertain about any extracted data (e.g., due to poor image quality or handwriting), add a comment in the Markdown file indicating the uncertainty. Use the following format:
Generated markdown
<!-- Possible OCR error: "Part Number: AB1234?" - Please verify. -->


Original Form Layout: While the focus is on structured data, try to preserve the original form's layout as much as possible in the Markdown structure to improve readability.
Cleanliness: Ensure that the final Markdown file is clean, well-formatted, and easy to read. Remove any unnecessary characters or formatting.

Constraints:
You should not add any information that is not present in the image.
Focus on accuracy and completeness.
Prioritize readability in the Markdown output.
Do not put ```markdown and ``` at starting and ending of markdown data ! important follow this
at last I need to combine multiple md files together , there will be a issue with this approach
Output:
A complete Markdown file containing all the extracted data from the image, formatted according to the guidelines above.
    """
    
    chat.send_message(initial_guidelines)
    
    output_dir = "images"
    if not os.path.exists(output_dir):
        print(f"Error: Output directory '{output_dir}' not found.")
        return False
    
    # Get all PNG files from output directory
    image_files = [f for f in os.listdir(output_dir) if f.lower().endswith('.png')]
    
    if not image_files:
        print("No image files found in the output directory.")
        return False
    
    # Sort files by page number
    image_files.sort(key=lambda x: int(x.split('_page_')[1].split('.')[0]))
    
    # Process each image and collect results
    all_results = []
    
    for image_file in image_files:
        image_path = os.path.join(output_dir, image_file)
        print(f"\nProcessing {image_file}...")
        
        try:
            # Load and process image
            image = Image.open(image_path)
            
            # Generate content using Gemini with chat context
            response = chat.send_message(image)
            
            # Add page number and content to results
            page_num = image_file.split('_page_')[1].split('.')[0]
            all_results.append(f"## Page {page_num}\n\n{response.text}\n\n---\n")
            
        except Exception as e:
            print(f"Error processing {image_file}: {str(e)}")
            return False
    
    # Combine all results into a single markdown file
    if all_results:
        combined_md = "\n".join(all_results)
        
        # Create output directory if it doesn't exist
        output_dir_main = "output"
        if not os.path.exists(output_dir_main):
            os.makedirs(output_dir_main)
        
        # Create a folder with the PDF filename
        pdf_folder = os.path.join(output_dir_main, pdf_filename)
        if not os.path.exists(pdf_folder):
            os.makedirs(pdf_folder)
        
        # Save the markdown file in the PDF folder
        output_md_path = os.path.join(pdf_folder, "output_all_pages.md")
        
        with open(output_md_path, "w", encoding="utf-8") as f:
            f.write(combined_md)
        
        print(f"\nSuccessfully created markdown report at: {output_md_path}")
        return True
    else:
        print("No results were generated from the images.")
        return False

def convert_pdf_to_images(pdf_path):
    """
    Convert a PDF file to images from the given path.
    
    Args:
        pdf_path (str): Full path to the PDF file to be converted
        
    Returns:
        bool: True if conversion and processing was successful, False otherwise
    """
    # Validate if the PDF file exists
    if not os.path.exists(pdf_path):
        print(f"Error: PDF file not found at path: {pdf_path}")
        return False
    
    # Validate if the file is a PDF
    if not pdf_path.lower().endswith('.pdf'):
        print(f"Error: File is not a PDF: {pdf_path}")
        return False
    
    # Create output directory if it doesn't exist
    output_dir = "images"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    try:
        # Get the PDF filename without extension
        pdf_filename = os.path.splitext(os.path.basename(pdf_path))[0]
        
        # Open PDF with PyMuPDF
        pdf_document = fitz.open(pdf_path)
        total_pages = len(pdf_document)
        
        print(f"Converting {pdf_path} to images...")
        print(f"Total pages: {total_pages}")
        
        # Convert each page to image
        for page_num in range(total_pages):
            # Get the page
            page = pdf_document[page_num]
            
            # Convert page to image with higher resolution (no cropping)
            pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
            
            # Save the full image without cropping
            output_path = os.path.join(output_dir, f"{pdf_filename}_page_{page_num+1}.png")
            pix.save(output_path)
            print(f"Saved page {page_num+1} as {output_path}")
            
            # Clean up
            pix = None
        
        # Close the PDF
        pdf_document.close()
        print(f"Successfully converted {total_pages} pages to PNG format")
        
        # Process images with Gemini
        success = process_images_with_gemini(pdf_filename)
        return success
        
    except Exception as e:
        print(f"Error occurred: {str(e)}")
        return False

if __name__ == "__main__":
    print("Starting PDF processing with completion tracking...")
    print("=" * 50)
    process_input_folder()
    print("=" * 50)
    print("PDF processing completed!")

