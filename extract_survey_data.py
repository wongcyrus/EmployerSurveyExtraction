import os
import zipfile
import pandas as pd
import json
from tqdm import tqdm
import base64
import vertexai
from vertexai.generative_models import GenerativeModel, Part

# --- GCP Configuration ---
# IMPORTANT: Please update these values with your GCP project and region
GCP_PROJECT_ID = "ai-invigilator-hkiit"
GCP_REGION = "us-central1"

# --- File Configuration ---
ZIP_FILE_PATH = "Data/ITE4116M_IT_ICT_kcheung_1-(3) Appendix6(for Company)-119607.zip"
EXTRACT_DIR = "Data/extracted"
EXCEL_OUTPUT_PATH = "survey_data.xlsx"
JSON_OUTPUT_DIR = "Data/extracted_json"
FIELDS_FILE_PATH = "fields.txt"

# --- Vertex AI Initialization ---
# This script uses Application Default Credentials (ADC).
# Before running, authenticate via the gcloud CLI:
# gcloud auth application-default login
try:
    vertexai.init(project=GCP_PROJECT_ID, location=GCP_REGION)
except Exception as e:
    print(f"ERROR: Vertex AI initialization failed. Have you run 'gcloud auth application-default login'?")
    print(f"Please also ensure you have set your GCP_PROJECT_ID and GCP_REGION in the script.")
    print(f"Error details: {e}")
    exit()

# --- Field Extraction Definition ---
def load_fields_from_file(file_path):
    """Loads and parses the fields to extract from a text file."""
    print(f"Loading fields from {file_path}...")
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        fields = []
        for line in content.splitlines():
            for field in line.split('\t'):
                cleaned_field = field.strip()
                if cleaned_field:
                    fields.append(cleaned_field)
        
        print(f"Found {len(fields)} fields to extract.")
        return fields
    except FileNotFoundError:
        print(f"ERROR: Fields file not found at '{file_path}'. Please create it.")
        print("The file should contain the fields to extract, separated by newlines or tabs.")
        exit()

FIELDS_TO_EXTRACT = load_fields_from_file(FIELDS_FILE_PATH)

def unzip_file(zip_path, extract_to):
    """Unzips a file to a specified directory."""
    print(f"Unzipping {zip_path} to {extract_to}...")
    if not os.path.exists(extract_to):
        os.makedirs(extract_to)
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)
    print("Unzipping complete.")

def find_pdfs(directory):
    """Recursively finds all PDF files in a directory."""
    pdf_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    print(f"Found {len(pdf_files)} PDF files.")
    return pdf_files

def extract_data_from_pdf(pdf_path, model):
    """Uses Vertex AI Gemini to extract structured data from a PDF."""
    print(f"Processing {os.path.basename(pdf_path)}...")
    try:
        # Read and encode the PDF file
        with open(pdf_path, "rb") as f:
            pdf_content = f.read()
        
        pdf_part = Part.from_data(
            data=pdf_content,
            mime_type="application/pdf"
        )

        # Create the prompt for Gemini
        prompt = f"""
        Please analyze the provided PDF document, which is a completed employer survey form.
        Extract the following fields and return the data as a single, minified JSON object.
        The field names in the JSON output should be exactly as listed below.

        Follow these rules for extracting values:
        1. For fields that are ratings on a numerical scale (like "Communication skills", "Teamwork", etc., which are typically on a 1-10 scale), you must extract the selected numerical value. If a value for such a rating is truly not present or cannot be determined, you must use the neutral value "5" as a reasonable guess.
        2. For all other fields, if a value is not found or is empty, use the string "N/A".


        Fields to extract:
        {json.dumps(FIELDS_TO_EXTRACT, indent=2, ensure_ascii=False)}

        Your output should be only the JSON object, with no other text before or after it.
        """

        # Call the Gemini API
        response = model.generate_content([pdf_part, prompt])

        # Clean up the response and parse the JSON
        cleaned_response = response.text.strip().replace('```json', '').replace('```', '')
        data = json.loads(cleaned_response)
        return data

    except Exception as e:
        print(f"  Error processing {os.path.basename(pdf_path)}: {e}")
        return None

def main():
    """Main function to run the extraction process."""
    # 1. Unzip the data file
    unzip_file(ZIP_FILE_PATH, EXTRACT_DIR)

    # 2. Find all PDF files
    pdf_files = find_pdfs(EXTRACT_DIR)

    if not pdf_files:
        print("No PDF files found. Exiting.")
        return

    # 3. Initialize Gemini model on Vertex AI
    model = GenerativeModel("gemini-2.5-flash")

    # 4. Process each PDF, save intermediate JSON results
    os.makedirs(JSON_OUTPUT_DIR, exist_ok=True)
    for pdf_path in tqdm(pdf_files, desc="Extracting data from PDFs"):
        pdf_basename = os.path.basename(pdf_path)
        # Create a unique filename from the relative path to avoid collisions
        relative_path = os.path.relpath(pdf_path, EXTRACT_DIR)
        unique_filename_base = relative_path.replace(os.sep, '_').rsplit('.', 1)[0]
        json_filename = f"{unique_filename_base}.json"
        json_path = os.path.join(JSON_OUTPUT_DIR, json_filename)

        if os.path.exists(json_path):
            print(f"Skipping {pdf_basename}, JSON output already exists at {json_path}")
            continue

        extracted_data = extract_data_from_pdf(pdf_path, model)
        if extracted_data:
            try:
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(extracted_data, f, ensure_ascii=False, indent=4)
                print(f"  Successfully saved data for {pdf_basename} to {json_path}")
            except Exception as e:
                print(f"  Error saving JSON for {pdf_basename}: {e}")


    # 5. Consolidate all JSON data
    all_records = []
    json_file_paths = [os.path.join(JSON_OUTPUT_DIR, f) for f in os.listdir(JSON_OUTPUT_DIR) if f.endswith('.json')]
    
    if not json_file_paths:
        print("No JSON data files found to consolidate. Exiting.")
        return

    for json_path in tqdm(json_file_paths, desc="Consolidating JSON data"):
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                all_records.append(data)
        except Exception as e:
            print(f"Error loading {json_path}: {e}")


    if not all_records:
        print("No data could be extracted. Exiting.")
        return

    # 6. Export to Excel
    print(f"Exporting {len(all_records)} records to {EXCEL_OUTPUT_PATH}...")
    df = pd.DataFrame(all_records)
    
    # Reorder columns to match the requested field list
    # Ensure all columns are present, adding missing ones with None
    for col in FIELDS_TO_EXTRACT:
        if col not in df.columns:
            df[col] = None
            
    df = df[FIELDS_TO_EXTRACT]
    
    df.to_excel(EXCEL_OUTPUT_PATH, index=False, engine='openpyxl')
    print("Export complete.")

if __name__ == "__main__":
    main()