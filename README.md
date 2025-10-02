# Employer Survey Extraction

## Overview

This project automates the extraction of data from employer survey forms in PDF format. It uses Google's Gemini Pro vision model to analyze the PDFs, extract structured data based on a dynamic list of fields, and saves the results in both intermediate JSON files and a final consolidated Excel spreadsheet.

## Features

- **Automated Data Extraction**: Extracts data from hundreds of PDF survey forms.
- **AI-Powered**: Leverages the Gemini Pro vision model for intelligent data extraction.
- **Configurable Fields**: The fields to be extracted can be easily configured in an external text file (`fields.txt`).
- **Robust and Resumable**: Saves the result of each processed PDF individually, allowing the process to be resumed without losing progress. It automatically skips already-processed files.
- **Intelligent Handling of Missing Data**:
    - For numerical rating fields, it can be configured to infer a neutral value if one is not found.
    - For other fields, it uses a clear "N/A" placeholder for missing data.
- **Consolidated Output**: Aggregates all extracted data into a single, easy-to-use Excel file.

## Prerequisites

Before you begin, ensure you have the following installed:

- Python 3.x
- Google Cloud SDK (`gcloud`)

You will also need a Google Cloud Platform (GCP) project with the Vertex AI API enabled.

## Configuration

1.  **GCP Authentication**: Authenticate your local environment with GCP by running the following command in your terminal:
    ```bash
    gcloud auth application-default login
    ```

2.  **Python Dependencies**: Install the required Python libraries from `requirements.txt`:
    ```bash
    pip install -r requirements.txt
    ```

3.  **Script Configuration**: Open the `extract_survey_data.py` file and update the following GCP configuration variables at the top of the file:
    - `GCP_PROJECT_ID`: Your Google Cloud project ID.
    - `GCP_REGION`: The GCP region for your project (e.g., "us-central1").

4.  **Define Extraction Fields**: Open the `fields.txt` file. List all the fields you want to extract from the PDFs. The fields can be separated by newlines or tabs. The script will parse this file to determine what data to look for.

5.  **Input Data**: Place the zip file containing your PDF surveys into the `Data/` directory. The script is currently configured to look for `Data/ITE4116M_IT_ICT_kcheung_1-(3) Appendix6(for Company)-119607.zip`. You can change the `ZIP_FILE_PATH` variable in the script if your file has a different name.

## Usage

Once the configuration is complete, run the extraction script from your terminal:

```bash
python extract_survey_data.py
```

The script will display its progress as it unzips the files, processes each PDF, and saves the data.

## Output

The script generates the following outputs:

-   **Intermediate JSON Files**: For each PDF processed, a corresponding `.json` file is created in the `Data/extracted_json/` directory. These files contain the structured data extracted from a single PDF. This allows for easy debugging and prevents data loss if the script is interrupted.
-   **Final Excel File**: A file named `survey_data.xlsx` is created in the root directory. This file contains all the extracted data from all the PDFs, consolidated into a single spreadsheet for analysis.
