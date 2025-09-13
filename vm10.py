import pytesseract
from PIL import Image
import requests
import json
import gspread
from google.oauth2.service_account import Credentials
import os
from pdf2image import convert_from_path
import re
import shutil
import pandas as pd

# === OCR Setup ===
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# === Paths and Config ===
unprocessed_image_folder = r"C:\******\*******\*******\********\*** ********\***********\unprocessed images"
processed_image_folder = r"C:\******\*******\******\********\***** ******\***********\processed images"

contracts_path = r"C:***********.txt"
specialties_path = r"C:***********.txt"
service_account_path = r"***********.json"
sheet_name = "ACPN AI Automation Sheet"

# === Load Specialty List ===
with open(specialties_path, "r", encoding="utf-8") as f:
    specialties_list = [line.strip() for line in f if line.strip()]

# === Google Sheets Auth ===
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file(service_account_path, scopes=SCOPES)
client = gspread.authorize(creds)
worksheet = client.open(sheet_name).sheet1

# === Get all image files ===
image_paths = [os.path.join(unprocessed_image_folder, f) for f in os.listdir(unprocessed_image_folder)
              if f.lower().endswith(('.png', '.jpg', '.jpeg', '.pdf', '.xlsx'))]

def remove_underscores(s):
    return s.replace("_", "") if isinstance(s, str) else s

import re

def flatten_ocr_text(ocr_text: str) -> str:
    lines = [line.strip() for line in ocr_text.splitlines() if line.strip()]
    flattened_lines = []
    i = 0

    while i < len(lines):
        current = lines[i]
        next_line = lines[i + 1] if i + 1 < len(lines) else ""

        # Match name line followed by label line
        if re.match(r'^[A-Z][a-z]+\s+[A-Z][a-z]+(?:\s+[A-Z])?\s+[A-Z]+$', current) and \
           'Provider Last/First' in next_line:
            parts = current.split()
            last, first = parts[0], parts[1]
            middle = parts[2] if len(parts) == 4 else None
            degree = parts[-1]

            flattened_lines.append(f"Provider's First Name: {first}")
            flattened_lines.append(f"Provider's Last Name: {last}")
            flattened_lines.append(f"Provider's Middle Initial: {middle or 'null'}")
            flattened_lines.append(f"Degree: {degree}")
            i += 2
            continue

        # Facility name
        if 'Practice/Facility Name' in next_line:
            flattened_lines.append(f"Facility/Group Name: {current}")
            i += 2
            continue

        # Tax ID, NPI, CAQH (3 numeric values in one line)
        if re.match(r'^\d{2,}-\d{2,}\s+\d{9}\s+\d{8}$', current):
            parts = current.split()
            flattened_lines.append(f"Provider Tax ID: {parts[0].replace('-', '')}")
            flattened_lines.append(f"Provider Billing NPI: {parts[1]}")
            flattened_lines.append(f"CAQH Number: {parts[2]}")
            i += 1
            continue

        # Claim Format
        if 'CLAIM TYPE SUBMISSION' in current:
            match = re.search(r'(HCFA|UB|BOTH)', current, re.IGNORECASE)
            if match:
                flattened_lines.append(f"Claim Format: {match.group(1).upper()}")
            i += 1
            continue

        # Email extraction
        email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', current)
        if email_match:
            flattened_lines.append(f"Location Email: {email_match.group(0)}")

        # Billing address line
        if current.upper().startswith("PO BOX") or "P.O. BOX" in current.upper():
            flattened_lines.append(f"Billing Address: {current}")
            i += 1
            continue

        # Zipcode/City/State line
        if re.search(r'\b[A-Z]{2}\s+\d{5}$', current):
            city_state_zip = current.split()
            flattened_lines.append(f"Billing City: {' '.join(city_state_zip[:-2])}")
            flattened_lines.append(f"Billing State: {city_state_zip[-2]}")
            flattened_lines.append(f"Billing Zipcode: {city_state_zip[-1]}")
            i += 1
            continue

        # Phone line
        if re.match(r'\d{3}-\d{3}-\d{4}', current.replace(" ", "")):
            flattened_lines.append(f"Billing Phone Number: {current}")
            i += 1
            continue

        # Otherwise, keep the original
        flattened_lines.append(current)
        i += 1

    return "\n".join(flattened_lines)


def write_into_cell(worksheet, next_row, entry, data_name, column):
    try:
        value = entry.get(data_name, "null")
        if value and value.lower() != "null":
            worksheet.update_cell(next_row, column, remove_underscores(value))
    except Exception as e:
        print(f"Error writing '{data_name}' to cell {column}: {e}")

for img_path in image_paths:
    try:
        print(f"\n Processing image: {img_path}")
        # === OCR ===
        # image = Image.open(img_path)
        # text = pytesseract.image_to_string(image)
        image_pages = []
        all_text = ""

        if img_path.lower().endswith(".xlsx"):
            print(f"\n Processing Excel file: {img_path}")
            excel_data = pd.read_excel(img_path, sheet_name=None)
            for sheet_name, df in excel_data.items():
                all_text += f"\n--- Sheet: {sheet_name} ---\n"
                for _, row in df.iterrows():
                    row_text = " | ".join([f"{col}: {val}" for col, val in row.items() if pd.notna(val)])
                    all_text += row_text + "\n"
            hint = "This was a sheet of information and each row is a different person"
            all_text = all_text

        elif img_path.lower().endswith(".pdf"):
            image_pages = convert_from_path(
                img_path, dpi=300,
                poppler_path=r"C:\Users\bkmed\OneDrive\Desktop\ACPN Automation\Release-24.08.0-0\poppler-24.08.0\Library\bin"
            )
            for page_num, image in enumerate(image_pages, start=1):
                text = pytesseract.image_to_string(image)
                all_text += f"\n--- Page {page_num} ---\n{text}"

        elif img_path.lower().endswith(('.png', '.jpg', '.jpeg')):
            image = Image.open(img_path)
            text = pytesseract.image_to_string(image)
            all_text += text

        else:
            print("⚠️ Unsupported file type.")
            continue
        print("file processed")
        print(all_text)
        
        # === Build LLM Extraction Prompt ===
        extract_prompt = f"""
        You are a document data extractor for a health care provider network. Here's the OCR text from a form:

        {all_text}

        Extract the following fields (return every field given to you, if you can not find a field DO NOT WRITE ANYTHING, do not alter the fields name in any way, remove any undersocre _ from extraction):
        - Email of Sender: 
        - Provider's First Name: (Must Find)
        - Provider's Middle Initial: (usually just one letter, may not have one)
        - Provider's Last Name: (Must Find)
        - Facility/Group Name:
        - Date of Birth: (as a ##/##/##, may not be listed)
        - Degree: (example: MD, NO, PA, CRNA, DO, NP,FNP etc)
        - Gender: MUST be Male or Female do not say null. If the gender is not explicitly mentioned in the OCR text, YOU MUST MAKE A BEST GUESS based on the first name. Use common name-to-gender associations (e.g., 'Emily' is typically female, 'John' is typically male). DO NOT return "Unknown".
        - Contract Name:
        - Effective Date:
        - Provider Tax ID: (remove the -)
        - Group NPI: (different number than provider billing NPI)
        - Provider Billing NPI: (different number than group NPI)
        - Provider Type: (Hospital, Facility, or Physician — if not provided, make your best guess; most likely is a Physician)
        - Claim Format: (HAS TO BE either HCFA, UB, or Both — may not be mentioned)

        - Location Address 1:
        - Location Suite 1: 
        - Location Zipcode 1:
        - Location City 1:
        - Location State (Abbreviation) 1:
        - Location Phone Number 1:
        - Location Fax 1:
        - Location Email 1: (may not be provided)

            (There may only be one locations if so skip the rest of the locations)

        - Location Address 2:
        - Location Suite 2:
        - Location Zipcode 2:
        - Location City 2:
        - Location State (Abbreviation) 2 :
        - Location Phone Number 2: 
        - Location Fax 2:
        - Location Email 2: (may not be provided)

        - Location Address 3:
        - Location Suite 3:
        - Location Zipcode 3:
        - Location City 3:
        - Location State (Abbreviation) 3:
        - Location Phone Number 3:
        - Location Fax 3:
        - Location Email 3: (may not be provided)

        - Location Address 4:
        - Location Suite 4:
        - Location Zipcode 4:
        - Location City 4:
        - Location State (Abbreviation) 4:
        - Location Phone Number 4:
        - Location Fax 4:
        - Location Email 4: (may not be provided)

        - Location Address 5:
        - Location Suite 5:
        - Location Zipcode 5:
        - Location City 5:
        - Location State (Abbreviation) 5:
        - Location Phone Number 5:
        - Location Fax 5:
        - Location Email 5: (may not be provided)

        - Billing Address: (usually a P.O. Box — include only the P.O. Box or address)
        - Billing Zipcode:
        - Billing City:
        - Billing State (Abbreviation):
        - Billing Phone Number:
        - Billing Fax:
        - Billing Email:

        - Primary Specialty: (If not stated decided based of their other information, for example their degree; if NP then Nurse Practitioner, if FNP then Family Nurse Practitioner-Certified)
        - Secondary Specialty: (may not be provided)
        - Tertiary Specialty: (may not be provided)


        Return the result as **strict JSON only**. Do not include any comments, explanations, no commas, or extra text — only the JSON object.
        """

        try:
            print(" Sending request to LLM")
            response = requests.post(
                'http://localhost:11434/api/generate',
                json={'model': 'mistral', 'prompt': extract_prompt, 'temperature': 0, 'stream': False},
                timeout=60  # prevent freezing forever
            )
            print("LLM responded")
        except requests.exceptions.Timeout:
            print("Request to LLM timed out.")
            continue
        except requests.exceptions.RequestException as e:
            print(f"Request to LLM failed: {e}")
            continue

        print("extracted")

        if response.status_code != 200:
            print(" LLM extraction failed.")
            continue
        
        raw_response = response.json()['response']
        print(" Raw LLM Response:\n", raw_response)

        entry = json.loads(response.json()['response'])

        # === Specialty Normalization ===
        specialty_prompt = f"""
        You are a data normalizer. Your job is to take raw specialty names and match them to a provided list of standard specialties.

        Here is the list of standard specialties:
        {json.dumps(specialties_list, indent=2)}

        Now, for the following raw specialties, provide the best match from the list above. If there is no good match, return the original value.
        IF THE SPECIALTY IS NULL DO NOT TRY TO MATCH AND LEAVE IT null.

        Respond in this format as JSON:
        {{
          "Primary Specialty": "...",
          "Secondary Specialty": "...",
          "Tertiary Specialty": "..."
        }}

        Raw specialties:
        Primary Specialty: {entry.get("Primary Specialty", "null")}
        Secondary Specialty: {entry.get("Secondary Specialty", "null")}
        Tertiary Specialty: {entry.get("Tertiary Specialty", "null")}
        """

        specialty_response = requests.post(
            'http://localhost:11434/api/generate',
            json={'model': 'mistral', 'prompt': specialty_prompt, 'temperature': 0, 'stream': False}
        )

        if specialty_response.status_code == 200:
            matched = json.loads(specialty_response.json()['response'])
            entry["Primary Specialty"] = matched.get("Primary Specialty", entry.get("Primary Specialty"))
            entry["Secondary Specialty"] = matched.get("Secondary Specialty", entry.get("Secondary Specialty"))
            entry["Tertiary Specialty"] = matched.get("Tertiary Specialty", entry.get("Tertiary Specialty"))
        else:
            print(" Specialty matching failed.")

        # === Write to Sheet ===
        next_row = len(worksheet.get_all_values()) + 1
        write_map = {
            "Email of Sender": 3, "Provider's Last Name": 4, "Provider's First Name": 5, "Provider's Middle Initial": 6, "Degree": 7, "Effective Date": 8,
            "Gender": 9, "Provider Tax ID": 10, "Group NPI": 11, "Provider Billing NPI": 12,
            "Provider Type": 13, "Claim Format": 14, "Location Address 1": 15, "Location Suite 1": 16,
            "Location Zipcode 1": 17, "Location City 1": 18, "Location State 1": 19,
            "Location Phone Number 1": 20, "Location Fax 1": 21, "Billing Address": 22,
            "Billing Zipcode": 23, "Billing City": 24, "Billing State": 25, "Billing Phone Number": 26,
            "Billing Fax": 27, "Primary Specialty": 28, "Secondary Specialty": 29, "Secondary Specialty": 30,
            "Facility/Group Name": 31, "Location Address 2": 33, "Location Suite 2": 34,
            "Location Zipcode 2": 35, "Location City 2": 36, "Location State 2": 37,
            "Location Phone Number 2": 38, "Location Fax 2": 39, "Location Address 3": 40,
            "Location Suite 3": 41, "Location Zipcode 3": 42, "Location City 3": 43,
            "Location State 3": 44, "Location Phone Number 3": 45, "Location Fax 3": 46, "Location Address 4": 47,
            "Location Suite 4": 49, "Location Zipcode 4": 49, "Location City 4": 50,
            "Location State 4": 51, "Location Phone Number 4": 52, "Location Fax 4": 53, "Location Address 5": 54,
            "Location Suite 5": 55, "Location Zipcode 5": 56, "Location City 5": 57,
            "Location State 5": 58, "Location Phone Number 5": 59, "Location Fax 5": 60
        }

        for field, col in write_map.items():
            write_into_cell(worksheet, next_row, entry, field, col)

        print("Finished writing to sheet.")

        # moves images to processed folder
        try:
            filename = os.path.basename(img_path)
            dest_path = os.path.join(processed_image_folder, filename)
            shutil.move(img_path, dest_path)
            print(f" Moved to processed folder: {dest_path}")
        except Exception as move_err:
            print(f" Failed to move image: {img_path} -> {processed_image_folder}\n{move_err}")

    except Exception as e:
        print(f" Error processing {img_path}: {e}")


