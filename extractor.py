import os
import argparse
from docx import Document
import csv
import re
from datetime import datetime

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

def extract_from_docx(file_path):
    """Extract structured entries from a .docx file."""
    doc = Document(file_path)
    lines = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    return extract_structured_data(lines)

def extract_from_pdf(file_path):
    """Extract structured entries from a .pdf file."""
    if not pdfplumber:
        print("❌ pdfplumber is not installed. Please run: pip install pdfplumber")
        return []
    lines = []
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines.extend([line.strip() for line in text.split("\n") if line.strip()])
    return extract_structured_data(lines)

def extract_structured_data(lines):
    """Extract entries using labeled or block-based formats."""
    data = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if re.match(r"(?i)^name[\s:.-]*", line):
            entry = {"Full Name": "", "Job Title": "", "Company": "", "Email": ""}
            entry["Full Name"] = re.sub(r"(?i)^name[\s:.-]*", "", line).strip()

            for j in range(i + 1, len(lines)):
                if re.match(r"(?i)^name[\s:.-]*", lines[j]):
                    break  # stop if next block begins
                elif re.match(r"(?i)^(title|role|position)[\s:.-]*", lines[j]):
                    entry["Job Title"] = re.sub(r"(?i)^(title|role|position)[\s:.-]*", "", lines[j]).strip()
                elif re.match(r"(?i)^(company|organization)[\s:.-]*", lines[j]):
                    entry["Company"] = re.sub(r"(?i)^(company|organization)[\s:.-]*", "", lines[j]).strip()
                elif re.match(r"(?i)^email[\s:.-]*", lines[j]) or "@" in lines[j]:
                    email_candidate = re.sub(r"(?i)^email[\s:.-]*", "", lines[j]).strip()
                    if "@" in email_candidate:
                        entry["Email"] = email_candidate

            data.append(entry)
            i = j  # jump to next potential block
        else:
            i += 1

    return data

def save_to_csv(data, base_name="output"):
    """Save the extracted data to a timestamped CSV file."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{base_name}_{timestamp}.csv"
    with open(filename, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=["Full Name", "Job Title", "Company", "Email"])
        writer.writeheader()
        writer.writerows(data)
    print(f"\n✅ Extracted {len(data)} entries and saved to: {filename}")

def main():
    parser = argparse.ArgumentParser(description="Extract structured data from PDF or DOCX into a CSV.")
    parser.add_argument("filepath", help="Path to the PDF or DOCX file")
    args = parser.parse_args()

    file_path = args.filepath
    if not os.path.isfile(file_path):
        print("❌ File does not exist.")
        return

    if file_path.lower().endswith(".docx"):
        print("✅ DOCX file detected. Ready for extraction.")
        data = extract_from_docx(file_path)

    elif file_path.lower().endswith(".pdf"):
        print("✅ PDF file detected. Ready for extraction.")
        data = extract_from_pdf(file_path)

    else:
        print("❌ Unsupported file type. Use .docx or .pdf")
        return

    if data:
        save_to_csv(data)
    else:
        print("⚠️ No entries found in the document.")

if __name__ == "__main__":
    main()


