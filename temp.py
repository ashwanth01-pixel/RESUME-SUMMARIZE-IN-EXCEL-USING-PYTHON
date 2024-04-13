import os
import docx
import pandas as pd
import re
from PyPDF2 import PdfReader

def extract_text_from_docx(docx_path):
    doc = docx.Document(docx_path)
    text = ""
    for paragraph in doc.paragraphs:
        text += paragraph.text
    return text

def extract_text_from_pdf(pdf_path):
    text = ""
    with open(pdf_path, 'rb') as file:
        reader = PdfReader(file)
        for page in reader.pages:
            text += page.extract_text()
    return text

def extract_email(text):
    email_pattern = r'[\w\.-]+@[\w\.-]+'
    emails = re.findall(email_pattern, text)
    return emails[0] if emails else None

def extract_phone_number(text):
    phone_pattern = r'(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4})'
    phones = re.findall(phone_pattern, text)
    return phones[0] if phones else None

def extract_summary(text):
    # Implement your logic for extracting summary here
    # For example, you can extract the first few sentences or paragraphs
    summary = text[:200]  # Example: Extracting first 200 characters
    return summary

def process_cv(cv_path):
    if cv_path.endswith('.docx'):
        text = extract_text_from_docx(cv_path)
    elif cv_path.endswith('.pdf'):
        text = extract_text_from_pdf(cv_path)
    else:
        print("Unsupported file format")
        return None, None, None

    email = extract_email(text)
    phone = extract_phone_number(text)
    summary = extract_summary(text)
    return email, phone, summary

def main():
    dataset_folder = 'dataset'  # Folder containing resumes
    output_excel = 'cv_data.xlsx'  # Output Excel file

    # Initialize lists to store data
    emails = []
    phones = []
    summaries = []

    # Iterate over files in the dataset folder
    for filename in os.listdir(dataset_folder):
        if filename.endswith('.docx') or filename.endswith('.pdf'):
            cv_path = os.path.join(dataset_folder, filename)
            email, phone, summary = process_cv(cv_path)
            emails.append(email)
            phones.append(phone)
            summaries.append(summary)

    # Create a DataFrame from the extracted data
    df = pd.DataFrame({
        'Email': emails,
        'Phone': phones,
        'Summary': summaries
    })

    # Save the DataFrame to an Excel file
    df.to_excel(output_excel, index=False)

if __name__ == "__main__":
    main()
    print("sucessfully extracted and saved in excel")
