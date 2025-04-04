import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
import re

def extract_text_from_pdf(pdf_path):
    """Extract raw text from a PDF using PyMuPDF (fitz)."""
    doc = fitz.open(pdf_path)
    text_data = []
    
    for page in doc:
        text = page.get_text("text")  # Extract text from page
        text_data.append(text)
    
    return "\n".join(text_data)  # Combine text from all pages

def extract_tables_from_pdf(pdf_path):
    """Extract tables from a PDF using pdfplumber."""
    extracted_tables = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()  # Detect tables
            for table in tables:
                structured_table = [[cell if cell else "" for cell in row] for row in table]  # Ensure no None values
                extracted_tables.append(structured_table)

    return extracted_tables

def process_text_into_table(text_data):
    """Convert unstructured text into tabular format (for PDFs without proper tables)."""
    lines = text_data.split("\n")
    structured_data = []

    for line in lines:
        line = line.strip()
        if line:  
            row = re.split(r'\s{2,}', line)  # Split on multiple spaces
            structured_data.append([cell.encode("utf-8", "ignore").decode("utf-8") for cell in row])  # Ensure UTF-8

    return structured_data

def save_to_excel(data_list, output_path):
    """Save extracted tables to an Excel file with UTF-8 encoding."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for i, data in enumerate(data_list):
            df = pd.DataFrame(data)
            
            # Use .map() instead of .applymap()
            df = df.map(lambda x: x.encode("utf-8", "ignore").decode("utf-8") if isinstance(x, str) else x)  

            sheet_name = f"Table_{i+1}"
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    print(f"✅ Data saved to {output_path}")

def main(pdf_path, output_excel):
    """Main function to extract structured tables from any given PDF."""
    print("🔄 Extracting text-based tables...")
    text_data = extract_text_from_pdf(pdf_path)
    text_tables = process_text_into_table(text_data)

    print("🔄 Extracting embedded tables...")
    structured_tables = extract_tables_from_pdf(pdf_path)

    all_data = structured_tables + [text_tables]  # Combine extracted tables

    print("🔄 Saving to Excel with UTF-8 encoding...")
    save_to_excel(all_data, output_excel)

    print("✅ Table extraction complete!")

# Example Usage
pdf_path = r"C:\Users\Naman\Downloads\test3 (1) (1).pdf"  # Replace with actual PDF
output_excel = "extracted_tables.xlsx"
main(pdf_path, output_excel)
