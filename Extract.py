import os
import re
import PyPDF2
import pandas as pd
from tqdm import tqdm

def extract_text_from_pdf(pdf_path):
    # Open the PDF file
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        # Extract text from each page
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            text += page.extract_text()
    return text

def extract_invoice_details(text):
    # Extract Order Information
    order_id = re.search(r"Order ID:\s*(\d+)", text)
    customer_id = re.search(r"Customer ID:\s*(\w+)", text)
    order_date = re.search(r"Order Date:\s*([\d-]+)", text)
    
    # Extract Customer Details
    contact_name = re.search(r"Contact Name:\s*(.+)", text)
    address = re.search(r"Address:\s*(.+)", text)
    city = re.search(r"City:\s*(.+)", text)
    postal_code = re.search(r"Postal Code:\s*(\d{5}-\d{3})", text)
    country = re.search(r"Country:\s*(.+)", text)
    phone = re.search(r"Phone:\s*(.+)", text)
    fax = re.search(r"Fax:\s*(.+)", text)
    
    # Extract Product Details (Handling variable number of products)
    product_pattern = r"(\d+)\s+(.+?)\s+(\d+)\s+([\d.]+)"  # For Product ID, Name, Quantity, Unit Price
    products = re.findall(product_pattern, text)
    
    # Extract Total Price
    total_price = re.search(r"TotalPrice\s+([\d.]+)", text)
    
    invoice_data = {
        "Order ID": order_id.group(1) if order_id else None,
        "Customer ID": customer_id.group(1) if customer_id else None,
        "Order Date": order_date.group(1) if order_date else None,
        "Contact Name": contact_name.group(1) if contact_name else None,
        "Address": address.group(1) if address else None,
        "City": city.group(1) if city else None,
        "Postal Code": postal_code.group(1) if postal_code else None,
        "Country": country.group(1) if country else None,
        "Phone": phone.group(1) if phone else None,
        "Fax": fax.group(1) if fax else None,
        "Products": products,
        "Total Price": total_price.group(1) if total_price else None
    }
    
    return invoice_data

def process_pdfs_in_directory(directory_path):
    all_invoices = []
    
    # List all files in the directory
    pdf_files = [f for f in os.listdir(directory_path) if f.endswith(".pdf")]
    
    # Loop through all PDF files in the directory with a progress bar
    for file_name in tqdm(pdf_files, desc="Processing PDFs"):
        file_path = os.path.join(directory_path, file_name)
        
        # Extract text from the PDF
        pdf_text = extract_text_from_pdf(file_path)
        
        # Extract invoice details
        invoice_details = extract_invoice_details(pdf_text)
        
        # Store each invoice's data
        all_invoices.append(invoice_details)
    
    return all_invoices

def save_to_excel(invoices, output_file):
    # Prepare the DataFrame to store the invoices
    data = []
    for invoice in invoices:
        for product in invoice["Products"]:
            product_id, product_name, quantity, unit_price = product
            data.append({
                "Order ID": invoice["Order ID"],
                "Customer ID": invoice["Customer ID"],
                "Order Date": invoice["Order Date"],
                "Contact Name": invoice["Contact Name"],
                "Address": invoice["Address"],
                "City": invoice["City"],
                "Postal Code": invoice["Postal Code"],
                "Country": invoice["Country"],
                "Phone": invoice["Phone"],
                "Fax": invoice["Fax"],
                "Product ID": product_id,
                "Product Name": product_name,
                "Quantity": quantity,
                "Unit Price": unit_price,
                "Total Price": invoice["Total Price"]
            })
    
    df = pd.DataFrame(data)
    
    # Save the DataFrame to an Excel file using the openpyxl engine
    df.to_excel(output_file, index=False, engine='openpyxl')

# Example usage
directory_path = "Documents/invoices"  # Replace with your directory path containing the PDF files
output_excel_file = "Output/All_invoices.xlsx"  # Replace with your desired Excel file path

# Process all PDFs in the directory
all_invoices = process_pdfs_in_directory(directory_path)

# Save the extracted information to an Excel file
save_to_excel(all_invoices, output_excel_file)

print(f"Extracted invoice data saved to {output_excel_file}")
