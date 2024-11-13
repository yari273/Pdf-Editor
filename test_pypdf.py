import PyPDF2
import pandas as pd
import re

# Function to extract all text from PDF 
def extract_pdf_text(pdf_path):
    all_text = ""
    with open(pdf_path, "rb") as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        for page in pdf_reader.pages:
            text = page.extract_text()
            if text:  # Check if text extraction was successful
                all_text += text + "\n"
    return all_text.strip()

# Process the extracted invoice data and create "Usage Type", "Item No", "Qty", "Amount", and "Description"
def process_invoice_data(extracted_text):
    invoice_info = []
    lines = extracted_text.split("\n")

    # Fields to collect for each item
    sub_id = qty = start_date = end_date = usage_type = end_user = tenant = item_no = amount = None
    description = []
    ri_billing_found = False  
    consumption_found = False  

    output = []  
    output.append("Processing lines for extraction...\n")
    capturing_description = False

    for i, line in enumerate(lines):
        output.append(f"Line {i}: {line}\n")

        if "RI Billing" in line:
            ri_billing_found = True
            output.append("Found usage type: Reservation (RI billing)\n")
        elif "Azure Billing" in line or "AZUREPLAN" in line:
            consumption_found = True
            output.append("Found usage type: Consumption (Azure Billing/AZUREPLAN)\n")

        if "Sub ID:" in line:
            sub_id = line.split(":")[1].strip()
            output.append(f"Extracted Sub ID: {sub_id}\n")
        if "End User:" in line:
            end_user = line.split(":")[1].strip()
            output.append(f"Extracted End User: {end_user}\n")
        if "Tenant:" in line:
            tenant = line.split(":")[1].strip()
            output.append(f"Extracted Tenant: {tenant}\n")
        
        # Adjusted section for Start and End Dates
        if re.search(r"Start\s*Date", line) or re.search(r"End\s*Date", line):
            date_match = re.findall(r'\d{2}\.\d{2}\.\d{4}', line)
            if date_match:
                if re.search(r"Start\s*Date", line):
                    start_date = date_match[0]
                    output.append(f"Extracted Start Date: {start_date}\n")
                if re.search(r"End\s*Date", line):
                    end_date = date_match[1] if len(date_match) > 1 else date_match[0]
                    output.append(f"Extracted End Date: {end_date}\n")

        # Extract Item No, Qty, and Amount from the line using regex
        item_match = re.match(r'^(\d{4})\s+(.+?)\s+(\d+)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$', line.strip())
        if item_match:
            # If a previous item is detected, save the current record and reset fields
            if item_no:
                usage_type = "Reservation" if ri_billing_found else ("Consumption" if consumption_found else "N/A")
                record = {
                    "Sub ID": sub_id,
                    "Qty": qty,
                    "Start Date": start_date,
                    "End Date": end_date,
                    "Usage Type": usage_type,
                    "End User": end_user,
                    "Tenant": tenant,
                    "Item No": item_no,
                    "Amount": amount,
                    "Description": " ".join(description).strip() if description else ""
                }
                invoice_info.append(record)
                output.append(f"Record saved: {record}\n")

                # Reset fields for the next item
                sub_id = start_date = end_date = usage_type = end_user = tenant = item_no = amount = None
                description = []

            # Capture new item information
            item_no = item_match.group(1).strip()
            qty = item_match.group(3).strip()
            amount = item_match.group(4).strip()
            output.append(f"Extracted Item No: {item_no}, Qty: {qty}, Amount: {amount}\n")
            
            # Capturing description lines
            capturing_description = True

        elif capturing_description:
            if any(keyword in line for keyword in ["End User:", "Sub ID:", "Tenant:", "Start Date", "End Date"]):
                capturing_description = False
            else:
                description.append(line.strip())

    # Save the last item if present
    if item_no:
        usage_type = "Reservation" if ri_billing_found else ("Consumption" if consumption_found else "N/A")
        record = {
            "Sub ID": sub_id,
            "Qty": qty,
            "Start Date": start_date,
            "End Date": end_date,
            "Usage Type": usage_type,
            "End User": end_user,
            "Tenant": tenant,
            "Item No": item_no,
            "Amount": amount,
            "Description": " ".join(description).strip() if description else ""
        }
        invoice_info.append(record)
        output.append(f"Record saved: {record}\n")

    return invoice_info, ''.join(output)

# Save the data into an Excel file
def save_to_excel(invoice_data, excel_path):
    if invoice_data:
        df = pd.DataFrame(invoice_data)
        df.to_excel(excel_path, index=False)
        print(f"Invoice data has been extracted and saved to {excel_path}.")
    else:
        print("No valid data to save.")

# Paths for PDF and Excel file
pdf_path = "C:/Users/RishithGatty/Downloads/2215138376_ZIGS.PDF"
excel_path = "C:/Users/RishithGatty/Downloads/Extracted_Invoice_Datatest376.xlsx"

# Extract text from the PDF
text_data = extract_pdf_text(pdf_path)

# Process the extracted text to gather relevant data
invoice_data, output = process_invoice_data(text_data)

# Display the output in the terminal
print(output)

# Save the extracted data to a new Excel file
save_to_excel(invoice_data, excel_path)
