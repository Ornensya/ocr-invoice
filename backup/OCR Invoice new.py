import streamlit as st
import pytesseract
import json
import pandas as pd
from io import BytesIO
from pdf2image import convert_from_bytes
from openai import OpenAI
import os
from streamlit_pdf_viewer import pdf_viewer
from streamlit import session_state as ss
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

load_dotenv()
# Initialize OpenAI API
api_key = os.getenv("OPENAI_API_KEY")
client = None
if api_key:
    client = OpenAI(api_key=api_key)

poppler_path = r"C:\Program Files\poppler-24.07.0\Library\bin"


# # Fungsi OCR untuk ekstraksi teks dari PDF (multi-halaman)
# def extract_text_from_pdf(pdf_file):
#     images = convert_from_bytes(pdf_file.read())
#     extracted_text = "\n".join([pytesseract.image_to_string(img) for img in images])
#     return extracted_text

def extract_text_from_pdf(pdf_file):
    images = convert_from_bytes(pdf_file.read())
    st.write("### Hasil Gambar dari PDF")
    for i, img in enumerate(images):
        st.image(img, caption=f'Halaman {i+1}', use_container_width=True)
    text = "\n".join([pytesseract.image_to_string(img) for img in images])
    print("ISI TEKS EXTRACT", text)
    return text


def save_to_excel(structured_invoice_data):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    # Extract data
    vendor = structured_invoice_data.get("vendor", {})
    invoice = structured_invoice_data.get("invoice", {})
    customer = structured_invoice_data.get("customer", {})
    items = structured_invoice_data.get("items", [])
    payment = structured_invoice_data.get("payment", {})
    subtotal = structured_invoice_data.get("subtotal", "")
    tax = structured_invoice_data.get("tax", "")
    total = structured_invoice_data.get("total", "")
    currency = structured_invoice_data.get("currency", "")

    # Write Vendor section
    ws.append(["Vendor Information"])
    ws.append(["Name", vendor.get("name", "")])
    ws.append(["Address", vendor.get("address", "")])
    ws.append(["Phone", vendor.get("phone", "")])
    ws.append(["Email", vendor.get("email", "")])
    ws.append([])

    # Write Invoice section
    ws.append(["Invoice Information"])
    ws.append(["Invoice Number", invoice.get("invoice_number", "")])
    ws.append(["Invoice Date", invoice.get("invoice_date", "")])
    ws.append(["Due Date", invoice.get("idue_date", "")])
    ws.append([])

    # Write Customer section
    ws.append(["Customer Information"])
    ws.append(["Name", customer.get("name", "")])
    ws.append(["Address", customer.get("address", "")])
    ws.append(["Phone", customer.get("phone", "")])
    ws.append(["Email", customer.get("email", "")])
    ws.append([])

    # Write Items table
    if items:
        ws.append(["Items"])
        item_df = pd.DataFrame(items)
        for row in dataframe_to_rows(item_df, index=False, header=True):
            ws.append(row)
        ws.append([])

    # Write Payment & Totals
    ws.append(["Payment Information"])
    ws.append(["Bank", payment.get("bank", "")])
    ws.append(["Account Number", payment.get("account_number", "")])
    ws.append([])

    ws.append(["Subtotal", subtotal])
    ws.append(["Tax", tax])
    ws.append(["Total", total])
    ws.append(["Currency", currency])

    # Save to BytesIO
    wb.save(output)
    output.seek(0)
    return output


# Function to structure CV data using OpenAI
def structure_invoice_data(extracted_text):
    if not client:
        return {"error": "OpenAI API is not available. Make sure the API key is properly configured."}
    prompt = f"""
    You are a financial assistant. Based on the following extracted invoice text, convert it into a clean and structured JSON format.

    Extract the following information:
    - Vendor identity: name, address, phone, and email
    - Invoice information: invoice number and invoice date
    - Customer identity (bill to): name, address, phone, and email
    - List of items including: description, quantity, unit price, and total per item
    - Subtotal, tax, total, and currency

    Use this JSON format as a reference for your output:

    {{
    "vendor": {{
        "name": "PT Sumber Makmur",
        "address": "Jl. Merdeka No. 123, Jakarta",
        "phone": "+62-21-12345678",
        "email": "info@sumbermakmur.co.id"
    }},
    "invoice": {{
        "invoice_number": "INV-2025-0001",
        "invoice_date": "2025-05-06",
        "idue_date": "2025-05-09"
    }},
    "customer": {{
        "name": "PT Sentosa Abadi",
        "address": "Jl. Sudirman No. 88, Bandung",
        "phone": "+62-22-98765432",
        "email": "purchasing@sentosaabadi.com"
    }},
    "items": [
        {{
        "description": "Jasa Instalasi Sistem Keamanan",
        "quantity": 1,
        "unit_price": 5000000,
        "total": 5000000
        }},
        {{
        "description": "Kamera CCTV HD",
        "quantity": 4,
        "unit_price": 750000,
        "total": 3000000
        }}
    ],
    "payment": {{
        "bank": "Bank Mandiri",
        "account_name": "PT. XYZ",
        "account_number": 0700010202428
    }},
    "subtotal": 8000000,
    "tax": 800000,
    "total": 8800000,
    "currency": "IDR"
    }}

    Only output the final JSON object. Do not add explanations or other text.

    Invoice Text:
    \"\"\"
    {extracted_text}
    \"\"\"
    """
    response = client.chat.completions.create(
        model="gpt-4",
        messages=[
            {"role": "system", "content": "You are an assistant that extracts information from Invoice."},
            {"role": "user", "content": prompt}
        ]
    )

    structured_data = response.choices[0].message.content.strip()
    print("ISI JSON", structured_data)

    try:
        return json.loads(structured_data)
    except json.JSONDecodeError:
        return {"error": "Failed to properly structure the data."}


# Declare variable.
if 'pdf_ref' not in ss:
    ss.pdf_ref = None

# Streamlit UI
st.title("WiratekAI - Smart OCR")
st.caption("Upload CV PDF file for information extraction.")

uploaded_file = st.file_uploader("Select PDF file", type=["pdf"], accept_multiple_files= True)

if uploaded_file:
    if st.button("Run Modeling"):
        st.session_state.results = []

        for idx, uploaded_file in enumerate(uploaded_file):
            pdf_bytes = uploaded_file.getvalue()
            # Viewer PDF
            pdf_viewer(input=pdf_bytes, width=700)

            # Ekstraksi dan struktur data
            extracted_text = extract_text_from_pdf(uploaded_file)
            structured_data = structure_invoice_data(extracted_text)

            st.session_state.results.append({
                "idx": idx + 1,
                "image": pdf_bytes,
                "data": structured_data
            })

            # Tampilkan hasil
            # st.subheader("Hasil Ekstraksi OCR (JSON):")
            # st.json(structured_data)

if "results" in st.session_state:
    for result in st.session_state.results:
        idx = result["idx"]
        structured_invoice_data = result["data"]

        st.subheader(f"OCR Extraction Result (JSON) - File {idx}:")
        st.json(structured_invoice_data)
        excel_file = save_to_excel(structured_invoice_data)
        st.download_button(
            label=f"Download File Excel untuk Invoice {idx}",
            data=excel_file,
            file_name=f"invoice_data_{idx}.xlsx",
            mime="application/vnd.ms-excel",
            key=f"download_result_{idx}"
        )
