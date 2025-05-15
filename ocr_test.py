import streamlit as st
from paddleocr import PaddleOCR, draw_ocr
from pdf2image import convert_from_bytes
from PIL import Image
import numpy as np
import os
import json
import pandas as pd
from io import BytesIO
from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Load .env (for OpenAI API key)
load_dotenv()

# Konfigurasi path
POPPLER_PATH = r"C:\Program Files\poppler-24.07.0\Library\bin"
FONT_PATH = "C:/Windows/Fonts/arial.ttf"

# Inisialisasi PaddleOCR
ocr = PaddleOCR(use_angle_cls=True, lang='en')

# OpenAI API Key
api_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=api_key) if api_key else None

# Streamlit UI
st.title("🔍 Smart Invoice OCR - PaddleOCR + OpenAI")
uploaded_file = st.file_uploader("📄 Upload file PDF Invoice", type="pdf", accept_multiple_files=True)

# --- Fungsi Ekstraksi Teks dari PaddleOCR ---
def extract_text_with_paddleocr(pdf_file):
    images = convert_from_bytes(pdf_file.read())
    extracted_text = ""
    for idx, image in enumerate(images):
        st.subheader(f"🖼️ Halaman {idx + 1}")
        st.image(image, caption="Gambar Asli", use_container_width=True)
        image_np = np.array(image)
        result = ocr.ocr(image_np, cls=True)

        boxes = [line[0] for line in result[0]]
        txts = [line[1][0] for line in result[0]]
        scores = [line[1][1] for line in result[0]]

        annotated_image = draw_ocr(image_np, boxes, txts, scores, font_path=FONT_PATH)
        annotated_image = Image.fromarray(annotated_image)
        st.image(annotated_image, caption="🔎 Hasil Deteksi OCR")

        extracted_text += "\n".join(txts) + "\n"

    return extracted_text

# --- Fungsi Strukturkan JSON dari OpenAI ---
def structure_invoice_data(extracted_text):
    if not client:
        return {"error": "OpenAI API key belum dikonfigurasi."}

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
        "due_date": "2025-05-09"
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
    \"\"\"{extracted_text}\"\"\"
    """

    response = client.chat.completions.create(
        model="gpt-4",
        messages=[
            {"role": "system", "content": "You are an assistant that extracts information from Invoice."},
            {"role": "user", "content": prompt}
        ]
    )

    structured_data = response.choices[0].message.content.strip()

    try:
        return json.loads(structured_data)
    except json.JSONDecodeError:
        return {"error": "❌ Gagal parsing JSON dari LLM."}

# --- Fungsi Simpan ke Excel ---
def save_to_excel(structured_invoice_data):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    vendor = structured_invoice_data.get("vendor", {})
    invoice = structured_invoice_data.get("invoice", {})
    customer = structured_invoice_data.get("customer", {})
    items = structured_invoice_data.get("items", [])
    payment = structured_invoice_data.get("payment", {})
    subtotal = structured_invoice_data.get("subtotal", "")
    tax = structured_invoice_data.get("tax", "")
    total = structured_invoice_data.get("total", "")
    currency = structured_invoice_data.get("currency", "")

    ws.append(["Vendor Information"])
    ws.append(["Name", vendor.get("name", "")])
    ws.append(["Address", vendor.get("address", "")])
    ws.append(["Phone", vendor.get("phone", "")])
    ws.append(["Email", vendor.get("email", "")])
    ws.append([])

    ws.append(["Invoice Information"])
    ws.append(["Invoice Number", invoice.get("invoice_number", "")])
    ws.append(["Invoice Date", invoice.get("invoice_date", "")])
    ws.append(["Due Date", invoice.get("due_date", "")])
    ws.append([])

    ws.append(["Customer Information"])
    ws.append(["Name", customer.get("name", "")])
    ws.append(["Address", customer.get("address", "")])
    ws.append(["Phone", customer.get("phone", "")])
    ws.append(["Email", customer.get("email", "")])
    ws.append([])

    if items:
        ws.append(["Items"])
        item_df = pd.DataFrame(items)
        for row in dataframe_to_rows(item_df, index=False, header=True):
            ws.append(row)
        ws.append([])

    ws.append(["Payment Information"])
    ws.append(["Bank", payment.get("bank", "")])
    ws.append(["Account Number", payment.get("account_number", "")])
    ws.append([])

    ws.append(["Subtotal", subtotal])
    ws.append(["Tax", tax])
    ws.append(["Total", total])
    ws.append(["Currency", currency])

    wb.save(output)
    output.seek(0)
    return output

# --- Jalankan Pipeline ---
# if uploaded_file and st.button("🚀 Jalankan OCR + Ekstraksi JSON"):
#     with st.spinner("Sedang memproses..."):
#         extracted_text = extract_text_with_paddleocr(uploaded_file)
#         structured_data = structure_invoice_data(extracted_text)

#         st.subheader("📋 Teks Hasil Ekstraksi:")
#         st.text(extracted_text)

#         st.subheader("🧾 Hasil JSON Terstruktur:")
#         st.json(structured_data)

#         excel_file = save_to_excel(structured_data)
#         st.download_button("⬇️ Download Hasil Excel", data=excel_file, file_name="invoice_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if uploaded_file:
    if st.button("🚀 Jalankan OCR"):
        st.session_state.results = []

        for idx, uploaded_file in enumerate(uploaded_file):
            pdf_bytes = uploaded_file.getvalue()
            # Viewer PDF
            # pdf_viewer(input=pdf_bytes, width=700)

            # Ekstraksi dan struktur data
            extracted_text = extract_text_with_paddleocr(uploaded_file)
            structured_data = structure_invoice_data(extracted_text)

            st.subheader("📋 Teks Hasil Ekstraksi:")
            st.text(extracted_text)

            st.subheader("🧾 Hasil JSON Terstruktur:")
            st.json(structured_data)

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

        # st.subheader(f"OCR Extraction Result (JSON) - File {idx}:")
        # st.json(structured_invoice_data)
        excel_file = save_to_excel(structured_invoice_data)
        st.download_button(
            label=f"Download File Excel untuk Invoice {idx}",
            data=excel_file,
            file_name=f"invoice_data_{idx}.xlsx",
            mime="application/vnd.ms-excel",
            key=f"download_result_{idx}"
        )
