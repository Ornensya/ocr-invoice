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
# FONT_PATH = "C:/Windows/Fonts/arial.ttf"
FONT_PATH = os.path.join(os.path.dirname(__file__), "..", "fonts", "arial.ttf")


# Inisialisasi PaddleOCR
# ocr = PaddleOCR(use_angle_cls=True, lang='en')
@st.cache_resource
def load_ocr_model():
    return PaddleOCR(use_angle_cls=True, lang='en')
ocr = load_ocr_model()


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
        # st.subheader(f"🖼️ Halaman {idx + 1}")
        # st.image(image, caption="Gambar Asli", use_container_width=True)
        image_np = np.array(image)
        result = ocr.ocr(image_np, cls=True)

        boxes = [line[0] for line in result[0]]
        txts = [line[1][0] for line in result[0]]
        scores = [line[1][1] for line in result[0]]

        annotated_image = draw_ocr(image_np, boxes, txts, scores, font_path=FONT_PATH)
        annotated_image = Image.fromarray(annotated_image)
        # st.image(annotated_image, caption="🔎 Hasil Deteksi OCR")

        extracted_text += "\n".join(txts) + "\n"

    return extracted_text

# --- Fungsi Strukturkan JSON dari OpenAI ---
def structure_invoice_data(extracted_text):
    if not client:
        return {"error": "OpenAI API key belum dikonfigurasi."}

    prompt = f"""
    You are a financial assistant. Based on the following extracted invoice text, convert it into a clean and structured JSON format.
    Extract structured data from the invoice document using the following rules and output it strictly in the provided JSON format.

    RULES:
    - Fields marked as "mandatory" must always be filled based on the content found in the invoice.
    - Fields marked as "not mandatory" must ONLY be filled if the exact information is found in the document. If not available, set them to `null`.
    - DO NOT guess, infer, or hallucinate values that are not explicitly stated in the document.
    - Use proper data types: strings for text, integers for amounts, ISO 8601 format (YYYY-MM-DD) for dates.
    - Show all readable items
    - Return ONLY a valid JSON object, no explanation or surrounding text.

    JSON FORMAT:
    {{
    "seller_identity": {{
        "company_name": "...",
        "address": "...",
        "email_address": "...",
        "phone": "... or null",
        "company_npwp_tin": "... or null"
    }},
    "buyer_identity": {{
        "company_name": "...",
        "address": "...",
        "email_address": "...",
        "phone": "... or null",
        "company_npwp_tin": "... or null",
        "attention": "... or null"
    }},
    "invoice_details": {{
        "invoice_no": "...",
        "invoice_date": "...",
        "order_po_number": "... or null",
        "term_of_payment_due_date": "... or null"
    }},
    "item_details": [
        {{
        "item_description": "...",
        "quantity": ...,
        "unit_price": ...,
        "amount": ...
        }}
    ],
    "subtotal_invoice": ...,
    "discount": ... or null,
    "vat": ... or null,
    "invoice_total": ...,
    "bank_details": {{
        "account_no": "...",
        "account_name": "...",
        "beneficiary_bank": "...",
        "branch": "... or null",
        "swift_code": "... or null"
    }},
    "currency": "IDR"
    }}

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

# --- Fungsi Perhitungan Tambahan (DPP, VAT) ---
def calculate_invoice_fields(data):
    try:
        subtotal = data.get("subtotal_invoice", 0)
        vat = data.get("vat", None)
        # discount = data.get("discount", 0)

        # Jika discount tidak valid (misalnya None), jadikan 0
        # if discount is None:
        #     discount = 0

        # Hitung subtotal setelah diskon
        after_discount = subtotal
        # after_discount = max(after_discount, 0)  # Hindari negatif
        dpp = round((100 / 111) * after_discount, 2)
        calculated_vat = round(0.11 * dpp, 2)

        return {
            "subtotal_sebelum_diskon": subtotal,
            # "diskon": discount,
            # "subtotal_setelah_diskon": after_discount,
            "dpp": dpp,
            "ppn_11_persen": calculated_vat
        }
    except Exception as e:
        return {"error": f"Gagal menghitung: {str(e)}"}


def save_to_excel(structured_invoice_data, calculated_fields):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    seller_identity = structured_invoice_data.get("seller_identity", {})
    buyer_identity = structured_invoice_data.get("buyer_identity", {})
    invoice_details = structured_invoice_data.get("invoice_details", {})
    item_details = structured_invoice_data.get("item_details", [])
    subtotal_invoice = structured_invoice_data.get("subtotal_invoice", "")
    discount = structured_invoice_data.get("discount", "")
    vat = structured_invoice_data.get("vat", "")
    invoice_total = structured_invoice_data.get("invoice_total", "")
    currency = structured_invoice_data.get("currency", "")
    bank_details = structured_invoice_data.get("bank_details", {})
    

    ws.append(["Seller Identity"])
    ws.append(["Company Name", seller_identity.get("company_name", "")])
    ws.append(["Address", seller_identity.get("address", "")])
    ws.append(["Email Address", seller_identity.get("email_address", "")])
    ws.append(["Phone", seller_identity.get("phone", "")])
    ws.append(["Company NPWP/TIN", seller_identity.get("company_npwp_tin", "")])
    ws.append([])

    ws.append(["Buyer Identity"])
    ws.append(["Company Name", buyer_identity.get("company_name", "")])
    ws.append(["Address", buyer_identity.get("address", "")])
    ws.append(["Email Address", buyer_identity.get("email_address", "")])
    ws.append(["Phone", buyer_identity.get("phone", "")])
    ws.append(["Company NPWP/TIN", buyer_identity.get("company_npwp_tin", "")])
    ws.append(["Attention", buyer_identity.get("attention", "")])
    ws.append([])

    ws.append(["Invoice Details"])
    ws.append(["Invoice No", invoice_details.get("invoice_no", "")])
    ws.append(["Invoice Date", invoice_details.get("invoice_date", "")])
    ws.append(["Order/PO Number", invoice_details.get("order_po_number", "")])
    ws.append(["Term of Payment/Due Date", invoice_details.get("term_of_payment_due_date", "")])
    ws.append([])

    if item_details:
        ws.append(["item_details"])
        item_df = pd.DataFrame(item_details)
        for row in dataframe_to_rows(item_df, index=False, header=True):
            ws.append(row)
        ws.append([])

    ws.append(["Subtotal Invoice", subtotal_invoice])
    ws.append(["Discount", discount])
    ws.append(["VAT", vat])
    ws.append(["Invoice Total", invoice_total])
    ws.append(["Currency", currency])

    ws.append(["Bank Details"])
    ws.append(["Account No", bank_details.get("account_no", "")])
    ws.append(["Account Name", bank_details.get("account_name", "")])
    ws.append(["Benecifiary Bank", bank_details.get("beneficiary_bank", "")])
    ws.append(["Branch", bank_details.get("branch", "")])
    ws.append(["SWIFT Code", bank_details.get("swift_code", "")])
    ws.append([])

    

    wb.save(output)
    output.seek(0)
    return output

# --- Streamlit Logic ---
if uploaded_file:
    if st.button("🚀 Jalankan OCR"):
        st.session_state.results = []

        for idx, uploaded in enumerate(uploaded_file):
            pdf_bytes = uploaded.getvalue()

            # 1. Ekstrak teks dari OCR
            extracted_text = extract_text_with_paddleocr(uploaded)

            # 2. Strukturkan data via OpenAI
            structured_data = structure_invoice_data(extracted_text)

            # 3. Hitung nilai tambahan
            calculated_fields = calculate_invoice_fields(structured_data)

            # 4. Isi nilai VAT jika kosong dari hasil perhitungan
            # if structured_data.get("vat") is None and "ppn_11_persen" in calculated_fields:
            #     structured_data["vat"] = calculated_fields["ppn_11_persen"]

            # 4. Periksa apakah invoice menyebutkan "Price including VAT"
            if "price including vat" in extracted_text.lower():
                structured_data["vat"] = calculated_fields.get("ppn_11_persen", None)


            # 5. Tampilkan hasil
            # st.subheader("📋 Teks Hasil Ekstraksi:")
            # st.text(extracted_text)

            st.subheader("🧾 Hasil JSON Terstruktur:")
            st.json(structured_data)

            st.subheader(f"🧮 Perhitungan Tambahan - Invoice {idx+1}")
            st.json(calculated_fields)

            st.session_state.results.append({
                "idx": idx + 1,
                "image": pdf_bytes,
                "data": structured_data,
                "calculation": calculated_fields
            })

if "results" in st.session_state:
    for result in st.session_state.results:
        idx = result["idx"]
        structured_invoice_data = result["data"]
        calculated_fields = result.get("calculation", None)

        excel_file = save_to_excel(structured_invoice_data, calculated_fields)
        st.download_button(
            label=f"📥 Download File Excel untuk Invoice {idx}",
            data=excel_file,
            file_name=f"invoice_data_{idx}.xlsx",
            mime="application/vnd.ms-excel",
            key=f"download_result_{idx}"
        )
