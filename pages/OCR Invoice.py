import streamlit as st
from PIL import Image
import pytesseract
import json
from openai import OpenAI
import os
import pandas as pd
from io import BytesIO
import re

# Initialize OpenAI API
api_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=api_key) if api_key else None

# Set path Tesseract (Sesuaikan dengan sistem Anda)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Fungsi untuk meratakan struktur data menjadi key-value pairs
def flatten_data(data, parent_key=""):
    flat_data = []
    
    if isinstance(data, dict):
        for key, value in data.items():
            new_key = f"{parent_key} - {key}" if parent_key else key
            flat_data.extend(flatten_data(value, new_key))
    
    elif isinstance(data, list):
        for idx, item in enumerate(data):
            new_key = f"{parent_key} [{idx + 1}]"
            flat_data.extend(flatten_data(item, new_key))
    
    else:
        flat_data.append((parent_key, data))
    
    return flat_data

# Fungsi untuk menyimpan hasil ke dalam Excel
def save_to_excel(data):
    flat_data = flatten_data(data)
    df = pd.DataFrame(flat_data, columns=["Kunci", "Nilai"])

    excel_file = BytesIO()
    df.to_excel(excel_file, index=False)
    excel_file.seek(0)

    return excel_file

# Fungsi OCR menggunakan Tesseract
def process_invoice_image_without_model(image_path):
    image = Image.open(image_path).convert("RGB")
    extracted_text = pytesseract.image_to_string(image)
    return extracted_text

# # Fungsi untuk menyusun data menggunakan OpenAI
# def structure_invoice_data_with_llm(extracted_text):
#     print(extracted_text)
#     if not client:
#         return {"error": "API OpenAI tidak tersedia. Pastikan API key sudah dikonfigurasi."}

#     prompt = f"Please structure the following unstructured invoice data into key-value pairs:\n\n{extracted_text}\n\nReturn a JSON object with keys like 'Invoice Number', 'Date', 'Amount', etc."

#     response = client.chat.completions.create(
#         model="gpt-4",
#         messages=[
#             {"role": "system", "content": "You are an assistant that structures invoice data."},
#             {"role": "user", "content": prompt}
#         ]
#     )

#     structured_data = response.choices[0].message.content.strip()

#     # try:
#     return json.loads(structured_data)
#     # except json.JSONDecodeError:
#     #     return {"error": "Failed to structure the data correctly."}

def structure_invoice_data_with_llm(extracted_text):
    print("Extracted Text:\n", extracted_text)  # Debugging

    if not client:
        return {"error": "API OpenAI tidak tersedia. Pastikan API key sudah dikonfigurasi."}

    prompt = f"Please structure the following unstructured invoice data into key-value pairs:\n\n{extracted_text}\n\nReturn a JSON object. Return JSON only, without any explanation or extra text."

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are an assistant that extracts and structures invoice data from AlfaMart stores. Format the data in a clear and organized manner, including details such as invoice number, date, item name, quantity, price per item, total price, and other relevant details."},
                {"role": "user", "content": prompt}
            ]
        )

        structured_data = response.choices[0].message.content.strip()
        print("Raw Response:\n", structured_data)  # Debugging

        # Validasi apakah output adalah JSON valid
        try:
            parsed_data = json.loads(structured_data)
            return parsed_data
        except json.JSONDecodeError:
            print("Attempting to extract JSON from response...")

            # Gunakan regex untuk mengekstrak JSON dari teks
            match = re.search(r"\{.*\}", structured_data, re.DOTALL)
            if match:
                json_str = match.group(0)
                try:
                    return json.loads(json_str)
                except json.JSONDecodeError:
                    return {"error": "Failed to parse extracted JSON.", "raw_response": json_str}
            
            return {"error": "Failed to parse JSON. Response was:\n" + structured_data}

    except Exception as e:
        return {"error": f"API call failed: {str(e)}"}

# Streamlit UI
st.title("WiratekAI - Smart OCR")
st.write("Upload receipt image to extract data.")

uploaded_file = st.file_uploader("Select image file", type=["jpg", "jpeg", "png"])

if uploaded_file is not None:
    # Menampilkan gambar yang diunggah
    image = Image.open(uploaded_file)
    st.image(image, caption="Uploaded image", width=400)

    # Proses OCR
    extracted_text = process_invoice_image_without_model(uploaded_file)

    # Proses LLM untuk struktur data
    structured_invoice_data = structure_invoice_data_with_llm(extracted_text)

    # Menampilkan hasil dalam JSON
    st.subheader("OCR Extraction Result (JSON):")
    st.json(structured_invoice_data)

    # # Menampilkan hasil dalam tabel
    # st.subheader("Hasil Ekstraksi OCR dalam Tabel:")

    # if isinstance(structured_invoice_data, dict) and "error" not in structured_invoice_data:
    #     flat_data = flatten_data(structured_invoice_data)
    #     df_main = pd.DataFrame(flat_data, columns=["Kunci", "Nilai"])
    #     st.table(df_main)
    # else:
    #     st.warning("Gagal menampilkan data dalam tabel.")

    # Simpan ke Excel
    if structured_invoice_data:
        excel_file = save_to_excel(structured_invoice_data)
        st.download_button(
            label="Download File Excel",
            data=excel_file,
            file_name="invoice_data.xlsx",
            mime="application/vnd.ms-excel"
        )
