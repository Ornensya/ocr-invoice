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
    text = "\n".join([pytesseract.image_to_string(img) for img in images])
    return text

# # Fungsi untuk menyusun data CV dengan OpenAI
# def structure_cv_data(extracted_text):
#     if not client:
#         return {"error": "API OpenAI tidak tersedia. Pastikan API key sudah dikonfigurasi."}

#     prompt = f"""Ekstrak informasi berikut dari teks CV yang tidak terstruktur:
#     - Nama
#     - Kontak (Email, Telepon)
#     - Pendidikan
#     - Pengalaman kerja
#     - Keterampilan
#     - Sertifikasi (jika ada)
    
#     Tampilkan dalam format JSON.
#     \n\nTeks CV:
#     {extracted_text}
#     """

#     response = client.chat.completions.create(
#         model="gpt-4",
#         messages=[
#             {"role": "system", "content": "Anda adalah asisten yang mengekstrak informasi dari CV."},
#             {"role": "user", "content": prompt}
#         ]
#     )

#     structured_data = response.choices[0].message.content.strip()

#     try:
#         return json.loads(structured_data)
#     except json.JSONDecodeError:
#         return {"error": "Gagal menyusun data dengan benar."}

# Function to structure CV data using OpenAI
def structure_cv_data(extracted_text):
    if not client:
        return {"error": "OpenAI API is not available. Make sure the API key is properly configured."}

    prompt = f"""Extract the following information from the unstructured CV text:
    - Name
    - About Me (if any)
    - Contact (Email, Phone)
    - Education
    - Work Experience
    - Skills
    
    Present the data in JSON format.
    \n\nCV Text:
    {extracted_text}
    """
    response = client.chat.completions.create(
        model="gpt-4",
        messages=[
            {"role": "system", "content": "You are an assistant that extracts information from CVs."},
            {"role": "user", "content": prompt}
        ]
    )

    structured_data = response.choices[0].message.content.strip()

    try:
        return json.loads(structured_data)
    except json.JSONDecodeError:
        return {"error": "Failed to properly structure the data."}


# Fungsi untuk menyimpan hasil ke dalam Excel
def save_to_excel(data):
    df = pd.DataFrame([{k: v for k, v in data.items()}])
    excel_file = BytesIO()
    df.to_excel(excel_file, index=False)
    excel_file.seek(0)
    return excel_file

# Declare variable.
if 'pdf_ref' not in ss:
    ss.pdf_ref = None

# Streamlit UI
st.title("WiratekAI - Smart OCR")
st.caption("Upload CV PDF file for information extraction.")

uploaded_file = st.file_uploader("Select PDF file", type=["pdf"])

if uploaded_file is not None:

    pdf_bytes = uploaded_file.getvalue()

    # Simpan ke session_state agar bisa diakses ulang
    ss.pdf_ref = pdf_bytes

    # Tampilkan PDF Viewer
    pdf_viewer(input=pdf_bytes, width=700)

    # st.write("Mengekstrak teks dari CV...") #Nanti coba pake st.spiner
    extracted_text = extract_text_from_pdf(uploaded_file)
    
    # st.write("Menyusun data CV...")
    structured_cv_data = structure_cv_data(extracted_text)
    
    # Menampilkan hasil dalam JSON
    st.subheader("OCR Extraction Result (JSON):")
    st.json(structured_cv_data)
    
    # # Menampilkan hasil dalam tabel
    # if isinstance(structured_cv_data, dict) and "error" not in structured_cv_data:
    #     df_main = pd.DataFrame([structured_cv_data])
    #     st.subheader("Hasil Ekstraksi dalam Tabel:")
    #     st.table(df_main)
    # else:
    #     st.warning("Gagal menampilkan data dalam tabel.")
    
    # Simpan ke Excel
    if structured_cv_data:
        excel_file = save_to_excel(structured_cv_data)
        st.download_button(
            label="Download File Excel",
            data=excel_file,
            file_name="cv_data.xlsx",
            mime="application/vnd.ms-excel"
        )
