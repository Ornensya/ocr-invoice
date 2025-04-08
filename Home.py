import streamlit as st

st.set_page_config(page_title="Menu Utama", layout="wide")

st.title("📊 Aplikasi OCR CV & INVOICE")

st.markdown("____")
container = st.container(border=True)
container.write('''Selamat datang di aplikasi OCR CV & Invoice Analyzer – solusi cerdas untuk mengubah dokumen cetak menjadi data digital yang mudah diproses! 
                Aplikasi ini dirancang untuk membantu perusahaan, tim HR, dan bagian keuangan dalam memproses dokumen seperti Curriculum Vitae (CV) dan Invoice (faktur) 
                secara otomatis menggunakan teknologi Optical Character Recognition (OCR).''')