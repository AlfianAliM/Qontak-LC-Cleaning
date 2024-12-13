import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Proses data
def process_data(uploaded_file):
    # Read file
    df = pd.read_excel(uploaded_file)

    # Drop duplicates based on 'handler'
    df_cleaned = df.drop_duplicates(subset='handler', keep='last')

    # Ekstrak status lead
    def extract_status_lead(tag):
        tag_lower = str(tag).lower()
        if 'cold' in tag_lower:
            return 'cold'
        elif 'warm' in tag_lower:
            return 'warm'
        elif 'hot' in tag_lower:
            return 'hot'
        return None

    # Ekstrak grade
    def extract_grade(tag):
        match = re.search(r'\bgrade\s*[a-z0-9]+\b', str(tag), re.IGNORECASE)
        if match:
            return match.group(0)
        return "tidak ada grade"

    # Daftar cabang LC
    cabang_list = ['Pare', 'Bogor', 'Bandung', 'Jogja', 'Serang', 'Lampung', 'Medan', 'Makassar']

    # Ekstrak cabang
    def extract_cabang(tag):
        for cabang in cabang_list:
            if cabang.lower() in str(tag).lower():
                return cabang
        return "tidak ada cabang"

    # Keyword keterangan
    keterangan_list = [
        'no respon (sudah ada conversation)', 'payment', 'Tanya Harga', 'terkendala biaya', 
        'diskusi dulu', 'DP', 'Mengisi Form Pendaftaran', 'Pelunasan', 'Pembayaran DP'
    ]

    # Ekstrak keterangan
    def extract_keterangan(tag):
        for keterangan in keterangan_list:
            if keterangan.lower() in str(tag).lower():
                return keterangan
        return "tidak ada keterangan"

    # Daftar Program LC
    program_list = [
        'reguler sm', 'desember ceria', 'integrated speaking', 'emp', 'em', 'camp', 'non camp', 
        'intensive', 'rombongan', 'reguler iep', 'toefl', 'ielts', 'esp', 'private'
    ]

    # Ekstrak program
    def extract_program(tag):
        tag_lower = str(tag).lower()
        for program in program_list:
            if program in tag_lower:
                return program.capitalize()
        return "tidak ada program"

    # Ekstrak jenis program
    def extract_online_offline(tag):
        tag_lower = str(tag).lower()
        if 'online' in tag_lower:
            return 'online'
        elif 'offline' in tag_lower:
            return 'offline'
        return "tidak ada data offline/online"

    # Kolom baru
    df_cleaned['status lead'] = df_cleaned['tag'].apply(extract_status_lead)
    df_cleaned['grade'] = df_cleaned['tag'].apply(extract_grade)
    df_cleaned['cabang'] = df_cleaned['tag'].apply(extract_cabang)
    df_cleaned['keterangan'] = df_cleaned['tag'].apply(extract_keterangan)
    df_cleaned['program'] = df_cleaned['tag'].apply(extract_program)
    df_cleaned['online/offline'] = df_cleaned['tag'].apply(extract_online_offline)

    # Kolom Respons
    df_cleaned['response'] = (
        df_cleaned['cabang'] + ', ' +
        df_cleaned['program'] + ', ' +
        df_cleaned['grade'] + ', ' +
        df_cleaned['keterangan'] + ', ' +
        df_cleaned['online/offline']
    )

    # Kolom final (6 kolom yang diminta)
    df_final = df_cleaned[[
        'cabang', 'program', 'grade', 'response', 'online/offline', 'note'
    ]]

    # Convert all columns to lowercase
    df_final.columns = df_final.columns.str.lower()

    return df_final

# Dataframe to excel
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    return processed_data

# Streamlit
st.title("Cleaning Data Qontak")

# Upload
uploaded_file = st.file_uploader("Upload file xlsx", type=["xlsx"])

if uploaded_file is not None:
    # Proses data
    processed_data = process_data(uploaded_file)

    # Tampilkan preview hasil
    st.write("Preview Hasil:")
    st.dataframe(processed_data)

    # Ke Excel
    excel_data = convert_df_to_excel(processed_data)

    # Download
    st.download_button(
        label="Download Data",
        data=excel_data,
        file_name="processed_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
