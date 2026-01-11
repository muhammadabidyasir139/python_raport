import streamlit as st
from docxtpl import DocxTemplate
import pandas as pd
import os
import zipfile
import io
import shutil

# --- Konfigurasi Halaman ---
st.set_page_config(page_title="Sistem Generate Rapor", layout="wide")

# --- Fungsi Backend ---
def generate_rapor_logic(dataframe, template_file, output_folder):
    """
    Fungsi inti untuk generate raport.
    :param dataframe: pandas DataFrame berisi data siswa
    :param template_file: Objek file template (.docx) dari uploader
    :param output_folder: Path folder untuk menyimpan hasil
    :return: List of generated filenames, List of errors
    """
    generated_files = []
    errors = []

    # Pastikan folder output ada
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Simpan template sementara agar bisa dibaca oleh DocxTemplate
    temp_template_path = os.path.join(output_folder, "temp_template.docx")
    with open(temp_template_path, "wb") as f:
        f.write(template_file.getbuffer())

    doc = DocxTemplate(temp_template_path)

    for index, row in dataframe.iterrows():
        try:
            # Mengambil data baris sebagai dictionary
            context = row.to_dict()
            
            # --- LOGIKA OTOMATIS (CONTOH) ---
            # Membersihkan nama file dari karakter ilegal
            safe_name = "".join([c for c in str(context.get('nama_lengkap', 'TanpaNama')) if c.isalpha() or c.isdigit() or c==' ']).strip()
            safe_kelas = "".join([c for c in str(context.get('Kelas', 'Umum')) if c.isalpha() or c.isdigit() or c==' ']).strip()
            
            # Contoh logika tambahan: Hitung Rata-rata jika kolom nilai ada (opsional)
            # context['rata_rata'] = (row['nilai1'] + row['nilai2']) / 2
            
            # Render template
            doc.render(context)
            
            # Simpan file
            file_name = f"Rapor_{safe_name}_{safe_kelas}.docx"
            file_path = os.path.join(output_folder, file_name)
            doc.save(file_path)
            
            generated_files.append(file_path)
        except Exception as e:
            errors.append(f"Gagal memproses baris {index+1} ({row.get('nama_lengkap', 'Unknown')}): {str(e)}")

    # Hapus template sementara
    if os.path.exists(temp_template_path):
        os.remove(temp_template_path)
        
    return generated_files, errors

def create_zip(file_paths, zip_name):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for file_path in file_paths:
            zip_file.write(file_path, os.path.basename(file_path))
    return zip_buffer

# --- Frontend (Streamlit) ---
st.title("📄 Sistem Generate Rapor Otomatis")
st.markdown("Upload template Word dan data Excel/Input Manual untuk membuat rapor secara massal.")

# 1. Sidebar untuk Template
with st.sidebar:
    st.header("1. Upload Template")
    uploaded_template = st.file_uploader("Upload File Template (.docx)", type=["docx"])
    st.info("Pastikan template menggunakan format Jinja2. Contoh: `{{ nama_lengkap }}`")

# 2. Main Area untuk Data
st.header("2. Input Data Nilai")

data_source = st.radio("Pilih Sumber Data:", ("Upload Excel", "Input Manual / Edit Data"))

df = pd.DataFrame()

if data_source == "Upload Excel":
    uploaded_excel = st.file_uploader("Upload File Excel (.xlsx)", type=["xlsx"])
    if uploaded_excel:
        try:
            df = pd.read_excel(uploaded_excel)
        except Exception as e:
            st.error(f"Error membaca Excel: {e}")
else:
    st.write("Masukkan data siswa di bawah ini (Klik tombol '+' untuk tambah baris):")
    # Template data awal jika kosong
    default_data = {
        'nama_lengkap': ['Budi Santoso', 'Siti Aminah'],
        'Kelas': ['10A', '10B'],
        'matematika': [85, 90],
        'bahasa_indonesia': [88, 92],
        'catatan_wali': ['Tingkatkan prestasi', 'Pertahankan prestasi']
    }
    df = pd.DataFrame(default_data)

# Tampilkan Editor Data (Bisa diedit user langsung di browser)
if not df.empty:
    edited_df = st.data_editor(df, num_rows="dynamic")
    
    # 3. Tombol Generate
    st.header("3. Generate Dokumen")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        generate_btn = st.button("🚀 Generate Rapor", type="primary")
    
    if generate_btn:
        if uploaded_template is None:
            st.error("⚠️ Harap upload file Template (.docx) terlebih dahulu di Sidebar!")
        elif edited_df.empty:
            st.error("⚠️ Data nilai kosong!")
        else:
            with st.spinner("Sedang memproses dokumen..."):
                output_folder = "hasil_rapor"
                
                # Bersihkan folder hasil sebelumnya jika ada (opsional)
                if os.path.exists(output_folder):
                    shutil.rmtree(output_folder)
                
                files, error_logs = generate_rapor_logic(edited_df, uploaded_template, output_folder)
                
                if error_logs:
                    for err in error_logs:
                        st.error(err)
                
                if files:
                    st.success(f"✅ Berhasil membuat {len(files)} dokumen!")
                    
                    # Buat tombol download ZIP
                    zip_buffer = create_zip(files, "semua_rapor.zip")
                    st.download_button(
                        label="⬇️ Download Semua (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name="semua_rapor.zip",
                        mime="application/zip"
                    )
                    
                    # Tampilkan list file
                    with st.expander("Lihat detail file yang dibuat"):
                        for f in files:
                            st.write(os.path.basename(f))
