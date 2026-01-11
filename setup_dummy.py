import pandas as pd
import os

def create_sample_excel():
    data = {
        'nama_lengkap': ['Ahmad Dahlan', 'Dewi Sartika', 'Ki Hajar Dewantara'],
        'nis': ['1001', '1002', '1003'],
        'Kelas': ['XII IPA 1', 'XII IPA 1', 'XII IPA 2'],
        'semester': ['Ganjil', 'Ganjil', 'Ganjil'],
        'tahun_ajaran': ['2023/2024', '2023/2024', '2023/2024'],
        'nilai_matematika': [85, 90, 78],
        'nilai_bahasa': [88, 92, 85],
        'nilai_inggris': [75, 88, 90],
        'catatan_wali': ['Pertahankan prestasimu.', 'Sangat baik, tingkatkan keaktifan.', 'Perlu belajar lebih giat.']
    }
    
    df = pd.DataFrame(data)
    filename = "data_nilai_sample.xlsx"
    df.to_excel(filename, index=False)
    print(f"File sample berhasil dibuat: {filename}")
    print("Gunakan file ini sebagai referensi kolom untuk Template Word Anda.")

if __name__ == "__main__":
    create_sample_excel()
