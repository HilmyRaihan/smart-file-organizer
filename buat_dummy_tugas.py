import os
from docx import Document
from fpdf import FPDF

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FOLDER_DUMMY = os.path.join(BASE_DIR, "file_tes_dummy")
os.makedirs(FOLDER_DUMMY, exist_ok=True)

# 1. File DOCX dengan NAMA JELAS ('Tugas_Enterpreneur_Kelompok2.docx')
doc1 = Document()
doc1.add_heading("Tugas Mata Kuliah Kewirausahaan", 0)
doc1.add_paragraph("Ini adalah dokumen tugas kelompok mengenai ide bisnis startup kuliner.")
doc1.save(os.path.join(FOLDER_DUMMY, "Tugas_Enterpreneur_Kelompok2.docx"))

# 2. File DOCX dengan NAMA ACAK ('dokumen_revisi_final.docx'), tapi ISINYA tentang Basis Data
doc2 = Document()
section = doc2.sections[0]
header = section.header
header.paragraphs[0].text = "Praktikum Basis Data & Perancangan ERD"
doc2.add_paragraph("Pada pertemuan ini kita membahas normalisasi tabel 1NF, 2NF, dan 3NF menggunakan SQL.")
doc2.save(os.path.join(FOLDER_DUMMY, "dokumen_revisi_final.docx"))

# 3. File PDF dengan NAMA ACAK ('scan_tugas_0912.pdf'), tapi ISINYA tentang Algoritma
pdf1 = FPDF()
pdf1.add_page()
pdf1.set_font("Arial", size=12)
pdf1.cell(200, 10, txt="Laporan Praktikum Algoritma dan Pemrograman", ln=True)
pdf1.cell(200, 10, txt="Materi: Pembuatan Flowchart dan Pseudocode Perulangan Python.", ln=True)
pdf1.output(os.path.join(FOLDER_DUMMY, "scan_tugas_0912.pdf"))

# 4. File TXT acak yang TIDAK COCOK aturan mana pun (untuk tes folder Umum_Lainnya)
with open(os.path.join(FOLDER_DUMMY, "catatan_belanja_harian.txt"), "w", encoding="utf-8") as f:
    f.write("Beli kopi, gula, dan kertas HVS A4.")

print(f"Berhasil membuat 4 file dummy di folder:\n{FOLDER_DUMMY}")