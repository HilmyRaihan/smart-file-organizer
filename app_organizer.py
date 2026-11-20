import json
import os
from datetime import datetime
import pandas as pd
import streamlit as st
from docx import Document
from pypdf import PdfReader

# ==========================================
# 1. KONFIGURASI DIREKTORI & FILE CONFIG
# ==========================================
st.set_page_config(
    page_title="Smart File & Assignment Organizer",
    page_icon=":material/folder_managed:",
    layout="wide"
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FOLDER_TUJUAN_UTAMA = os.path.join(BASE_DIR, "Arsip_Terorganisir")
CONFIG_FILE = os.path.join(BASE_DIR, "aturan_folder.json")
LOG_FILE = os.path.join(BASE_DIR, "riwayat_pemindahan.csv")

os.makedirs(FOLDER_TUJUAN_UTAMA, exist_ok=True)

DEFAULT_ATURAN = {
    "Enterpreneurship": ["enterpreneur", "entrepreneur", "kewirausahaan", "bisnis plan", "bmc"],
    "Algoritma_Pemrograman": ["algo", "algoritma", "flowchart", "pseudocode", "python"],
    "Basis_Data": ["basis_data", "basis data", "database", "normalisasi", "sql", "erd"],
    "Pemrograman_Web": ["html", "css", "javascript", "react", "frontend", "web"],
    "Sistem_Informasi": ["sistem informasi", "uml", "use case", "sdlc", "analisis sistem"],
}

# ==========================================
# 2. FUNGSI MANAJEMEN ATURAN & LOG
# ==========================================
def muat_aturan():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    simpan_aturan(DEFAULT_ATURAN)
    return DEFAULT_ATURAN

def simpan_aturan(aturan_dict):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(aturan_dict, f, indent=4)

def catat_log(nama_asli, nama_baru, folder_tujuan, metode):
    waktu = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    baris_baru = pd.DataFrame([{
        "Waktu": waktu,
        "Nama File Asli": nama_asli,
        "Nama File Tersimpan": nama_baru,
        "Folder Tujuan": folder_tujuan,
        "Metode Deteksi": metode
    }])
    if os.path.exists(LOG_FILE):
        baris_baru.to_csv(LOG_FILE, mode="a", header=False, index=False)
    else:
        baris_baru.to_csv(LOG_FILE, index=False)

# ==========================================
# 3. FUNGSI EKSTRAKSI TEKS & FILTERING
# ==========================================
def ekstrak_teks_pdf(file_obj):
    teks = ""
    try:
        reader = PdfReader(file_obj)
        for i in range(min(2, len(reader.pages))):
            hasil = reader.pages[i].extract_text()
            if hasil:
                teks += hasil + " "
    except Exception:
        pass
    return teks

def ekstrak_teks_docx(file_obj):
    teks = ""
    try:
        doc = Document(file_obj)
        for section in doc.sections:
            for p in section.header.paragraphs:
                teks += p.text + " "
        for p in doc.paragraphs[:30]:
            teks += p.text + " "
    except Exception:
        pass
    return teks

def cari_folder_cocok(teks, aturan_dict):
    teks_kecil = teks.lower()
    for nama_folder, daftar_keyword in aturan_dict.items():
        for kw in daftar_keyword:
            if kw.lower() in teks_kecil:
                return nama_folder, kw
    return None, None

def proses_file_upload(uploaded_file, aturan_dict, tambah_tanggal):
    nama_asli = uploaded_file.name
    ekstensi = os.path.splitext(nama_asli)[1].lower()

    # Tahap 1: Cek berdasarkan Nama File
    folder_target, kw_cocok = cari_folder_cocok(nama_asli, aturan_dict)
    metode = f"Nama File ('{kw_cocok}')" if folder_target else ""

    # Tahap 2: Jika nama file tidak cocok, baca isi dokumen (PDF / DOCX / TXT)
    if not folder_target:
        uploaded_file.seek(0)
        isi_teks = ""
        if ekstensi == ".pdf":
            isi_teks = ekstrak_teks_pdf(uploaded_file)
        elif ekstensi == ".docx":
            isi_teks = ekstrak_teks_docx(uploaded_file)
        elif ekstensi in [".txt", ".md"]:
            isi_teks = uploaded_file.read().decode("utf-8", errors="ignore")

        if isi_teks:
            folder_target, kw_cocok = cari_folder_cocok(isi_teks, aturan_dict)
            if folder_target:
                metode = f"Isi Dokumen ('{kw_cocok}')"

    if not folder_target:
        folder_target = "Umum_Lainnya"
        metode = "Default (Tidak ada kata kunci cocok)"

    path_folder = os.path.join(FOLDER_TUJUAN_UTAMA, folder_target)
    os.makedirs(path_folder, exist_ok=True)

    if tambah_tanggal:
        prefix = datetime.now().strftime("%Y-%m-%d_")
        nama_final = prefix + nama_asli if not nama_asli.startswith(prefix) else nama_asli
    else:
        nama_final = nama_asli

    path_simpan = os.path.join(path_folder, nama_final)
    base, ext = os.path.splitext(path_simpan)
    counter = 1
    while os.path.exists(path_simpan):
        path_simpan = f"{base}_{counter}{ext}"
        counter += 1

    nama_tersimpan = os.path.basename(path_simpan)

    uploaded_file.seek(0)
    with open(path_simpan, "wb") as f:
        f.write(uploaded_file.getbuffer())

    catat_log(nama_asli, nama_tersimpan, folder_target, metode)
    return folder_target, metode, nama_tersimpan

# ==========================================
# 4. ANTARMUKA WEB (CLEAN MATERIAL UI)
# ==========================================
aturan_aktif = muat_aturan()

st.sidebar.markdown("### :material/tune: Pengaturan & Kategori")
tambah_tgl = st.sidebar.checkbox("Tambahkan tanggal otomatis di depan nama file", value=True)

st.sidebar.divider()
st.sidebar.markdown("#### :material/create_new_folder: Tambah Kategori Baru")
with st.sidebar.form("form_tambah_kategori", clear_on_submit=True):
    nama_folder_baru = st.text_input("Nama Folder", placeholder="Contoh: Kecerdasan_Buatan")
    keywords_baru = st.text_input("Kata Kunci (pisahkan koma)", placeholder="ai, machine learning, neural")
    submit_kat = st.form_submit_button("Simpan Aturan", icon=":material/add_circle:", use_container_width=True)

    if submit_kat and nama_folder_baru and keywords_baru:
        folder_bersih = nama_folder_baru.strip().replace(" ", "_")
        list_kw = [k.strip().lower() for k in keywords_baru.split(",") if k.strip()]
        if folder_bersih in aturan_aktif:
            aturan_aktif[folder_bersih].extend(list_kw)
            aturan_aktif[folder_bersih] = list(set(aturan_aktif[folder_bersih]))
        else:
            aturan_aktif[folder_bersih] = list_kw
        simpan_aturan(aturan_aktif)
        st.sidebar.success(f"Aturan untuk '{folder_bersih}' tersimpan.", icon=":material/check_circle:")
        st.rerun()

st.title(":material/folder_managed: Smart File & Assignment Organizer")
st.markdown("Unggah file tugas atau dokumen apa pun. Sistem otomatis memfilter berdasarkan **Nama File** maupun **Isi Dokumen (PDF/DOCX)** dan merapikannya ke folder yang sesuai.")

tab1, tab2, tab3 = st.tabs([
    ":material/upload_file: Upload & Sortir File",
    ":material/folder_open: Jelajahi Folder Tersimpan",
    ":material/rule: Daftar Aturan Kata Kunci"
])

# --- TAB 1: UPLOAD & RIWAYAT ---
with tab1:
    files_uploaded = st.file_uploader(
        "Seret & lepas banyak file sekaligus di sini (PDF, DOCX, XLSX, PPTX, Gambar, dll.)",
        accept_multiple_files=True
    )

    if files_uploaded:
        if st.button("Sortir & Simpan Sekarang", type="primary", icon=":material/auto_fix_high:"):
            for file in files_uploaded:
                folder_tujuan, metode, nama_baru = proses_file_upload(file, aturan_aktif, tambah_tgl)
                st.success(
                    f"**{file.name}** masuk ke folder **`{folder_tujuan}`** ({metode})",
                    icon=":material/task_alt:"
                )

    st.divider()
    st.subheader(":material/history: Riwayat Pemindahan File")
    if os.path.exists(LOG_FILE):
        df_log = pd.read_csv(LOG_FILE)
        st.dataframe(df_log.iloc[::-1], use_container_width=True, hide_index=True)
    else:
        st.info("Belum ada riwayat file yang dipindahkan.", icon=":material/info:")

# --- TAB 2: JELAJAHI ISI FOLDER ---
with tab2:
    daftar_folder = [d for d in os.listdir(FOLDER_TUJUAN_UTAMA) if os.path.isdir(os.path.join(FOLDER_TUJUAN_UTAMA, d))]
    if not daftar_folder:
        st.info("Folder arsip masih kosong. Silakan unggah file di Tab pertama.", icon=":material/info:")
    else:
        col1, col2 = st.columns([1, 2])
        with col1:
            pilih_folder = st.selectbox("Pilih Folder Mata Kuliah / Kategori:", sorted(daftar_folder))
        
        path_terpilih = os.path.join(FOLDER_TUJUAN_UTAMA, pilih_folder)
        isi_files = sorted(os.listdir(path_terpilih))

        with col2:
            st.metric(f"Jumlah File di {pilih_folder}", f"{len(isi_files)} file")

        st.divider()

        if isi_files:
            for item in isi_files:
                item_path = os.path.join(path_terpilih, item)
                c_nama, c_aksi = st.columns([4, 1])
                c_nama.markdown(f":material/description: `{item}`")
                with open(item_path, "rb") as f:
                    c_aksi.download_button(
                        "Unduh",
                        f.read(),
                        file_name=item,
                        key=f"dl_{pilih_folder}_{item}",
                        icon=":material/download:",
                        use_container_width=True
                    )
        else:
            st.write("Folder ini masih kosong.")


# --- TAB 3: KELOLA ATURAN ---
with tab3:
    st.subheader(":material/label: Daftar Kata Kunci Aktif")
    st.caption("Jika nama file atau isi halaman awal dokumen mengandung kata kunci berikut, file akan otomatis masuk ke folder terkait.")
    
    for folder, kws in aturan_aktif.items():
        with st.container(border=True):
            st.markdown(f"**:material/folder: {folder}**")
            st.markdown(f"Kata kunci: `{', '.join(kws)}`")

# ==========================================
# 5. FOOTER
# ==========================================
st.divider()
st.markdown(
    """
    <div style='text-align: center; color: gray; font-size: 0.9em; padding: 10px 0;'>
        Built with Python & Streamlit • <b>by Hilmy</b>
    </div>
    """,
    unsafe_allow_html=True
)