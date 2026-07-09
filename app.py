import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import subprocess, os, json, locale
from datetime import date

# ---------- CONFIG ----------
TEMPLATE_PATH = "template-placeholder.docx"
DATA_PEGAWAI_PATH = "data_pegawai.csv"   # fallback lokal jika Google Sheets gagal diakses
COUNTER_PATH = "nomor_surat_counter.json"
OUTPUT_DIR = "generated"
os.makedirs(OUTPUT_DIR, exist_ok=True)

BULAN_ID = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
    7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"
}

def format_tanggal_indo(d: date) -> str:
    return f"{d.day:02d}-{BULAN_ID[d.month]}-{d.year}"

# ---------- NOMOR SURAT AUTO INCREMENT ----------
def get_next_nomor():
    if not os.path.exists(COUNTER_PATH):
        with open(COUNTER_PATH, "w") as f:
            json.dump({"last": 0}, f)
    with open(COUNTER_PATH, "r") as f:
        data = json.load(f)
    return data["last"] + 1

def commit_nomor(nomor):
    with open(COUNTER_PATH, "w") as f:
        json.dump({"last": nomor}, f)

# ---------- LOAD DATA PEGAWAI (GOOGLE SHEETS) ----------
GOOGLE_SHEET_ID = "1bNy8AurgGFLvRaKh3reeYNSoURCLzR6T3IKj1x0yHVk"
GOOGLE_SHEET_CSV_URL = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/export?format=csv&gid=0"

@st.cache_data(ttl=300)
def load_pegawai():
    # Kolom wajib pada sheet: nama, nip, jabatan, atasan, nip_atasan
    try:
        return pd.read_csv(GOOGLE_SHEET_CSV_URL)
    except Exception:
        return pd.read_csv(DATA_PEGAWAI_PATH)

def convert_docx_to_pdf(docx_path, out_dir):
    # Membutuhkan LibreOffice terpasang di server (soffice)
    subprocess.run([
        "soffice", "--headless", "--convert-to", "pdf",
        "--outdir", out_dir, docx_path
    ], check=True)
    return os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")

# ---------- STREAMLIT UI ----------
st.set_page_config(page_title="Formulir Cuti - Mail Merge", layout="centered")
st.title("📄 Generator Formulir Cuti (Mail Merge PDF)")

df_pegawai = load_pegawai()

with st.form("form_cuti"):
    st.subheader("I. Data Pegawai (otomatis)")
    nama_pilihan = st.selectbox("Pilih Nama Pegawai", df_pegawai["nama"].tolist())
    row = df_pegawai[df_pegawai["nama"] == nama_pilihan].iloc[0]

    nip_pegawai = row["nip"]
    jabatan = row["jabatan"]
    atasan_langsung = row["atasan"]
    nip_atasan = row["nip_atasan"]

    st.text_input("NIP Pegawai", value=str(nip_pegawai), disabled=True)
    st.text_input("Jabatan", value=jabatan, disabled=True)
    st.text_input("Atasan Langsung", value=atasan_langsung, disabled=True)
    st.text_input("NIP Atasan", value=str(nip_atasan), disabled=True)

    st.subheader("II. Detail Surat")
    tanggal_surat = st.date_input("Tanggal Surat", value=date.today())
    nomor_preview = get_next_nomor()
    st.info(f"Nomor Surat berikutnya: **{nomor_preview}** (auto-increment)")

    st.subheader("III. Input Manual")
    masa_kerja = st.text_input("Masa Kerja (contoh: 5 Tahun 3 Bulan)")
    jumlah_hari = st.text_input("Jumlah Hari Cuti")
    cuti_sisa_1 = st.text_input("Sisa Cuti Tahunan 2025")
    cuti_sisa_2 = st.text_input("Sisa Cuti Tahunan 2026")
    cuti_tambahan_sisa = st.text_input("Sisa Cuti Tahunan Tambahan 2026")
    tanggal_mulai = st.date_input("Tanggal Mulai Cuti")
    tanggal_selesai = st.date_input("Tanggal Selesai Cuti")
    alasan_cuti = st.text_area("Alasan Cuti")
    alamat_cuti = st.text_area("Alamat Selama Cuti")
    telp_cuti = st.text_input("No. Telp Selama Cuti")

    submitted = st.form_submit_button("Generate PDF")

if submitted:
    nomor_surat = get_next_nomor()

    context = {
        "tanggalSurat": format_tanggal_indo(tanggal_surat),
        "nomorSurat": str(nomor_surat),
        "namaPegawai": nama_pilihan,
        "nipPegawai": str(nip_pegawai),
        "jabatan": jabatan,
        "masaKerja": masa_kerja,
        "atasanLangsung": atasan_langsung,
        "nipAtasan": str(nip_atasan),
        "jumlahHari": jumlah_hari,
        "tanggalMulai": format_tanggal_indo(tanggal_mulai),
        "tanggalSelesai": format_tanggal_indo(tanggal_selesai),
        "alasanCuti": alasan_cuti,
        "cutiTahunanSisa1": cuti_sisa_1,
        "cutiTahunanSisa2": cuti_sisa_2,
        "cutiTahunanTambahanSisa": cuti_tambahan_sisa,
        "alamatCuti": alamat_cuti,
        "telpCuti": telp_cuti,
    }

    doc = DocxTemplate(TEMPLATE_PATH)
    doc.render(context)

    filename_base = f"Cuti_{nama_pilihan.replace(' ', '_')}_{nomor_surat}"
    docx_out = os.path.join(OUTPUT_DIR, filename_base + ".docx")
    doc.save(docx_out)

    try:
        pdf_out = convert_docx_to_pdf(docx_out, OUTPUT_DIR)
        commit_nomor(nomor_surat)  # commit nomor hanya jika sukses generate

        st.success(f"Formulir cuti berhasil dibuat dengan Nomor Surat: {nomor_surat}")
        with open(pdf_out, "rb") as f:
            st.download_button(
                label="⬇️ Download PDF",
                data=f,
                file_name=filename_base + ".pdf",
                mime="application/pdf"
            )
    except Exception as e:
        st.error(f"Gagal konversi ke PDF: {e}")
        st.info("Pastikan LibreOffice (soffice) terinstall di server deployment.")