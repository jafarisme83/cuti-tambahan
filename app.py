import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from docxtpl import DocxTemplate
import subprocess, os, json
from datetime import date

# ---------- CONFIG ----------
TEMPLATE_PATH = "template-placeholder.docx"
OUTPUT_DIR = "generated"
os.makedirs(OUTPUT_DIR, exist_ok=True)

GOOGLE_SHEET_ID = "1bNy8AurgGFLvRaKh3reeYNSoURCLzR6T3IKj1x0yHVk"
SHEET_PEGAWAI = "DataPegawai"
SHEET_KUOTA_PREFIX = "Kuota Cuti Tb "

BULAN_ID = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
    7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"
}

def format_tanggal_indo(d: date) -> str:
    return f"{d.day:02d}-{BULAN_ID[d.month]}-{d.year}"

# ---------- PAGE CONFIG & STYLING ----------
st.set_page_config(
    page_title="Formulir Cuti Digital",
    page_icon="🗂️",
    layout="centered",
    initial_sidebar_state="collapsed"
)

CUSTOM_CSS = """
<style>
    .main {
        background: linear-gradient(180deg, #f8fafc 0%, #eef2f7 100%);
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 780px;
    }
    .app-header {
        background: linear-gradient(135deg, #1e3a5f 0%, #2c5282 100%);
        padding: 28px 32px;
        border-radius: 16px;
        margin-bottom: 24px;
        box-shadow: 0 4px 14px rgba(30, 58, 95, 0.25);
    }
    .app-header h1 {
        color: white;
        font-size: 26px;
        font-weight: 700;
        margin: 0;
    }
    .app-header p {
        color: #cbd8e8;
        font-size: 14px;
        margin: 4px 0 0 0;
    }
    .section-card {
        background: white;
        padding: 20px 24px;
        border-radius: 12px;
        margin-bottom: 16px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        border-left: 4px solid #2c5282;
    }
    .section-title {
        font-size: 15px;
        font-weight: 700;
        color: #1e3a5f;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    div[data-testid="stFormSubmitButton"] button {
        background: linear-gradient(135deg, #2c5282 0%, #1e3a5f 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 12px 0;
        font-weight: 600;
        font-size: 15px;
        width: 100%;
        box-shadow: 0 3px 10px rgba(30, 58, 95, 0.3);
        transition: transform 0.15s ease;
    }
    div[data-testid="stFormSubmitButton"] button:hover {
        transform: translateY(-1px);
        box-shadow: 0 5px 14px rgba(30, 58, 95, 0.4);
    }
    .stDownloadButton button {
        border-radius: 10px;
        font-weight: 600;
    }
    .nomor-badge {
        background: #eaf3ff;
        color: #1e3a5f;
        padding: 10px 16px;
        border-radius: 10px;
        font-size: 14px;
        font-weight: 600;
        border: 1px solid #bcd6f5;
        margin-top: 4px;
    }
    div[data-testid="stTextInput"] input:disabled {
        background-color: #f4f6f8;
        color: #4a5568;
        font-weight: 500;
    }
    footer {visibility: hidden;}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

st.markdown("""
<div class="app-header">
    <h1>🗂️ Formulir Permintaan &amp; Pemberian Cuti</h1>
    <p>KPPN Tipe A2 Sungai Penuh · Generate PDF otomatis dari data pegawai</p>
</div>
""", unsafe_allow_html=True)

# ---------- GOOGLE SHEETS CLIENT ----------
@st.cache_resource
def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    if "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
    else:
        with open("service_account.json") as f:
            creds_dict = json.load(f)
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

def get_spreadsheet():
    client = get_gspread_client()
    return client.open_by_key(GOOGLE_SHEET_ID)

@st.cache_data(ttl=300)
def load_pegawai():
    sh = get_spreadsheet()
    ws = sh.worksheet(SHEET_PEGAWAI)
    records = ws.get_all_records()
    return pd.DataFrame(records)

def get_or_create_monitoring_ws(sh, tahun):
    sheet_name = tahun
    try:
        ws = sh.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows=200, cols=10)
        ws.append_row(["No.", "Tanggal Surat", "Nomor Surat", "Nama", "NIP",
                        "Lama Cuti (Hari)", "Tanggal Mulai Cuti", "Tanggal Akhir Cuti"])
    return ws

def get_next_nomor(tahun):
    sh = get_spreadsheet()
    ws = get_or_create_monitoring_ws(sh, tahun)
    values = ws.get_all_values()
    if len(values) <= 1:
        return 1
    nomor_list = []
    for row in values[1:]:
        if len(row) >= 3 and row[2].strip().isdigit():
            nomor_list.append(int(row[2]))
    return (max(nomor_list) + 1) if nomor_list else 1

def append_monitoring_row(tahun, nomor_surat, tanggal_surat, nama, nip, jumlah_hari, tanggal_mulai, tanggal_selesai):
    sh = get_spreadsheet()
    ws = get_or_create_monitoring_ws(sh, tahun)
    values = ws.get_all_values()
    next_no = len(values)
    new_row = [
        next_no, format_tanggal_indo(tanggal_surat), nomor_surat, nama, str(nip),
        jumlah_hari, format_tanggal_indo(tanggal_mulai), format_tanggal_indo(tanggal_selesai),
    ]
    ws.append_row(new_row, value_input_option="USER_ENTERED")

def get_or_create_kuota_ws(sh, tahun):
    sheet_name = f"{SHEET_KUOTA_PREFIX}{tahun}"
    try:
        ws = sh.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows=100, cols=4)
        ws.append_row([f"Kuota Cuti Tambahan Pegawai Tahun {tahun}", "", "", ""])
        ws.append_row(["Nama", "Kuota", "Terpakai", "Sisa"])
    return ws

def update_kuota(tahun, nama, jumlah_hari):
    sh = get_spreadsheet()
    ws = get_or_create_kuota_ws(sh, tahun)
    values = ws.get_all_values()
    header_row_idx = None
    for i, row in enumerate(values):
        if row and row[0].strip().lower() == "nama":
            header_row_idx = i
            break
    if header_row_idx is None:
        return
    target_row_idx = None
    for i in range(header_row_idx + 1, len(values)):
        if values[i] and values[i][0].strip().lower() == nama.strip().lower():
            target_row_idx = i
            break
    if target_row_idx is None:
        return
    row_num = target_row_idx + 1
    row_data = values[target_row_idx]
    def to_num(v):
        try:
            return float(v) if v not in (None, "", "-") else 0
        except ValueError:
            return 0
    kuota = to_num(row_data[1]) if len(row_data) > 1 else 0
    terpakai = to_num(row_data[2]) if len(row_data) > 2 else 0
    terpakai_baru = terpakai + jumlah_hari
    sisa_baru = kuota - terpakai_baru
    ws.update_cell(row_num, 3, terpakai_baru)
    ws.update_cell(row_num, 4, sisa_baru)

def convert_docx_to_pdf(docx_path, out_dir):
    subprocess.run([
        "soffice", "--headless", "--convert-to", "pdf",
        "--outdir", out_dir, docx_path
    ], check=True)
    return os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")

# ---------- DATA PEGAWAI ----------
df_pegawai = load_pegawai()

with st.form("form_cuti"):
    st.markdown('<div class="section-card"><div class="section-title">👤 Data Pegawai</div>', unsafe_allow_html=True)
    nama_pilihan = st.selectbox("Pilih Nama Pegawai", df_pegawai["nama"].tolist())
    row = df_pegawai[df_pegawai["nama"] == nama_pilihan].iloc[0]
    nip_pegawai = row["nip"]
    jabatan = row["jabatan"]
    atasan_langsung = row["atasan"]
    nip_atasan = row["nip_atasan"]

    c1, c2 = st.columns(2)
    with c1:
        st.text_input("NIP Pegawai", value=str(nip_pegawai), disabled=True)
        st.text_input("Atasan Langsung", value=atasan_langsung, disabled=True)
    with c2:
        st.text_input("Jabatan", value=jabatan, disabled=True)
        st.text_input("NIP Atasan", value=str(nip_atasan), disabled=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card"><div class="section-title">📅 Detail Surat</div>', unsafe_allow_html=True)
    tanggal_surat = st.date_input("Tanggal Surat", value=date.today())
    tahun_aktif = str(tanggal_surat.year)
    nomor_preview = get_next_nomor(tahun_aktif)
    st.markdown(f'<div class="nomor-badge">📌 Nomor Surat berikutnya: {nomor_preview} (tahun {tahun_aktif})</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card"><div class="section-title">✍️ Input Manual</div>', unsafe_allow_html=True)
    c3, c4 = st.columns(2)
    with c3:
        masa_kerja = st.text_input("Masa Kerja", placeholder="contoh: 5 Tahun 3 Bulan")
        jumlah_hari = st.number_input("Jumlah Hari Cuti", min_value=1, step=1)
        tanggal_mulai = st.date_input("Tanggal Mulai Cuti")
        cuti_sisa_1 = st.text_input("Sisa Cuti Tahunan 2025")
        cuti_tambahan_sisa = st.text_input("Sisa Cuti Tahunan Tambahan 2026")
    with c4:
        telp_cuti = st.text_input("No. Telp Selama Cuti")
        tanggal_selesai = st.date_input("Tanggal Selesai Cuti")
        cuti_sisa_2 = st.text_input("Sisa Cuti Tahunan 2026")

    alasan_cuti = st.text_area("Alasan Cuti", height=80)
    alamat_cuti = st.text_area("Alamat Selama Cuti", height=80)
    st.markdown('</div>', unsafe_allow_html=True)

    submitted = st.form_submit_button("🚀 Generate Formulir PDF")

if submitted:
    with st.spinner("Sedang membuat dokumen dan memperbarui data..."):
        tahun_aktif = str(tanggal_surat.year)
        nomor_surat = get_next_nomor(tahun_aktif)

        context = {
            "tanggalSurat": format_tanggal_indo(tanggal_surat),
            "nomorSurat": str(nomor_surat),
            "namaPegawai": nama_pilihan,
            "nipPegawai": str(nip_pegawai),
            "jabatan": jabatan,
            "masaKerja": masa_kerja,
            "atasanLangsung": atasan_langsung,
            "nipAtasan": str(nip_atasan),
            "jumlahHari": str(jumlah_hari),
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
            try:
                append_monitoring_row(
                    tahun_aktif, nomor_surat, tanggal_surat, nama_pilihan, nip_pegawai,
                    int(jumlah_hari), tanggal_mulai, tanggal_selesai
                )
                update_kuota(tahun_aktif, nama_pilihan, int(jumlah_hari))
                st.toast("Google Sheets Monitoring & Kuota Cuti berhasil diupdate ✅", icon="✅")
            except Exception as e_sheets:
                st.warning(f"PDF berhasil dibuat, namun gagal update Google Sheets: {e_sheets}")

            st.success(f"✅ Formulir cuti berhasil dibuat — Nomor Surat: {nomor_surat}")
            with open(pdf_out, "rb") as f:
                st.download_button(
                    label="⬇️ Download PDF Formulir Cuti",
                    data=f,
                    file_name=filename_base + ".pdf",
                    mime="application/pdf"
                )
        except Exception as e:
            st.error(f"Gagal konversi ke PDF: {e}")
            st.info("Pastikan LibreOffice (soffice) terinstall di server deployment.")
