import io
from datetime import datetime

import streamlit as st
import gspread
from google.oauth2 import service_account
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from pypdf import PdfReader, PdfWriter


# ================== KONEKSI GOOGLE SHEETS ==================

@st.cache_resource
def get_gspread_client():
    service_info = st.secrets["gcp_service_account"]

    credentials = service_account.Credentials.from_service_account_info(
        dict(service_info),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    return gspread.authorize(credentials)


# Ganti dengan ID spreadsheet Anda
SPREADSHEET_ID = "1kR7zbXQC5CQ_yO6C388JaIbgdlU03Tv85D6vrHPkELw"
PEG_SHEET_NAME = "pegawai"
COUNTER_SHEET_NAME = "counter"


@st.cache_resource
def get_worksheets():
    client = get_gspread_client()
    sh = client.open_by_key(SPREADSHEET_ID)
    peg_sheet = sh.worksheet(PEG_SHEET_NAME)
    counter_sheet = sh.worksheet(COUNTER_SHEET_NAME)
    return peg_sheet, counter_sheet


def get_auto_increment_number():
    _, counter_sheet = get_worksheets()
    current_val = counter_sheet.acell("A2").value
    if not current_val:
        current = 1
    else:
        current = int(current_val)
    next_val = current + 1
    counter_sheet.update_acell("A2", str(next_val))
    return current


def get_pegawai_list():
    peg_sheet, _ = get_worksheets()
    data = peg_sheet.get_all_records()
    return [row["Nama"] for row in data]


def get_nip_by_name(nama):
    peg_sheet, _ = get_worksheets()
    data = peg_sheet.get_all_records()
    for row in data:
        if row["Nama"] == nama:
            return str(row["NIP"])
    return ""


def format_tanggal_indonesia(date_obj):
    try:
        return date_obj.strftime("%-d %B %Y")
    except ValueError:
        return date_obj.strftime("%d %B %Y")


# ================== GENERATE PDF DARI TEMPLATE ==================

TEMPLATE_PATH = "Form-Cuti-Tambahan-2025-SIC-Tb-23-Jafar.pdf"
PAGE_SIZE = A4


def fill_cuti_on_template(
    output_buffer,
    tanggal_form,
    nomor_surat,
    nama_pegawai,
    nip_pegawai,
    jabatan,
    masa_kerja,
    unit_kerja,
    jenis_cuti,
    alasan_cuti,
    lama_cuti,
    tanggal_mulai,
    tanggal_selesai,
    sisa_cuti_2023,
    sisa_cuti_2024,
    sisa_cuti_2025,
    sisa_cuti_tambahan_2025,
    alamat_cuti,
    telp,
    nama_atasan_langsung,
    nip_atasan_langsung,
    nama_pejabat_berwenang,
    nip_pejabat_berwenang,
):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=PAGE_SIZE)
    width, height = PAGE_SIZE

    # Header tanggal & nomor
    can.setFont("Helvetica", 10)
    can.drawString(150 * mm, 285 * mm, tanggal_form)
    can.drawCentredString(width / 2, 268 * mm, f"NOMOR {nomor_surat}")

    # I. DATA PEGAWAI
    can.setFont("Helvetica", 10)
    can.drawString(50 * mm, 252 * mm, nama_pegawai)
    can.drawString(50 * mm, 246 * mm, nip_pegawai)
    can.drawString(50 * mm, 240 * mm, jabatan)
    can.drawString(50 * mm, 234 * mm, masa_kerja)
    can.drawString(50 * mm, 228 * mm, unit_kerja)

    # II. JENIS CUTI
    jenis_list = [
        "Cuti tahunan",
        "Cuti Besar",
        "Cuti Sakit",
        "Cuti Melahirkan",
        "Cuti Karena Alasan Penting",
        "Cuti di Luar Tanggungan Negara",
    ]
    y_jenis = 210
    can.setFont("Helvetica", 12)
    for jn in jenis_list:
        if jn.lower() == jenis_cuti.lower():
            can.drawString(70 * mm, y_jenis * mm, "✓")
        y_jenis -= 6

    # III. ALASAN CUTI
    can.setFont("Helvetica", 10)
    can.drawString(20 * mm, 174 * mm, alasan_cuti)

    # IV. LAMANYA CUTI
    can.drawString(40 * mm, 160 * mm, lama_cuti)
    can.drawString(60 * mm, 154 * mm, tanggal_mulai)
    can.drawString(120 * mm, 154 * mm, tanggal_selesai)

    # V. CATATAN CUTI
    can.setFont("Helvetica", 9)
    can.drawString(40 * mm, 128 * mm, str(sisa_cuti_2023))
    can.drawString(40 * mm, 122 * mm, str(sisa_cuti_2024))
    can.drawString(40 * mm, 116 * mm, str(sisa_cuti_2025))
    can.drawString(120 * mm, 128 * mm, str(sisa_cuti_tambahan_2025))

    # VI. ALAMAT CUTI
    can.setFont("Helvetica", 10)
    can.drawString(20 * mm, 100 * mm, alamat_cuti)
    can.drawString(35 * mm, 94 * mm, telp)

    # TTD pemohon
    can.drawString(140 * mm, 84 * mm, nama_pegawai)
    can.setFont("Helvetica", 8)
    can.drawString(140 * mm, 80 * mm, f"NIP {nip_pegawai}")

    # VII. Atasan langsung
    can.setFont("Helvetica", 12)
    can.drawString(60 * mm, 66 * mm, "✓")
    can.setFont("Helvetica", 10)
    can.drawString(140 * mm, 48 * mm, nama_atasan_langsung)
    can.setFont("Helvetica", 8)
    can.drawString(140 * mm, 44 * mm, f"NIP {nip_atasan_langsung}")

    # VIII. Pejabat berwenang
    can.setFont("Helvetica", 12)
    can.drawString(60 * mm, 30 * mm, "✓")
    can.setFont("Helvetica", 10)
    can.drawString(140 * mm, 16 * mm, nama_pejabat_berwenang)
    can.setFont("Helvetica", 8)
    can.drawString(140 * mm, 12 * mm, f"NIP {nip_pejabat_berwenang}")

    can.save()
    packet.seek(0)

    overlay_pdf = PdfReader(packet)
    template_pdf = PdfReader(open(TEMPLATE_PATH, "rb"))
    template_page = template_pdf.pages[0]
    overlay_page = overlay_pdf.pages[0]
    template_page.merge_page(overlay_page)

    writer = PdfWriter()
    writer.add_page(template_page)
    writer.write(output_buffer)
    output_buffer.seek(0)


# ================== STREAMLIT UI ==================

def main():
    st.title("Form Mail Merge Cuti Tambahan (PDF)")

    # Koneksi awal ke sheets (untuk memastikan kredensial OK)
    get_worksheets()

    # Tanggal form
    tgl_form = st.date_input("Tanggal Formulir (kuning)", value=datetime.today())
    tgl_form_str = format_tanggal_indonesia(tgl_form)

    # Nomor auto
    if "nomor_auto" not in st.session_state:
        st.session_state["nomor_auto"] = get_auto_increment_number()

    col_num1, col_num2 = st.columns([1, 2])
    with col_num1:
        if st.button("Ambil nomor baru"):
            st.session_state["nomor_auto"] = get_auto_increment_number()
    nomor_auto = st.session_state["nomor_auto"]
    st.write(f"Nomor urut: {nomor_auto}")
    nomor_surat_default = f"SICTb-{nomor_auto}/KPN.0603/2025"
    nomor_surat = st.text_input("Nomor Surat (hijau)", value=nomor_surat_default)

    # Nama & NIP
    pegawai_list = get_pegawai_list()
    nama_pegawai = st.selectbox("Nama Pegawai (abu-abu)", options=pegawai_list)
    nip_pegawai = get_nip_by_name(nama_pegawai)
    st.text_input("NIP Pegawai (linked, abu-abu)", value=nip_pegawai, disabled=True)

    # Input pink
    jabatan = st.text_input("Jabatan (pink)", value="")
    masa_kerja = st.text_input("Masa Kerja (pink)", value="")
    unit_kerja = st.text_area("Unit Kerja (pink)", value="")
    jenis_cuti = st.selectbox(
        "Jenis Cuti (pink)",
        [
            "Cuti tahunan",
            "Cuti Besar",
            "Cuti Sakit",
            "Cuti Melahirkan",
            "Cuti Karena Alasan Penting",
            "Cuti di Luar Tanggungan Negara",
        ],
        index=0,
    )
    alasan_cuti = st.text_area("Alasan Cuti (pink)", value="")
    lama_cuti = st.text_input("Lamanya Cuti (contoh: 6 hari, pink)", value="")
    tgl_mulai = st.date_input("Tanggal Mulai Cuti (pink)", value=datetime.today())
    tgl_selesai = st.date_input("Tanggal Selesai Cuti (pink)", value=datetime.today())
    tgl_mulai_str = format_tanggal_indonesia(tgl_mulai)
    tgl_selesai_str = format_tanggal_indonesia(tgl_selesai)

    col_sisa1, col_sisa2 = st.columns(2)
    with col_sisa1:
        sisa_cuti_2023 = st.text_input("Sisa Cuti 2023 (pink)", value="")
        sisa_cuti_2024 = st.text_input("Sisa Cuti 2024 (pink)", value="")
        sisa_cuti_2025 = st.text_input("Sisa Cuti 2025 (pink)", value="")
    with col_sisa2:
        sisa_cuti_tambahan_2025 = st.text_input("Sisa Cuti Tambahan 2025 (pink)", value="")

    alamat_cuti = st.text_area("Alamat Selama Cuti (pink)", value="")
    telp = st.text_input("Telp (pink)", value="")

    st.subheader("Data Atasan dan Pejabat Berwenang")
    nama_atasan_langsung = st.text_input("Nama Atasan Langsung", value="Setiyono")
    nip_atasan_langsung = st.text_input("NIP Atasan Langsung", value="197311161996021001")
    nama_pejabat_berwenang = st.text_input("Nama Pejabat Berwenang", value="Setiyono")
    nip_pejabat_berwenang = st.text_input("NIP Pejabat Berwenang", value="197311161996021001")

    if st.button("Generate PDF"):
        buffer = io.BytesIO()
        fill_cuti_on_template(
            buffer,
            tanggal_form=tgl_form_str,
            nomor_surat=nomor_surat,
            nama_pegawai=nama_pegawai,
            nip_pegawai=nip_pegawai,
            jabatan=jabatan,
            masa_kerja=masa_kerja,
            unit_kerja=unit_kerja,
            jenis_cuti=jenis_cuti,
            alasan_cuti=alasan_cuti,
            lama_cuti=lama_cuti,
            tanggal_mulai=tgl_mulai_str,
            tanggal_selesai=tgl_selesai_str,
            sisa_cuti_2023=sisa_cuti_2023,
            sisa_cuti_2024=sisa_cuti_2024,
            sisa_cuti_2025=sisa_cuti_2025,
            sisa_cuti_tambahan_2025=sisa_cuti_tambahan_2025,
            alamat_cuti=alamat_cuti,
            telp=telp,
            nama_atasan_langsung=nama_atasan_langsung,
            nip_atasan_langsung=nip_atasan_langsung,
            nama_pejabat_berwenang=nama_pejabat_berwenang,
            nip_pejabat_berwenang=nip_pejabat_berwenang,
        )

        file_name = f"Form-Cuti-{nama_pegawai.replace(' ', '-')}-{nomor_auto}.pdf"
        st.download_button(
            label="Download PDF",
            data=buffer,
            file_name=file_name,
            mime="application/pdf",
        )


if __name__ == "__main__":
    main()
