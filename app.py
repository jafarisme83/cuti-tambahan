# FORM CUTI GENERATOR - STREAMLIT + DOCXTPL + PDF
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from docxtpl import DocxTemplate
from datetime import datetime
import io
import os
import json
import subprocess

try:
    from docx2pdf import convert as docx2pdf_convert
    DOCX2PDF_AVAILABLE = True
except Exception:
    DOCX2PDF_AVAILABLE = False

# ============================================
# PAGE CONFIG
# ============================================
st.set_page_config(
    page_title="Form Cuti Generator",
    page_icon="📄",
    layout="wide"
)

# ============================================
# CONFIGURATION
# ============================================
TEMPLATE_DOCX = "template placeholder.docx"
COUNTER_FILE = "nomor_surat_counter.txt"

# ============================================
# FUNCTIONS
# ============================================
@st.cache_resource
def setup_gsheets(credentials_dict, spreadsheet_id):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(
        credentials_dict, scope
    )
    client = gspread.authorize(creds)
    sheet = client.open_by_key(spreadsheet_id).worksheet("DataPegawai")
    return sheet


def get_pegawai_data(sheet):
    data = sheet.get_all_records()
    pegawai_dict = {}
    for row in data:
        pegawai_dict[row["Nama"]] = {
            "nip": str(row["NIP"]),
            "jabatan": row["Jabatan"],
            "atasan": row["Atasan Langsung"],
            "nip_atasan": str(row["NIP Atasan"]),
        }
    return pegawai_dict


def get_next_nomor_surat():
    if os.path.exists(COUNTER_FILE):
        with open(COUNTER_FILE, "r") as f:
            current = int(f.read().strip())
    else:
        current = 0
    next_number = current + 1
    with open(COUNTER_FILE, "w") as f:
        f.write(str(next_number))
    return next_number


def generate_docx_from_template(pegawai_data, form_data):
    """Generate DOCX dari template DOCX dengan placeholder Jinja2"""
    nama = form_data["nama_pegawai"]
    if nama not in pegawai_data:
        raise ValueError(f"Pegawai '{nama}' tidak ditemukan")

    pegawai_info = pegawai_data[nama]
    nomor_surat = get_next_nomor_surat()

    context = {
        "tanggalSurat": form_data["tanggal_surat"],
        "nomorSurat": f"{nomor_surat:04d}",
        "namaPegawai": nama,
        "nipPegawai": pegawai_info["nip"],
        "jabatan": pegawai_info["jabatan"],
        "atasanLangsung": pegawai_info["atasan"],
        "nipAtasan": pegawai_info["nip_atasan"],
        "masaKerja": form_data["masa_kerja"],
        "jumlahHari": form_data["jumlah_hari"],
        "tanggalMulai": form_data["tanggal_mulai"],
        "tanggalSelesai": form_data["tanggal_selesai"],
        "alasanCuti": form_data["alasan_cuti"],
        "cutiTahunanSisa1": form_data["cuti_tahunan_sisa1"],
        "cutiTahunanSisa2": form_data["cuti_tahunan_sisa2"],
        "cutiTahunanTambahanSisa": form_data["cuti_tambahan_sisa"],
        "alamatCuti": form_data["alamat_cuti"],
        "telpCuti": form_data["telp_cuti"],
    }

    if not os.path.exists(TEMPLATE_DOCX):
        raise FileNotFoundError(
            f"Template DOCX '{TEMPLATE_DOCX}' tidak ditemukan di folder project."
        )

    doc = DocxTemplate(TEMPLATE_DOCX)
    doc.render(context)

    # Simpan ke file sementara di disk untuk konversi PDF
    os.makedirs("tmp_docs", exist_ok=True)
    docx_filename = (
        f"tmp_docs/Cuti_{nama.replace(' ', '_')}_{context['nomorSurat']}.docx"
    )
    doc.save(docx_filename)

    # Juga simpan ke buffer untuk download DOCX langsung
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer, docx_filename, context


def convert_docx_to_pdf(docx_path: str) -> bytes:
    """
    Convert DOCX ke PDF.
    Urutan:
    1) Coba LibreOffice (libreoffice / soffice)
    2) Fallback ke docx2pdf (jika tersedia)
    Return: bytes PDF (untuk download). Raise Exception kalau gagal.
    """
    base_dir = os.path.dirname(docx_path)
    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(base_dir, f"{base_name}.pdf")

    # 1. Coba LibreOffice headless
    for exe in ["libreoffice", "soffice"]:
        try:
            result = subprocess.run(
                [
                    exe,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    base_dir,
                    docx_path,
                ],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=60,
            )
            if result.returncode == 0 and os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f:
                    pdf_bytes = f.read()
                return pdf_bytes
        except FileNotFoundError:
            # Executable tidak ada, lanjut ke metode lain
            pass
        except Exception as e:
            # Kalau error lain, lanjut coba metode berikutnya
            print("LibreOffice error:", e)

    # 2. Fallback ke docx2pdf (Windows/macOS dengan MS Word)
    if DOCX2PDF_AVAILABLE:
        try:
            docx2pdf_convert(docx_path, pdf_path)
            if os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f:
                    pdf_bytes = f.read()
                return pdf_bytes
        except Exception as e:
            print("docx2pdf error:", e)

    raise RuntimeError(
        "Konversi DOCX → PDF gagal. Pastikan LibreOffice atau MS Word (docx2pdf) terpasang di server."
    )

# ============================================
# STREAMLIT APP
# ============================================
def main():
    st.title("📄 Form Cuti Generator")
    st.markdown("**Mail Merge Application - Generate Formulir Cuti ke PDF (dari template DOCX)**")

    with st.sidebar:
        st.header("⚙️ Konfigurasi")

        st.subheader("1. Google Sheets Credentials")
        creds_file = st.file_uploader(
            "Upload JSON Credentials",
            type=["json"],
            help="Service account credentials dari Google Cloud",
        )

        spreadsheet_id = st.text_input(
            "Spreadsheet ID",
            help="ID dari Google Sheets Anda (bukan URL penuh)",
        )

        st.subheader("2. Template DOCX")
        st.info(
            f"Pastikan file **{TEMPLATE_DOCX}** sudah ada di folder project.\n"
            "Template menggunakan placeholder Jinja2, misalnya {{namaPegawai}}, {{nipPegawai}}, dst."
        )

    if creds_file and spreadsheet_id:
        try:
            creds_dict = json.load(creds_file)

            with st.spinner("Menghubungkan ke Google Sheets..."):
                sheet = setup_gsheets(creds_dict, spreadsheet_id)
                pegawai_data = get_pegawai_data(sheet)

            st.success(f"✅ Berhasil memuat data {len(pegawai_data)} pegawai")

            st.header("📝 Isi Formulir Cuti")

            col1, col2 = st.columns(2)

            with col1:
                st.subheader("Data Pegawai")
                nama_pegawai = st.selectbox(
                    "Nama Pegawai *",
                    options=list(pegawai_data.keys()),
                )

                if nama_pegawai:
                    info = pegawai_data[nama_pegawai]
                    st.info(
                        f"**NIP:** {info['nip']}\n\n"
                        f"**Jabatan:** {info['jabatan']}\n\n"
                        f"**Atasan:** {info['atasan']}\n\n"
                        f"**NIP Atasan:** {info['nip_atasan']}"
                    )

                tanggal_surat = st.date_input("Tanggal Surat *")
                tanggal_formatted = tanggal_surat.strftime("%d-%B-%Y")

                masa_kerja = st.text_input(
                    "Masa Kerja *",
                    placeholder="Contoh: 5 tahun 3 bulan",
                )

                jumlah_hari = st.number_input(
                    "Jumlah Hari Cuti *",
                    min_value=1,
                    value=1,
                )

                col_tgl1, col_tgl2 = st.columns(2)
                with col_tgl1:
                    tanggal_mulai = st.date_input("Tanggal Mulai Cuti *")
                with col_tgl2:
                    tanggal_selesai = st.date_input("Tanggal Selesai Cuti *")

                alasan_cuti = st.text_area(
                    "Alasan Cuti *",
                    placeholder="Tulis alasan cuti",
                )

            with col2:
                st.subheader("Sisa Cuti & Kontak")
                cuti_tahunan_sisa1 = st.number_input(
                    "Sisa Cuti Tahunan 2025 *",
                    min_value=0,
                    value=0,
                )

                cuti_tahunan_sisa2 = st.number_input(
                    "Sisa Cuti Tahunan 2026 *",
                    min_value=0,
                    value=0,
                )

                cuti_tambahan_sisa = st.number_input(
                    "Sisa Cuti Tambahan 2026 *",
                    min_value=0,
                    value=0,
                )

                alamat_cuti = st.text_area(
                    "Alamat Selama Cuti *",
                    placeholder="Masukkan alamat lengkap",
                )

                telp_cuti = st.text_input(
                    "Telepon Selama Cuti *",
                    placeholder="Contoh: 081234567890",
                )

            st.divider()

            if st.button("🎯 Generate PDF", type="primary", use_container_width=True):
                required = [
                    nama_pegawai,
                    masa_kerja,
                    alasan_cuti,
                    alamat_cuti,
                    telp_cuti,
                ]
                if not all(required):
                    st.error("⚠️ Mohon lengkapi semua field yang wajib diisi (*)")
                else:
                    with st.spinner("⏳ Membuat DOCX dari template..."):
                        form_data = {
                            "nama_pegawai": nama_pegawai,
                            "tanggal_surat": tanggal_formatted,
                            "masa_kerja": masa_kerja,
                            "jumlah_hari": str(jumlah_hari),
                            "tanggal_mulai": tanggal_mulai.strftime("%d-%B-%Y"),
                            "tanggal_selesai": tanggal_selesai.strftime("%d-%B-%Y"),
                            "alasan_cuti": alasan_cuti,
                            "cuti_tahunan_sisa1": str(cuti_tahunan_sisa1),
                            "cuti_tahunan_sisa2": str(cuti_tahunan_sisa2),
                            "cuti_tambahan_sisa": str(cuti_tambahan_sisa),
                            "alamat_cuti": alamat_cuti,
                            "telp_cuti": telp_cuti,
                        }

                        try:
                            docx_buffer, docx_path, complete_data = generate_docx_from_template(
                                pegawai_data, form_data
                            )

                            st.success("✅ DOCX berhasil dibuat. Mengonversi ke PDF...")

                            pdf_bytes = None
                            pdf_error = None
                            try:
                                pdf_bytes = convert_docx_to_pdf(docx_path)
                            except Exception as e:
                                pdf_error = str(e)

                            if pdf_bytes:
                                st.success("✅ PDF berhasil dibuat!")

                                with st.expander("📋 Data yang digunakan"):
                                    st.json(complete_data)

                                pdf_filename = (
                                    f"Cuti_{nama_pegawai.replace(' ', '_')}_"
                                    f"{complete_data['nomorSurat']}.pdf"
                                )
                                st.download_button(
                                    label="📥 Download PDF",
                                    data=pdf_bytes,
                                    file_name=pdf_filename,
                                    mime="application/pdf",
                                    use_container_width=True,
                                )

                                docx_download_name = (
                                    f"Cuti_{nama_pegawai.replace(' ', '_')}_"
                                    f"{complete_data['nomorSurat']}.docx"
                                )
                                st.download_button(
                                    label="📥 Download DOCX (opsional)",
                                    data=docx_buffer,
                                    file_name=docx_download_name,
                                    mime=(
                                        "application/"
                                        "vnd.openxmlformats-officedocument.wordprocessingml.document"
                                    ),
                                    use_container_width=True,
                                )
                            else:
                                st.error(
                                    "❌ Konversi DOCX → PDF gagal di server.\n"
                                    "Anda masih bisa download file DOCX di bawah ini."
                                )
                                if pdf_error:
                                    st.caption(f"Detail error: {pdf_error}")

                                docx_download_name = (
                                    f"Cuti_{nama_pegawai.replace(' ', '_')}_"
                                    f"{complete_data['nomorSurat']}.docx"
                                )
                                st.download_button(
                                    label="📥 Download DOCX",
                                    data=docx_buffer,
                                    file_name=docx_download_name,
                                    mime=(
                                        "application/"
                                        "vnd.openxmlformats-officedocument.wordprocessingml.document"
                                    ),
                                    use_container_width=True,
                                )

                        except Exception as e:
                            st.error(f"❌ Error: {str(e)}")

        except Exception as e:
            st.error(f"❌ Gagal menghubungkan ke Google Sheets: {str(e)}")

    else:
        st.info("👈 Silakan lengkapi konfigurasi di sidebar terlebih dahulu")


if __name__ == "__main__":
    main()
