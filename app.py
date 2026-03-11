# FORM CUTI GENERATOR - STREAMLIT + DOCXTPL VERSION
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from docxtpl import DocxTemplate
from datetime import datetime
import io
import os
import json

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
    """Setup Google Sheets connection"""
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
    """Get pegawai data from Google Sheets"""
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
    """Auto-increment nomor surat (persist di file teks)"""
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

    # Context harus sama persis dengan placeholder di template DOCX
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

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer, context

# ============================================
# STREAMLIT APP
# ============================================
def main():
    st.title("📄 Form Cuti Generator")
    st.markdown("**Mail Merge Application - Generate Formulir Cuti ke DOCX (template BKN)**")

    # Sidebar - Configuration
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

    # Main Form
    if creds_file and spreadsheet_id:
        try:
            creds_dict = json.load(creds_file)

            with st.spinner("Menghubungkan ke Google Sheets..."):
                sheet = setup_gsheets(creds_dict, spreadsheet_id)
                pegawai_data = get_pegawai_data(sheet)

            st.success(f"✅ Berhasil memuat data {len(pegawai_data)} pegawai")

            # Form inputs
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

            if st.button("🎯 Generate DOCX", type="primary", use_container_width=True):
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
                    with st.spinner("⏳ Generating DOCX dari template..."):
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
                            docx_buffer, complete_data = generate_docx_from_template(
                                pegawai_data, form_data
                            )

                            st.success("✅ DOCX berhasil dibuat!")

                            with st.expander("📋 Data yang digunakan"):
                                st.json(complete_data)

                            filename = (
                                f"Cuti_{nama_pegawai.replace(' ', '_')}_"
                                f"{complete_data['nomorSurat']}.docx"
                            )
                            st.download_button(
                                label="📥 Download DOCX",
                                data=docx_buffer,
                                file_name=filename,
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
