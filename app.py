# FORM CUTI GENERATOR - STREAMLIT VERSION
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
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
COUNTER_FILE = 'nomor_surat_counter.txt'

# ============================================
# FUNCTIONS
# ============================================
@st.cache_resource
def setup_gsheets(credentials_dict, spreadsheet_id):
    """Setup Google Sheets connection"""
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(spreadsheet_id).worksheet('DataPegawai')
    
    return sheet

def get_pegawai_data(sheet):
    """Get pegawai data from Google Sheets"""
    data = sheet.get_all_records()
    
    pegawai_dict = {}
    for row in data:
        pegawai_dict[row['Nama']] = {
            'nip': str(row['NIP']),
            'jabatan': row['Jabatan'],
            'atasan': row['Atasan Langsung'],
            'nip_atasan': str(row['NIP Atasan'])
        }
    
    return pegawai_dict

def get_next_nomor_surat():
    """Auto-increment nomor surat"""
    if os.path.exists(COUNTER_FILE):
        with open(COUNTER_FILE, 'r') as f:
            current = int(f.read().strip())
    else:
        current = 0
    
    next_number = current + 1
    
    with open(COUNTER_FILE, 'w') as f:
        f.write(str(next_number))
    
    return next_number

def create_overlay_pdf(data):
    """Create overlay PDF with data"""
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    width, height = A4
    
    can.setFont("Helvetica", 10)
    
    # Placeholder positions (adjust based on your template)
    positions = {
        'tanggalSurat': (450, height - 50),
        'nomorSurat': (280, height - 88),
        'namaPegawai': (150, height - 145),
        'nipPegawai': (150, height - 165),
        'jabatan': (150, height - 185),
        'masaKerja': (150, height - 205),
        'jumlahHari': (100, height - 375),
        'cutiTahunanSisa1': (220, height - 467),
        'cutiTahunanSisa2': (220, height - 487),
        'cutiTahunanTambahanSisa': (220, height - 567),
        'alamatCuti': (100, height - 625),
        'telpCuti': (150, height - 645),
        'atasanLangsung': (150, height - 735),
        'nipAtasan': (150, height - 755)
    }
    
    for key, (x, y) in positions.items():
        if key in data:
            can.drawString(x, y, str(data[key]))
    
    can.save()
    packet.seek(0)
    
    return packet

def merge_pdfs(template_path, overlay_packet):
    """Merge template with overlay"""
    template_pdf = PdfReader(template_path)
    overlay_pdf = PdfReader(overlay_packet)
    
    output = PdfWriter()
    
    page = template_pdf.pages[0]
    page.merge_page(overlay_pdf.pages[0])
    output.add_page(page)
    
    for page_num in range(1, len(template_pdf.pages)):
        output.add_page(template_pdf.pages[page_num])
    
    return output

def generate_pdf(pegawai_data, form_data):
    """Generate final PDF"""
    nama = form_data['nama_pegawai']
    if nama not in pegawai_data:
        raise ValueError(f"Pegawai '{nama}' tidak ditemukan")
    
    pegawai_info = pegawai_data[nama]
    nomor_surat = get_next_nomor_surat()
    
    complete_data = {
        'tanggalSurat': form_data['tanggal_surat'],
        'nomorSurat': f"{nomor_surat:04d}",
        'namaPegawai': nama,
        'nipPegawai': pegawai_info['nip'],
        'jabatan': pegawai_info['jabatan'],
        'atasanLangsung': pegawai_info['atasan'],
        'nipAtasan': pegawai_info['nip_atasan'],
        'masaKerja': form_data['masa_kerja'],
        'jumlahHari': form_data['jumlah_hari'],
        'cutiTahunanSisa1': form_data['cuti_tahunan_sisa1'],
        'cutiTahunanSisa2': form_data['cuti_tahunan_sisa2'],
        'cutiTahunanTambahanSisa': form_data['cuti_tambahan_sisa'],
        'alamatCuti': form_data['alamat_cuti'],
        'telpCuti': form_data['telp_cuti']
    }
    
    result_pdf = fill_pdf_form(TEMPLATE_PDF, complete_data)
    
    # Save to BytesIO for download
    pdf_buffer = io.BytesIO()
    result_pdf.write(pdf_buffer)
    pdf_buffer.seek(0)
    
    return pdf_buffer, complete_data

# ============================================
# STREAMLIT APP
# ============================================
def main():
    st.title("📄 Form Cuti Generator")
    st.markdown("**Mail Merge Application - Generate Formulir Cuti ke PDF**")
    
    # Sidebar - Configuration
    with st.sidebar:
        st.header("⚙️ Konfigurasi")
        
        st.subheader("1. Google Sheets Credentials")
        creds_file = st.file_uploader(
            "Upload JSON Credentials",
            type=['json'],
            help="Service account credentials dari Google Cloud"
        )
        
        spreadsheet_id = st.text_input(
            "Spreadsheet ID",
            help="ID dari Google Sheets Anda"
        )
        
        st.info("ℹ️ PDF akan di-generate otomatis tanpa template")
    
    # Main Form
    if creds_file and spreadsheet_id:
        try:
            # Load credentials
            creds_dict = json.load(creds_file)
            
            # Connect to Google Sheets
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
                    options=list(pegawai_data.keys())
                )
                
                if nama_pegawai:
                    info = pegawai_data[nama_pegawai]
                    st.info(f"""
                    **NIP:** {info['nip']}  
                    **Jabatan:** {info['jabatan']}  
                    **Atasan:** {info['atasan']}  
                    **NIP Atasan:** {info['nip_atasan']}
                    """)
                
                tanggal_surat = st.date_input("Tanggal Surat *")
                tanggal_formatted = tanggal_surat.strftime("%d-%B-%Y")
                
                masa_kerja = st.text_input(
                    "Masa Kerja *",
                    placeholder="Contoh: 5 tahun 3 bulan"
                )
                
                jumlah_hari = st.number_input(
                    "Jumlah Hari Cuti *",
                    min_value=1,
                    value=1
                )
            
            with col2:
                st.subheader("Data Cuti")
                cuti_tahunan_sisa1 = st.number_input(
                    "Sisa Cuti Tahunan 2025 *",
                    min_value=0,
                    value=0
                )
                
                cuti_tahunan_sisa2 = st.number_input(
                    "Sisa Cuti Tahunan 2026 *",
                    min_value=0,
                    value=0
                )
                
                cuti_tambahan_sisa = st.number_input(
                    "Sisa Cuti Tambahan 2026 *",
                    min_value=0,
                    value=0
                )
                
                alamat_cuti = st.text_area(
                    "Alamat Selama Cuti *",
                    placeholder="Masukkan alamat lengkap"
                )
                
                telp_cuti = st.text_input(
                    "Telepon Selama Cuti *",
                    placeholder="Contoh: 081234567890"
                )
            
            # Generate button
            st.divider()
            
            if st.button("🎯 Generate PDF", type="primary", use_container_width=True):
                if not all([nama_pegawai, masa_kerja, alamat_cuti, telp_cuti]):
                    st.error("⚠️ Mohon lengkapi semua field yang wajib diisi (*)")
                else:
                    with st.spinner("⏳ Generating PDF..."):
                        form_data = {
                            'nama_pegawai': nama_pegawai,
                            'tanggal_surat': tanggal_formatted,
                            'masa_kerja': masa_kerja,
                            'jumlah_hari': str(jumlah_hari),
                            'cuti_tahunan_sisa1': str(cuti_tahunan_sisa1),
                            'cuti_tahunan_sisa2': str(cuti_tahunan_sisa2),
                            'cuti_tambahan_sisa': str(cuti_tambahan_sisa),
                            'alamat_cuti': alamat_cuti,
                            'telp_cuti': telp_cuti
                        }
                        
                        try:
                            pdf_buffer, complete_data = generate_pdf(pegawai_data, form_data)
                            
                            st.success("✅ PDF berhasil dibuat!")
                            
                            # Show data preview
                            with st.expander("📋 Data yang digunakan"):
                                st.json(complete_data)
                            
                            # Download button
                            filename = f"Cuti_{nama_pegawai.replace(' ', '_')}_{complete_data['nomorSurat']}.pdf"
                            st.download_button(
                                label="📥 Download PDF",
                                data=pdf_buffer,
                                file_name=filename,
                                mime="application/pdf",
                                use_container_width=True
                            )
                            
                        except Exception as e:
                            st.error(f"❌ Error: {str(e)}")
        
        except Exception as e:
            st.error(f"❌ Gagal menghubungkan ke Google Sheets: {str(e)}")
    
    else:
        st.info("👈 Silakan lengkapi konfigurasi di sidebar terlebih dahulu")

if __name__ == "__main__":
    main()
