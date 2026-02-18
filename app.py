import streamlit as st
from datetime import datetime
from main import (  # jika Anda taruh fungsi di file lain, sesuaikan import
    get_auto_increment_number,
    get_pegawai_list,
    get_nip_by_name,
    format_tanggal_indonesia,
    generate_cuti_pdf,
)

def main():
    st.title("Form Mail Merge Cuti (PDF)")

    # 1. Tanggal kuning (dd-Month-YYYY) -> pakai date_input lalu format
    tgl_form = st.date_input("Tanggal Form (kuning)", value=datetime.today())
    tgl_form_str = tgl_form.strftime("%d-%B-%Y")
    tgl_form_fmt = format_tanggal_indonesia(tgl_form_str)

    # 2. Nomor auto-increment (hijau)
    if st.button("Ambil Nomor Otomatis"):
        nomor_auto = get_auto_increment_number()
        st.session_state["nomor_auto"] = nomor_auto
    nomor_auto = st.session_state.get("nomor_auto", 1)
    st.write(f"Nomor urut (auto): {nomor_auto}")

    # Misal: format nomor surat SICTb-{nomor_auto}/KPN.0603/2025
    nomor_surat = st.text_input(
        "Nomor Surat (otomatis boleh di-edit, hijau)",
        value=f"SICTb-{nomor_auto}/KPN.0603/2025"
    )

    # 3. Nama–NIP (abu-abu, lookup dari Google Sheets)
    pegawai_list = get_pegawai_list()
    nama_pegawai = st.selectbox("Nama Pegawai (abu-abu)", options=pegawai_list)
    nip_pegawai = get_nip_by_name(nama_pegawai)
    st.text_input("NIP Pegawai (linked, abu-abu)", value=nip_pegawai, disabled=True)

    # 4. Input manual (pink)
    jabatan = st.text_input("Jabatan (pink)")
    masa_kerja = st.text_input("Masa Kerja (pink)")
    unit_kerja = st.text_area("Unit Kerja (pink)")
    jenis_cuti = st.selectbox(
        "Jenis Cuti (pink)",
        ["Cuti tahunan", "Cuti Besar", "Cuti Sakit", "Cuti Melahirkan", "Cuti Karena Alasan Penting", "Cuti di Luar Tanggungan Negara"],
        index=0
    )
    alasan_cuti = st.text_area("Alasan Cuti (pink)")
    lama_cuti = st.text_input("Lamanya Cuti (contoh: 6 hari, pink)")

    tgl_mulai = st.date_input("Tanggal Mulai Cuti (pink)")
    tgl_selesai = st.date_input("Tanggal Selesai Cuti (pink)")

    tgl_mulai_str = format_tanggal_indonesia(tgl_mulai.strftime("%d-%B-%Y"))
    tgl_selesai_str = format_tanggal_indonesia(tgl_selesai.strftime("%d-%B-%Y"))

    sisa_cuti_2023 = st.text_input("Sisa Cuti 2023 (pink)", value="")
    sisa_cuti_2024 = st.text_input("Sisa Cuti 2024 (pink)", value="")
    sisa_cuti_2025 = st.text_input("Sisa Cuti 2025 (pink)", value="")
    sisa_cuti_tambahan_2025 = st.text_input("Sisa Cuti Tambahan 2025 (pink)", value="")

    alamat_cuti = st.text_area("Alamat Selama Cuti (pink)")
    telp = st.text_input("Telp (pink)")

    nama_atasan_langsung = st.text_input("Nama Atasan Langsung (pink)", value="Setiyono")
    nip_atasan_langsung = st.text_input("NIP Atasan Langsung (pink)", value="197311161996021001")
    nama_pejabat_berwenang = st.text_input("Nama Pejabat Berwenang (pink)", value="Setiyono")
    nip_pejabat_berwenang = st.text_input("NIP Pejabat Berwenang (pink)", value="197311161996021001")

    if st.button("Generate PDF"):
        output_file = f"Form-Cuti-{nama_pegawai.replace(' ', '-')}-{nomor_auto}.pdf"
        generate_cuti_pdf(
            output_file,
            tanggal_form=tgl_form_fmt,
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
        with open(output_file, "rb") as f:
            st.download_button("Download PDF", f, file_name=output_file, mime="application/pdf")

if __name__ == "__main__":
    main()
