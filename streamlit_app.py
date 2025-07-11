import streamlit as st
import pandas as pd
from datetime import date
from spj_generator import buat_spj, generate_register_code

st.set_page_config(page_title="Sistem SPJ Desa Keling", layout="wide")
st.title("📝 Sistem SPJ Desa Keling, Kec. Kepung")

st.markdown("---")
st.subheader("📌 Data Kegiatan")

with st.form("form_kegiatan"):
    lembaga = st.selectbox("Lembaga Pelaksana", ["TPK", "PPKAD", "PPS", "Karang Taruna", "LPMD"])
    nama_kegiatan = st.text_input("Nama Kegiatan")
    tgl_pelaksanaan = st.date_input("Tanggal Pelaksanaan", date.today())
    lokasi = st.text_input("Lokasi")
    sumber_dana = st.selectbox("Sumber Dana", ["DD", "ADD", "BKK", "Swadaya", "Lainnya"])

    kode_register = generate_register_code()
    st.info(f"📑 Kode Register Otomatis: {kode_register}")

    st.markdown("### 🧾 RAB dan Realisasi")
    rab_df = st.data_editor(pd.DataFrame({
        "Uraian": [""],
        "Volume": [1],
        "Satuan": [""],
        "Harga Satuan": [0],
        "Realisasi": [0],
        "Pajak (%)": [10],
        "Diskon (Rp)": [0],
    }), num_rows="dynamic", use_container_width=True)

    st.markdown("### 🖋️ Data Penandatangan")
    nama_kades = st.text_input("Nama Kepala Desa", "Sutrisno, S.E.")
    nama_ketua_bpd = st.text_input("Nama Ketua BPD", "Misdi")
    nama_ketua_lembaga = st.text_input(f"Nama Ketua {lembaga}", "Budi Santoso")
    nama_bendahara = st.text_input("Nama Bendahara", "Rina Puspitasari")

    bukti_upload = st.file_uploader("📎 Upload Bukti Transaksi (PDF/JPG)", accept_multiple_files=True)

    submit = st.form_submit_button("✅ Buat Surat Pertanggungjawaban")

if submit:
    with st.spinner("Sedang memproses dokumen SPJ..."):
        try:
            file_path = buat_spj(
                lembaga, nama_kegiatan, tgl_pelaksanaan, lokasi, sumber_dana,
                rab_df, bukti_upload, nama_kades, nama_ketua_bpd, nama_ketua_lembaga,
                nama_bendahara, kode_register
            )
            st.success("✅ SPJ berhasil dibuat!")
            with open(file_path, "rb") as f:
                st.download_button("📥 Unduh SPJ (DOCX)", f.read(), file_name="SPJ_Kegiatan.docx")
        except Exception as e:
            st.error(f"❌ Terjadi kesalahan: {e}")

st.markdown("---")
st.markdown(
    "<div style='text-align: center; font-size: 13px; color: gray;'>"
    "Developed by <strong>CV. Mitra Utama Consultindo (MUC)</strong>"
    "</div>",
    unsafe_allow_html=True
)
