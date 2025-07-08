# =========================
# File: streamlit_app.py
# =========================

import streamlit as st
import pandas as pd
from datetime import date
from spj_generator import buat_spj

st.set_page_config(page_title="SPJ Desa", layout="wide")
st.title("📄 Sistem Penerbitan Surat Pertanggungjawaban (SPJ) Desa")

with st.form("form_kegiatan"):
    st.subheader("📌 Data Kegiatan")
    lembaga = st.selectbox("Lembaga Pelaksana", [
        "TPK", "PPKAD", "PPS", "Karang Taruna", "LPMD", "PKK", "Posyandu", "TSBD", "Karang Wreda", "Pokmas"
    ])
    nama_kegiatan = st.text_input("Nama Kegiatan")
    tgl_pelaksanaan = st.date_input("Tanggal Pelaksanaan", date.today())
    lokasi = st.text_input("Lokasi")
    kode_register = st.text_input("Kode Register SPJ", value="470/SPJ-DESA")

    st.markdown("### 💰 Data RAB dan Realisasi")
    rab_awal = st.number_input("Total RAB (Rp)", 0)
    realisasi = st.number_input("Realisasi (Rp)", 0)
    selisih = rab_awal - realisasi
    status = "✅ Sesuai" if selisih == 0 else ("🔺 Lebih" if selisih < 0 else "🔻 Kurang")
    st.markdown(f"**Selisih:** Rp {selisih:,.0f} ({status})")

    sumber_dana = st.selectbox("Sumber Dana", ["DD", "ADD", "BKK", "Swadaya", "Lainnya"])

    st.subheader("📎 Upload & Identitas Pejabat")
    bukti_upload = st.file_uploader("Upload Bukti Transaksi (PDF/JPG)", accept_multiple_files=True)

    st.subheader("✍️ Data Pejabat Pengesah")
    nama_kades = st.text_input("Nama Kepala Desa", "Sutrisno, S.E.")
    nama_ketua_bpd = st.text_input("Nama Ketua BPD", "Misdi")
    nama_ketua_lembaga = st.text_input(f"Nama Ketua {lembaga}", "")
    nama_bendahara = st.text_input("Nama Bendahara", "")

    submit = st.form_submit_button("✅ Buat SPJ")

if submit:
    with st.spinner("📄 Membuat dokumen SPJ..."):
        try:
            file_path = buat_spj(
                lembaga, nama_kegiatan, tgl_pelaksanaan, lokasi,
                rab_awal, realisasi, sumber_dana, bukti_upload,
                nama_kades, nama_ketua_bpd, nama_ketua_lembaga, nama_bendahara,
                kode_register
            )
            st.success("✅ SPJ berhasil dibuat!")
            with open(file_path, "rb") as f:
                st.download_button("📥 Unduh SPJ (DOCX)", data=f.read(), file_name="SPJ_Kegiatan.docx")
        except Exception as e:
            st.error(f"❌ Terjadi kesalahan saat membuat dokumen SPJ:\n\n{e}")
