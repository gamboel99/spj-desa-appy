import streamlit as st
import pandas as pd
from datetime import date
from spj_generator import buat_spj

st.set_page_config(page_title="Sistem SPJ Desa Keling", layout="wide")

# ✅ Header aplikasi
st.title("📝 Sistem SPJ Desa Keling, Kec. Kepung")

# Form dan konten lain di sini seperti biasa...
# ...

# ✅ Footer aplikasi
st.markdown("---")
st.markdown(
    "<div style='text-align: center; font-size: 13px; color: gray;'>"
    "Developed by <strong>CV. Mitra Utama Consultindo</strong>"
    "</div>",
    unsafe_allow_html=True
)
