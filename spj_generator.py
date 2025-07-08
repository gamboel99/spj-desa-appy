import os
import qrcode
from docx import Document
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
from datetime import datetime

def buat_spj(lembaga, nama_kegiatan, tgl, lokasi, sumber_dana, rab_df, bukti_upload,
             nama_kades, nama_ketua_bpd, nama_ketua_lembaga, nama_bendahara, kode_register):

    doc = Document()

    # Logo (opsional)
    logo_path = os.path.join(os.path.dirname(__file__), "logo_desa.png")
    if os.path.exists(logo_path):
        doc.add_picture(logo_path, width=Inches(1.2))

    # Kop
    doc.add_paragraph("PEMERINTAH DESA KELING", style='Heading 1').alignment = 1
    doc.add_paragraph("KECAMATAN KEPUNG, KABUPATEN KEDIRI", style='Heading 2').alignment = 1
    doc.add_paragraph("Jl. Raya Keling, Bukaan, Keling, Kediri, Jawa Timur 64293", style='Normal').alignment = 1
    doc.add_paragraph("______________________________________________________________")

    doc.add_paragraph(f"\nNomor: {kode_register}/{tgl.month:02d}/{tgl.year}")

    doc.add_paragraph("\nSURAT PERTANGGUNGJAWABAN KEGIATAN", style='Title').alignment = 1

    doc.add_paragraph(f"\nNama Kegiatan        : {nama_kegiatan}")
    doc.add_paragraph(f"Tanggal Pelaksanaan  : {tgl.strftime('%d-%m-%Y')}")
    doc.add_paragraph(f"Lembaga Pelaksana    : {lembaga}")
    doc.add_paragraph(f"Lokasi               : {lokasi}")
    doc.add_paragraph(f"Sumber Dana          : {sumber_dana}\n")

    doc.add_paragraph("Tabel RAB dan Realisasi:")

    table = doc.add_table(rows=1, cols=9)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers = [
        "No", "Uraian", "Volume", "Satuan", "Harga Satuan",
        "Realisasi", "Potongan Pajak", "Diskon/Bunga", "Realisasi Bersih"
    ]
    for i, h in enumerate(headers):
        table.cell(0, i).text = h

    total_anggaran = 0
    total_realisasi_bersih = 0

    for i, row in rab_df.iterrows():
        try:
            volume = float(row["Volume"])
            harga = float(row["Harga Satuan"])
            realisasi = float(row["Realisasi"])
            pajak = float(row["Potongan Pajak"])
            diskon = float(row["Diskon/Bunga"])
        except:
            volume = harga = realisasi = pajak = diskon = 0

        jumlah = volume * harga
        realisasi_bersih = realisasi - pajak - diskon
        total_anggaran += jumlah
        total_realisasi_bersih += realisasi_bersih

        data = [
            str(i + 1),
            str(row["Uraian"]),
            str(volume),
            str(row["Satuan"]),
            f"Rp {harga:,.0f}",
            f"Rp {realisasi:,.0f}",
            f"Rp {pajak:,.0f}",
            f"Rp {diskon:,.0f}",
            f"Rp {realisasi_bersih:,.0f}"
        ]

        row_cells = table.add_row().cells
        for j, val in enumerate(data):
            row_cells[j].text = val

    doc.add_paragraph(f"\nTotal Anggaran       : Rp {total_anggaran:,.0f}")
    doc.add_paragraph(f"Total Realisasi Bersih: Rp {total_realisasi_bersih:,.0f}")
    doc.add_paragraph(f"Selisih              : Rp {total_anggaran - total_realisasi_bersih:,.0f}")

    doc.add_paragraph(f"\nDesa Keling, {tgl.strftime('%d-%m-%Y')}\n")

    # Tanda tangan (tanpa logo atau gambar)
    doc.add_paragraph(f"Mengetahui,\nKepala Desa\n\n{nama_kades}", style="Normal")
    doc.add_paragraph(f"\nKetua {lembaga}\n\n{nama_ketua_lembaga}", style="Normal")
    doc.add_paragraph(f"\nBendahara\n\n{nama_bendahara}", style="Normal")
    doc.add_paragraph(f"\nMengesahkan,\nKetua BPD\n\n{nama_ketua_bpd}")

    # QR Code (opsional, jika sudah tersedia)
    qr_path = os.path.join(os.path.dirname(__file__), "qr_temp.png")
    if os.path.exists(qr_path):
        doc.add_picture(qr_path, width=Inches(1.2))
        os.remove(qr_path)

    output_path = os.path.join(os.path.dirname(__file__), "SPJ_Kegiatan.docx")
    doc.save(output_path)

    return output_path
