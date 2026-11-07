import streamlit as st
import pandas as pd
import io
from datetime import date
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference

st.set_page_config(
    page_title="Audit COBIT 5",
    page_icon="📋",
    layout="wide"
)

# ==============================
# HEADER APLIKASI
# ==============================

st.title("AUDIT SISTEM INFORMASI PENJUALAN")
st.subheader("CV Seleksi Alam Muria")

st.markdown("""
Aplikasi Audit Sistem Informasi Penjualan berbasis COBIT 5
merupakan platform digital untuk membantu proses evaluasi
pengelolaan sistem informasi penjualan pada CV Seleksi Alam
Muria (Alammu).

Aplikasi ini menyediakan 30 pertanyaan audit berdasarkan lima
domain COBIT 5, yaitu EDM, APO, BAI, DSS, dan MEA. Pengguna
dapat mengisi identitas, memberikan jawaban audit, melihat
rekapitulasi dalam bentuk tabel dan grafik, serta mengunduh
hasil audit dalam format Excel.
""")

st.divider()

# ==============================
# DATA PERTANYAAN AUDIT
# ==============================

data = [
    ("EDM", "EDM01", "Apakah pihak yang bertanggung jawab melakukan evaluasi terhadap pengelolaan sistem informasi penjualan secara berkala?"),
    ("EDM", "EDM01", "Apakah peran, tanggung jawab, dan wewenang dalam pengelolaan data penjualan sudah ditentukan dengan jelas?"),
    ("EDM", "EDM02", "Apakah manfaat pengelolaan sistem informasi penjualan telah dievaluasi sesuai kebutuhan bisnis?"),
    ("EDM", "EDM03", "Apakah risiko kesalahan, kehilangan, atau perubahan data penjualan tanpa izin telah dipertimbangkan?"),
    ("EDM", "EDM04", "Apakah kebutuhan sumber daya untuk mendukung pengelolaan informasi penjualan telah dipertimbangkan?"),
    ("EDM", "EDM05", "Apakah laporan penjualan disampaikan kepada pihak terkait sesuai kebutuhan?"),

    ("APO", "APO01", "Apakah terdapat prosedur atau pedoman dalam pengelolaan sistem informasi penjualan?"),
    ("APO", "APO01", "Apakah prosedur pengelolaan informasi penjualan telah disampaikan kepada pihak yang terlibat?"),
    ("APO", "APO02", "Apakah kebutuhan sistem informasi penjualan telah dipertimbangkan dalam perencanaan bisnis?"),
    ("APO", "APO07", "Apakah pihak yang mengelola informasi penjualan memiliki keterampilan yang sesuai?"),
    ("APO", "APO12", "Apakah risiko dalam pencatatan, penyimpanan, dan penggunaan data penjualan telah diidentifikasi?"),
    ("APO", "APO13", "Apakah terdapat perlindungan untuk mencegah akses tidak sah terhadap informasi penjualan?"),

    ("BAI", "BAI02", "Apakah kebutuhan pengguna dalam pencatatan dan pelaporan penjualan telah diidentifikasi?"),
    ("BAI", "BAI04", "Apakah kapasitas dan ketersediaan sarana pencatatan penjualan telah dievaluasi?"),
    ("BAI", "BAI06", "Apakah perubahan pada sistem atau metode pencatatan penjualan dikendalikan dan disetujui?"),
    ("BAI", "BAI07", "Apabila terdapat perubahan sistem, apakah perubahan tersebut diuji sebelum digunakan?"),
    ("BAI", "BAI08", "Apakah tersedia dokumentasi atau panduan kerja terkait pengelolaan informasi penjualan?"),
    ("BAI", "BAI09", "Apakah aset yang digunakan untuk mendukung informasi penjualan dicatat dan dipelihara?"),

    ("DSS", "DSS01", "Apakah transaksi penjualan dicatat secara rutin sesuai prosedur?"),
    ("DSS", "DSS02", "Apakah terdapat prosedur untuk menangani masalah atau kesalahan dalam pencatatan transaksi penjualan?"),
    ("DSS", "DSS03", "Apakah kesalahan pencatatan yang terjadi berulang dianalisis dan ditindaklanjuti?"),
    ("DSS", "DSS04", "Apakah terdapat langkah untuk menjaga dan memulihkan data penjualan apabila terjadi gangguan?"),
    ("DSS", "DSS05", "Apakah akses terhadap informasi penjualan dibatasi berdasarkan peran pengguna?"),
    ("DSS", "DSS06", "Apakah terdapat pengendalian untuk memastikan transaksi penjualan tercatat dengan benar dan dapat ditelusuri?"),

    ("MEA", "MEA01", "Apakah kinerja pengelolaan informasi penjualan dipantau secara berkala?"),
    ("MEA", "MEA01", "Apakah hasil pengelolaan informasi penjualan dievaluasi berdasarkan kebutuhan dan target bisnis?"),
    ("MEA", "MEA02", "Apakah pengendalian internal terhadap data penjualan diperiksa secara berkala?"),
    ("MEA", "MEA02", "Apakah data transaksi penjualan diperiksa kesesuaiannya dengan bukti transaksi?"),
    ("MEA", "MEA02", "Apakah kelemahan pengendalian yang ditemukan ditindaklanjuti?"),
    ("MEA", "MEA03", "Apakah pengelolaan informasi penjualan ditinjau berdasarkan ketentuan eksternal yang relevan?")
]

# ==============================
# FORM IDENTITAS DAN AUDIT
# ==============================

st.header("Identitas Pengisi Form")

with st.form("form_audit"):

    nama = st.text_input("Nama Lengkap")
    jabatan = st.text_input("Jabatan/Posisi")
    divisi = st.text_input("Bagian/Divisi")

    tanggal = st.date_input(
        "Tanggal Pengisian",
        value=date.today()
    )

    st.divider()

    st.header("Form Pertanyaan Audit")
    st.write(
        "Pilih jawaban berdasarkan hasil pemeriksaan "
        "dan bukti pendukung."
    )

    jawaban = []

    for domain in ["EDM", "APO", "BAI", "DSS", "MEA"]:

        st.subheader(f"Domain {domain}")

        for nomor, (d, kode, pertanyaan) in enumerate(data, start=1):

            if d == domain:

                st.write(
                    f"**{nomor}. [{kode}] {pertanyaan}**"
                )

                pilihan = st.radio(
                    "Pilih jawaban:",
                    ["Ya", "Tidak", "N/A"],
                    key=f"jawaban_{nomor}",
                    horizontal=True
                )

                jawaban.append({
                    "No": nomor,
                    "Domain": domain,
                    "Kode Proses": kode,
                    "Pertanyaan": pertanyaan,
                    "Jawaban": pilihan
                })

    submit = st.form_submit_button(
        "Simpan dan Tampilkan Hasil Audit",
        type="primary"
    )

# ==============================
# SIMPAN HASIL KE SESSION
# ==============================

if submit:

    if not nama.strip():
        st.error("Nama pengisi form wajib diisi.")

    else:
        st.session_state["identitas_audit"] = {
            "Nama": nama,
            "Jabatan/Posisi": jabatan,
            "Bagian/Divisi": divisi,
            "Tanggal Pengisian": tanggal.strftime("%d-%m-%Y")
        }

        st.session_state["jawaban_audit"] = pd.DataFrame(jawaban)

        st.success("Jawaban audit berhasil diproses!")

# ==============================
# TAMPILKAN HASIL AUDIT
# ==============================

if "jawaban_audit" in st.session_state:

    df = st.session_state["jawaban_audit"]
    identitas = st.session_state["identitas_audit"]

    st.divider()

    st.header("Identitas Pengisi")

    col1, col2, col3 = st.columns(3)

    col1.write("**Nama**")
    col1.write(identitas["Nama"])

    col2.write("**Jabatan/Posisi**")
    col2.write(identitas["Jabatan/Posisi"] or "-")

    col3.write("**Bagian/Divisi**")
    col3.write(identitas["Bagian/Divisi"] or "-")

    st.write(
        f"**Tanggal Pengisian:** "
        f"{identitas['Tanggal Pengisian']}"
    )

    # ==============================
    # REKAPITULASI
    # ==============================

    rekap = []

    for domain in ["EDM", "APO", "BAI", "DSS", "MEA"]:

        bagian = df[df["Domain"] == domain]

        rekap.append({
            "Domain": domain,
            "Jumlah Pertanyaan": len(bagian),
            "Ya": (bagian["Jawaban"] == "Ya").sum(),
            "Tidak": (bagian["Jawaban"] == "Tidak").sum(),
            "N/A": (bagian["Jawaban"] == "N/A").sum()
        })

    df_rekap = pd.DataFrame(rekap)

    ya = (df["Jawaban"] == "Ya").sum()
    tidak = (df["Jawaban"] == "Tidak").sum()
    na = (df["Jawaban"] == "N/A").sum()

    total_dinilai = ya + tidak

    persentase = (
        ya / total_dinilai * 100
        if total_dinilai > 0 else 0
    )

    # ==============================
    # DASHBOARD
    # ==============================

    st.header("Dashboard Hasil Audit")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Total Pertanyaan", len(df))
    col2.metric("Jawaban Ya", int(ya))
    col3.metric("Jawaban Tidak", int(tidak))
    col4.metric("Jawaban N/A", int(na))

    st.metric(
        "Indikator Jawaban Ya",
        f"{persentase:.2f}%"
    )

    st.caption(
        "Persentase dihitung dari jumlah jawaban Ya "
        "dibagi jumlah jawaban Ya dan Tidak. "
        "Indikator ini bukan capability level resmi COBIT 5."
    )

    # ==============================
    # GRAFIK
    # ==============================

    st.header("Grafik Rekapitulasi per Domain")

    grafik = df_rekap.set_index("Domain")[
        ["Ya", "Tidak", "N/A"]
    ]

    st.bar_chart(grafik)

    # ==============================
    # TABEL REKAPITULASI
    # ==============================

    st.header("Tabel Rekapitulasi Hasil Audit")

    st.dataframe(
        df_rekap,
        use_container_width=True,
        hide_index=True
    )

    # ==============================
    # TABEL JAWABAN
    # ==============================

    st.header("Detail Jawaban Audit")

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )

    
# ==============================
# DOWNLOAD LAPORAN AUDIT LENGKAP
# ==============================

import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference

output = io.BytesIO()

wb = Workbook()
ws = wb.active
ws.title = "Laporan Audit"

# JUDUL
ws.merge_cells("A1:E1")
ws["A1"] = "LAPORAN AUDIT SISTEM INFORMASI PENJUALAN COBIT 5"
ws["A1"].font = Font(bold=True, size=14)

# IDENTITAS
ws["A3"] = "IDENTITAS PENGISI"
ws["A3"].font = Font(bold=True, size=12)

for row, (key, value) in enumerate(identitas.items(), start=4):
    ws.cell(row=row, column=1, value=key)
    ws.cell(row=row, column=2, value=value)

# TABEL JAWABAN
start = 10
ws.cell(row=start, column=1, value="HASIL JAWABAN AUDIT")
ws.cell(row=start, column=1).font = Font(bold=True, size=12)

for col, header in enumerate(df.columns, start=1):
    cell = ws.cell(row=start + 1, column=col, value=header)
    cell.font = Font(bold=True)
    cell.fill = PatternFill("solid", fgColor="D9EAF7")

for r, row in enumerate(df.itertuples(index=False, name=None), start=start + 2):
    for c, value in enumerate(row, start=1):
        ws.cell(row=r, column=c, value=value)

# REKAPITULASI
rekap_start = start + len(df) + 4

ws.cell(row=rekap_start, column=1, value="REKAPITULASI HASIL AUDIT")
ws.cell(row=rekap_start, column=1).font = Font(bold=True, size=12)

for col, header in enumerate(df_rekap.columns, start=1):
    cell = ws.cell(row=rekap_start + 1, column=col, value=header)
    cell.font = Font(bold=True)
    cell.fill = PatternFill("solid", fgColor="D9EAF7")

for r, row in enumerate(
    df_rekap.itertuples(index=False, name=None),
    start=rekap_start + 2
):
    for c, value in enumerate(row, start=1):
        ws.cell(row=r, column=c, value=value)

# GRAFIK DI SHEET KEDUA
grafik_ws = wb.create_sheet("Grafik")

grafik_ws.append(["Domain", "Ya", "Tidak", "N/A"])

for _, row in df_rekap.iterrows():
    grafik_ws.append([
        row["Domain"],
        int(row["Ya"]),
        int(row["Tidak"]),
        int(row["N/A"])
    ])

chart = BarChart()
chart.type = "col"
chart.title = "Grafik Hasil Audit per Domain"
chart.y_axis.title = "Jumlah Jawaban"
chart.x_axis.title = "Domain COBIT 5"

chart_data = Reference(
    grafik_ws,
    min_col=2,
    max_col=4,
    min_row=1,
    max_row=6
)

categories = Reference(
    grafik_ws,
    min_col=1,
    min_row=2,
    max_row=6
)

chart.add_data(chart_data, titles_from_data=True)
chart.set_categories(categories)
grafik_ws.add_chart(chart, "F2")

# RAPATKAN KOLOM
ws.column_dimensions["A"].width = 18
ws.column_dimensions["B"].width = 22
ws.column_dimensions["C"].width = 18
ws.column_dimensions["D"].width = 75
ws.column_dimensions["E"].width = 18

for row in ws.iter_rows():
    for cell in row:
        cell.alignment = Alignment(
            vertical="center",
            wrap_text=True
        )

# SIMPAN FILE
wb.save(output)
output.seek(0)

st.download_button(
    label="📥 Unduh Laporan Audit Excel",
    data=output.getvalue(),
    file_name="Laporan_Audit_COBIT5.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)  
