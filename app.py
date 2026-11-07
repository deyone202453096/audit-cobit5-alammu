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
    # DOWNLOAD EXCEL DENGAN GRAFIK
    # ==============================

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        pd.DataFrame(
            [identitas]
        ).to_excel(
            writer,
            sheet_name="Identitas",
            index=False
        )

        df.to_excel(
            writer,
            sheet_name="Jawaban Audit",
            index=False
        )

        df_rekap.to_excel(
            writer,
            sheet_name="Rekapitulasi",
            index=False
        )

    # Tambahkan sheet grafik Excel
    output.seek(0)
    workbook = load_workbook(output)

    ws = workbook.create_sheet("Grafik")

    ws["A1"] = "REKAPITULASI HASIL AUDIT COBIT 5"

    headers = list(df_rekap.columns)

    for col_num, header in enumerate(headers, start=1):
        ws.cell(row=3, column=col_num, value=header)

    for row_num, row in enumerate(
        df_rekap.itertuples(index=False, name=None),
        start=4
    ):
        for col_num, value in enumerate(row, start=1):
            ws.cell(row=row_num, column=col_num, value=int(value) if hasattr(value, "item") else value)

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Grafik Jawaban Audit per Domain"
    chart.y_axis.title = "Jumlah Jawaban"
    chart.x_axis.title = "Domain COBIT 5"

    chart_data = Reference(
        ws,
        min_col=3,
        max_col=5,
        min_row=3,
        max_row=8
    )

    categories = Reference(
        ws,
        min_col=1,
        min_row=4,
        max_row=8
    )

    chart.add_data(
        chart_data,
        titles_from_data=True
    )

    chart.set_categories(categories)
    chart.height = 9
    chart.width = 18

    ws.add_chart(chart, "G3")

    # Lebar kolom
    for column in ws.columns:
        max_length = max(
            len(str(cell.value or ""))
            for cell in column
        )
        ws.column_dimensions[
            column[0].column_letter
        ].width = min(max_length + 3, 35)

    final_output = io.BytesIO()
    workbook.save(final_output)
    final_output.seek(0)

    st.divider()

    st.header("Unduh Laporan Audit")

    st.download_button(
        label="📥 Unduh Hasil Audit Excel",
        data=final_output.getvalue(),
        file_name="Hasil_Audit_COBIT5.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
