import streamlit as st
import pandas as pd
import io

from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

# ==================================================
# KONFIGURASI APLIKASI
# ==================================================

st.set_page_config(
    page_title="Audit COBIT 5 Alammu",
    page_icon="📋",
    layout="wide"
)

# ==================================================
# HEADER DAN DESKRIPSI
# ==================================================

st.title("AUDIT SISTEM INFORMASI PENJUALAN")
st.subheader("CV Seleksi Alam Muria (Alammu)")

st.markdown("""
Aplikasi Audit Sistem Informasi Penjualan berbasis COBIT 5
merupakan platform digital yang digunakan untuk membantu
proses evaluasi pengelolaan sistem informasi penjualan
pada CV Seleksi Alam Muria (Alammu).

Aplikasi ini menyediakan 30 pertanyaan audit berdasarkan
lima domain COBIT 5, yaitu EDM, APO, BAI, DSS, dan MEA.
Pengguna dapat mengisi identitas, memberikan jawaban,
melihat hasil audit melalui tabel dan grafik, serta
mengunduh laporan dalam format Excel.
""")

st.info(
    "Petunjuk: Isi identitas pengisi dan jawab seluruh "
    "30 pertanyaan dengan memilih Ya, Tidak, atau N/A."
)

st.divider()

# ==================================================
# DATA PERTANYAAN AUDIT
# ==================================================

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

# ==================================================
# FORM IDENTITAS DAN PERTANYAAN
# ==================================================

st.header("1. Identitas Pengisi Form")

with st.form("form_audit"):

    nama = st.text_input("Nama Lengkap *")

    jabatan = st.text_input("Jabatan/Posisi")

    divisi = st.text_input("Bagian/Divisi")

    tanggal = st.date_input(
        "Tanggal Pengisian",
        value=date.today(),
        format="DD/MM/YYYY"
    )

    st.divider()

    st.header("2. Form Pertanyaan Audit")

    st.write(
        "Pilih satu jawaban untuk setiap pertanyaan. "
        "Seluruh pertanyaan wajib dijawab."
    )

    jawaban = []

    for domain in ["EDM", "APO", "BAI", "DSS", "MEA"]:

        st.subheader(f"Domain {domain}")

        for nomor, (d, kode, pertanyaan) in enumerate(data, start=1):

            if d == domain:

                st.markdown(
                    f"**{nomor}. [{kode}] {pertanyaan}**"
                )

                pilihan = st.radio(
                    "Pilih jawaban:",
                    options=["Ya", "Tidak", "N/A"],
                    index=None,
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

    st.divider()

    submit = st.form_submit_button(
        "Simpan dan Tampilkan Hasil Audit",
        type="primary",
        use_container_width=True
    )

# ==================================================
# VALIDASI DAN PENYIMPANAN HASIL
# ==================================================

if submit:

    # Hapus hasil sebelumnya jika ada pengiriman baru
    st.session_state.pop("hasil_audit", None)

    if not nama.strip():

        st.error("Nama lengkap wajib diisi.")

    elif len(jawaban) != 30:

        st.error("Data pertanyaan tidak lengkap.")

    elif any(item["Jawaban"] is None for item in jawaban):

        jumlah_kosong = sum(
            item["Jawaban"] is None for item in jawaban
        )

        st.error(
            f"Masih ada {jumlah_kosong} pertanyaan "
            "yang belum dijawab. Silakan lengkapi semuanya."
        )

    else:

        st.session_state["hasil_audit"] = {
            "identitas": {
                "Nama": nama.strip(),
                "Jabatan/Posisi": jabatan.strip(),
                "Bagian/Divisi": divisi.strip(),
                "Tanggal Pengisian": tanggal.strftime("%d-%m-%Y")
            },
            "jawaban": pd.DataFrame(jawaban)
        }

        st.success("Seluruh jawaban berhasil disimpan!")

# ==================================================
# TAMPILKAN HASIL AUDIT
# ==================================================

if "hasil_audit" in st.session_state:

    hasil = st.session_state["hasil_audit"]

    identitas = hasil["identitas"]
    df = hasil["jawaban"]

    st.divider()

    st.header("3. Hasil Audit")

    # IDENTITAS
    st.subheader("Identitas Pengisi")

    col1, col2, col3 = st.columns(3)

    col1.metric("Nama", identitas["Nama"])
    col2.metric(
        "Jabatan",
        identitas["Jabatan/Posisi"] or "-"
    )
    col3.metric(
        "Bagian/Divisi",
        identitas["Bagian/Divisi"] or "-"
    )

    st.write(
        f"**Tanggal Pengisian:** "
        f"{identitas['Tanggal Pengisian']}"
    )

    # HITUNG REKAPITULASI
    rekap = []

    for domain in ["EDM", "APO", "BAI", "DSS", "MEA"]:

        bagian = df[df["Domain"] == domain]

        rekap.append({
            "Domain": domain,
            "Jumlah Pertanyaan": len(bagian),
            "Ya": int((bagian["Jawaban"] == "Ya").sum()),
            "Tidak": int((bagian["Jawaban"] == "Tidak").sum()),
            "N/A": int((bagian["Jawaban"] == "N/A").sum())
        })

    df_rekap = pd.DataFrame(rekap)

    ya = int((df["Jawaban"] == "Ya").sum())
    tidak = int((df["Jawaban"] == "Tidak").sum())
    na = int((df["Jawaban"] == "N/A").sum())

    total_dinilai = ya + tidak

    persentase = (
        ya / total_dinilai * 100
        if total_dinilai > 0 else 0
    )

    # DASHBOARD
    st.subheader("Ringkasan Jawaban")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Total Pertanyaan", len(df))
    col2.metric("Jawaban Ya", ya)
    col3.metric("Jawaban Tidak", tidak)
    col4.metric("Jawaban N/A", na)

    st.metric(
        "Persentase Jawaban Ya",
        f"{persentase:.2f}%"
    )

    st.caption(
        "Persentase dihitung dari jumlah jawaban Ya "
        "dibagi jumlah jawaban Ya dan Tidak. "
        "Indikator ini bukan capability level resmi COBIT 5."
    )

    # GRAFIK
    st.subheader("Grafik Hasil Audit per Domain")

    grafik = df_rekap.set_index("Domain")[
        ["Ya", "Tidak", "N/A"]
    ]

    st.bar_chart(
        grafik,
        y=["Ya", "Tidak", "N/A"],
        use_container_width=True
    )

    # REKAPITULASI
    st.subheader("Rekapitulasi Hasil Audit")

    st.dataframe(
        df_rekap,
        use_container_width=True,
        hide_index=True
    )

    # DETAIL JAWABAN
    st.subheader("Detail Jawaban 30 Pertanyaan")

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )

    # ==================================================
    # MEMBUAT FILE EXCEL LENGKAP
    # ==================================================

    def buat_excel(identitas, df, df_rekap, ya, tidak, na, persentase):

        wb = Workbook()

        ws = wb.active
        ws.title = "Laporan Audit"

        # JUDUL
        ws.merge_cells("A1:E1")
        ws["A1"] = "LAPORAN AUDIT SISTEM INFORMASI PENJUALAN"
        ws["A1"].font = Font(
            bold=True,
            size=14,
            color="FFFFFF"
        )
        ws["A1"].fill = PatternFill(
            "solid",
            fgColor="1F4E78"
        )
        ws["A1"].alignment = Alignment(
            horizontal="center"
        )

        ws.merge_cells("A2:E2")
        ws["A2"] = "CV Seleksi Alam Muria (Alammu)"
        ws["A2"].alignment = Alignment(
            horizontal="center"
        )

        # IDENTITAS
        ws["A4"] = "IDENTITAS PENGISI"
        ws["A4"].font = Font(bold=True, size=12)

        for row, (key, value) in enumerate(
            identitas.items(), start=5
        ):
            ws.cell(row=row, column=1, value=key)
            ws.cell(row=row, column=2, value=value)

        # TABEL JAWABAN
        ws["A10"] = "DETAIL JAWABAN AUDIT"
        ws["A10"].font = Font(bold=True, size=12)

        header_row = 11

        for col, header in enumerate(df.columns, start=1):

            cell = ws.cell(
                row=header_row,
                column=col,
                value=header
            )

            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )

            cell.fill = PatternFill(
                "solid",
                fgColor="4472C4"
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

        for row_num, row_data in enumerate(
            df.itertuples(index=False, name=None),
            start=header_row + 1
        ):

            for col_num, value in enumerate(
                row_data, start=1
            ):

                cell = ws.cell(
                    row=row_num,
                    column=col_num,
                    value=value
                )

                cell.alignment = Alignment(
                    vertical="top",
                    wrap_text=True
                )

        # REKAPITULASI
        recap_start = header_row + len(df) + 3

        ws.cell(
            row=recap_start,
            column=1,
            value="REKAPITULASI HASIL AUDIT"
        ).font = Font(bold=True, size=12)

        recap_header = recap_start + 1

        for col, header in enumerate(
            df_rekap.columns, start=1
        ):

            cell = ws.cell(
                row=recap_header,
                column=col,
                value=header
            )

            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )

            cell.fill = PatternFill(
                "solid",
                fgColor="4472C4"
            )

        for row_num, row_data in enumerate(
            df_rekap.itertuples(index=False, name=None),
            start=recap_header + 1
        ):

            for col_num, value in enumerate(
                row_data, start=1
            ):

                ws.cell(
                    row=row_num,
                    column=col_num,
                    value=value
                )

        # TOTAL
        total_row = recap_header + len(df_rekap) + 1

        ws.cell(total_row, 1, "TOTAL")
        ws.cell(total_row, 2, len(df))
        ws.cell(total_row, 3, ya)
        ws.cell(total_row, 4, tidak)
        ws.cell(total_row, 5, na)

        for cell in ws[total_row]:
            cell.font = Font(bold=True)

        # PERSENTASE
        percentage_row = total_row + 2

        ws.cell(
            percentage_row,
            1,
            "Persentase Jawaban Ya"
        )

        ws.cell(
            percentage_row,
            2,
            f"{persentase:.2f}%"
        )

        # GRAFIK PADA SHEET KEDUA
        grafik_ws = wb.create_sheet("Grafik")

        grafik_ws.append([
            "Domain",
            "Ya",
            "Tidak",
            "N/A"
        ])

        for _, row in df_rekap.iterrows():

            grafik_ws.append([
                row["Domain"],
                int(row["Ya"]),
                int(row["Tidak"]),
                int(row["N/A"])
            ])

        for cell in grafik_ws[1]:

            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )

            cell.fill = PatternFill(
                "solid",
                fgColor="4472C4"
            )

        chart = BarChart()

        chart.type = "col"
        chart.style = 10
        chart.title = "Grafik Hasil Audit COBIT 5"
        chart.y_axis.title = "Jumlah Jawaban"
        chart.x_axis.title = "Domain"

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

        chart.add_data(
            chart_data,
            titles_from_data=True
        )

        chart.set_categories(categories)

        chart.height = 10
        chart.width = 20

        grafik_ws.add_chart(chart, "F2")

        # ATUR LEBAR KOLOM
        lebar = {
            "A": 12,
            "B": 14,
            "C": 14,
            "D": 85,
            "E": 18
        }

        for kolom, ukuran in lebar.items():
            ws.column_dimensions[kolom].width = ukuran

        ws.freeze_panes = "A12"
        ws.auto_filter.ref = f"A11:E{header_row + len(df)}"

        # SIMPAN
        output = io.BytesIO()

        wb.save(output)
        output.seek(0)

        return output.getvalue()

    # GENERATE EXCEL
    file_excel = buat_excel(
        identitas,
        df,
        df_rekap,
        ya,
        tidak,
        na,
        persentase
    )

    # TOMBOL DOWNLOAD
    st.divider()

    st.header("4. Unduh Laporan Audit")

    st.download_button(
        label="📥 Unduh Laporan Audit Excel",
        data=file_excel,
        file_name="Laporan_Audit_COBIT5_Alammu.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.success(
        "Laporan Excel berisi identitas, 30 pertanyaan, "
        "jawaban, rekapitulasi, dan grafik."
    )
