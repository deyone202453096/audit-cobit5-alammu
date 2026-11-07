

import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Audit COBIT 5",
    page_icon="📋",
    layout="wide"
)

st.title("AUDIT SISTEM INFORMASI PENJUALAN")
st.subheader("CV Seleksi Alam Muria")
st.write("Implementasi Framework COBIT 5")

st.divider()

# DATA PERTANYAAN AUDIT
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

# FORM INPUT
st.header("Form Pertanyaan Audit")
st.write("Pilih jawaban berdasarkan hasil pemeriksaan dan bukti pendukung.")

with st.form("form_audit"):
    jawaban = []

    for domain in ["EDM", "APO", "BAI", "DSS", "MEA"]:
        st.subheader(f"Domain {domain}")

        pertanyaan_domain = [
            (kode, pertanyaan)
            for d, kode, pertanyaan in data if d == domain
        ]

        for i, (kode, pertanyaan) in enumerate(pertanyaan_domain):
            nomor = len([
                x for x in data[:data.index((domain, kode, pertanyaan))]
            ]) + 1

            st.write(f"**{nomor}. [{kode}] {pertanyaan}**")

            pilihan = st.radio(
                "Pilih jawaban:",
                ["Ya", "Tidak", "N/A"],
                key=f"jawaban_{domain}_{i}",
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

# HASIL AUDIT
if submit:
    df = pd.DataFrame(jawaban)

    st.success("Jawaban audit berhasil diproses!")

    st.header("Hasil Jawaban Audit")
    st.dataframe(df, use_container_width=True)

    # REKAPITULASI
    st.header("Rekapitulasi Hasil Audit")

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
    st.dataframe(df_rekap, use_container_width=True)

    # RINGKASAN
    ya = (df["Jawaban"] == "Ya").sum()
    tidak = (df["Jawaban"] == "Tidak").sum()
    na = (df["Jawaban"] == "N/A").sum()

    total_dinilai = ya + tidak
    persentase = (ya / total_dinilai * 100) if total_dinilai else 0

    st.header("Ringkasan Audit")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Total Pertanyaan", len(df))
    col2.metric("Jawaban Ya", ya)
    col3.metric("Jawaban Tidak", tidak)
    col4.metric("Jawaban N/A", na)

    st.metric(
        "Indikator Kepatuhan Sederhana",
        f"{persentase:.2f}%"
    )

    st.caption(
        "Indikator ini bukan capability level resmi COBIT 5."
    )

    # DOWNLOAD EXCEL
    import io

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Jawaban Audit", index=False)
        df_rekap.to_excel(writer, sheet_name="Rekapitulasi", index=False)

    st.download_button(
        label="Unduh Hasil Audit Excel",
        data=output.getvalue(),
        file_name="Hasil_Audit_COBIT5.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
