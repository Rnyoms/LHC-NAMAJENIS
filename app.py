import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# ------------------ KONFIGURASI JENIS ------------------
JENIS_POHON = {
    "Kelompok Meranti": ["KENARI", "NYATOH", "PULAI", "MERSAWA", "RESAK", "MATOA", "MERAWAN", "PENJALIN"],
    "Kelompok Rimba Campuran": [
        "BENUANG", "TERENTANG", "KETAPANG", "JAMBU - JAMBU", "KEDONDONG HUTAN", "SESENDOK",
        "MENDARAHAN", "BINTANGUR", "TERAP", "BUGIS", "BIPA", "DUABANGA", "LANCAT", "KENANGA",
        "MEMPISANG", "SENGON", "SURIAN", "TENGGAYUN", "MEDANG", "JABON", "PELAWAN", "KAYU BATU",
        "LARA", "SIMPUR", "LASI", "MAHANG", "GOPASA"
    ],
    "Kelompok Kayu Indah": ["CEMPAKA", "DAHU", "MELUR", "SINDUR"]
}

# ------------------ UI INPUT ------------------
st.title("🌲 Simulasi Acak Jenis Pohon Berdasarkan Persentase")

nama_simulasi = st.text_input("Nama Simulasi", "Simulasi-Pohon")
jumlah_pohon = st.number_input("Jumlah Pohon", min_value=1, value=100, step=1)

st.subheader("🎯 Persentase Jenis Pohon")
persen_jenis = {}
total_input = 0
sisa_input = []

for kelompok, jenis_list in JENIS_POHON.items():
    st.markdown(f"### {kelompok}")
    for jenis in jenis_list:
        col1, col2 = st.columns([3, 1])
        with col1:
            persen = st.slider(jenis, min_value=0, max_value=100, value=0, step=1)
        persen_jenis[jenis] = persen
        if persen > 0:
            total_input += persen
        else:
            sisa_input.append(jenis)

# ------------------ PENGACAKAN ------------------
def acak_jenis(jumlah, persen_jenis):
    # Hitung total input dan distribusi ke sisa
    total_input = sum([p for p in persen_jenis.values() if p > 0])
    kosong = [j for j, p in persen_jenis.items() if p == 0]
    sisa = 100 - total_input
    auto_persen = {j: sisa / len(kosong) for j in kosong} if kosong else {}

    # Gabungkan
    final_persen = {
        j: (p if p > 0 else auto_persen.get(j, 0))
        for j, p in persen_jenis.items()
    }

    jenis_list = list(final_persen.keys())
    probs = [final_persen[j] / 100 for j in jenis_list]
    hasil = np.random.choice(jenis_list, size=jumlah, p=probs)
    return pd.DataFrame({"Jenis": hasil})

# ------------------ PROSES ------------------
if st.button("🚀 Jalankan Simulasi"):
    if total_input > 100:
        st.error(f"Total persentase melebihi 100% ({total_input}%). Harap koreksi input.")
        st.stop()

    df_hasil = acak_jenis(jumlah_pohon, persen_jenis)

    # Rekap hasil
    rekap = df_hasil.value_counts().reset_index()
    rekap.columns = ["Jenis", "Jumlah"]
    rekap["Persentase"] = rekap["Jumlah"] / jumlah_pohon * 100

    st.success("✅ Simulasi selesai!")
    st.dataframe(rekap)

    # Simpan ke Excel
    wb = Workbook()
    ws_data = wb.active
    ws_data.title = "DataPohon"
    for r in dataframe_to_rows(df_hasil, index=False, header=True):
        ws_data.append(r)

    ws_rekap = wb.create_sheet("Rekap")
    for r in dataframe_to_rows(rekap, index=False, header=True):
        ws_rekap.append(r)

    filename = f"{nama_simulasi}.xlsx"
    wb.save(filename)

    st.download_button("⬇️ Unduh Excel", open(filename, "rb"), file_name=filename)
