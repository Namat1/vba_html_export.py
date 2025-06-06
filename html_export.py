import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import calendar

# Hilfsfunktion: Kalenderwoche berechnen
def get_kw(datum):
    return datum.isocalendar()[1]

# Streamlit UI
st.set_page_config(page_title="Touren-CSV-Export", layout="centered")
st.title("Tourenplan als CSV exportieren")

uploaded_file = st.file_uploader("Excel-Datei hochladen (Blatt 'Touren')", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Touren", skiprows=4, engine="openpyxl")

        daten = []
        for idx, row in df.iterrows():
            datum = row.iloc[14]  # Spalte O
            tour = row.iloc[15]   # Spalte P
            uhrzeit = row.iloc[8] # Spalte I

            nachname = row.iloc[3] if pd.notna(row.iloc[3]) else row.iloc[6]  # D oder G
            vorname = row.iloc[4] if pd.notna(row.iloc[4]) else row.iloc[7]  # E oder H

            if pd.isna(datum) or pd.isna(tour):
                continue

            try:
                datum_dt = pd.to_datetime(datum)
            except:
                continue

            kw = get_kw(datum_dt)
            wochentag = calendar.day_name[datum_dt.weekday()]

            daten.append({
                "Kalenderwoche": kw,
                "Nachname": nachname,
                "Vorname": vorname,
                "Datum": datum_dt.strftime("%d.%m.%Y"),
                "Wochentag": wochentag,
                "Uhrzeit": str(uhrzeit) if pd.notna(uhrzeit) else "",
                "Tour": str(tour).strip()
            })

        export_df = pd.DataFrame(daten)

        csv = export_df.to_csv(index=False).encode("utf-8")
        st.success(f"{len(export_df)} Einträge exportiert.")
        st.download_button("CSV herunterladen", data=csv, file_name="touren_export.csv", mime="text/csv")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten: {e}")
