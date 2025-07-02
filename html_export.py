import streamlit as st
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv  # für .env Support
import os

# .env laden (falls später FTP/andere Funktionen ergänzt werden sollen)
load_dotenv()

st.set_page_config(page_title="Touren-Export als JSON", layout="centered")
st.title("Dienstplan als JSON exportieren")

uploaded_files = st.file_uploader(
    "Excel-Dateien hochladen (Blatt 'Touren')",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    json_records = []
    exclusion_keywords = [
        "zippel", "insel", "paasch", "meyer",
        "ihde", "devies", "insellogistik"
    ]

    for file in uploaded_files:
        df = pd.read_excel(
            file,
            sheet_name="Touren",
            skiprows=4,
            engine="openpyxl"
        )
        for _, row in df.iterrows():
            datum = row.iloc[14]
            tour = row.iloc[15]
            uhrzeit = row.iloc[8]

            # Datum überspringen wenn ungültig
            if pd.isna(datum):
                continue
            try:
                datum_dt = pd.to_datetime(datum)
            except:
                continue

            # Uhrzeit formatieren
            if pd.isna(uhrzeit):
                uhrzeit_str = "–"
            elif isinstance(uhrzeit, (int, float)) and uhrzeit == 0:
                uhrzeit_str = "00:00"
            elif isinstance(uhrzeit, datetime):
                uhrzeit_str = uhrzeit.strftime("%H:%M")
            else:
                try:
                    uhrzeit_parsed = pd.to_datetime(uhrzeit)
                    uhrzeit_str = uhrzeit_parsed.strftime("%H:%M")
                except:
                    uhrzeit_str = str(uhrzeit).strip()
                    if ":" in uhrzeit_str:
                        uhrzeit_str = ":".join(uhrzeit_str.split(":")[:2])

            # Fahrer-Namen auslesen
            fahrer_name = None
            for pos in [(3, 4), (6, 7)]:
                last = row.iloc[pos[0]]
                first = row.iloc[pos[1]]
                if pd.notna(last) or pd.notna(first):
                    nachname = str(last).strip().title() if pd.notna(last) else ""
                    vorname = str(first).strip().title() if pd.notna(first) else ""
                    fahrer_name = f"{nachname}, {vorname}".strip(', ')
                    break
            if not fahrer_name:
                continue

            # Eintrag zum JSON-Output
            record = {
                "Fahrer": fahrer_name,
                "Datum": datum_dt.date().isoformat(),
                "Uhrzeit": uhrzeit_str,
                "Tour/Aufgabe": str(tour).strip()
            }
            # Ausschluss nach Schlüsselwörtern
            filename_lower = f"{fahrer_name}_{datum_dt.date().isoformat()}".lower()
            if any(keyword in filename_lower for keyword in exclusion_keywords):
                continue
            json_records.append(record)

    # JSON-Export anbieten
    if json_records:
        df_export = pd.DataFrame(json_records)
        json_str = df_export.to_json(orient='records', force_ascii=False, indent=2)
        json_bytes = json_str.encode('utf-8')
        st.download_button(
            label="JSON mit allen Touren herunterladen",
            data=json_bytes,
            file_name="touren_export.json",
            mime="application/json"
        )
    else:
        st.warning("Keine gültigen Einträge gefunden.")
