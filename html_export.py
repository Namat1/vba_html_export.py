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

        fahrer_dict = {}

        for idx, row in df.iterrows():
            datum = row.iloc[14]  # Spalte O
            tour = row.iloc[15]   # Spalte P
            uhrzeit = row.iloc[8] # Spalte I

            if pd.isna(datum):
                continue

            try:
                datum_dt = pd.to_datetime(datum)
            except:
                continue

            kw = get_kw(datum_dt)
            wochentag = datum_dt.strftime("%A")

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

            eintrag = f"{datum_dt.strftime('%d.%m.%Y')} ({wochentag}): {uhrzeit_str} – {str(tour).strip()}"

            # Fahrer 1 (D/E)
            if pd.notna(row.iloc[3]) or pd.notna(row.iloc[4]):
                nachname = str(row.iloc[3]).strip().title() if pd.notna(row.iloc[3]) else ""
                vorname = str(row.iloc[4]).strip().title() if pd.notna(row.iloc[4]) else ""
                if nachname or vorname:
                    fahrer_key = f"{nachname}, {vorname} | KW{kw}"
                    if fahrer_key not in fahrer_dict:
                        fahrer_dict[fahrer_key] = []
                    fahrer_dict[fahrer_key].append(eintrag)

            # Fahrer 2 (G/H)
            if pd.notna(row.iloc[6]) or pd.notna(row.iloc[7]):
                nachname = str(row.iloc[6]).strip().title() if pd.notna(row.iloc[6]) else ""
                vorname = str(row.iloc[7]).strip().title() if pd.notna(row.iloc[7]) else ""
                if nachname or vorname:
                    fahrer_key = f"{nachname}, {vorname} | KW{kw}"
                    if fahrer_key not in fahrer_dict:
                        fahrer_dict[fahrer_key] = []
                    fahrer_dict[fahrer_key].append(eintrag)

        export_rows = []
        for fahrer, eintraege in fahrer_dict.items():
            export_rows.append({
                "Fahrer": fahrer,
                "Einsätze": " | ".join(eintraege)
            })

        export_df = pd.DataFrame(export_rows)

        csv = export_df.to_csv(index=False, encoding="utf-8-sig")
        st.success(f"{len(export_df)} Fahrer-Einträge exportiert.")
        st.download_button("CSV herunterladen", data=csv, file_name="touren_kompakt.csv", mime="text/csv")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten: {e}")
