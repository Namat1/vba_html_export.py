import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import calendar

# Deutsche Übersetzung für strftime
wochentage_deutsch_map = {
    "Monday": "Montag",
    "Tuesday": "Dienstag",
    "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag",
    "Friday": "Freitag",
    "Saturday": "Samstag",
    "Sunday": "Sonntag"
}

# Kalenderwoche berechnen
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
        fahrer_kw_map = {}

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
            wochentag_en = datum_dt.strftime("%A")
            wochentag = wochentage_deutsch_map.get(wochentag_en, wochentag_en)

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

            eintrag_text = f"{uhrzeit_str} – {str(tour).strip()}"

            for pos in [(3, 4), (6, 7)]:  # D/E und G/H
                nachname = str(row.iloc[pos[0]]).strip().title() if pd.notna(row.iloc[pos[0]]) else ""
                vorname = str(row.iloc[pos[1]]).strip().title() if pd.notna(row.iloc[pos[1]]) else ""
                if nachname or vorname:
                    fahrer_name = f"{nachname}, {vorname}"
                    if fahrer_name not in fahrer_dict:
                        fahrer_dict[fahrer_name] = {}
                    if datum_dt.date() not in fahrer_dict[fahrer_name]:
                        fahrer_dict[fahrer_name][datum_dt.date()] = []
                    if eintrag_text not in fahrer_dict[fahrer_name][datum_dt.date()]:
                        fahrer_dict[fahrer_name][datum_dt.date()].append(eintrag_text)

        export_rows = []
        for fahrer_name, eintraege in fahrer_dict.items():
            if not eintraege:
                continue

            start_datum = min(eintraege.keys())
            start_sonntag = start_datum - pd.Timedelta(days=(start_datum.weekday() + 1) % 7)
            kw = get_kw(start_sonntag)

            wochen_eintraege = []
            for i in range(7):
                tag_datum = start_sonntag + pd.Timedelta(days=i)
                wochentag_en = tag_datum.strftime("%A")
                wochentag = wochentage_deutsch_map.get(wochentag_en, wochentag_en)

                if tag_datum in eintraege:
                    for eintrag in eintraege[tag_datum]:
                        zeile = f"{tag_datum.strftime('%d.%m.%Y')} ({wochentag}): {eintrag}"
                        wochen_eintraege.append(zeile)
                else:
                    zeile = f"{tag_datum.strftime('%d.%m.%Y')} ({wochentag}): –"
                    wochen_eintraege.append(zeile)

            export_rows.append({
                "Fahrer": fahrer_name,
                "Kalenderwoche": f"KW{kw}",
                "Einsätze": " | ".join(wochen_eintraege)
            })

        export_df = pd.DataFrame(export_rows)

        csv = export_df.to_csv(index=False, encoding="utf-8-sig")
        st.success(f"{len(export_df)} Fahrer-Einträge exportiert.")
        st.download_button("CSV herunterladen", data=csv, file_name="touren_kompakt.csv", mime="text/csv")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten: {e}")
