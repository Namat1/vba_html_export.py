import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import calendar

# Deutsche Wochentage
wochentage_deutsch = {
    0: "Montag",
    1: "Dienstag",
    2: "Mittwoch",
    3: "Donnerstag",
    4: "Freitag",
    5: "Samstag",
    6: "Sonntag"
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
            wochentag_idx = datum_dt.weekday()
            wochentag = wochentage_deutsch.get(wochentag_idx, datum_dt.strftime("%A"))

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

            eintrag = {
                "datum": datum_dt.strftime('%d.%m.%Y'),
                "wochentag": wochentag,
                "uhrzeit": uhrzeit_str,
                "tour": str(tour).strip()
            }

            for pos in [(3, 4), (6, 7)]:  # D/E und G/H
                nachname = str(row.iloc[pos[0]]).strip().title() if pd.notna(row.iloc[pos[0]]) else ""
                vorname = str(row.iloc[pos[1]]).strip().title() if pd.notna(row.iloc[pos[1]]) else ""
                if nachname or vorname:
                    fahrer_key = f"{nachname}, {vorname} | KW{kw}"
                    if fahrer_key not in fahrer_dict:
                        fahrer_dict[fahrer_key] = {}
                    fahrer_dict[fahrer_key][datum_dt.date()] = eintrag

        export_rows = []
        for fahrer, eintraege in fahrer_dict.items():
            # Startdatum bestimmen
            alle_daten = list(eintraege.keys())
            if not alle_daten:
                continue
            start_datum = min(alle_daten)
            start_montag = start_datum - pd.Timedelta(days=start_datum.weekday())
            eintragsliste = []
            for i in range(7):
                tag_datum = start_montag + pd.Timedelta(days=i)
                if tag_datum in eintraege:
                    e = eintraege[tag_datum]
                    eintragsliste.append(f"{e['datum']} ({e['wochentag']}): {e['uhrzeit']} – {e['tour']}")
                else:
                    wtag = wochentage_deutsch[i]
                    eintragsliste.append(f"{(start_montag + pd.Timedelta(days=i)).strftime('%d.%m.%Y')} ({wtag}): –")

            export_rows.append({
                "Fahrer": fahrer,
                "Einsätze": " | ".join(eintragsliste)
            })

        export_df = pd.DataFrame(export_rows)

        csv = export_df.to_csv(index=False, encoding="utf-8-sig")
        st.success(f"{len(export_df)} Fahrer-Einträge exportiert.")
        st.download_button("CSV herunterladen", data=csv, file_name="touren_kompakt.csv", mime="text/csv")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten: {e}")
