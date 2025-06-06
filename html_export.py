import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
import tempfile
import os

# Deutsche Wochentage
wochentage_deutsch_map = {
    "Monday": "Montag", "Tuesday": "Dienstag", "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag", "Friday": "Freitag", "Saturday": "Samstag", "Sunday": "Sonntag"
}

# Kalenderwoche berechnen
def get_kw(datum):
    return datum.isocalendar()[1]

# HTML-Template
def generate_html(fahrer_name, eintraege, kw, start_date, css_styles):
    html = f"""<!DOCTYPE html>
<html lang="de">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>KW{kw} – {fahrer_name}</title>
<style>{css_styles}</style></head><body>
<div class="container-outer"><div class="container">
<div class="headline-block"><div class="headline-kw">KW{kw}</div>
<div class="headline-period">{start_date.strftime('%d.%m.%Y')} – {(start_date + pd.Timedelta(days=6)).strftime('%d.%m.%Y')}</div></div>"""
    for eintrag in eintraege:
        date_text, content = eintrag.split(": ", 1)
        date_obj = pd.to_datetime(date_text.split(" ")[0], format="%d.%m.%Y")
        weekday = date_text.split("(")[-1].replace(")", "")
        html += f"""<div class="daycard"><div class="header-row">
<div class="prominent-date">{date_obj.strftime('%d.%m.%Y')}</div>
<div class="weekday">{weekday}</div>
<div class="prominent-name">{fahrer_name}</div></div>
<div class="info"><div class="label">Tour:</div><div class="value">{content}</div></div></div>"""
    html += "</div></div></body></html>"
    return html

# CSS-Stile
css_styles = """
body {font-family: 'Inter', Arial, sans-serif; background: #f5f8fb; margin: 0; color: #23262f; min-height: 100vh; font-size: 14px;}
.container-outer {max-width: 430px; width: 94vw; margin: 16px auto; background: linear-gradient(112deg, #eaf1ff 0%, #f9fafd 100%);
  border-radius: 16px; box-shadow: 0 4px 19px rgba(40,60,130,0.14); padding: 10px 4px 12px 4px; border: 2px solid #a9bcdf;}
.container {max-width: 390px; margin: 0 auto; padding: 0;}
.headline-block {display: flex; flex-direction: column; align-items: center; margin-bottom: 0.8em;}
.headline-kw {font-size: 1.5rem; font-weight: 900; background: linear-gradient(92deg,#2464e4 25%,#63b3ff 100%);
  color: #fff; border-radius: 10px; padding: 0.2em 0.9em; border: 2px solid #2564e4; text-shadow: 0 1px 3px black;}
.headline-period {font-size: 0.78rem; color: #2564e4; background: #eaf2fe; border-radius: 14px; font-weight: 700;
  padding: 0.25em 0.9em; border: 1px solid #2564e4; margin-top: 0.2em;}
.daycard {background: #fff; border-radius: 10px; box-shadow: 0 1px 4px rgba(60,80,160,0.1); margin-bottom: 6px;
  padding: 10px 6px; border: 1.5px solid #2564e4;}
.header-row {display: flex; justify-content: space-between; align-items: center; gap: 4px; margin-bottom: 6px;}
.prominent-name, .prominent-date, .weekday {
  flex: 1 1 0; text-align: center; font-size: 0.95rem; font-weight: 800; color: #f0f4ff;
  padding: 3px 10px; border-radius: 6px; text-shadow: 0 1px 2px black;}
.prominent-name {background: linear-gradient(90deg,#2564e4 40%,#4aa8ff 100%); border: 1px solid #2564e4;}
.prominent-date {background: linear-gradient(90deg,#e03244,#fc8282); border: 1px solid #fc1d1d;}
.weekday {background: linear-gradient(90deg,#3abc3a,#a4ffb5); border: 1px solid #218a16;}
.label {font-size: 0.8rem; color: #2564e4; font-weight: 700;}
.value {font-weight: 800; font-size: 0.9rem; background: #f6faff; padding: 3px 6px; border: 1px solid #aac3e9;
  border-radius: 4px; display: inline-block;}
"""

# Streamlit UI
st.set_page_config(page_title="Touren-Export", layout="centered")
st.title("Tourenplan als CSV + HTML exportieren")

uploaded_file = st.file_uploader("Excel-Datei hochladen (Blatt 'Touren')", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Touren", skiprows=4, engine="openpyxl")

        fahrer_dict = {}
        for _, row in df.iterrows():
            datum = row.iloc[14]  # Spalte O
            tour = row.iloc[15]   # Spalte P
            uhrzeit = row.iloc[8] # Spalte I
            if pd.isna(datum): continue
            try: datum_dt = pd.to_datetime(datum)
            except: continue
            if pd.isna(uhrzeit): uhrzeit_str = "–"
            elif isinstance(uhrzeit, (int, float)) and uhrzeit == 0: uhrzeit_str = "00:00"
            elif isinstance(uhrzeit, datetime): uhrzeit_str = uhrzeit.strftime("%H:%M")
            else:
                try: uhrzeit_parsed = pd.to_datetime(uhrzeit)
                except: uhrzeit_str = str(uhrzeit).strip()
                else: uhrzeit_str = uhrzeit_parsed.strftime("%H:%M")
            eintrag_text = f"{uhrzeit_str} – {str(tour).strip()}"
            for pos in [(3, 4), (6, 7)]:
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
        html_files = {}
        for fahrer_name, eintraege in fahrer_dict.items():
            if not eintraege: continue
            start_datum = min(eintraege.keys())
            start_sonntag = start_datum - pd.Timedelta(days=(start_datum.weekday() + 1) % 7)
            kw = get_kw(start_sonntag)
            wochen_eintraege = []
            for i in range(7):
                tag_datum = start_sonntag + pd.Timedelta(days=i)
                wochentag = wochentage_deutsch_map.get(tag_datum.strftime("%A"), tag_datum.strftime("%A"))
                if tag_datum in eintraege:
                    for eintrag in eintraege[tag_datum]:
                        wochen_eintraege.append(f"{tag_datum.strftime('%d.%m.%Y')} ({wochentag}): {eintrag}")
                else:
                    wochen_eintraege.append(f"{tag_datum.strftime('%d.%m.%Y')} ({wochentag}): –")
            export_rows.append({"Fahrer": fahrer_name, "Einsätze": " | ".join(wochen_eintraege)})

            html_code = generate_html(fahrer_name, wochen_eintraege, kw, start_sonntag, css_styles)
            name_html = fahrer_name.split(",")[0].replace(" ", "_")
            filename = f"KW{kw:02d}_{name_html}.html"
            html_files[filename] = html_code

        export_df = pd.DataFrame(export_rows)
        csv = export_df.to_csv(index=False, encoding="utf-8-sig")

        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "touren_html_export.zip")
            with ZipFile(zip_path, "w") as zipf:
                for name, content in html_files.items():
                    filepath = os.path.join(tmpdir, name)
                    with open(filepath, "w", encoding="utf-8") as f:
                        f.write(content)
                    zipf.write(filepath, arcname=name)
            with open(zip_path, "rb") as f:
                zip_bytes = f.read()

        st.success(f"{len(export_df)} Fahrer-Einträge verarbeitet.")
        st.download_button("CSV herunterladen", data=csv, file_name="touren_kompakt.csv", mime="text/csv")
        st.download_button("HTML-Archiv herunterladen", data=zip_bytes, file_name="touren_html.zip", mime="application/zip")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten: {e}")
