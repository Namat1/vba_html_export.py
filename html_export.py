import streamlit as st
import pandas as pd
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
    return datum.isocalendar()[1] + 1  # KW +1 wie gewünscht

# HTML-Template generieren
def generate_html(fahrer_name, eintraege, kw, start_date, css_styles):
    html = f"""<!DOCTYPE html>
<html lang="de">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>KW{kw} – {fahrer_name}</title>
  <style>{css_styles}</style>
</head>
<body>
<div class="container-outer"><div class="container">
<div class="headline-block">
  <div class="headline-kw-box">
    <div class="headline-kw">KW {kw}</div>
    <div class="headline-period">{start_date.strftime('%d.%m.%Y')} – {(start_date + pd.Timedelta(days=6)).strftime('%d.%m.%Y')}</div>
  </div>
</div>"""

    for eintrag in eintraege:
        date_text, content = eintrag.split(": ", 1)
        date_obj = pd.to_datetime(date_text.split(" ")[0], format="%d.%m.%Y")
        weekday = date_text.split("(")[-1].replace(")", "")

        if "–" in content:
            uhrzeit, tour = [x.strip() for x in content.split("–", 1)]
        else:
            uhrzeit, tour = "–", content.strip()

        card_class = "daycard"
        if weekday == "Samstag":
            card_class += " samstag"
        elif weekday == "Sonntag":
            card_class += " sonntag"

        html += f"""
<div class="{card_class}">
  <div class="header-row">
    <div class="prominent-date">{date_obj.strftime('%d.%m.%Y')}</div>
    <div class="weekday">{weekday}</div>
    <div class="prominent-name">{fahrer_name}</div>
  </div>
  <div class="info">
    <div class="info-block">
      <div class="label">Tour:</div>
      <div class="value">{tour}</div>
    </div>
    <div class="info-block">
      <div class="label">Uhrzeit:</div>
      <div class="value">{uhrzeit}</div>
    </div>
  </div>
</div>"""

    html += "</div></div></body></html>"
    return html

# Stildefinition (dein zuletzt gepostetes CSS)
css_styles = """body { font-family: 'Inter', Arial, sans-serif; background: #e3e7ee; margin: 0; padding: 0; color: #1c1c1c; font-size: 15px; }
.container-outer { max-width: 460px; margin: 28px auto; background: #f9fafc; border-radius: 16px; border: 3px solid #5c6f8c; box-shadow: 0 6px 20px rgba(0, 0, 0, 0.08); padding: 20px 16px; }
.headline-block { display: flex; justify-content: center; margin-bottom: 20px; }
.headline-kw-box { text-align: center; border: 3px solid #2a4a9c; border-radius: 14px; padding: 10px 18px; background: #dde5f3; }
.headline-kw { font-size: 1.7rem; font-weight: 900; color: #1e3f85; margin-bottom: 4px; }
.headline-period { font-size: 0.95rem; font-weight: 600; color: #2f528e; }
.daycard { background: #ffffff; border: 3px solid #6c7d99; border-radius: 12px; padding: 16px; margin-bottom: 18px; box-shadow: 0 3px 10px rgba(30, 30, 30, 0.05); }
.daycard.samstag { background: #fef4e8; border-color: #e89b3d; }
.daycard.sonntag { background: #fceaea; border-color: #d65656; }
.header-row { display: flex; flex-wrap: wrap; justify-content: space-between; gap: 8px; margin-bottom: 12px; }
.prominent-date, .weekday, .prominent-name { flex: 1 1 30%; font-size: 1rem; font-weight: 800; text-align: center; color: #ffffff; padding: 7px 10px; border-radius: 8px; text-shadow: 0 1px 2px rgba(0,0,0,0.25); }
.prominent-date { background: linear-gradient(to right, #933636, #c74e4e); }
.weekday { background: linear-gradient(to right, #427a42, #63a663); }
.prominent-name { background: linear-gradient(to right, #355eb5, #5f8de3); }
.info { display: flex; flex-wrap: wrap; justify-content: space-between; gap: 10px; margin-top: 6px; }
.info-block { flex: 1 1 48%; }
.label { font-size: 0.8rem; color: #2c436e; font-weight: 600; margin-bottom: 2px; }
.value { font-size: 0.94rem; font-weight: 700; background: #ecf0f8; padding: 6px 10px; border: 2px solid #889ab5; border-radius: 6px; display: inline-block; min-width: 40px; }
@media (max-width: 480px) { .prominent-date, .weekday, .prominent-name, .info-block { flex: 1 1 100%; } }"""
"""  # Oder so wie du es schon oben integriert hast

# Streamlit UI
st.set_page_config(page_title="Touren-Export", layout="centered")
st.title("Tourenplan als HTML je KW exportieren")

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

            html_code = generate_html(fahrer_name, wochen_eintraege, kw, start_sonntag, css_styles)
            kw_folder = f"KW{kw:02d}"
            if kw_folder not in html_files:
                html_files[kw_folder] = {}
            name_html = fahrer_name.split(",")[0].replace(" ", "_")
            filename = f"{name_html}.html"
            html_files[kw_folder][filename] = html_code

        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "touren_html_export.zip")
            with ZipFile(zip_path, "w") as zipf:
                for kw_folder, file_dict in html_files.items():
                    for filename, html_content in file_dict.items():
                        folder_path = os.path.join(tmpdir, kw_folder)
                        os.makedirs(folder_path, exist_ok=True)
                        file_path = os.path.join(folder_path, filename)
                        with open(file_path, "w", encoding="utf-8") as f:
                            f.write(html_content)
                        zipf.write(file_path, arcname=os.path.join(kw_folder, filename))
            with open(zip_path, "rb") as f:
                zip_bytes = f.read()

        st.success("HTML-Dateien wurden erzeugt und nach KW sortiert.")
        st.download_button("ZIP herunterladen", data=zip_bytes, file_name="touren_nach_kw.zip", mime="application/zip")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten: {e}")
