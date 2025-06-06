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
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>KW{kw} – {fahrer_name}</title>
  <style>{css_styles}</style>
</head>
<body>
<div class="container-outer"><div class="container">
<div class="headline-block">
  <div class="headline-kw">KW{kw}</div>
  <div class="headline-period">{start_date.strftime('%d.%m.%Y')} – {(start_date + pd.Timedelta(days=6)).strftime('%d.%m.%Y')}</div>
</div>"""

    for eintrag in eintraege:
        date_text, content = eintrag.split(": ", 1)
        date_obj = pd.to_datetime(date_text.split(" ")[0], format="%d.%m.%Y")
        weekday = date_text.split("(")[-1].replace(")", "")

        # Uhrzeit & Tour trennen
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


# CSS-Stile
css_styles = """
body {
  font-family: 'Inter', Arial, sans-serif;
  background: #f4f6fb;
  margin: 0;
  padding: 0;
  color: #1c1c1c;
  font-size: 15px;
}
.container-outer {
  max-width: 460px;
  margin: 24px auto;
  background: #ffffff;
  border-radius: 16px;
  box-shadow: 0 4px 20px rgba(0, 0, 0, 0.04);
  border: 1px solid #dce3ec;
  padding: 18px 14px;
}
.headline-block {
  text-align: center;
  margin-bottom: 20px;
}
.headline-kw {
  font-size: 1.6rem;
  font-weight: 800;
  background: linear-gradient(92deg, #2564e4, #4b91f4);
  color: #ffffff;
  padding: 6px 16px;
  border-radius: 12px;
  display: inline-block;
  text-shadow: 0 1px 1px rgba(0,0,0,0.2);
}
.headline-period {
  margin-top: 6px;
  font-size: 0.9rem;
  font-weight: 600;
  background: #eef3fb;
  color: #1f4d91;
  padding: 4px 10px;
  border-radius: 10px;
  border: 1px solid #cfdcf0;
  display: inline-block;
}

.daycard {
  background: #ffffff;
  border: 1px solid #d8e0eb;
  border-radius: 10px;
  padding: 14px;
  margin-bottom: 14px;
  box-shadow: 0 2px 6px rgba(60, 60, 60, 0.03);
}

.header-row {
  display: flex;
  flex-wrap: wrap;
  justify-content: space-between;
  gap: 6px;
  margin-bottom: 10px;
}
.prominent-date,
.weekday,
.prominent-name {
  flex: 1 1 30%;
  font-size: 0.95rem;
  font-weight: 700;
  text-align: center;
  color: #ffffff;
  padding: 6px 10px;
  border-radius: 8px;
  text-shadow: 0 1px 1px rgba(0,0,0,0.2);
}
.prominent-date {
  background: linear-gradient(to right, #b14a4a, #d86565);
}
.weekday {
  background: linear-gradient(to right, #529e52, #7acc7a);
}
.prominent-name {
  background: linear-gradient(to right, #3a6fd4, #6ea8f8);
}

.info {
  display: flex;
  flex-wrap: wrap;
  justify-content: space-between;
  gap: 10px;
  margin-top: 6px;
}
.info-block {
  flex: 1 1 48%;
}
.label {
  font-size: 0.8rem;
  color: #375a91;
  font-weight: 600;
  margin-bottom: 2px;
}
.value {
  font-size: 0.92rem;
  font-weight: 600;
  background: #f3f6fc;
  padding: 6px 10px;
  border: 1px solid #c6d2e5;
  border-radius: 6px;
  display: inline-block;
  min-width: 40px;
}

@media (max-width: 480px) {
  .prominent-date, .weekday, .prominent-name {
    flex: 1 1 100%;
  }
  .info-block {
    flex: 1 1 100%;
  }
}

/* Samstag & Sonntag außerhalb des media-queries! */
.daycard.samstag {
  background: #fff7ef;
  border-color: #f4be83;
}
.daycard.sonntag {
  background: #fff0f0;
  border-color: #ee8d8d;
}


}

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
