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
<div class="page-wrap">
  <div class="headline">
    <div class="kw-box">
      <div class="kw-title">KW {kw}</div>
      <div class="kw-range">{start_date.strftime('%d.%m.%Y')} – {(start_date + pd.Timedelta(days=6)).strftime('%d.%m.%Y')}</div>
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

        # CSS-Klasse für Samstag oder Sonntag zuweisen
        card_class = "daycard"
        if weekday == "Samstag":
            card_class += " samstag"
        elif weekday == "Sonntag":
            card_class += " sonntag"

        html += f"""
  <div class="{card_class}">
    <div class="daycard-header">
      <div class="daycard-date">{date_obj.strftime('%d.%m.%Y')}</div>
      <div class="daycard-weekday">{weekday}</div>
      <div class="daycard-name">{fahrer_name}</div>
    </div>
    <div class="daycard-info">
      <div><strong>Tour:</strong> <span>{tour}</span></div>
      <div><strong>Uhrzeit:</strong> <span>{uhrzeit}</span></div>
    </div>
  </div>"""

    html += "</div></body></html>"
    return html




# CSS-Stile 
css_styles = """
body {
  margin: 0;
  padding: 0;
  background: #f5f7fa;
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
  color: #1d1d1f;
  font-size: 15px;
}

.page-wrap {
  max-width: 500px;
  margin: 28px auto;
  padding: 0 16px;
}

.headline {
  text-align: center;
  margin-bottom: 24px;
}

.kw-box {
  background: #eef2f9;
  border-radius: 16px;
  padding: 12px 18px;
  border: 2px solid #a8b4cc;
  box-shadow: 0 2px 6px rgba(0,0,0,0.06);
}

.kw-title {
  font-size: 1.6rem;
  font-weight: 700;
  color: #1b3a7a;
  margin-bottom: 4px;
}

.kw-range {
  font-size: 0.95rem;
  color: #3e567f;
}

.daycard {
  background: #ffffff;
  border-radius: 14px;
  padding: 16px 18px;
  margin-bottom: 18px;
  border: 2px solid #b4bcc9;
  box-shadow: 0 3px 8px rgba(0,0,0,0.08);
  transition: box-shadow 0.2s;
}

.daycard:hover {
  box-shadow: 0 4px 14px rgba(0,0,0,0.12);
}

/* Samstag & Sonntag gleich, aber deutlich hervorgehoben */
.daycard.samstag,
.daycard.sonntag {
  background: #fff5dc;
  border: 2px solid #e6aa00;
  box-shadow: inset 0 0 0 5px #ffe18c, 0 4px 10px rgba(0,0,0,0.1);
}

.daycard.samstag .daycard-header,
.daycard.sonntag .daycard-header {
  background: #fce9b5;
  border-radius: 8px;
  padding: 8px 10px;
  margin: -16px -18px 12px -18px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  border-bottom: 2px solid #e6aa00;
}

.daycard.samstag .daycard-date,
.daycard.sonntag .daycard-date {
  color: #a85c00;
  font-weight: 700;
}

.daycard.samstag .daycard-weekday,
.daycard.sonntag .daycard-weekday {
  color: #7a4e00;
  font-weight: 700;
}

.daycard.samstag .daycard-name,
.daycard.sonntag .daycard-name {
  color: #1a3662;
  font-weight: 700;
}

.daycard-header {
  display: flex;
  justify-content: space-between;
  flex-wrap: wrap;
  margin-bottom: 10px;
  font-weight: 600;
  font-size: 0.95rem;
  color: #2a2a2a;
}

.daycard-weekday {
  color: #5e8f64;
}

.daycard-date {
  color: #bb4444;
}

.daycard-name {
  color: #3568b5;
}

.daycard-info {
  display: flex;
  justify-content: space-between;
  flex-wrap: wrap;
  gap: 10px;
  font-size: 0.9rem;
  padding-top: 4px;
}

.daycard-info div {
  background: #f4f6fb;
  padding: 6px 10px;
  border-radius: 8px;
  border: 1.5px solid #9ca7bc;
  flex: 1 1 45%;
  display: flex;
  justify-content: space-between;
}

/* Responsiv */
@media (max-width: 480px) {
  .daycard-header {
    flex-direction: column;
    align-items: flex-start;
    gap: 2px;
  }
  .daycard-info {
    flex-direction: column;
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
        namensliste = {}  # Nachnamen-Zähler initialisieren

        for fahrer_name, eintraege in fahrer_dict.items():
            if not eintraege:
                continue

            start_datum = min(eintraege.keys())
            start_sonntag = start_datum - pd.Timedelta(days=(start_datum.weekday() + 1) % 7)
            kw = get_kw(start_sonntag) + 1

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

            # === Dateinamen mit Nachname und Duplikat-Suffix ===
            nachname = fahrer_name.split(",")[0].strip().replace(" ", "_")

            if nachname not in namensliste:
                namensliste[nachname] = 0
                dateiname = nachname
            else:
                namensliste[nachname] += 1
                dateiname = f"{nachname}_{namensliste[nachname]}"

            if dateiname == "Fechner_1":
                filename = f"KW{kw:02d}_KFechner.html"
            elif dateiname == "Scheil_1":
               filename = f"KW{kw:02d}_RScheil.html"
            else:
               filename = f"KW{kw:02d}_{dateiname}.html"


            html_code = generate_html(fahrer_name, wochen_eintraege, kw, start_sonntag, css_styles)
            html_files[filename] = html_code




        export_df = pd.DataFrame(export_rows)
        csv = export_df.to_csv(index=False, encoding="utf-8-sig")

        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "touren_html_export.zip")
            kw_folder = f"KW{kw:02d}"
            with ZipFile(zip_path, "w") as zipf:
                for name, content in html_files.items():
                    subpath = os.path.join(tmpdir, kw_folder, name)
                    os.makedirs(os.path.dirname(subpath), exist_ok=True)
                    with open(subpath, "w", encoding="utf-8") as f:
                        f.write(content)
                    zipf.write(subpath, arcname=os.path.join(kw_folder, name))

            with open(zip_path, "rb") as f:
                zip_bytes = f.read()

        st.success(f"{len(export_df)} Fahrer-Einträge verarbeitet.")
        st.download_button("CSV herunterladen", data=csv, file_name="touren_kompakt.csv", mime="text/csv")
        st.download_button("HTML-Archiv herunterladen", data=zip_bytes, file_name="touren_html.zip", mime="application/zip")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten: {e}")
