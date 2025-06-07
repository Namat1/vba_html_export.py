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
  <div class="header-row compact">
    <div class="header-left">
      <div class="date-line">{date_obj.strftime('%d.%m.%Y')} ({weekday})</div>
    </div>
    <div class="header-right">
      <div class="fahrer-name">{fahrer_name}</div>
    </div>
  </div>
  <div class="info">
    <div><span class="label">Tour:</span> <span class="value">{tour}</span></div>
    <div><span class="label">Uhrzeit:</span> <span class="value">{uhrzeit}</span></div>
  </div>
</div>"""

    html += "</div></div></body></html>"
    return html


# CSS-Stile 
css_styles = """
body {
  font-family: 'Inter', Arial, sans-serif;
  background: #f4f6f9;
  margin: 0;
  padding: 0;
  color: #1c1c1c;
  font-size: 15px;
}
.container-outer {
  max-width: 460px;
  margin: 28px auto;
  background: #ffffff;
  border-radius: 12px;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.08);
  padding: 18px 14px;
}
.headline-block {
  text-align: center;
  margin-bottom: 16px;
}
.headline-kw-box {
  border: 2px solid #355eb5;
  border-radius: 12px;
  padding: 8px 16px;
  background: #e7edfa;
}
.headline-kw {
  font-size: 1.6rem;
  font-weight: 800;
  color: #1e3f85;
  margin-bottom: 2px;
}
.headline-period {
  font-size: 0.95rem;
  font-weight: 600;
  color: #2f528e;
}

.daycard {
  background: #fdfdfd;
  border: 1px solid #ccc;
  border-radius: 10px;
  padding: 12px 14px;
  margin-bottom: 14px;
  box-shadow: 0 2px 5px rgba(0,0,0,0.05);
}
.daycard.samstag {
  background: #fff6e6;
  border-color: #e89b3d;
}
.daycard.sonntag {
  background: #ffecec;
  border-color: #d65656;
}

.header-row.compact {
  display: flex;
  justify-content: space-between;
  flex-wrap: wrap;
  margin-bottom: 6px;
  border-bottom: 1px solid #dcdcdc;
  padding-bottom: 4px;
}
.header-left .date-line {
  font-size: 0.95rem;
  font-weight: 700;
  color: #2a2a2a;
}
.header-right .fahrer-name {
  font-size: 0.95rem;
  font-weight: 700;
  color: #355eb5;
  text-align: right;
}

.info {
  display: flex;
  justify-content: space-between;
  flex-wrap: wrap;
  font-size: 0.9rem;
  gap: 8px;
  margin-top: 6px;
}
.label {
  font-weight: 600;
  color: #555;
  margin-right: 4px;
}
.value {
  font-weight: 600;
  color: #222;
  background: #f0f3f8;
  padding: 4px 8px;
  border-radius: 6px;
  border: 1px solid #aaa;
  display: inline-block;
}

@media (max-width: 480px) {
  .header-row.compact {
    flex-direction: column;
    align-items: flex-start;
  }
  .info {
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
