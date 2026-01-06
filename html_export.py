# Erweiterungen oben ergänzen
import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
import tempfile
import os
from ftplib import FTP
from dotenv import load_dotenv  # für .env Support

# .env laden
load_dotenv()
FTP_HOST = os.getenv("FTP_HOST")
FTP_USER = os.getenv("FTP_USER")
FTP_PASS = os.getenv("FTP_PASS")
FTP_BASE_DIR = os.getenv("FTP_BASE_DIR", "/")

# Deutsche Wochentage
wochentage_deutsch_map = {
    "Monday": "Montag", "Tuesday": "Dienstag", "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag", "Friday": "Freitag", "Saturday": "Samstag", "Sunday": "Sonntag"
}

def get_kw(datum):
    return datum.isocalendar()[1]

def upload_folder_to_ftp_with_progress(local_dir, ftp_dir):
    ftp = FTP()
    ftp.connect(FTP_HOST, 21)
    ftp.login(FTP_USER, FTP_PASS)

    all_files = []
    for root, _, files in os.walk(local_dir):
        for file in files:
            rel_dir = os.path.relpath(root, local_dir)
            all_files.append((os.path.join(root, file), os.path.join(ftp_dir, rel_dir, file)))

    total = len(all_files)
    uploaded = 0

    progress_bar = st.progress(0)
    status_text = st.empty()

    for local_path, remote_path in all_files:
        remote_dir = os.path.dirname(remote_path).replace("\\", "/")

        parts = remote_dir.split("/")
        path_built = ""
        for part in parts:
            if part:
                path_built += "/" + part
                try:
                    ftp.mkd(path_built)
                except:
                    pass

        with open(local_path, "rb") as f:
            ftp.cwd(remote_dir)
            ftp.storbinary(f"STOR {os.path.basename(local_path)}", f)

        uploaded += 1
        progress = uploaded / total
        progress_bar.progress(progress)
        status_text.info(f"Hochgeladen: {uploaded}/{total} – {os.path.basename(local_path)}")

    ftp.quit()
    status_text.success("Alle Dateien erfolgreich hochgeladen.")

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
<div class="container-outer">
  <div class="headline-block">
    <div class="headline-kw-box">
      <div class="headline-kw">KW {kw}</div>
      <div class="headline-period">{start_date.strftime('%d.%m.%Y')} – {(start_date + pd.Timedelta(days=6)).strftime('%d.%m.%Y')}</div>
      <div class="headline-name">{fahrer_name}</div>
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

        empty_day = (tour.strip() == "–" and uhrzeit.strip() == "–")
        empty_class = " is-empty" if empty_day else ""

        html += f"""
  <div class="{card_class}{empty_class}">
    <div class="header-row">
      <div class="prominent-date">{date_obj.strftime('%d.%m.%Y')}</div>

      <div class="pill pill-day" title="{weekday}">
        <div class="pill-label">TAG</div>
        <div class="pill-value">{weekday}</div>
      </div>
    </div>

    <div class="info">
      <div class="info-block">
        <span class="label">Tour / Aufgabe:</span>
        <span class="value">{tour}</span>
      </div>

      <div class="pill pill-time" title="{uhrzeit}">
        <div class="pill-label">UHR</div>
        <div class="pill-value">{uhrzeit}</div>
      </div>
    </div>
  </div>"""

    html += "</div></body></html>"
    return html

css_styles = """
:root{
  --bg:#f5f7fa;
  --text:#1d1d1f;
  --muted:#647083;

  --card:#ffffff;
  --line:#b4bcc9;

  --good:#4b7a52;
  --good-bg:#e9f6ef;
  --good-line:#cbe8d6;

  --weekend:#7a4e00;
  --weekend-bg:#fff2da;
  --weekend-line:#f0d39a;

  --shadow:0 2px 5px rgba(0,0,0,0.06);
  --shadow-hover:0 3px 10px rgba(0,0,0,0.10);

  --radius:12px;
}

*{box-sizing:border-box}

body{
  margin:0;
  padding:0;
  background:var(--bg);
  font-family:'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
  color:var(--text);
  font-size:14px;
}

.container-outer{
  max-width:500px;
  margin:18px auto;
  padding:0 12px;
}

.headline-block{
  text-align:center;
  margin-bottom:12px;
}

.headline-kw-box{
  background:#eef2f9;
  border-radius:12px;
  padding:8px 14px;
  border:2px solid #a8b4cc;
  box-shadow:0 2px 5px rgba(0,0,0,0.05);
}

.headline-kw{
  font-size:1.25rem;
  font-weight:800;
  color:#1b3a7a;
  margin-bottom:2px;
}

.headline-period{
  font-size:0.83rem;
  color:#3e567f;
}

.headline-name{
  font-size:0.93rem;
  font-weight:700;
  color:#1a3662;
  margin-top:2px;
}

/* Cards */
.daycard{
  background:var(--card);
  border-radius:var(--radius);
  padding:8px 10px;
  margin-bottom:10px;
  border:1.5px solid var(--line);
  box-shadow:var(--shadow);
  transition:box-shadow .2s;
}

.daycard:hover{
  box-shadow:var(--shadow-hover);
}

/* Wochenende */
.daycard.samstag,
.daycard.sonntag{
  background:#fff3cc;
  border:1.5px solid #e5aa00;
  box-shadow: inset 0 0 0 2px #ffd566, var(--shadow);
  overflow:hidden;
}

.daycard.samstag .header-row,
.daycard.sonntag .header-row{
  background:#ffedb0;
  padding:4px 6px;
  margin:-8px -10px 6px -10px;
  border-bottom:1px solid #e5aa00;
}

/* Header */
.header-row{
  display:flex;
  justify-content:space-between;
  align-items:center;
  gap:8px;
  flex-wrap:nowrap;
  font-weight:700;
  font-size:0.9rem;
  color:#2a2a2a;
  padding:2px 0;
  margin-bottom:6px;
}

.prominent-date{
  color:#bb4444;
  font-weight:800;
}

/* Info Bereich */
.info{
  display:flex;
  justify-content:space-between;
  align-items:stretch;
  flex-wrap:wrap;
  gap:8px;
  font-size:0.85rem;
}

/* Tour Box kompakt */
.info-block{
  flex:1 1 calc(100% - 104px); /* Platz für Pill */
  background:#f4f6fb;
  padding:6px 8px;
  border-radius:8px;
  border:1px solid #9ca7bc;
  display:flex;
  justify-content:space-between;
  align-items:center;
  gap:8px;
  min-height:34px;
}

.label{
  font-weight:700;
  color:#555;
  font-size:0.78rem;
  white-space:nowrap;
}

.value{
  font-weight:800;
  color:#222;
  font-size:0.84rem;
  text-align:right;
  overflow:hidden;
  text-overflow:ellipsis;
  white-space:nowrap;
}

/* ===== Pills kompakt (unterteilt) ===== */
.pill{
  width:96px;      /* alle gleich breit */
  height:34px;     /* kompakt */

  display:flex;
  flex-direction:column;
  align-items:center;
  justify-content:center;

  border-radius:999px;
  border:1px solid var(--line);
  background:#f5f7fb;

  text-align:center;
  line-height:1.0;
  user-select:none;
}

.pill-label{
  font-size:0.58rem;
  font-weight:900;
  letter-spacing:0.08em;
  color:var(--muted);
  opacity:0.95;
}

.pill-value{
  margin-top:1px;
  font-size:0.80rem;
  font-weight:950;
  color:var(--text);
  white-space:nowrap;
  overflow:hidden;
  text-overflow:ellipsis;
  max-width:86px;
}

/* Tag-Pill */
.pill-day{
  background:var(--good-bg);
  border-color:var(--good-line);
}

/* Wochenende Tag-Pill */
.daycard.samstag .pill-day,
.daycard.sonntag .pill-day{
  background:var(--weekend-bg);
  border-color:var(--weekend-line);
}

/* Zeit-Pill */
.pill-time{
  background:#f5f7fb;
  font-variant-numeric: tabular-nums;
}

/* optional: leere Tage dezenter */
.is-empty .value,
.is-empty .pill-value{
  color:#7d8796;
}

@media (max-width: 440px){
  .header-row{
    flex-wrap:wrap;
  }
  .info{
    flex-direction:column;
  }
  .info-block{
    flex:1 1 auto;
  }
}
"""

# Streamlit UI für Mehrfach-Upload
st.set_page_config(page_title="Touren-Export", layout="centered")
st.title("Dienstplan aktualisieren")

uploaded_files = st.file_uploader("Excel-Dateien hochladen (Blatt 'Touren')", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "gesamt_export.zip")
            with ZipFile(zip_path, "w") as zipf:

                ausschluss_stichwoerter = ["zippel", "insel", "paasch", "meyer", "ihde", "devies", "insellogistik"]

                for file in uploaded_files:
                    df = pd.read_excel(file, sheet_name="Touren", skiprows=4, engine="openpyxl")

                    fahrer_dict = {}
                    for _, row in df.iterrows():
                        datum = row.iloc[14]
                        tour = row.iloc[15]
                        uhrzeit = row.iloc[8]

                        if pd.isna(datum):
                            continue
                        try:
                            datum_dt = pd.to_datetime(datum)
                        except:
                            continue

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

                        try:
                            nachname, vorname = [s.strip() for s in fahrer_name.split(",")]
                        except ValueError:
                            nachname, vorname = fahrer_name.strip(), ""

                        sonder_dateien = {
                            ("fechner", "klaus"): "KFechner",
                            ("fechner", "danny"): "Fechner",
                            ("scheil", "rene"): "RScheil",
                            ("scheil", "eric"): "Scheil",
                            ("schulz", "julian"): "Schulz",
                            ("schulz", "stephan"): "STSchulz",
                            ("lewandowski", "kamil"): "Lewandowski",
                            ("lewandowski", "dominik"): "DLewandowski",
                        }

                        n_clean = nachname.strip().lower()
                        v_clean = vorname.strip().lower()
                        filename_part = sonder_dateien.get((n_clean, v_clean), nachname.replace(" ", "_"))
                        filename = f"KW{kw:02d}_{filename_part}.html"

                        html_code = generate_html(fahrer_name, wochen_eintraege, kw, start_sonntag, css_styles)

                        folder_name = f"KW{kw:02d}"
                        full_path = os.path.join(tmpdir, folder_name, filename)
                        os.makedirs(os.path.dirname(full_path), exist_ok=True)
                        with open(full_path, "w", encoding="utf-8") as f:
                            f.write(html_code)

                        filename_lower = filename.lower()
                        if "ch._holtz" in filename_lower or any(stichwort in filename_lower for stichwort in ausschluss_stichwoerter):
                            os.remove(full_path)
                            continue

                        zipf.write(full_path, arcname=os.path.join(folder_name, filename))

            with open(zip_path, "rb") as f:
                zip_bytes = f.read()

            if st.checkbox("Automatisch auf FTP hochladen", value=False):
                if not all([FTP_HOST, FTP_USER, FTP_PASS]):
                    st.warning("FTP-Zugangsdaten fehlen in .env")
                else:
                    st.info("Starte FTP-Upload...")
                    upload_folder_to_ftp_with_progress(tmpdir, FTP_BASE_DIR)

            st.success(f"{len(uploaded_files)} Dateien verarbeitet.")
            st.download_button("ZIP mit allen HTML-Dateien herunterladen", data=zip_bytes, file_name="gesamt_export.zip", mime="application/zip")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten: {e}")
