# ===============================
# Imports & Setup
# ===============================
import streamlit as st
import pandas as pd
from zipfile import ZipFile
from datetime import datetime
import tempfile
import os
from ftplib import FTP
from dotenv import load_dotenv

# ===============================
# .env laden (FTP)
# ===============================
load_dotenv()
FTP_HOST = os.getenv("FTP_HOST")
FTP_USER = os.getenv("FTP_USER")
FTP_PASS = os.getenv("FTP_PASS")
FTP_BASE_DIR = os.getenv("FTP_BASE_DIR", "/")

# ===============================
# Hilfsfunktionen
# ===============================
wochentage_deutsch_map = {
    "Monday": "Montag", "Tuesday": "Dienstag", "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag", "Friday": "Freitag",
    "Saturday": "Samstag", "Sunday": "Sonntag"
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

    progress_bar = st.progress(0.0)
    status = st.empty()

    for i, (local_path, remote_path) in enumerate(all_files, start=1):
        remote_dir = os.path.dirname(remote_path).replace("\\", "/")
        path = ""
        for part in remote_dir.split("/"):
            if part:
                path += "/" + part
                try:
                    ftp.mkd(path)
                except:
                    pass

        ftp.cwd(remote_dir)
        with open(local_path, "rb") as f:
            ftp.storbinary(f"STOR {os.path.basename(local_path)}", f)

        progress_bar.progress(i / len(all_files))
        status.info(f"Hochgeladen {i}/{len(all_files)}: {os.path.basename(local_path)}")

    ftp.quit()
    status.success("FTP-Upload abgeschlossen")

# ===============================
# HTML Generator
# ===============================
def generate_html(fahrer_name, eintraege, kw, start_date, css):
    html = f"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>KW {kw} – {fahrer_name}</title>
<style>{css}</style>
</head>
<body>
<div class="container-outer">

<div class="headline-block">
  <div class="headline-kw-box">
    <div class="headline-top">
      <div class="headline-kw">KW {kw}</div>
      <div class="headline-period">
        {start_date.strftime('%d.%m.%Y')} – {(start_date + pd.Timedelta(days=6)).strftime('%d.%m.%Y')}
      </div>
    </div>
    <div class="headline-name">{fahrer_name}</div>
  </div>
</div>
"""

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

        empty = (tour == "–" and uhrzeit == "–")
        empty_class = " is-empty" if empty else ""

        html += f"""
<div class="{card_class}{empty_class}">
  <div class="header-row">
    <div class="prominent-date">{date_obj.strftime('%d.%m.%Y')}</div>
    <div class="weekday-badge">{weekday}</div>
  </div>

  <div class="info">
    <div>
      <div class="tour-title">{tour}</div>
      <div class="tour-sub">Tour / Aufgabe</div>
    </div>
    <div class="chip">🕒 {uhrzeit}</div>
  </div>
</div>
"""

    html += """
</div>
</body>
</html>
"""
    return html

# ===============================
# Helles CSS Theme (FINAL)
# ===============================
css_styles = """
:root{
  --bg:#f2f4f8;
  --card:#ffffff;
  --line:#d7dbe2;
  --text:#1a1d23;
  --muted:#5f6b7a;
  --accent:#1b66b3;
  --good:#1f8a4c;
  --weekend:#b07200;
  --shadow:0 4px 14px rgba(0,0,0,.08);
  --radius:18px;
}

*{box-sizing:border-box}

body{
  margin:0;
  background:linear-gradient(180deg,#eef2f7,#f7f9fc);
  font-family:Inter,system-ui,sans-serif;
  color:var(--text);
  font-size:14px;
}

.container-outer{
  max-width:560px;
  margin:18px auto 28px;
  padding:0 14px;
}

.headline-block{
  position:sticky;
  top:10px;
  z-index:50;
  margin-bottom:14px;
}

.headline-kw-box{
  background:#fff;
  border:1px solid var(--line);
  border-radius:var(--radius);
  padding:14px 16px;
  box-shadow:var(--shadow);
}

.headline-top{
  display:flex;
  justify-content:space-between;
  gap:10px;
}

.headline-kw{
  font-weight:800;
  font-size:1.15rem;
}

.headline-period{
  color:var(--muted);
  font-weight:600;
  font-size:.85rem;
}

.headline-name{
  margin-top:8px;
  font-weight:700;
  color:var(--accent);
}

.daycard{
  background:var(--card);
  border:1px solid var(--line);
  border-radius:var(--radius);
  box-shadow:var(--shadow);
  padding:12px;
  margin-bottom:12px;
}

.header-row{
  display:flex;
  justify-content:space-between;
  margin-bottom:10px;
}

.prominent-date{
  font-weight:800;
}

.weekday-badge{
  padding:5px 10px;
  border-radius:999px;
  font-size:.78rem;
  font-weight:800;
  background:#e9f6ef;
  color:var(--good);
  border:1px solid #cbe8d6;
}

.daycard.samstag .weekday-badge,
.daycard.sonntag .weekday-badge{
  background:#fff2da;
  color:var(--weekend);
  border-color:#f0d39a;
}

.info{
  display:grid;
  grid-template-columns:1fr auto;
  gap:10px;
}

.tour-title{
  font-weight:800;
}

.tour-sub{
  font-size:.82rem;
  color:var(--muted);
}

.chip{
  padding:7px 10px;
  border-radius:999px;
  border:1px solid var(--line);
  background:#f5f7fb;
  font-weight:800;
}

.is-empty .tour-title{color:var(--muted)}
.is-empty .chip{opacity:.7}

@media(max-width:420px){
  .info{grid-template-columns:1fr}
}
"""

# ===============================
# Streamlit UI
# ===============================
st.set_page_config(page_title="Dienstplan Export", layout="centered")
st.title("Dienstplan aktualisieren")

uploaded_files = st.file_uploader(
    "Excel-Dateien hochladen (Blatt: Touren)",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "dienstplan_export.zip")
            with ZipFile(zip_path, "w") as zipf:

                for file in uploaded_files:
                    df = pd.read_excel(file, sheet_name="Touren", skiprows=4)

                    fahrer_dict = {}

                    for _, row in df.iterrows():
                        datum = row.iloc[14]
                        tour = row.iloc[15]
                        uhrzeit = row.iloc[8]

                        if pd.isna(datum):
                            continue

                        datum_dt = pd.to_datetime(datum)
                        uhrzeit_str = "–" if pd.isna(uhrzeit) else str(uhrzeit)[:5]

                        eintrag = f"{uhrzeit_str} – {str(tour).strip()}"

                        for pos in [(3,4),(6,7)]:
                            nn = str(row.iloc[pos[0]]).strip().title() if pd.notna(row.iloc[pos[0]]) else ""
                            vn = str(row.iloc[pos[1]]).strip().title() if pd.notna(row.iloc[pos[1]]) else ""
                            if nn:
                                name = f"{nn}, {vn}"
                                fahrer_dict.setdefault(name, {}).setdefault(datum_dt.date(), []).append(eintrag)

                    for fahrer, daten in fahrer_dict.items():
                        start = min(daten.keys())
                        sonntag = start - pd.Timedelta(days=(start.weekday()+1)%7)
                        kw = get_kw(sonntag) + 1

                        wochen_eintraege = []
                        for i in range(7):
                            tag = sonntag + pd.Timedelta(days=i)
                            wd = wochentage_deutsch_map[tag.strftime("%A")]
                            if tag in daten:
                                for e in daten[tag]:
                                    wochen_eintraege.append(f"{tag.strftime('%d.%m.%Y')} ({wd}): {e}")
                            else:
                                wochen_eintraege.append(f"{tag.strftime('%d.%m.%Y')} ({wd}): –")

                        html = generate_html(fahrer, wochen_eintraege, kw, sonntag, css_styles)
                        fname = f"KW{kw:02d}_{fahrer.split(',')[0]}.html"
                        path = os.path.join(tmpdir, fname)

                        with open(path, "w", encoding="utf-8") as f:
                            f.write(html)

                        zipf.write(path, fname)

            with open(zip_path, "rb") as f:
                st.download_button(
                    "ZIP mit HTML-Dateien herunterladen",
                    f.read(),
                    file_name="dienstplan_export.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"Fehler: {e}")
