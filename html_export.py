# Erweiterungen oben ergänzen
import streamlit as st
import pandas as pd
from zipfile import ZipFile
from datetime import datetime
import tempfile
import os
from ftplib import FTP
from dotenv import load_dotenv

# .env laden
load_dotenv()
FTP_HOST = os.getenv("FTP_HOST")
FTP_USER = os.getenv("FTP_USER")
FTP_PASS = os.getenv("FTP_PASS")
FTP_BASE_DIR = os.getenv("FTP_BASE_DIR", "/")

# Deutsche Wochentage
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

    progress = st.progress(0.0)
    info = st.empty()

    for i, (lp, rp) in enumerate(all_files, start=1):
        rdir = os.path.dirname(rp).replace("\\", "/")
        path = ""
        for part in rdir.split("/"):
            if part:
                path += "/" + part
                try:
                    ftp.mkd(path)
                except:
                    pass

        with open(lp, "rb") as f:
            ftp.cwd(rdir)
            ftp.storbinary(f"STOR {os.path.basename(lp)}", f)

        progress.progress(i / len(all_files))
        info.info(f"Hochgeladen {i}/{len(all_files)} – {os.path.basename(lp)}")

    ftp.quit()
    info.success("FTP-Upload abgeschlossen")

# ==========================
# HTML Generator (clean)
# ==========================
def generate_html(fahrer_name, eintraege, kw, start_date, css):
    html = f"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>KW{kw} – {fahrer_name}</title>
<style>{css}</style>
</head>
<body>
<div class="container">

<div class="header">
  <div class="kw">KW {kw}</div>
  <div class="period">{start_date.strftime('%d.%m.%Y')} – {(start_date + pd.Timedelta(days=6)).strftime('%d.%m.%Y')}</div>
  <div class="name">{fahrer_name}</div>
</div>
"""

    for eintrag in eintraege:
        date_text, content = eintrag.split(": ", 1)
        date_obj = pd.to_datetime(date_text.split(" ")[0], format="%d.%m.%Y")
        weekday = date_text.split("(")[-1].replace(")", "")

        if "–" in content:
            time, tour = [x.strip() for x in content.split("–", 1)]
        else:
            time, tour = "–", content.strip()

        card_class = "card"
        if weekday in ("Samstag", "Sonntag"):
            card_class += " weekend"

        html += f"""
<div class="{card_class}">
  <div class="row">
    <div class="date">{date_obj.strftime('%d.%m.%Y')}</div>

    <div class="pills">
      <div class="pill">{weekday}</div>
      <div class="pill pill-time">{time}</div>
      <div class="pill pill-tour" title="{tour}">{tour}</div>
    </div>
  </div>
</div>
"""

    html += "</div></body></html>"
    return html

# ==========================
# CSS – clean & ruhig
# ==========================
css_styles = """
:root{
  --bg:#f5f7fa;
  --card:#ffffff;
  --line:#d3d9e3;
  --text:#1d1d1f;
  --muted:#6b7280;
  --accent:#1b66b3;
  --shadow:0 2px 10px rgba(0,0,0,0.06);

  --pill-h:26px;
  --pill-w:92px;
  --pill-w-tour:160px;
}

*{box-sizing:border-box}

body{
  margin:0;
  background:var(--bg);
  font-family:Inter,system-ui,-apple-system,Segoe UI,sans-serif;
  color:var(--text);
  font-size:14px;
}

.container{
  max-width:560px;
  margin:18px auto;
  padding:0 12px;
}

/* Header */
.header{
  text-align:center;
  background:#fff;
  border:1px solid var(--line);
  border-radius:14px;
  padding:10px 14px;
  box-shadow:var(--shadow);
  margin-bottom:12px;
}

.kw{
  font-size:1.2rem;
  font-weight:800;
  color:#173a7a;
}

.period{
  font-size:.85rem;
  color:var(--muted);
}

.name{
  font-size:.95rem;
  font-weight:800;
  color:var(--accent);
  margin-top:4px;
}

/* Cards */
.card{
  background:var(--card);
  border:1px solid var(--line);
  border-radius:14px;
  padding:10px 12px;
  margin-bottom:8px;
  box-shadow:var(--shadow);
}

.card.weekend{
  border-color:#e6b657;
}

/* Row */
.row{
  display:flex;
  justify-content:space-between;
  align-items:center;
  gap:10px;
}

.date{
  font-weight:800;
  color:#bb4444;
  white-space:nowrap;
}

/* Pills */
.pills{
  display:flex;
  gap:8px;
  align-items:center;
}

.pill{
  height:var(--pill-h);
  width:var(--pill-w);
  display:flex;
  align-items:center;
  justify-content:center;
  border-radius:999px;
  border:1px solid var(--line);
  background:#f1f4f9;
  font-weight:800;
  font-size:.82rem;
  white-space:nowrap;
  overflow:hidden;
  text-overflow:ellipsis;
}

.pill-time{
  font-variant-numeric: tabular-nums;
}

.pill-tour{
  width:var(--pill-w-tour);
}

/* Mobile */
@media (max-width:480px){
  .row{
    flex-wrap:wrap;
  }
  .pills{
    width:100%;
    justify-content:flex-end;
  }
}
"""

# ==========================
# Streamlit UI
# ==========================
st.set_page_config(page_title="Touren-Export", layout="centered")
st.title("Dienstplan aktualisieren")

uploaded_files = st.file_uploader(
    "Excel-Dateien hochladen (Blatt 'Touren')",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "gesamt_export.zip")
            with ZipFile(zip_path, "w") as zipf:

                for file in uploaded_files:
                    df = pd.read_excel(file, sheet_name="Touren", skiprows=4)

                    fahrer = {}
                    for _, row in df.iterrows():
                        if pd.isna(row.iloc[14]):
                            continue
                        d = pd.to_datetime(row.iloc[14]).date()
                        t = str(row.iloc[15]).strip()
                        u = "–" if pd.isna(row.iloc[8]) else str(row.iloc[8])[:5]
                        e = f"{u} – {t}"

                        for pos in [(3,4),(6,7)]:
                            n = str(row.iloc[pos[0]]).strip().title() if pd.notna(row.iloc[pos[0]]) else ""
                            v = str(row.iloc[pos[1]]).strip().title() if pd.notna(row.iloc[pos[1]]) else ""
                            if n:
                                name = f"{n}, {v}"
                                fahrer.setdefault(name, {}).setdefault(d, []).append(e)

                    for name, days in fahrer.items():
                        start = min(days.keys())
                        sunday = start - pd.Timedelta(days=(start.weekday()+1)%7)
                        kw = get_kw(sunday) + 1

                        entries = []
                        for i in range(7):
                            day = sunday + pd.Timedelta(days=i)
                            wd = wochentage_deutsch_map[day.strftime("%A")]
                            if day in days:
                                for e in days[day]:
                                    entries.append(f"{day.strftime('%d.%m.%Y')} ({wd}): {e}")
                            else:
                                entries.append(f"{day.strftime('%d.%m.%Y')} ({wd}): –")

                        html = generate_html(name, entries, kw, sunday, css_styles)
                        fname = f"KW{kw:02d}_{name.split(',')[0]}.html"
                        path = os.path.join(tmpdir, fname)

                        with open(path, "w", encoding="utf-8") as f:
                            f.write(html)

                        zipf.write(path, fname)

            with open(zip_path, "rb") as f:
                st.download_button(
                    "ZIP mit HTML-Dateien herunterladen",
                    f.read(),
                    file_name="gesamt_export.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"Fehler: {e}")
