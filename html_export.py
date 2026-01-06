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

        html += f"""
  <div class="daycard">
    <div class="header-row">
      <div class="prominent-date">{date_obj.strftime('%d.%m.%Y')}</div>

      <div class="pill-row">
        <div class="pill pill-day">{weekday}</div>
        <div class="pill pill-time">{uhrzeit}</div>
        <div class="pill pill-tour">{tour}</div>
      </div>
    </div>
  </div>"""

    html += "</div></body></html>"
    return html

css_styles = """
:root{
  --bg:#f5f7fa;
  --card:#ffffff;
  --text:#1d1d1f;
  --muted:#5f6b7a;
  --line:#c9d1de;
  --shadow:0 2px 8px rgba(0,0,0,0.06);

  --pill-h:26px;
  --pill-pad-x:10px;

  /* sanfte Farben */
  --pill-day-bg:#edf6f1;
  --pill-day-bd:#cfe4da;

  --pill-time-bg:#eef2f7;
  --pill-time-bd:#d4dbe6;

  --pill-tour-bg:#e9edf3;
  --pill-tour-bd:#c7cfdb;
}

*{box-sizing:border-box}

body{
  margin:0;
  background:var(--bg);
  font-family:'Inter',-apple-system,BlinkMacSystemFont,sans-serif;
  color:var(--text);
  font-size:14px;
}

.container-outer{
  max-width:500px;
  margin:20px auto;
  padding:0 12px;
}

.headline-block{text-align:center;margin-bottom:14px;}

.headline-kw-box{
  background:#fff;
  border-radius:12px;
  padding:10px 14px;
  border:1px solid var(--line);
  box-shadow:var(--shadow);
}

.headline-kw{font-size:1.25rem;font-weight:800;color:#1b3a7a;}
.headline-period{font-size:.85rem;color:var(--muted);}
.headline-name{font-size:.95rem;font-weight:800;color:#1a3662;margin-top:4px;}

.daycard{
  background:var(--card);
  border-radius:12px;
  padding:8px 12px;
  margin-bottom:10px;
  border:1px solid var(--line);
  box-shadow:var(--shadow);
}

.header-row{
  display:flex;
  justify-content:space-between;
  gap:10px;
  flex-wrap:wrap;
}

.prominent-date{
  color:#bb4444;
  font-weight:800;
  white-space:nowrap;
}

/* Pills */
.pill-row{
  display:flex;
  gap:8px;
  flex-wrap:wrap;
  flex:1 1 260px;
  justify-content:flex-end;
}

.pill{
  height:var(--pill-h);
  border-radius:999px;
  padding:0 var(--pill-pad-x);
  display:flex;
  align-items:center;
  justify-content:center;
  font-size:.82rem;
  font-weight:650;
  white-space:nowrap;
  overflow:hidden;
  text-overflow:ellipsis;
}

/* leicht differenziert */
.pill-day{
  width:88px;
  background:var(--pill-day-bg);
  border:1px solid var(--pill-day-bd);
}

.pill-time{
  width:70px;
  background:var(--pill-time-bg);
  border:1px solid var(--pill-time-bd);
  font-variant-numeric:tabular-nums;
}

.pill-tour{
  flex:1 1 160px;
  min-width:140px;
  background:var(--pill-tour-bg);
  border:1px solid var(--pill-tour-bd);
}

@media(max-width:440px){
  .pill-row{justify-content:flex-start;}
  .pill-tour{flex:1 1 100%;min-width:100%;}
}
"""

# Streamlit UI
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

                ausschluss = ["zippel","insel","paasch","meyer","ihde","devies","insellogistik"]

                for file in uploaded_files:
                    df = pd.read_excel(file, sheet_name="Touren", skiprows=4)

                    fahrer_dict = {}
                    for _, row in df.iterrows():
                        datum = row.iloc[14]
                        tour = row.iloc[15]
                        uhrzeit = row.iloc[8]

                        if pd.isna(datum): continue
                        datum_dt = pd.to_datetime(datum, errors="coerce")
                        if pd.isna(datum_dt): continue

                        if pd.isna(uhrzeit):
                            uhrzeit_str="–"
                        elif isinstance(uhrzeit,datetime):
                            uhrzeit_str=uhrzeit.strftime("%H:%M")
                        else:
                            uhrzeit_str=str(uhrzeit)[:5]

                        eintrag=f"{uhrzeit_str} – {str(tour).strip()}"

                        for pos in [(3,4),(6,7)]:
                            n=row.iloc[pos[0]]
                            v=row.iloc[pos[1]]
                            if pd.notna(n) or pd.notna(v):
                                name=f"{str(n).title()}, {str(v).title()}"
                                fahrer_dict.setdefault(name,{}).setdefault(datum_dt.date(),[])
                                if eintrag not in fahrer_dict[name][datum_dt.date()]:
                                    fahrer_dict[name][datum_dt.date()].append(eintrag)

                    for name, tage in fahrer_dict.items():
                        start=min(tage)
                        start_sonntag=start-pd.Timedelta(days=(start.weekday()+1)%7)
                        kw=get_kw(start_sonntag)+1

                        eintraege=[]
                        for i in range(7):
                            d=start_sonntag+pd.Timedelta(days=i)
                            wt=wochentage_deutsch_map.get(d.strftime("%A"),d.strftime("%A"))
                            if d in tage:
                                for e in tage[d]:
                                    eintraege.append(f"{d.strftime('%d.%m.%Y')} ({wt}): {e}")
                            else:
                                eintraege.append(f"{d.strftime('%d.%m.%Y')} ({wt}): –")

                        filename=f"KW{kw:02d}_{name.split(',')[0]}.html"
                        if any(x in filename.lower() for x in ausschluss): continue

                        html=generate_html(name,eintraege,kw,start_sonntag,css_styles)
                        path=os.path.join(tmpdir,f"KW{kw:02d}",filename)
                        os.makedirs(os.path.dirname(path),exist_ok=True)
                        open(path,"w",encoding="utf-8").write(html)
                        zipf.write(path,arcname=os.path.join(f"KW{kw:02d}",filename))

            st.download_button(
                "ZIP mit allen HTML-Dateien herunterladen",
                open(zip_path,"rb").read(),
                "gesamt_export.zip",
                "application/zip"
            )
    except Exception as e:
        st.error(f"Fehler: {e}")
