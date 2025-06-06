import streamlit as st
from datetime import datetime
from io import BytesIO
import pandas as pd
import zipfile
from openpyxl import load_workbook
from pathlib import Path
import tempfile

# Funktion: Kalenderwoche
def get_calendar_week(date):
    return date.isocalendar()[1]

# Funktion: sichere Dateinamen
def clean_filename(name):
    return "".join(c for c in name if c.isalnum() or c in "._- ").strip().replace(" ", "_")

# HTML-Exportfunktion
def create_html_content(name, kw):
    return f"""<!DOCTYPE html>
<html lang='de'>
<head>
  <meta charset='UTF-8'>
  <meta name='viewport' content='width=device-width, initial-scale=1.0'>
  <title>KW{kw} – {name}</title>
  <link href='https://fonts.googleapis.com/css2?family=Inter:wght@500;700&display=swap' rel='stylesheet'>
  <style>
    body {{
      font-family: 'Inter', Arial, sans-serif;
      background: #f5f8fb;
      margin: 0;
      color: #23262f;
      min-height: 100vh;
    }}
    .container-outer {{
      max-width: 540px;
      margin: 26px auto 36px auto;
      background: linear-gradient(112deg, #eaf1ff 0%, #f9fafd 100%);
      border-radius: 28px;
      box-shadow: 0 2px 6px rgba(0,0,0,0.08);
      padding: 20px;
    }}
  </style>
</head>
<body>
  <div class="container-outer">
    <h1>KW{kw} – {name}</h1>
    <p>Hier steht dein Plan.</p>
  </div>
</body>
</html>"""

# Hauptverarbeitung
def process_excel_to_zip(excel_bytes):
    in_memory_zip = BytesIO()
    with tempfile.TemporaryDirectory() as tmpdirname:
        wb = load_workbook(filename=BytesIO(excel_bytes), data_only=True)
        ws = wb["Druck Fahrer"]

        datum = ws["E2"].value
        if not isinstance(datum, datetime):
            datum = pd.to_datetime(datum)

        kw = get_calendar_week(datum)
        name_dict = {}

        for i in range(11, ws.max_row + 1):
            name_cell = ws[f"B{i}"].value
            if name_cell:
                base_name = name_cell.strip()
                count = name_dict.get(base_name, 0)
                name_dict[base_name] = count + 1
                final_name = base_name if count == 0 else f"{base_name}_{count}"
                safe_name = clean_filename(final_name)
                html = create_html_content(final_name, kw)

                html_path = Path(tmpdirname) / f"KW{kw}_{safe_name}.html"
                with open(html_path, "w", encoding="utf-8") as f:
                    f.write(html)

        with zipfile.ZipFile(in_memory_zip, "w", zipfile.ZIP_DEFLATED) as zipf:
            for html_file in Path(tmpdirname).glob("*.html"):
                zipf.write(html_file, html_file.name)

    in_memory_zip.seek(0)
    return in_memory_zip

# Streamlit Interface
st.title("Excel zu HTML Exporter (wie VBA)")
uploaded_file = st.file_uploader("Excel-Datei mit 'Druck Fahrer' hochladen", type=["xlsx", "xlsm"])

if uploaded_file:
    zip_bytes = process_excel_to_zip(uploaded_file.read())
    st.success("Export erfolgreich.")
    st.download_button(
        label="ZIP-Datei herunterladen",
        data=zip_bytes,
        file_name="html_export.zip",
        mime="application/zip"
    )
