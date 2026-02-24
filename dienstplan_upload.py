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
<html lang=\"de\">
<head>
  <meta charset=\"UTF-8\">
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">
  <title>KW{kw} – {fahrer_name}</title>
  <style>{css_styles}</style>
</head>
<body>
<div class=\"container-outer\">
  <div class=\"headline-block\">
    <div class=\"headline-kw-box\">
      <div class=\"headline-kw\">KW {kw}</div>
      <div class=\"headline-period\">{start_date.strftime('%d.%m.%Y')} – {(start_date + pd.Timedelta(days=6)).strftime('%d.%m.%Y')}</div>
      <div class=\"headline-name\">{fahrer_name}</div>
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
  <div class=\"{card_class}\">
    <div class=\"header-row\">
      <div class=\"prominent-date\">{date_obj.strftime('%d.%m.%Y')}</div>
      <div class=\"weekday\">{weekday}</div>
    </div>
    <div class=\"info\">
      <div class=\"info-block\">
        <span class=\"label\">Tour / Aufgabe:</span>
        <span class=\"value\">{tour}</span>
      </div>
      <div class=\"info-block\">
        <span class=\"label\">Uhrzeit:</span>
        <span class=\"value\">{uhrzeit}</span>
      </div>
    </div>
  </div>"""

    html += "</div></body></html>"
    return html

css_styles = """
body {
  margin: 0;
  padding: 0;
  background: #f5f7fa;
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
  color: #1d1d1f;
  font-size: 14px;
}

.container-outer {
  max-width: 500px;
  margin: 20px auto;
  padding: 0 12px;
}

.headline-block {
  text-align: center;
  margin-bottom: 16px;
}

.headline-kw-box {
  background: #eef2f9;
  border-radius: 12px;
  padding: 8px 14px;
  border: 2px solid #a8b4cc;
  box-shadow: 0 2px 5px rgba(0,0,0,0.05);
}

.headline-kw {
  font-size: 1.3rem;
  font-weight: 700;
  color: #1b3a7a;
  margin-bottom: 2px;
}

.headline-period {
  font-size: 0.85rem;
  color: #3e567f;
}

.headline-name {
  font-size: 0.95rem;
  font-weight: 600;
  color: #1a3662;
  margin-top: 2px;
}

.daycard {
  background: #ffffff;
  border-radius: 12px;
  padding: 8px 12px;
  margin-bottom: 12px;
  border: 1.5px solid #b4bcc9;
  box-shadow: 0 2px 5px rgba(0,0,0,0.06);
  transition: box-shadow 0.2s;
}

.daycard:hover {
  box-shadow: 0 3px 10px rgba(0,0,0,0.1);
}

.daycard.samstag,
.daycard.sonntag {
  background: #fff3cc;
  border: 1.5px solid #e5aa00;
  box-shadow: inset 0 0 0 3px #ffd566, 0 3px 8px rgba(0, 0, 0, 0.06);
  border-radius: 12px;
  overflow: hidden;
}

.daycard.samstag .header-row,
.daycard.sonntag .header-row {
  background: #ffedb0;
  padding: 4px 0;
  margin-bottom: 6px;
  border-bottom: 1px solid #e5aa00;
}

.daycard.samstag .prominent-date,
.daycard.sonntag .prominent-date {
  color: #8c5a00;
  font-weight: 700;
}

.daycard.samstag .weekday,
.daycard.sonntag .weekday {
  color: #7a4e00;
  font-weight: 700;
}

.header-row {
  display: flex;
  justify-content: space-between;
  align-items: center;
  flex-wrap: nowrap;
  font-weight: 600;
  font-size: 0.9rem;
  color: #2a2a2a;
  padding: 4px 0;
  margin-bottom: 6px;
}

.weekday {
  color: #5e8f64;
  font-weight: 600;
  margin-left: 8px;
}

.prominent-date {
  color: #bb4444;
  font-weight: 600;
}

.info {
  display: flex;
  justify-content: space-between;
  flex-wrap: wrap;
  gap: 8px;
  font-size: 0.85rem;
  padding-top: 4px;
}

.info-block {
  flex: 1 1 48%;
  background: #f4f6fb;
  padding: 4px 6px;
  border-radius: 6px;
  border: 1px solid #9ca7bc;
  display: flex;
  justify-content: space-between;
  align-items: center;
  flex-direction: row;
  gap: 6px;
}

.label {
  font-weight: 600;
  color: #555;
  font-size: 0.8rem;
  margin-bottom: 0;
}

.value {
  font-weight: 600;
  color: #222;
  font-size: 0.85rem;
}

@media (max-width: 440px) {
  .header-row {
    flex-direction: row;
    flex-wrap: wrap;
    gap: 4px;
  }
  .info {
    flex-direction: column;
  }
}
"""

# Der restliche Code (Excel-Verarbeitung + generate_html-Aufruf) bleibt wie gehabt und verwendet jetzt das neue Design.


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
