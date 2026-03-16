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
    "Monday": "Montag",
    "Tuesday": "Dienstag",
    "Wednesday": "Mittwoch",
    "Thursday": "Donnerstag",
    "Friday": "Freitag",
    "Saturday": "Samstag",
    "Sunday": "Sonntag"
}

def get_kw(datum):
    return datum.isocalendar()[1]

def ensure_ftp_dirs(ftp, remote_dir):
    remote_dir = remote_dir.replace("\\", "/")
    parts = remote_dir.split("/")
    path_built = ""

    for part in parts:
        if part:
            path_built += "/" + part
            try:
                ftp.mkd(path_built)
            except:
                pass

def upload_folder_to_ftp_with_progress(local_dir, ftp_dir):
    ftp = FTP()
    ftp.connect(FTP_HOST, 21)
    ftp.login(FTP_USER, FTP_PASS)

    all_files = []
    for root, _, files in os.walk(local_dir):
        for file in files:
            local_path = os.path.join(root, file)
            rel_path = os.path.relpath(local_path, local_dir).replace("\\", "/")
            remote_path = os.path.join(ftp_dir, rel_path).replace("\\", "/")
            all_files.append((local_path, remote_path))

    total = len(all_files)
    uploaded = 0

    progress_bar = st.progress(0)
    status_text = st.empty()

    for local_path, remote_path in all_files:
        remote_dir = os.path.dirname(remote_path).replace("\\", "/")
        ensure_ftp_dirs(ftp, remote_dir)

        with open(local_path, "rb") as f:
            ftp.cwd(remote_dir)
            ftp.storbinary(f"STOR {os.path.basename(local_path)}", f)

        uploaded += 1
        progress = uploaded / total if total > 0 else 1
        progress_bar.progress(progress)
        status_text.info(f"Hochgeladen: {uploaded}/{total} – {os.path.basename(local_path)}")

    ftp.quit()
    status_text.success("Alle Dateien erfolgreich hochgeladen.")

def normalize_driver_name(nachname, vorname):
    nachname = str(nachname).strip().title() if pd.notna(nachname) else ""
    vorname = str(vorname).strip().title() if pd.notna(vorname) else ""

    if not nachname and not vorname:
        return ""

    return f"{nachname}, {vorname}".strip().strip(",")

def parse_uhrzeit(uhrzeit):
    if pd.isna(uhrzeit):
        return "–"
    elif isinstance(uhrzeit, (int, float)) and uhrzeit == 0:
        return "00:00"
    elif isinstance(uhrzeit, datetime):
        return uhrzeit.strftime("%H:%M")
    else:
        try:
            uhrzeit_parsed = pd.to_datetime(uhrzeit)
            return uhrzeit_parsed.strftime("%H:%M")
        except:
            uhrzeit_str = str(uhrzeit).strip()
            if not uhrzeit_str or uhrzeit_str.lower() == "nan":
                return "–"
            if ":" in uhrzeit_str:
                return ":".join(uhrzeit_str.split(":")[:2])
            return uhrzeit_str

def parse_tour(tour):
    if pd.isna(tour):
        return "–"

    tour_str = str(tour).strip()
    if not tour_str or tour_str.lower() == "nan":
        return "–"

    return tour_str

def generate_html(fahrer_name, eintraege, kw, start_date, css_styles):
    html = f"""<!DOCTYPE html>
<html lang="de">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>KW{kw:02d} – {fahrer_name}</title>
  <style>{css_styles}</style>
</head>
<body>
<div class="container-outer">
  <div class="headline-block">
    <div class="headline-kw-box">
      <div class="headline-kw">KW {kw:02d}</div>
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

        if not uhrzeit:
            uhrzeit = "–"
        if not tour:
            tour = "–"

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
    </div>
    <div class="info">
      <div class="info-block">
        <span class="label">Tour / Aufgabe:</span>
        <span class="value">{tour}</span>
      </div>
      <div class="info-block">
        <span class="label">Uhrzeit:</span>
        <span class="value">{uhrzeit}</span>
      </div>
    </div>
  </div>"""

    html += """
</div>
</body>
</html>"""
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

st.set_page_config(page_title="Touren-Export", layout="centered")
st.title("Dienstplan aktualisieren")

uploaded_files = st.file_uploader(
    "Excel-Dateien hochladen (Blatt 'Touren' und 'a Fahrer')",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "gesamt_export.zip")

            with ZipFile(zip_path, "w") as zipf:
                ausschluss_stichwoerter = [
                    "zippel",
                    "insel",
                    "paasch",
                    "meyer",
                    "ihde",
                    "devies",
                    "insellogistik"
                ]

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

                erzeugte_dateien = 0

                for file in uploaded_files:
                    # Blatt "a Fahrer" laden: Spalte B = Nachname, Spalte C = Vorname
                    fahrer_df = pd.read_excel(file, sheet_name="a Fahrer", engine="openpyxl")

                    # Blatt "Touren" laden
                    touren_df = pd.read_excel(file, sheet_name="Touren", skiprows=4, engine="openpyxl")

                    # Alle Fahrer zuerst aus "a Fahrer" sammeln
                    fahrer_dict = {}
                    for _, r in fahrer_df.iterrows():
                        fahrer_name = normalize_driver_name(r.iloc[1], r.iloc[2])
                        if fahrer_name:
                            fahrer_dict[fahrer_name] = {}

                    # Alle gültigen Daten sammeln für Fallback-Woche
                    alle_gueltigen_daten = []

                    # Touren aus Blatt "Touren" einlesen
                    for _, row in touren_df.iterrows():
                        datum = row.iloc[14]
                        tour = row.iloc[15]
                        uhrzeit = row.iloc[8]

                        datum_dt = None
                        if pd.notna(datum):
                            try:
                                datum_dt = pd.to_datetime(datum)
                                alle_gueltigen_daten.append(datum_dt)
                            except:
                                datum_dt = None

                        uhrzeit_str = parse_uhrzeit(uhrzeit)
                        tour_str = parse_tour(tour)
                        eintrag_text = f"{uhrzeit_str} – {tour_str}"

                        for pos in [(3, 4), (6, 7)]:
                            fahrer_name = normalize_driver_name(row.iloc[pos[0]], row.iloc[pos[1]])
                            if not fahrer_name:
                                continue

                            # Falls im Tourenblatt ein Fahrer steht, der in "a Fahrer" fehlt:
                            if fahrer_name not in fahrer_dict:
                                fahrer_dict[fahrer_name] = {}

                            if datum_dt is not None:
                                tag = datum_dt.date()
                                if tag not in fahrer_dict[fahrer_name]:
                                    fahrer_dict[fahrer_name][tag] = []

                                if eintrag_text not in fahrer_dict[fahrer_name][tag]:
                                    fahrer_dict[fahrer_name][tag].append(eintrag_text)

                    # Fallback-Woche bestimmen
                    if alle_gueltigen_daten:
                        global_start_datum = min(alle_gueltigen_daten).date()
                    else:
                        global_start_datum = pd.Timestamp.today().date()

                    global_start_sonntag = global_start_datum - pd.Timedelta(days=(global_start_datum.weekday() + 1) % 7)
                    global_kw = get_kw(global_start_sonntag) + 1

                    # Alphabetisch sortieren
                    fahrer_dict = dict(sorted(fahrer_dict.items(), key=lambda x: x[0].lower()))

                    for fahrer_name, eintraege in fahrer_dict.items():
                        if eintraege:
                            start_datum = min(eintraege.keys())
                            start_sonntag = start_datum - pd.Timedelta(days=(start_datum.weekday() + 1) % 7)
                            kw = get_kw(start_sonntag) + 1
                        else:
                            start_sonntag = global_start_sonntag
                            kw = global_kw

                        wochen_eintraege = []
                        for i in range(7):
                            tag_datum = start_sonntag + pd.Timedelta(days=i)
                            wochentag = wochentage_deutsch_map.get(
                                tag_datum.strftime("%A"),
                                tag_datum.strftime("%A")
                            )

                            if tag_datum in eintraege and len(eintraege[tag_datum]) > 0:
                                for eintrag in eintraege[tag_datum]:
                                    wochen_eintraege.append(
                                        f"{tag_datum.strftime('%d.%m.%Y')} ({wochentag}): {eintrag}"
                                    )
                            else:
                                wochen_eintraege.append(
                                    f"{tag_datum.strftime('%d.%m.%Y')} ({wochentag}): –"
                                )

                        try:
                            nachname, vorname = [s.strip() for s in fahrer_name.split(",", 1)]
                        except ValueError:
                            nachname, vorname = fahrer_name.strip(), ""

                        n_clean = nachname.lower()
                        v_clean = vorname.lower()

                        filename_part = sonder_dateien.get(
                            (n_clean, v_clean),
                            nachname.replace(" ", "_")
                        )

                        filename = f"KW{kw:02d}_{filename_part}.html"
                        filename_lower = filename.lower()

                        if "ch._holtz" in filename_lower or any(
                            stichwort in filename_lower for stichwort in ausschluss_stichwoerter
                        ):
                            continue

                        html_code = generate_html(
                            fahrer_name=fahrer_name,
                            eintraege=wochen_eintraege,
                            kw=kw,
                            start_date=start_sonntag,
                            css_styles=css_styles
                        )

                        folder_name = f"KW{kw:02d}"
                        full_path = os.path.join(tmpdir, folder_name, filename)
                        os.makedirs(os.path.dirname(full_path), exist_ok=True)

                        with open(full_path, "w", encoding="utf-8") as f:
                            f.write(html_code)

                        zipf.write(full_path, arcname=os.path.join(folder_name, filename))
                        erzeugte_dateien += 1

            with open(zip_path, "rb") as f:
                zip_bytes = f.read()

            if st.checkbox("Automatisch auf FTP hochladen", value=False):
                if not all([FTP_HOST, FTP_USER, FTP_PASS]):
                    st.warning("FTP-Zugangsdaten fehlen in .env")
                else:
                    st.info("Starte FTP-Upload ...")
                    upload_folder_to_ftp_with_progress(tmpdir, FTP_BASE_DIR)

            st.success(
                f"{len(uploaded_files)} Datei(en) verarbeitet. "
                f"{erzeugte_dateien} HTML-Datei(en) erstellt."
            )

            st.download_button(
                "ZIP mit allen HTML-Dateien herunterladen",
                data=zip_bytes,
                file_name="gesamt_export.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten: {e}")
