import streamlit as st
import pandas as pd
from zipfile import ZipFile
from datetime import datetime
import tempfile
import os
import ftplib
from ftplib import FTP, FTP_TLS, error_perm
from dotenv import load_dotenv

# .env laden
load_dotenv()
FTP_HOST = os.getenv("FTP_HOST")
FTP_USER = os.getenv("FTP_USER")
FTP_PASS = os.getenv("FTP_PASS")
FTP_BASE_DIR = os.getenv("FTP_BASE_DIR", "/")
FTP_USE_TLS = os.getenv("FTP_TLS", "0") == "1"

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

def ensure_ftp_dirs(ftp, remote_dir, created_cache):
    remote_dir = remote_dir.replace("\\", "/")
    if remote_dir in created_cache:
        return
    path_built = ""
    for part in remote_dir.split("/"):
        if not part:
            continue
        path_built += "/" + part
        if path_built in created_cache:
            continue
        try:
            ftp.mkd(path_built)
        except error_perm:
            # existiert schon (550) -> ok; alles andere ist ein echter Fehler
            pass
        created_cache.add(path_built)
    created_cache.add(remote_dir)

def upload_folder_to_ftp_with_progress(local_dir, ftp_dir):
    # Dateien sammeln
    all_files = []
    for root, _, files in os.walk(local_dir):
        for file in files:
            local_path = os.path.join(root, file)
            rel_path = os.path.relpath(local_path, local_dir).replace("\\", "/")
            remote_path = os.path.join(ftp_dir, rel_path).replace("\\", "/")
            all_files.append((local_path, remote_path))

    total = len(all_files)
    if total == 0:
        st.warning("Keine Dateien zum Hochladen gefunden.")
        return

    # nach Zielverzeichnis sortieren -> minimiert cwd-Wechsel
    all_files.sort(key=lambda p: os.path.dirname(p[1]))

    ftp = FTP_TLS() if FTP_USE_TLS else FTP()
    ftp.connect(FTP_HOST, 21, timeout=30)
    ftp.login(FTP_USER, FTP_PASS)
    if FTP_USE_TLS:
        ftp.prot_p()
    ftp.set_pasv(True)

    progress_bar = st.progress(0)
    status_text = st.empty()
    created_cache = set()
    current_dir = None
    uploaded = 0

    try:
        for local_path, remote_path in all_files:
            remote_dir = os.path.dirname(remote_path).replace("\\", "/")
            ensure_ftp_dirs(ftp, remote_dir, created_cache)

            # cwd nur wechseln, wenn sich das Verzeichnis ändert
            if remote_dir != current_dir:
                ftp.cwd(remote_dir)
                current_dir = remote_dir

            with open(local_path, "rb") as f:
                ftp.storbinary(f"STOR {os.path.basename(remote_path)}", f)

            uploaded += 1
            progress_bar.progress(uploaded / total)
            status_text.info(f"Hochgeladen: {uploaded}/{total} – {os.path.basename(local_path)}")

        status_text.success(f"Alle {uploaded} Dateien erfolgreich hochgeladen.")
    finally:
        try:
            ftp.quit()
        except (*ftplib.all_errors, OSError):
            ftp.close()

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
    period_end = start_date + pd.Timedelta(days=6)

    html = f"""<!DOCTYPE html>
<html lang="de">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
  <title>KW{kw:02d} – {fahrer_name}</title>
  <style>{css_styles}</style>
</head>
<body>
<div class="container-outer">
  <div class="back-bar">
    <a href="../../../../plane.php" class="btn-back" id="btnBack">
      <span class="btn-back-arrow" aria-hidden="true">‹</span>
      <span>Zurück</span>
    </a>
  </div>

  <header class="plan-hero">
    <div class="hero-topline">
      <span class="hero-dot"></span>
      <span>Dienstplan</span>
    </div>

    <div class="hero-main">
      <div class="hero-kw">
        <span class="hero-kw-label">Woche</span>
        <span class="hero-kw-number">{kw:02d}</span>
      </div>

      <div class="hero-meta">
        <div class="headline-name">{fahrer_name}</div>
        <div class="headline-period">{start_date.strftime('%d.%m.%Y')} – {period_end.strftime('%d.%m.%Y')}</div>
      </div>
    </div>
  </header>

  <main class="week-list">"""

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
        badge_text = "Dienst"
        icon = "🚚"

        content_check = f"{uhrzeit} {tour}".lower()
        if weekday == "Samstag":
            card_class += " samstag"
            badge_text = "Samstag"
        elif weekday == "Sonntag":
            card_class += " sonntag"
            badge_text = "Sonntag"

        if tour == "–" and uhrzeit == "–":
            card_class += " leer"
            badge_text = "Kein Eintrag"
            icon = "—"
        elif "urlaub" in content_check:
            card_class += " frei"
            badge_text = "Urlaub"
            icon = "☀️"
        elif "frei" in content_check:
            card_class += " frei"
            badge_text = "Frei"
            icon = "🕊️"
        elif "ausgleich" in content_check:
            card_class += " frei"
            badge_text = "Ausgleich"
            icon = "🛌"
        elif "krank" in content_check:
            card_class += " krank"
            badge_text = "Krank"
            icon = "💊"

        html += f"""
    <section class="{card_class}">
      <div class="day-top">
        <div class="day-date">
          <div class="weekday">{weekday}</div>
          <div class="prominent-date">{date_obj.strftime('%d.%m.%Y')}</div>
        </div>
        <div class="day-badge"><span>{icon}</span>{badge_text}</div>
      </div>

      <div class="info">
        <div class="info-block tour-block">
          <span class="label">Tour / Aufgabe:</span>
          <span class="value">{tour}</span>
        </div>
        <div class="info-block time-block">
          <span class="label">Uhrzeit:</span>
          <span class="value">{uhrzeit}</span>
        </div>
      </div>
    </section>"""

    html += """
  </main>
</div>
<div class="browser-safe-spacer" aria-hidden="true"></div>
<script src="../../../../dienstplan.js"></script>
</body>
</html>"""
    return html


css_styles = """
:root {
  --bg-main: #e6eef7;
  --card: #ffffff;
  --border: #d7e0ec;
  --border-strong: #b8c6d8;
  --text: #0f172a;
  --muted: #64748b;
  --soft: #f8fafc;
  --soft-blue: #eef6ff;
  --blue: #1b66b3;
  --blue-dark: #1b3a7a;
  --green: #15803d;
  --green-soft: #ecfdf3;
  --yellow: #b45309;
  --yellow-soft: #fffbeb;
  --red: #b91c1c;
  --red-soft: #fff1f2;
  --shadow: 0 10px 24px rgba(15, 23, 42, .08);
  --shadow-soft: 0 2px 8px rgba(15, 23, 42, .06);
}

* {
  box-sizing: border-box;
}

html {
  -webkit-text-size-adjust: 100%;
  min-height: 100%;
  min-height: -webkit-fill-available;
  min-height: 100svh;
  min-height: 100dvh;
}

body {
  margin: 0;
  padding: 0;
  min-height: 100%;
  min-height: -webkit-fill-available;
  min-height: calc(100svh + 2px);
  min-height: calc(100dvh + 2px);
  padding-bottom: calc(150px + env(safe-area-inset-bottom) + var(--browser-bottom-gap, 0px));
  background:
    radial-gradient(circle at top left, rgba(27, 102, 179, .10), transparent 32rem),
    var(--bg-main);
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
  color: var(--text);
  font-size: 15px;
  line-height: 1.45;
  overflow-x: hidden;
}

.container-outer {
  width: min(560px, calc(100vw - 24px));
  margin: 16px auto 0;
  padding-bottom: calc(80px + env(safe-area-inset-bottom) + var(--browser-bottom-gap, 0px));
}

.browser-safe-spacer {
  height: calc(140px + env(safe-area-inset-bottom) + var(--browser-bottom-gap, 0px));
  pointer-events: none;
}

.back-bar {
  margin-bottom: 10px;
}

.btn-back {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  font-size: .82rem;
  font-weight: 700;
  text-decoration: none;
  color: var(--blue-dark);
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 999px;
  padding: 7px 14px 7px 11px;
  box-shadow: var(--shadow-soft);
  -webkit-tap-highlight-color: transparent;
  transition: background .15s ease, border-color .15s ease, transform .05s ease;
}

.btn-back:hover {
  background: var(--soft-blue);
  border-color: var(--border-strong);
}

.btn-back:active {
  transform: scale(.97);
}

.btn-back-arrow {
  font-size: 1.15rem;
  line-height: 1;
  margin-top: -1px;
}

@media print {
  .back-bar {
    display: none;
  }
}

.plan-hero {
  background: linear-gradient(135deg, #ffffff 0%, #eef6ff 100%);
  border: 1px solid var(--border);
  border-radius: 18px;
  padding: 14px;
  box-shadow: var(--shadow);
  margin-bottom: 10px;
  overflow: hidden;
  position: relative;
}

.plan-hero::after {
  content: "";
  position: absolute;
  width: 90px;
  height: 90px;
  right: -34px;
  top: -36px;
  border-radius: 999px;
  background: rgba(27, 102, 179, .12);
}

.hero-topline {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-size: .70rem;
  font-weight: 800;
  letter-spacing: .05em;
  text-transform: uppercase;
  color: var(--blue-dark);
  background: rgba(27, 102, 179, .08);
  border: 1px solid rgba(27, 102, 179, .12);
  border-radius: 999px;
  padding: 4px 9px;
  margin-bottom: 10px;
  position: relative;
  z-index: 1;
}

.hero-dot {
  width: 7px;
  height: 7px;
  border-radius: 50%;
  background: var(--blue);
  display: inline-block;
}

.hero-main {
  display: grid;
  grid-template-columns: auto 1fr;
  gap: 12px;
  align-items: center;
  position: relative;
  z-index: 1;
}

.hero-kw {
  width: 76px;
  min-height: 70px;
  border-radius: 16px;
  background: linear-gradient(135deg, var(--blue), var(--blue-dark));
  color: #ffffff;
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  box-shadow: 0 8px 18px rgba(27, 102, 179, .25);
}

.hero-kw-label {
  font-size: .52rem;
  font-weight: 800;
  text-transform: uppercase;
  letter-spacing: .05em;
  opacity: .78;
}

.hero-kw-number {
  font-size: 1.85rem;
  line-height: 1;
  font-weight: 900;
  letter-spacing: -.03em;
}

.hero-meta {
  min-width: 0;
}

.headline-name {
  font-size: 1.08rem;
  line-height: 1.15;
  font-weight: 900;
  color: var(--text);
  overflow-wrap: anywhere;
}

.headline-period {
  margin-top: 4px;
  font-size: .82rem;
  font-weight: 700;
  color: var(--muted);
}

.week-list {
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.daycard {
  background: rgba(255,255,255,.92);
  border: 1px solid var(--border);
  border-radius: 16px;
  padding: 10px;
  box-shadow: var(--shadow-soft);
  transition: transform .12s ease, box-shadow .12s ease, border-color .12s ease;
}

.daycard:hover {
  transform: translateY(-1px);
  box-shadow: 0 8px 18px rgba(15, 23, 42, .10);
  border-color: var(--border-strong);
}

.day-top {
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 10px;
  margin-bottom: 8px;
}

.day-date {
  min-width: 0;
}

.weekday {
  font-size: .78rem;
  font-weight: 900;
  color: var(--blue-dark);
  line-height: 1.1;
}

.prominent-date {
  font-size: .72rem;
  color: var(--muted);
  font-weight: 700;
  margin-top: 2px;
}

.day-badge {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  flex: 0 0 auto;
  font-size: .66rem;
  font-weight: 900;
  color: var(--blue-dark);
  background: var(--soft-blue);
  border: 1px solid #dbeafe;
  border-radius: 999px;
  padding: 4px 8px;
}

.info {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 118px;
  gap: 7px;
}

.info-block {
  min-width: 0;
  display: flex;
  flex-direction: column;
  gap: 2px;
  background: var(--soft);
  border: 1px solid #e2e8f0;
  border-radius: 12px;
  padding: 8px 9px;
}

.label {
  color: #94a3b8;
  font-size: .58rem;
  font-weight: 900;
  text-transform: uppercase;
  letter-spacing: .045em;
  white-space: nowrap;
}

.value {
  color: var(--text);
  font-size: .88rem;
  font-weight: 900;
  line-height: 1.2;
  overflow-wrap: anywhere;
}

.time-block .value {
  white-space: nowrap;
}

.daycard.samstag,
.daycard.sonntag {
  background: linear-gradient(180deg, #ffffff, var(--yellow-soft));
  border-color: #fde68a;
}

.daycard.samstag .day-badge,
.daycard.sonntag .day-badge {
  color: var(--yellow);
  background: #fef3c7;
  border-color: #fde68a;
}

.daycard.samstag .weekday,
.daycard.sonntag .weekday {
  color: var(--yellow);
}

.daycard.frei {
  background: linear-gradient(180deg, #ffffff, var(--green-soft));
  border-color: #bbf7d0;
}

.daycard.frei .day-badge {
  color: var(--green);
  background: #dcfce7;
  border-color: #bbf7d0;
}

.daycard.frei .weekday {
  color: var(--green);
}

.daycard.krank {
  background: linear-gradient(180deg, #ffffff, var(--red-soft));
  border-color: #fecdd3;
}

.daycard.krank .day-badge {
  color: var(--red);
  background: #ffe4e6;
  border-color: #fecdd3;
}

.daycard.krank .weekday {
  color: var(--red);
}

.daycard.leer {
  opacity: .78;
}

.daycard.leer .value {
  color: #94a3b8;
}

@media (max-width: 440px) {
  .container-outer {
    width: min(100vw - 16px, 560px);
    margin-top: 8px;
    margin-bottom: 0;
  }

  .plan-hero {
    border-radius: 16px;
    padding: 12px;
  }

  .hero-main {
    grid-template-columns: 74px 1fr;
    gap: 10px;
  }

  .hero-kw {
    width: 74px;
    min-height: 64px;
    border-radius: 14px;
  }

  .hero-kw-label {
    font-size: .54rem;
  }

  .hero-kw-number {
    font-size: 1.65rem;
  }

  .headline-name {
    font-size: 1rem;
  }

  .headline-period {
    font-size: .76rem;
  }

  .daycard {
    border-radius: 14px;
    padding: 9px;
  }

  /* Handy: Tour und Uhrzeit bleiben nebeneinander, solange genug Platz ist. */
  .info {
    grid-template-columns: minmax(0, 1fr) 92px;
    gap: 6px;
  }

  .info-block {
    padding: 7px 7px;
  }

  .label {
    font-size: .50rem;
    letter-spacing: .025em;
  }

  .value {
    font-size: .80rem;
  }

  .time-block .value {
    white-space: nowrap;
  }
}

@media (max-width: 340px) {
  .info {
    grid-template-columns: 1fr;
  }

  .time-block .value {
    white-space: normal;
  }
}

@media print {
  body {
    background: #ffffff;
  }

  .container-outer {
    width: 100%;
    margin: 0;
    padding-bottom: 0;
  }

  .browser-safe-spacer {
    display: none;
  }

  .plan-hero,
  .daycard {
    box-shadow: none;
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
