import streamlit as st
import pandas as pd
from zipfile import ZipFile
from datetime import datetime
import tempfile
import os
import time
import ftplib
import threading
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from ftplib import FTP, FTP_TLS, error_perm
from dotenv import load_dotenv
from urllib.parse import quote

# .env laden
load_dotenv()
FTP_HOST = os.getenv("FTP_HOST")
FTP_USER = os.getenv("FTP_USER")
FTP_PASS = os.getenv("FTP_PASS")
FTP_BASE_DIR = os.getenv("FTP_BASE_DIR", "/")
FTP_USE_TLS = os.getenv("FTP_TLS", "0") == "1"
FTP_PARALLEL = int(os.getenv("FTP_PARALLEL", "6"))

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

def get_plan_kw(start_sonntag):
    """Kalenderwoche der Planwoche Sonntag bis Samstag.

    Maßgeblich ist der Montag innerhalb dieser Planwoche. Dadurch wird
    der Jahreswechsel korrekt behandelt: 28.12.2025 bis 03.01.2026
    gehört zur Kalenderwoche 01/2026 und nicht zu Kalenderwoche 53.
    """
    montag = pd.Timestamp(start_sonntag) + pd.Timedelta(days=1)
    return int(montag.isocalendar().week)

def _new_ftp():
    """Eine frische, eingeloggte FTP(S)-Verbindung im Passive-Mode."""
    ftp = FTP_TLS() if FTP_USE_TLS else FTP()
    ftp.connect(FTP_HOST, 21, timeout=30)
    ftp.login(FTP_USER, FTP_PASS)
    if FTP_USE_TLS:
        ftp.prot_p()
    ftp.set_pasv(True)
    return ftp

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

    progress_bar = st.progress(0)
    status_text = st.empty()
    t0 = time.perf_counter()

    # 1) Alle Zielverzeichnisse EINMAL vorab anlegen (eine Verbindung).
    #    Danach müssen die Worker nur noch STOR machen, kein mkd.
    status_text.info("Lege Verzeichnisse an ...")
    setup = _new_ftp()
    created = set()
    try:
        for _, remote_path in all_files:
            ensure_ftp_dirs(setup, os.path.dirname(remote_path), created)
    finally:
        try:
            setup.quit()
        except (*ftplib.all_errors, OSError):
            setup.close()

    # 2) Dateien parallel hochladen. Jeder Worker-Thread hält seine eigene
    #    FTP-Verbindung (thread-local) und nutzt sie für alle seine Dateien wieder.
    tl = threading.local()
    conns = []
    conns_lock = threading.Lock()

    def get_conn():
        ftp = getattr(tl, "ftp", None)
        if ftp is None:
            ftp = _new_ftp()
            tl.ftp = ftp
            tl.cwd = None
            with conns_lock:
                conns.append(ftp)
        return ftp

    def upload_one(item):
        local_path, remote_path = item
        ftp = get_conn()
        remote_dir = os.path.dirname(remote_path).replace("\\", "/")
        # cwd pro Verbindung nur wechseln, wenn nötig
        if tl.cwd != remote_dir:
            ftp.cwd(remote_dir)
            tl.cwd = remote_dir
        with open(local_path, "rb") as f:
            ftp.storbinary(f"STOR {os.path.basename(remote_path)}", f)

    workers = max(1, min(FTP_PARALLEL, total))
    uploaded = 0
    errors = []

    with ThreadPoolExecutor(max_workers=workers) as ex:
        futures = {ex.submit(upload_one, it): it for it in all_files}
        for fut in as_completed(futures):
            local_path, remote_path = futures[fut]
            try:
                fut.result()
            except Exception as e:
                errors.append((os.path.basename(remote_path), str(e)))
            uploaded += 1
            progress_bar.progress(uploaded / total)
            status_text.info(f"Hochgeladen: {uploaded}/{total}")

    # 3) Alle Worker-Verbindungen schließen
    for ftp in conns:
        try:
            ftp.quit()
        except (*ftplib.all_errors, OSError):
            try:
                ftp.close()
            except Exception:
                pass

    dauer = time.perf_counter() - t0
    ok = total - len(errors)
    if errors:
        status_text.warning(
            f"{ok}/{total} hochgeladen, {len(errors)} Fehler "
            f"in {dauer:.1f}s ({workers} parallele Verbindungen)."
        )
        st.code("\n".join(f"{n}: {e}" for n, e in errors[:10]))
    else:
        status_text.success(
            f"Alle {total} Dateien in {dauer:.1f}s hochgeladen "
            f"({workers} parallele Verbindungen, Ø {dauer / total:.2f}s/Datei)."
        )

def normalize_driver_name(nachname, vorname):
    """Fahrernamen bereinigen; leere Excel-Zellen können als 0/0.0 erscheinen."""
    def clean_part(value):
        if pd.isna(value):
            return ""

        value_str = str(value).strip()
        value_lower = value_str.lower()

        if value_lower in {"", "nan", "none", "null", "nat", "-", "–"}:
            return ""
        if re.fullmatch(r"0+(?:[.,]0+)?", value_str):
            return ""

        return value_str.title()

    nachname = clean_part(nachname)
    vorname = clean_part(vorname)

    # Bekannte Schreibvarianten aus unterschiedlichen Excel-Dateien
    # auf eine einheitliche Fahreridentität zusammenführen.
    alias = {
        ("khalleefah", ""): ("Khalleefah", "Saed Awami Sayid"),
        ("alem", "mohamed"): ("Alem", "Mohammed"),
        ("maghraoui", "zakaria"): ("Maghraoui", "Zakariae"),
    }.get((nachname.casefold(), vorname.casefold()))
    if alias:
        nachname, vorname = alias

    if nachname and vorname:
        return f"{nachname}, {vorname}"
    if nachname:
        return nachname
    if vorname:
        return vorname
    return ""

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

def generate_shared_html(css_styles):
    """Erzeugt eine CSP-konforme HTML-Datei ohne eingebettetes JavaScript."""
    template = r'''<!DOCTYPE html>
<html lang="de">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
  <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
  <meta http-equiv="Pragma" content="no-cache">
  <meta http-equiv="Expires" content="0">
  <title>Dienstplan</title>
  <style>__CSS_STYLES__</style>
</head>
<body>
<div class="container-outer">
  <div class="back-bar">
    <a href="../../../plane.php" class="btn-back" id="btnBack">
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
        <span class="hero-kw-number" id="kwNumber">--</span>
      </div>

      <div class="hero-meta">
        <div class="headline-name" id="driverName">Dienstplan wird geladen</div>
        <div class="headline-period" id="periodText">Bitte einen Moment</div>
      </div>
    </div>
  </header>

  <main class="week-list" id="weekList"></main>
</div>
<div class="browser-safe-spacer" aria-hidden="true"></div>

<script src="dienstplan_app.js?v=2" defer></script>
<script src="../../../dienstplan.js" defer></script>
</body>
</html>'''
    return template.replace("__CSS_STYLES__", css_styles)


def generate_shared_js():
    """JavaScript der gemeinsamen Dienstplanseite als externe Datei."""
    return r'''(function () {
  "use strict";

  const params = new URLSearchParams(window.location.search);
  const requestedKw = params.get("kw");
  const requestedDriver = params.get("fahrer") || params.get("driver");

  const weekList = document.getElementById("weekList");
  const kwNumber = document.getElementById("kwNumber");
  const driverName = document.getElementById("driverName");
  const periodText = document.getElementById("periodText");

  function escapeHtml(value) {
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function parseCsv(text, delimiter = ";") {
    const rows = [];
    let row = [];
    let field = "";
    let quoted = false;

    for (let i = 0; i < text.length; i += 1) {
      const char = text[i];

      if (quoted) {
        if (char === '"') {
          if (text[i + 1] === '"') {
            field += '"';
            i += 1;
          } else {
            quoted = false;
          }
        } else {
          field += char;
        }
      } else if (char === '"') {
        quoted = true;
      } else if (char === delimiter) {
        row.push(field);
        field = "";
      } else if (char === "\n") {
        row.push(field.replace(/\r$/, ""));
        rows.push(row);
        row = [];
        field = "";
      } else {
        field += char;
      }
    }

    if (field.length > 0 || row.length > 0) {
      row.push(field.replace(/\r$/, ""));
      rows.push(row);
    }

    if (!rows.length) return [];

    const headers = rows.shift().map((header, index) =>
      index === 0 ? header.replace(/^\uFEFF/, "") : header
    );

    return rows
      .filter(values => values.some(value => value !== ""))
      .map(values => {
        const item = {};
        headers.forEach((header, index) => {
          item[header] = values[index] ?? "";
        });
        return item;
      });
  }

  function normalizeKw(value) {
    const number = Number.parseInt(value, 10);
    return Number.isFinite(number) ? String(number) : "";
  }

  function formatDate(isoDate) {
    const parts = String(isoDate).split("-");
    if (parts.length !== 3) return isoDate;
    return `${parts[2]}.${parts[1]}.${parts[0]}`;
  }

  function createDayCard(entry) {
    const weekday = entry.wochentag || "";
    const time = entry.uhrzeit || "–";
    const tour = entry.tour || "–";
    const contentCheck = `${time} ${tour}`.toLowerCase();

    let cardClass = "daycard";
    let badgeText = "Dienst";
    let icon = "🚚";

    if (weekday === "Samstag") {
      cardClass += " samstag";
      badgeText = "Samstag";
    } else if (weekday === "Sonntag") {
      cardClass += " sonntag";
      badgeText = "Sonntag";
    }

    if (tour === "–" && time === "–") {
      cardClass += " leer";
      badgeText = "Kein Eintrag";
      icon = "—";
    } else if (contentCheck.includes("urlaub")) {
      cardClass += " frei";
      badgeText = "Urlaub";
      icon = "☀️";
    } else if (contentCheck.includes("frei")) {
      cardClass += " frei";
      badgeText = "Frei";
      icon = "🕊️";
    } else if (contentCheck.includes("ausgleich")) {
      cardClass += " frei";
      badgeText = "Ausgleich";
      icon = "🛌";
    } else if (contentCheck.includes("krank")) {
      cardClass += " krank";
      badgeText = "Krank";
      icon = "💊";
    }

    return `
      <section class="${cardClass}">
        <div class="day-top">
          <div class="day-date">
            <div class="weekday">${escapeHtml(weekday)}</div>
            <div class="prominent-date">${escapeHtml(formatDate(entry.datum))}</div>
          </div>
          <div class="day-badge"><span>${icon}</span>${escapeHtml(badgeText)}</div>
        </div>

        <div class="info">
          <div class="info-block tour-block">
            <span class="label">Tour / Aufgabe:</span>
            <span class="value">${escapeHtml(tour)}</span>
          </div>
          <div class="info-block time-block">
            <span class="label">Uhrzeit:</span>
            <span class="value">${escapeHtml(time)}</span>
          </div>
        </div>
      </section>`;
  }

  function showMessage(title, detail) {
    driverName.textContent = title;
    periodText.textContent = detail;
    weekList.innerHTML = `
      <section class="daycard leer">
        <div class="info-block">
          <span class="value">${escapeHtml(detail)}</span>
        </div>
      </section>`;
  }

  if (!requestedKw || !requestedDriver) {
    showMessage(
      "Dienstplan nicht ausgewählt",
      "Der Aufruf benötigt die Angaben kw und fahrer, zum Beispiel: dienstplan.html?kw=26&fahrer=Mueller"
    );
    return;
  }

  fetch(`dienstplaene.csv?v=${Date.now()}`, { cache: "no-store" })
    .then(response => {
      if (!response.ok) {
        throw new Error(`CSV konnte nicht geladen werden (${response.status}).`);
      }
      return response.text();
    })
    .then(text => {
      const data = parseCsv(text);
      const selected = data
        .filter(entry =>
          normalizeKw(entry.kw) === normalizeKw(requestedKw) &&
          entry.fahrer_key === requestedDriver
        )
        .sort((a, b) => Number(a.reihenfolge) - Number(b.reihenfolge));

      if (!selected.length) {
        kwNumber.textContent = String(requestedKw).padStart(2, "0");
        showMessage(
          "Kein Dienstplan gefunden",
          `Für Kalenderwoche ${requestedKw} und Fahrer ${requestedDriver} sind keine Daten vorhanden.`
        );
        return;
      }

      const first = selected[0];
      const dates = selected.map(entry => entry.datum).filter(Boolean).sort();
      const startDate = dates[0];
      const endDate = dates[dates.length - 1];

      kwNumber.textContent = String(first.kw).padStart(2, "0");
      driverName.textContent = first.fahrer_name;
      periodText.textContent = `${formatDate(startDate)} – ${formatDate(endDate)}`;
      document.title = `KW${String(first.kw).padStart(2, "0")} – ${first.fahrer_name}`;
      weekList.innerHTML = selected.map(createDayCard).join("");
    })
    .catch(error => {
      showMessage("Fehler beim Laden", error.message);
    });
})();
'''


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
            html_path = os.path.join(tmpdir, "dienstplan.html")
            js_path = os.path.join(tmpdir, "dienstplan_app.js")
            csv_path = os.path.join(tmpdir, "dienstplaene.csv")

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
                # Zwei verschiedene Fahrer mit gleichem Nachnamen:
                # Rene behält aus Kompatibilitätsgründen den bisherigen Schlüssel.
                ("schlutt", "rene"): "Schlutt",
                ("schlutt", "hubert"): "HSchlutt",
            }

            csv_rows = []
            bekannte_zeilen = set()
            fahrer_wochen = set()

            for file in uploaded_files:
                fahrer_df = pd.read_excel(file, sheet_name="a Fahrer", engine="openpyxl")
                touren_df = pd.read_excel(file, sheet_name="Touren", skiprows=4, engine="openpyxl")

                fahrer_dict = {}
                for _, r in fahrer_df.iterrows():
                    fahrer_name = normalize_driver_name(r.iloc[1], r.iloc[2])
                    if fahrer_name:
                        fahrer_dict[fahrer_name] = {}

                alle_gueltigen_daten = []

                for _, row in touren_df.iterrows():
                    datum = row.iloc[14]
                    tour = row.iloc[15]
                    uhrzeit = row.iloc[8]

                    datum_dt = None
                    if pd.notna(datum):
                        try:
                            datum_dt = pd.to_datetime(datum)
                            alle_gueltigen_daten.append(datum_dt)
                        except Exception:
                            datum_dt = None

                    uhrzeit_str = parse_uhrzeit(uhrzeit)
                    tour_str = parse_tour(tour)
                    eintrag_text = f"{uhrzeit_str} – {tour_str}"

                    for pos in [(3, 4), (6, 7)]:
                        fahrer_name = normalize_driver_name(row.iloc[pos[0]], row.iloc[pos[1]])
                        if not fahrer_name:
                            continue

                        if fahrer_name not in fahrer_dict:
                            fahrer_dict[fahrer_name] = {}

                        if datum_dt is not None:
                            tag = datum_dt.date()
                            if tag not in fahrer_dict[fahrer_name]:
                                fahrer_dict[fahrer_name][tag] = []

                            if eintrag_text not in fahrer_dict[fahrer_name][tag]:
                                fahrer_dict[fahrer_name][tag].append(eintrag_text)

                if alle_gueltigen_daten:
                    global_start_datum = min(alle_gueltigen_daten).date()
                else:
                    global_start_datum = pd.Timestamp.today().date()

                global_start_sonntag = global_start_datum - pd.Timedelta(
                    days=(global_start_datum.weekday() + 1) % 7
                )
                global_kw = get_plan_kw(global_start_sonntag)

                fahrer_dict = dict(sorted(fahrer_dict.items(), key=lambda x: x[0].lower()))

                for fahrer_name, eintraege in fahrer_dict.items():
                    if eintraege:
                        start_datum = min(eintraege.keys())
                        start_sonntag = start_datum - pd.Timedelta(
                            days=(start_datum.weekday() + 1) % 7
                        )
                        kw = get_plan_kw(start_sonntag)
                    else:
                        start_sonntag = global_start_sonntag
                        kw = global_kw

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

                    alter_dateiname = f"KW{kw:02d}_{filename_part}.html"
                    filename_lower = alter_dateiname.lower()

                    if "ch._holtz" in filename_lower or any(
                        stichwort in filename_lower for stichwort in ausschluss_stichwoerter
                    ):
                        continue

                    fahrer_wochen.add((kw, filename_part))
                    reihenfolge = 0

                    for i in range(7):
                        tag_datum = start_sonntag + pd.Timedelta(days=i)
                        wochentag = wochentage_deutsch_map.get(
                            tag_datum.strftime("%A"),
                            tag_datum.strftime("%A")
                        )

                        tag_eintraege = eintraege.get(tag_datum, [])
                        if not tag_eintraege:
                            tag_eintraege = ["– – –"]

                        for eintrag in tag_eintraege:
                            reihenfolge += 1

                            if " – " in eintrag:
                                uhrzeit_str, tour_str = eintrag.split(" – ", 1)
                            else:
                                uhrzeit_str, tour_str = "–", eintrag

                            if not uhrzeit_str or uhrzeit_str.strip() == "–":
                                uhrzeit_str = "–"
                            if not tour_str or tour_str.strip() == "–":
                                tour_str = "–"

                            url = (
                                f"dienstplan.html?kw={kw:02d}"
                                f"&fahrer={quote(filename_part, safe='')}"
                            )

                            row_key = (
                                kw,
                                filename_part,
                                fahrer_name,
                                tag_datum.isoformat(),
                                uhrzeit_str,
                                tour_str,
                            )
                            if row_key in bekannte_zeilen:
                                continue
                            bekannte_zeilen.add(row_key)

                            csv_rows.append({
                                "kw": kw,
                                "fahrer_key": filename_part,
                                "fahrer_name": fahrer_name,
                                "datum": tag_datum.isoformat(),
                                "wochentag": wochentag,
                                "uhrzeit": uhrzeit_str,
                                "tour": tour_str,
                                "reihenfolge": reihenfolge,
                                "url": url,
                            })

            if not csv_rows:
                st.warning("Es wurden keine Dienstplandaten erzeugt.")
                st.stop()

            with open(html_path, "w", encoding="utf-8") as f:
                f.write(generate_shared_html(css_styles))

            with open(js_path, "w", encoding="utf-8") as f:
                f.write(generate_shared_js())

            csv_df = pd.DataFrame(csv_rows)

            # Bei überlappenden Quelldateien können für denselben Fahrer und Tag
            # sowohl ein echter Eintrag als auch ein leerer Platzhalter entstehen.
            # Der Platzhalter wird dann entfernt; mehrere echte Aufgaben bleiben erhalten.
            csv_df = csv_df.drop_duplicates(
                subset=[
                    "kw", "fahrer_key", "fahrer_name", "datum",
                    "wochentag", "uhrzeit", "tour"
                ],
                keep="first"
            ).copy()

            ist_leer = (
                csv_df["uhrzeit"].astype(str).str.strip().isin(["", "–"])
                & csv_df["tour"].astype(str).str.strip().isin(["", "–"])
            )
            hat_echten_eintrag = (~ist_leer).groupby([
                csv_df["kw"],
                csv_df["fahrer_key"],
                csv_df["datum"],
            ]).transform("any")
            csv_df = csv_df.loc[~(ist_leer & hat_echten_eintrag)].copy()

            csv_df["_datum_sort"] = pd.to_datetime(
                csv_df["datum"], errors="coerce"
            )
            csv_df["_reihenfolge_alt"] = pd.to_numeric(
                csv_df["reihenfolge"], errors="coerce"
            ).fillna(9999)

            csv_df = csv_df.sort_values(
                by=[
                    "kw", "fahrer_name", "_datum_sort",
                    "_reihenfolge_alt"
                ],
                kind="stable"
            )
            csv_df["reihenfolge"] = (
                csv_df.groupby(["kw", "fahrer_key"]).cumcount() + 1
            )
            csv_df = csv_df.drop(
                columns=["_datum_sort", "_reihenfolge_alt"]
            )

            csv_df.to_csv(
                csv_path,
                sep=";",
                index=False,
                encoding="utf-8-sig",
                lineterminator="\n"
            )

            with ZipFile(zip_path, "w") as zipf:
                zipf.write(html_path, arcname="dienstplan.html")
                zipf.write(js_path, arcname="dienstplan_app.js")
                zipf.write(csv_path, arcname="dienstplaene.csv")

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
                f"Eine HTML-Datei, eine JavaScript-Datei, eine CSV-Datei und "
                f"{len(fahrer_wochen)} Fahrer-Woche(n) erstellt."
            )

            st.info(
                "Aufruf künftig zum Beispiel: "
                "dienstplan.html?kw=26&fahrer=Mueller"
            )

            st.download_button(
                "ZIP mit HTML-, JavaScript- und CSV-Datei herunterladen",
                data=zip_bytes,
                file_name="gesamt_export.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten: {e}")
