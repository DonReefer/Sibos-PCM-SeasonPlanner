"""
SIBOS Season Planner 2026
- Planung oben (Rennen → Datum → Etappen → Belegung → X)
- Tabs: Rennprofil | Fahrerwerte | Fahreranalyse | Eignung | Formkurven | Regelwerk
- Import: Excel + Foto/Screenshot mit Anleitung
- Buttons: KI-Planung starten + Formkurven berechnen
Daten 1:1 aus PCM_Saisonplaner.xlsx
"""

import streamlit as st
import pandas as pd
import openpyxl
from pathlib import Path
import io
import re
from datetime import datetime

st.set_page_config(
    page_title="SIBOS Season Planner 2026",
    page_icon="🚴",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE = Path(__file__).parent
EXCEL = BASE / "PCM_Saisonplaner.xlsx"
LOGO = BASE / "logo.png"
MAX_RIDERS = 7
MAX_RENNTAGE = 70

st.markdown("""
<style>
  .stApp { background:#f1f5f9; color:#0f172a; }
  header[data-testid="stHeader"] { background:#f1f5f9; }
  .block-container { padding-top:0.5rem; padding-bottom:1.5rem; max-width:100%; }
  .stButton>button {
    background:#fff; color:#0f172a; border:1px solid #cbd5e1;
    border-radius:8px; font-weight:500; height:2.3rem;
  }
  .stButton>button[kind="primary"] {
    background:#2563eb !important; border-color:#2563eb !important;
    color:#fff !important; font-weight:600;
  }
  /* secondary formkurven button */
  div[data-testid="stHorizontalBlock"] button[kind="secondary"] {
    background:#e2e8f0 !important; border-color:#94a3b8 !important;
    color:#334155 !important; font-weight:500; font-size:0.85rem;
  }
  .stTabs [data-baseweb="tab-list"] {
    background:#fff; border-radius:10px; gap:3px; padding:4px;
    border:1px solid #e2e8f0;
  }
  .stTabs [data-baseweb="tab"] { color:#64748b; border-radius:8px; font-size:0.88rem; }
  .stTabs [aria-selected="true"] {
    background:#2563eb !important; color:#fff !important;
  }
  .occ-ok  { background:#dcfce7; color:#166534; padding:2px 6px; border-radius:4px; font-weight:600; font-size:0.75rem; white-space:nowrap; }
  .occ-full{ background:#bbf7d0; color:#14532d; padding:2px 6px; border-radius:4px; font-weight:700; font-size:0.75rem; white-space:nowrap; }
  .occ-over{ background:#fecaca; color:#991b1b; padding:2px 6px; border-radius:4px; font-weight:700; font-size:0.75rem; white-space:nowrap; }
  .import-hint {
    background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px;
    padding:12px 14px; font-size:0.9rem; color:#334155; margin-top:6px;
  }
</style>
""", unsafe_allow_html=True)


def parse_date_range(s, year=2026):
    if not s or not isinstance(s, str):
        return None, None
    s = s.strip().rstrip('.')
    m = re.match(r'^(\d{1,2})\.(\d{1,2})$', s)
    if m:
        try:
            d = datetime(year, int(m.group(2)), int(m.group(1)))
            return d, d
        except ValueError:
            return None, None
    m = re.match(r'^(\d{1,2})\.-(\d{1,2})\.(\d{1,2})$', s)
    if m:
        try:
            return datetime(year, int(m.group(3)), int(m.group(1))), datetime(year, int(m.group(3)), int(m.group(2)))
        except ValueError:
            return None, None
    m = re.match(r'^(\d{1,2})\.(\d{1,2})\.-(\d{1,2})\.(\d{1,2})$', s)
    if m:
        try:
            return datetime(year, int(m.group(2)), int(m.group(1))), datetime(year, int(m.group(4)), int(m.group(3)))
        except ValueError:
            return None, None
    return None, None


def ranges_overlap(a0, a1, b0, b1):
    if None in (a0, a1, b0, b1):
        return False
    return a0 <= b1 and b0 <= a1


@st.cache_data
def load_all(path: str):
    wb = openpyxl.load_workbook(path, data_only=True)

    # Fahrerwerte
    ws = wb["Fahrerwerte"]
    riders = []
    for r in range(4, 50):
        name = ws.cell(r, 1).value
        if not name:
            break
        riders.append({
            "Name": str(name).strip(),
            "EB": ws.cell(r, 2).value, "BE": ws.cell(r, 3).value,
            "MG": ws.cell(r, 4).value, "HÜG": ws.cell(r, 5).value,
            "ZF": ws.cell(r, 6).value, "PRL": ws.cell(r, 7).value,
            "KSP": ws.cell(r, 8).value, "SP": ws.cell(r, 9).value,
            "BES": ws.cell(r, 10).value, "ABF": ws.cell(r, 11).value,
            "ASR": ws.cell(r, 12).value, "AUS": ws.cell(r, 13).value,
            "ZÄH": ws.cell(r, 14).value, "REG": ws.cell(r, 15).value,
            "Fahrerrolle": ws.cell(r, 16).value,
            "Bevorzugte Rennen": ws.cell(r, 17).value,
        })
    df_fw = pd.DataFrame(riders)

    # Races from Planung
    ws = wb["Planung"]
    races = []
    for c in range(3, 60):
        name = ws.cell(3, c).value
        if name is None:
            break
        races.append({
            "name": str(name).strip(),
            "typ": ws.cell(2, c).value,
            "datum": str(ws.cell(4, c).value or ""),
            "etappen": int(ws.cell(5, c).value or 1),
            "col": c,
        })
    for race in races:
        s, e = parse_date_range(race["datum"])
        race["start"], race["end"] = s, e

    # Plan matrix – names from Fahrerwerte (Planung col A is empty in this file)
    plan_rows = []
    for i, rider in enumerate(riders):
        row_idx = 7 + i
        row = {"Fahrer": rider["Name"]}
        for race in races:
            val = ws.cell(row_idx, race["col"]).value
            row[race["name"]] = "X" if val in ("X", "x") else ""
        plan_rows.append(row)
    df_plan = pd.DataFrame(plan_rows)

    # Rennprofil
    ws = wb["Rennprofil"]
    rp = []
    for r in range(2, 60):
        name = ws.cell(r, 1).value
        if not name or str(name).strip() == "":
            continue
        rp.append({
            "Rennen": str(name).strip(),
            "Priorität": ws.cell(r, 3).value,
            "Datum": ws.cell(r, 4).value,
            "Kategorie": ws.cell(r, 5).value,
            "Profil": ws.cell(r, 6).value,
            "Profil-Ampel": ws.cell(r, 7).value,
            "Etappen": ws.cell(r, 8).value,
            "Gesamthärte": ws.cell(r, 10).value,
            "Bergankunft": ws.cell(r, 11).value,
            "Sprintankunft": ws.cell(r, 12).value,
            "Bergetappen": ws.cell(r, 13).value,
            "Mittelgebirge": ws.cell(r, 14).value,
            "Flachetappen": ws.cell(r, 15).value,
            "Kopfsteinpflaster": ws.cell(r, 16).value,
            "Höhenmeter": ws.cell(r, 17).value,
            "Bergfaktor": ws.cell(r, 18).value,
            "Sprintfaktor": ws.cell(r, 19).value,
            "Hügelfaktor": ws.cell(r, 20).value,
            "TT-Faktor": ws.cell(r, 21).value,
            "Windanfälligkeit": ws.cell(r, 22).value,
            "Empfohlene Fahrertypen": ws.cell(r, 23).value,
            "Teilgenommen 2025": ws.cell(r, 24).value,
        })
    df_rp = pd.DataFrame(rp)

    # Fahreranalyse – names from riders list, data by row index
    ws = wb["Fahreranalyse"]
    fa = []
    for i, rider in enumerate(riders):
        r = 4 + i
        fa.append({
            "Name": rider["Name"],
            "Fahrerrolle": ws.cell(r, 3).value,
            "Fahrertyp": ws.cell(r, 4).value,
            "Bevorzugte Rennen": ws.cell(r, 5).value,
            "Größe (cm)": ws.cell(r, 7).value,
            "Gewicht (kg)": ws.cell(r, 8).value,
            "Alter": ws.cell(r, 9).value,
            "Nationalität": ws.cell(r, 12).value,
            "Höhen-/Hitzetauglichkeit": ws.cell(r, 13).value,
            "Regenerationsbedarf n. GT": ws.cell(r, 14).value,
            "Formkurve/Saisonziel": ws.cell(r, 15).value,
            "Palmarès": ws.cell(r, 16).value,
        })
    df_fa = pd.DataFrame(fa)

    # Eignung
    ws = wb["Fahrerbewertung"]
    eign_races = []
    for c in range(2, 60):
        n = ws.cell(3, c).value
        if n is None:
            break
        eign_races.append(str(n).strip())
    eign_rows = []
    for i, rider in enumerate(riders):
        r = 4 + i
        row = {"Fahrer": rider["Name"]}
        for j, rn in enumerate(eign_races):
            val = ws.cell(r, 2 + j).value
            row[rn] = val if val is not None else ""
        eign_rows.append(row)
    df_eign = pd.DataFrame(eign_rows)

    # Regelwerk text (column B)
    ws = wb["KI-Regelwerk"]
    lines = []
    for r in range(1, min(80, ws.max_row + 1)):
        for col in (1, 2, 3):
            v = ws.cell(r, col).value
            if v and isinstance(v, str) and len(v.strip()) > 5:
                t = v.strip()
                if t not in lines:
                    lines.append(t)
    regel_text = "\n\n".join(lines[:60]) if lines else "KI-Regelwerk – siehe Excel."

    return {
        "fahrerwerte": df_fw,
        "rennprofil": df_rp,
        "planung": df_plan,
        "analyse": df_fa,
        "eignung": df_eign,
        "races": races,
        "regelwerk": regel_text,
    }


def compute_renntage(plan, races):
    days = []
    for _, row in plan.iterrows():
        total = sum(race["etappen"] for race in races if row.get(race["name"]) == "X")
        days.append(total)
    return pd.Series(days, index=plan.index)


def occupancy(plan, races):
    return {r["name"]: int((plan[r["name"]] == "X").sum()) for r in races}


def find_overlaps(plan, races):
    conflicts = set()
    for idx, row in plan.iterrows():
        assigned = [r for r in races if row.get(r["name"]) == "X"]
        for i in range(len(assigned)):
            for j in range(i + 1, len(assigned)):
                a, b = assigned[i], assigned[j]
                if ranges_overlap(a["start"], a["end"], b["start"], b["end"]):
                    conflicts.add((idx, a["name"]))
                    conflicts.add((idx, b["name"]))
    return conflicts


# ── Session ────────────────────────────────────────────
if "data" not in st.session_state:
    st.session_state.data = load_all(str(EXCEL))
    st.session_state.plan = st.session_state.data["planung"].copy()
    st.session_state.ai = "Grok (xAI)"
    st.session_state.msg = None
    st.session_state.show_import = False

data = st.session_state.data
races = data["races"]
plan = st.session_state.plan
occ = occupancy(plan, races)
conflicts = find_overlaps(plan, races)
renntage = compute_renntage(plan, races)


# ── TOP BAR ────────────────────────────────────────────
c1, c2, c3, c4, c5, c6 = st.columns([1.3, 1.5, 1.3, 1.0, 1.0, 1.6])

with c1:
    if LOGO.exists():
        st.image(str(LOGO), width=120)
    else:
        st.markdown("**SIBOS**")

with c2:
    st.text_input("Suche", placeholder="Suchen…", label_visibility="collapsed", key="search")

with c3:
    st.session_state.ai = st.selectbox(
        "KI",
        ["Grok (xAI)", "Claude (Anthropic)", "GPT-4o (OpenAI)", "Gemini (Google)", "Lokal (Ollama)"],
        label_visibility="collapsed",
    )

with c4:
    if st.button("Importieren", use_container_width=True):
        st.session_state.show_import = not st.session_state.show_import

with c5:
    if st.button("Exportieren", use_container_width=True):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            plan.to_excel(w, sheet_name="Planung", index=False)
            data["fahrerwerte"].to_excel(w, sheet_name="Fahrerwerte", index=False)
            data["rennprofil"].to_excel(w, sheet_name="Rennprofil", index=False)
            data["analyse"].to_excel(w, sheet_name="Fahreranalyse", index=False)
            data["eignung"].to_excel(w, sheet_name="Fahrerbewertung", index=False)
        st.download_button(
            "📥 Download Excel",
            data=buf.getvalue(),
            file_name="SIBOS_Saisonplanung_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

with c6:
    if st.button("⚡ KI-Planung starten", type="primary", use_container_width=True):
        st.session_state.msg = (
            f"**KI-Planung** mit {st.session_state.ai}: Regelwerk Teil A wird ausgeführt "
            "(alte X löschen → neue setzen → QC). Vollständige Automatik folgt in der nächsten Version."
        )
        st.toast("KI-Planung vorbereitet", icon="⚡")

# Formkurven-Button (kleiner, andere Farbe) direkt darunter
fc1, fc2, fc3 = st.columns([3.8, 1.6, 1.6])
with fc2:
    if st.button("📈 Formkurven berechnen", type="secondary", use_container_width=True):
        st.session_state.msg = (
            "**Formkurven:** Wird laut Regelwerk in einem separaten Schritt (Teil B) berechnet – "
            "nach der Saisonplanung. Umsetzung folgt."
        )
        st.toast("Formkurven-Schritt vorbereitet", icon="📈")


# ── IMPORT PANEL ───────────────────────────────────────
if st.session_state.show_import:
    st.markdown("---")
    st.markdown("### Daten importieren")
    ic1, ic2 = st.columns(2)
    with ic1:
        st.markdown("**📄 Excel-Datei**")
        up_x = st.file_uploader("Excel wählen", type=["xlsx"], key="up_xlsx", label_visibility="collapsed")
        if up_x is not None:
            tmp = BASE / "_upload.xlsx"
            tmp.write_bytes(up_x.getbuffer())
            st.session_state.data = load_all(str(tmp))
            st.session_state.plan = st.session_state.data["planung"].copy()
            st.session_state.msg = "Excel importiert."
            st.session_state.show_import = False
            st.rerun()
    with ic2:
        st.markdown("**📷 Foto / Screenshot**")
        up_img = st.file_uploader(
            "Bild wählen", type=["png", "jpg", "jpeg", "webp"], key="up_img", label_visibility="collapsed"
        )
        if up_img is not None:
            st.session_state.msg = (
                f"Foto „{up_img.name}“ empfangen. KI-OCR liest Namen + Werte und trägt sie in "
                "Fahrerwerte ein – folgt in der nächsten Version."
            )
            st.session_state.show_import = False

    st.markdown(
        """
        <div class="import-hint">
        <b>Anleitung Foto / Screenshot</b><br>
        Fotografiere den Bildschirm im Spiel (PCM), auf dem die <b>Fahrerwerte</b> stehen:<br>
        • linke Spalte = <b>Fahrernamen</b><br>
        • daneben die Zahlen: <b>EB, BE, MG, HÜG, ZF, PRL, KSP, SP, BES, ABF, ASR, AUS, ZÄH, REG</b><br>
        Am besten so, dass Namen und alle Zahlen gut lesbar sind (wie in deinem Beispiel-Screenshot).<br>
        Die App übernimmt die Namen und Werte dann automatisch ins Blatt <b>Fahrerwerte</b>
        und zeigt sie überall (Planung, Analyse, Eignung …).
        </div>
        """,
        unsafe_allow_html=True,
    )

if st.session_state.msg:
    st.info(st.session_state.msg)


# ── PLANUNG ────────────────────────────────────────────
st.markdown("---")
st.markdown("### Planung 2026")
st.caption(
    "Kopfzeilen: **Rennen → Datum → Etappen → Belegung (aktuell/max)** · "
    "Grün = Maximum erreicht · Rot = überschritten · "
    "Rote Markierung = zeitliche Überschneidung · Renntage ab 71 = über Maximum (70)"
)

badges = []
for race in races:
    cnt = occ[race["name"]]
    label = f"{cnt}/{MAX_RIDERS}"
    cls = "occ-over" if cnt > MAX_RIDERS else ("occ-full" if cnt == MAX_RIDERS else "occ-ok")
    short = race["name"][:13]
    badges.append(f'<span class="{cls}" title="{race["name"]}">{short} {label}</span>')
st.markdown(
    '<div style="display:flex;gap:4px;overflow-x:auto;padding:4px 0;flex-wrap:nowrap;">'
    + "".join(badges) + "</div>",
    unsafe_allow_html=True,
)

display = plan.copy()
display.insert(1, "Renntage", renntage)

short_map = {}
for race in races:
    cnt = occ[race["name"]]
    flag = " 🔴" if cnt > MAX_RIDERS else (" 🟢" if cnt == MAX_RIDERS else "")
    header = f"{race['name'][:15]}\n{race['datum']}\n{race['etappen']} Et. | {cnt}/{MAX_RIDERS}{flag}"
    short_map[race["name"]] = header
    display[header] = display[race["name"]].replace("", "·")
    display.drop(columns=[race["name"]], inplace=True)

if conflicts:
    names = sorted({plan.loc[i, "Fahrer"] for i, _ in conflicts})
    st.warning("⏱ Zeitliche Überschneidung bei: " + ", ".join(names))

edited = st.data_editor(
    display,
    use_container_width=True,
    height=500,
    hide_index=True,
    disabled=["Fahrer", "Renntage"],
    key="plan_editor",
)

if not edited.equals(display):
    new_plan = plan.copy()
    for race in races:
        h = short_map[race["name"]]
        if h in edited.columns:
            new_plan[race["name"]] = edited[h].apply(
                lambda v: "X" if str(v).strip().upper() in ("X", "×", "1") else ""
            )
    st.session_state.plan = new_plan
    st.rerun()


# ── TABS ───────────────────────────────────────────────
st.markdown("---")
st.markdown("### Weitere Tabellenblätter")

t1, t2, t3, t4, t5, t6 = st.tabs(
    ["Rennprofil", "Fahrerwerte", "Fahreranalyse", "Eignung", "Formkurven", "Regelwerk"]
)

with t1:
    st.caption("Excel-Blatt „Rennprofil“")
    st.dataframe(data["rennprofil"], use_container_width=True, height=480, hide_index=True)

with t2:
    st.caption("Excel-Blatt „Fahrerwerte“")
    st.dataframe(data["fahrerwerte"], use_container_width=True, height=480, hide_index=True)

with t3:
    st.caption("Excel-Blatt „Fahreranalyse“")
    st.dataframe(data["analyse"], use_container_width=True, height=480, hide_index=True)

with t4:
    st.caption("Excel-Blatt „Fahrerbewertung“ – Eignungsmatrix (0–100)")
    eign = data["eignung"]
    numeric_cols = [c for c in eign.columns if c != "Fahrer"]
    has_values = any(
        eign[c].apply(lambda x: isinstance(x, (int, float)) and not pd.isna(x)).any()
        for c in numeric_cols
    )
    if has_values:
        st.dataframe(eign, use_container_width=True, height=480, hide_index=True)
    else:
        st.info(
            "Die Eignungsmatrix ist im Excel aktuell leer (Formeln noch nicht berechnet). "
            "Vorschau der relevanten Fahrerwerte:"
        )
        st.dataframe(
            data["fahrerwerte"][["Name", "BE", "HÜG", "SP", "ZF", "KSP", "Fahrerrolle"]],
            use_container_width=True, height=400, hide_index=True,
        )

with t5:
    st.caption("Excel-Blatt „Formkurven“ – wird per Button „Formkurven berechnen“ befüllt (Teil B Regelwerk)")
    st.info(
        "Formkurven werden laut KI-Regelwerk in einem **separaten Schritt nach der Saisonplanung** "
        "berechnet. Button oben nutzen, sobald die Logik aktiv ist."
    )

with t6:
    st.caption("Excel-Blatt „KI-Regelwerk“ – Auszug")
    st.text_area("Regelwerk", value=data["regelwerk"], height=460, label_visibility="collapsed")

st.markdown("---")
st.caption("SIBOS Season Planner 2026 · Daten aus PCM_Saisonplaner.xlsx · Prototyp")
