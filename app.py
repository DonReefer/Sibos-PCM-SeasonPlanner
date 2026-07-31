"""
SIBOS / PCM Season Planner 2026
- Helles modernes Grau (Theme + CSS)
- Planung: HTML-Tabelle mit hochkanten Rennnamen/Daten, horizontal scrollbar
- X bearbeiten ohne schwarze Tabelle (Fahrer wählen + Rennen anhaken)
- Eignung: berechnete Matrix Fahrer × Rennen mit Farbskala
"""

import streamlit as st
import pandas as pd
import openpyxl
from pathlib import Path
import io
import re
from datetime import datetime
import html as html_lib

st.set_page_config(
    page_title="PCM Season Planner 2026",
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
  header[data-testid="stHeader"] { background: transparent !important; height: 0 !important; }
  [data-testid="stToolbar"], [data-testid="stDecoration"] { display: none !important; }
  #MainMenu, footer { visibility: hidden; }
  .stApp { background: #e8eaed !important; color: #1f2937 !important; }
  .block-container { padding-top: 1.6rem !important; padding-bottom: 2rem !important; max-width: 100% !important; }
  .stButton > button {
    background: #f3f4f6 !important; color: #1f2937 !important;
    border: 1.5px solid #9ca3af !important; border-radius: 8px !important;
    font-weight: 500 !important; height: 2.4rem !important; box-shadow: none !important;
  }
  .stButton > button:hover { background: #e5e7eb !important; border-color: #6b7280 !important; }
  .stButton > button[kind="primary"] {
    background: #4b5563 !important; border-color: #374151 !important; color: #f9fafb !important; font-weight: 600 !important;
  }
  .stButton > button[kind="secondary"] {
    background: #e5e7eb !important; border-color: #9ca3af !important; color: #374151 !important; font-size: 0.88rem !important;
  }
  .stTextInput input, .stSelectbox [data-baseweb="select"] > div,
  .stMultiSelect [data-baseweb="select"] > div {
    background: #f9fafb !important; border: 1.5px solid #9ca3af !important;
    border-radius: 8px !important; color: #1f2937 !important;
  }
  .stTabs [data-baseweb="tab-list"] {
    background: #f3f4f6; border-radius: 10px; gap: 3px; padding: 4px; border: 1.5px solid #9ca3af;
  }
  .stTabs [data-baseweb="tab"] { color: #4b5563; border-radius: 8px; font-size: 0.88rem; }
  .stTabs [aria-selected="true"] { background: #4b5563 !important; color: #f9fafb !important; }
  div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {
    background: #f3f4f6 !important; border: 1.5px solid #9ca3af !important; border-radius: 8px !important;
  }
  .import-hint {
    background: #f3f4f6; border: 1.5px solid #9ca3af; border-radius: 10px;
    padding: 12px 14px; font-size: 0.9rem; color: #374151; margin-top: 8px;
  }
  .plan-wrap {
    overflow-x: auto; overflow-y: visible; border: 1.5px solid #9ca3af;
    border-radius: 10px; background: #f3f4f6; margin-bottom: 0.5rem; padding-bottom: 4px;
  }
  .plan-table { border-collapse: collapse; font-size: 0.85rem; color: #1f2937; background: #f3f4f6; }
  .plan-table th, .plan-table td {
    border: 1px solid #c4c9d1; padding: 3px 5px; text-align: center; vertical-align: middle; background: #f3f4f6;
  }
  .plan-table thead th { background: #e5e7eb; font-weight: 600; }
  .plan-table th.vert {
    writing-mode: vertical-rl; transform: rotate(180deg); height: 160px;
    min-width: 34px; max-width: 42px; white-space: nowrap; font-size: 0.82rem;
    letter-spacing: 0.03em; padding: 8px 4px;
  }
  .plan-table th.horiz { min-width: 34px; max-width: 48px; font-size: 0.78rem; white-space: nowrap; }
  .plan-table th.corner { background: #e5e7eb; min-width: 140px; text-align: left; padding-left: 10px; font-size: 0.85rem; }
  .plan-table td.rider {
    text-align: left; font-weight: 500; white-space: nowrap; min-width: 140px; padding-left: 10px; background: #eef0f3;
  }
  .plan-table td.days { min-width: 42px; font-weight: 600; background: #eef0f3; }
  .plan-table td.days.over { background: #fecaca; color: #991b1b; }
  .plan-table td.xcell { min-width: 32px; max-width: 40px; font-weight: 700; font-size: 0.9rem; }
  .plan-table td.xcell.mark { background: #bbf7d0; color: #14532d; }
  .plan-table td.xcell.conflict { background: #fecaca; color: #991b1b; }
  .plan-table td.xcell.full { background: #d1fae5; }
  .plan-table th.occ-ok { background: #dcfce7; color: #166534; }
  .plan-table th.occ-full { background: #86efac; color: #14532d; font-weight: 700; }
  .plan-table th.occ-over { background: #fecaca; color: #991b1b; font-weight: 700; }
  .plan-footer { font-size: 0.8rem; color: #4b5563; padding: 6px 4px; }
  .eign-wrap {
    overflow-x: auto; border: 1.5px solid #9ca3af; border-radius: 10px; background: #f3f4f6; margin-bottom: 0.5rem;
  }
  .eign-table { border-collapse: collapse; font-size: 0.72rem; color: #1f2937; }
  .eign-table th, .eign-table td { border: 1px solid #c4c9d1; padding: 2px 4px; text-align: center; background: #f3f4f6; }
  .eign-table th.vert {
    writing-mode: vertical-rl; transform: rotate(180deg); height: 120px;
    min-width: 26px; max-width: 32px; font-size: 0.7rem; background: #e5e7eb;
  }
  .eign-table th.corner { min-width: 130px; text-align: left; padding-left: 8px; background: #e5e7eb; }
  .eign-table td.rider { text-align: left; white-space: nowrap; padding-left: 8px; background: #eef0f3; font-weight: 500; }
</style>
""", unsafe_allow_html=True)

def parse_date_range(s, year=2026):
    if not s or not isinstance(s, str):
        return None, None
    s = s.strip().rstrip(".")
    m = re.match(r"^(\d{1,2})\.(\d{1,2})$", s)
    if m:
        try:
            d = datetime(year, int(m.group(2)), int(m.group(1)))
            return d, d
        except ValueError:
            return None, None
    m = re.match(r"^(\d{1,2})\.-(\d{1,2})\.(\d{1,2})$", s)
    if m:
        try:
            return (datetime(year, int(m.group(3)), int(m.group(1))), datetime(year, int(m.group(3)), int(m.group(2))))
        except ValueError:
            return None, None
    m = re.match(r"^(\d{1,2})\.(\d{1,2})\.-(\d{1,2})\.(\d{1,2})$", s)
    if m:
        try:
            return (datetime(year, int(m.group(2)), int(m.group(1))), datetime(year, int(m.group(4)), int(m.group(3))))
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

    plan_rows = []
    for i, rider in enumerate(riders):
        row_idx = 7 + i
        row = {"Fahrer": rider["Name"]}
        for race in races:
            val = ws.cell(row_idx, race["col"]).value
            row[race["name"]] = "X" if val in ("X", "x") else ""
        plan_rows.append(row)
    df_plan = pd.DataFrame(plan_rows)

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
        "fahrerwerte": df_fw, "rennprofil": df_rp, "planung": df_plan,
        "analyse": df_fa, "races": races, "regelwerk": regel_text,
    }


def _num(v, default=0):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return default
        return float(v)
    except (TypeError, ValueError):
        return default


def compute_eignung(df_fw, df_rp, races):
    profile = {}
    for _, r in df_rp.iterrows():
        key = str(r["Rennen"]).strip()
        profile[key] = {
            "berg": _num(r.get("Bergfaktor"), 5),
            "sprint": _num(r.get("Sprintfaktor"), 5),
            "huegel": _num(r.get("Hügelfaktor"), 5),
            "tt": _num(r.get("TT-Faktor"), 3),
            "ksp": 8 if str(r.get("Kopfsteinpflaster", "")).lower() in ("ja", "teilweise") else 2,
        }

    def factors_for(race_name):
        if race_name in profile:
            return profile[race_name]
        rn = race_name.lower()
        for k, v in profile.items():
            if k.lower() in rn or rn in k.lower():
                return v
        return {"berg": 5, "sprint": 5, "huegel": 5, "tt": 3, "ksp": 2}

    rows = []
    for _, rider in df_fw.iterrows():
        be = _num(rider.get("BE"), 70)
        sp = _num(rider.get("SP"), 70)
        hug = _num(rider.get("HÜG"), 70)
        zf = _num(rider.get("ZF"), 70)
        ksp = _num(rider.get("KSP"), 70)
        row = {"Fahrer": rider["Name"]}
        for race in races:
            f = factors_for(race["name"])
            wsum = f["berg"] + f["sprint"] + f["huegel"] + f["tt"] + f["ksp"] or 1
            raw = (be * f["berg"] + sp * f["sprint"] + hug * f["huegel"] + zf * f["tt"] + ksp * f["ksp"]) / wsum
            row[race["name"]] = int(max(0, min(100, round(raw))))
        rows.append(row)
    return pd.DataFrame(rows)


def score_color(score):
    s = max(0, min(100, int(score)))
    if s >= 80:
        return "#86efac", "#14532d"
    if s >= 65:
        return "#bbf7d0", "#166534"
    if s >= 50:
        return "#fef08a", "#854d0e"
    if s >= 35:
        return "#fdba74", "#9a3412"
    return "#fecaca", "#991b1b"


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


def build_plan_html(plan, races, renntage, occ, conflicts):
    esc = html_lib.escape
    rows = []
    cells = ['<th class="corner">Fahrer</th>', '<th class="corner">Renntage</th>']
    for race in races:
        cells.append(f'<th class="vert" title="{esc(race["name"])}">{esc(race["name"])}</th>')
    rows.append("<tr>" + "".join(cells) + "</tr>")
    cells = ['<th class="corner"></th>', '<th class="corner"></th>']
    for race in races:
        cells.append(f'<th class="vert">{esc(race["datum"])}</th>')
    rows.append("<tr>" + "".join(cells) + "</tr>")
    cells = ['<th class="corner"></th>', '<th class="corner"></th>']
    for race in races:
        cells.append(f'<th class="horiz">{race["etappen"]} Et.</th>')
    rows.append("<tr>" + "".join(cells) + "</tr>")
    cells = ['<th class="corner"></th>', '<th class="corner"></th>']
    for race in races:
        cnt = occ[race["name"]]
        label = f"{cnt}/{MAX_RIDERS}"
        cls = "horiz occ-over" if cnt > MAX_RIDERS else ("horiz occ-full" if cnt == MAX_RIDERS else "horiz occ-ok")
        cells.append(f'<th class="{cls}">{label}</th>')
    rows.append("<tr>" + "".join(cells) + "</tr>")
    thead = "<thead>" + "".join(rows) + "</thead>"
    body_rows = []
    for idx, row in plan.iterrows():
        days = int(renntage.loc[idx]) if idx in renntage.index else 0
        day_cls = "days over" if days > MAX_RENNTAGE else "days"
        cells = [f'<td class="rider">{esc(str(row["Fahrer"]))}</td>', f'<td class="{day_cls}">{days}</td>']
        for race in races:
            marked = row.get(race["name"]) == "X"
            conf = (idx, race["name"]) in conflicts
            cnt = occ[race["name"]]
            classes = ["xcell"]
            if marked and conf:
                classes.append("conflict")
            elif marked and cnt >= MAX_RIDERS:
                classes.extend(["full", "mark"])
            elif marked:
                classes.append("mark")
            cells.append(f'<td class="{" ".join(classes)}">{"X" if marked else ""}</td>')
        body_rows.append("<tr>" + "".join(cells) + "</tr>")
    return f'<div class="plan-wrap"><table class="plan-table">{thead}<tbody>{"".join(body_rows)}</tbody></table></div>'


def build_eign_html(df_eign, races):
    esc = html_lib.escape
    cells = ['<th class="corner">Fahrer</th>']
    for race in races:
        cells.append(f'<th class="vert" title="{esc(race["name"])}">{esc(race["name"][:18])}</th>')
    thead = "<thead><tr>" + "".join(cells) + "</tr></thead>"
    body = []
    for _, row in df_eign.iterrows():
        cells = [f'<td class="rider">{esc(str(row["Fahrer"]))}</td>']
        for race in races:
            sc = int(row.get(race["name"], 0) or 0)
            bg, fg = score_color(sc)
            cells.append(f'<td style="background:{bg};color:{fg};font-weight:600;min-width:28px;">{sc}</td>')
        body.append("<tr>" + "".join(cells) + "</tr>")
    return f'<div class="eign-wrap"><table class="eign-table">{thead}<tbody>{"".join(body)}</tbody></table></div>'


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
df_eign = compute_eignung(data["fahrerwerte"], data["rennprofil"], races)

c_logo, c_search, c_ki, c_imp, c_exp, c_run, c_form = st.columns([2.0, 1.4, 1.2, 1.0, 1.0, 1.4, 1.4])
with c_logo:
    if LOGO.exists():
        st.image(str(LOGO), width=280)
    else:
        st.markdown("### SIBOS")
with c_search:
    st.text_input("Suche", placeholder="Suchen…", label_visibility="collapsed", key="search")
with c_ki:
    st.session_state.ai = st.selectbox(
        "KI",
        ["Grok (xAI)", "Claude (Anthropic)", "GPT-4o (OpenAI)", "Gemini (Google)", "Lokal (Ollama)"],
        label_visibility="collapsed",
    )
with c_imp:
    if st.button("Importieren", use_container_width=True):
        st.session_state.show_import = not st.session_state.show_import
with c_exp:
    if st.button("Exportieren", use_container_width=True):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            plan.to_excel(w, sheet_name="Planung", index=False)
            data["fahrerwerte"].to_excel(w, sheet_name="Fahrerwerte", index=False)
            data["rennprofil"].to_excel(w, sheet_name="Rennprofil", index=False)
            data["analyse"].to_excel(w, sheet_name="Fahreranalyse", index=False)
            df_eign.to_excel(w, sheet_name="Fahrerbewertung", index=False)
        st.download_button(
            "Download Excel", data=buf.getvalue(),
            file_name="SIBOS_Saisonplanung_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
with c_run:
    if st.button("KI-Planung starten", type="primary", use_container_width=True):
        st.session_state.msg = f"**KI-Planung** mit {st.session_state.ai}: Regelwerk Teil A. Vollständige Automatik folgt."
        st.toast("KI-Planung vorbereitet", icon="⚡")
with c_form:
    if st.button("Formkurven berechnen", type="secondary", use_container_width=True):
        st.session_state.msg = "**Formkurven:** Separater Schritt nach der Saisonplanung (Teil B)."
        st.toast("Formkurven-Schritt vorbereitet", icon="📈")

if st.session_state.show_import:
    st.markdown("---")
    st.markdown("### Daten importieren")
    ic1, ic2 = st.columns(2)
    with ic1:
        st.markdown("**Excel-Datei**")
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
        st.markdown("**Foto / Screenshot**")
        up_img = st.file_uploader("Bild wählen", type=["png", "jpg", "jpeg", "webp"], key="up_img", label_visibility="collapsed")
        if up_img is not None:
            st.session_state.msg = f"Foto „{up_img.name}“ empfangen. KI-OCR folgt in der nächsten Version."
            st.session_state.show_import = False
    st.markdown(
        '<div class="import-hint"><b>Anleitung Foto / Screenshot</b><br>'
        "Fotografiere den Bildschirm im Spiel (PCM) mit <b>Fahrernamen + Werte</b> "
        "(EB, BE, MG, HÜG, ZF, PRL, KSP, SP, …). Die App übernimmt sie ins Blatt Fahrerwerte.</div>",
        unsafe_allow_html=True,
    )

if st.session_state.msg:
    st.info(st.session_state.msg)

st.markdown("---")
st.markdown("### PCM Season Planner")
st.markdown(build_plan_html(plan, races, renntage, occ, conflicts), unsafe_allow_html=True)
st.markdown(
    '<div class="plan-footer">'
    "Kopfzeilen: <b>Rennen → Datum → Etappen → Belegung (aktuell/max)</b> · "
    "Grün = Maximum · Rot = überschritten / Überschneidung · Renntage ab 71 = über Maximum · "
    "Tabelle horizontal scrollbar"
    "</div>",
    unsafe_allow_html=True,
)
if conflicts:
    names = sorted({plan.loc[i, "Fahrer"] for i, _ in conflicts})
    st.warning("Zeitliche Überschneidung bei: " + ", ".join(names))

st.markdown("#### X setzen / entfernen")
ec1, ec2 = st.columns([1.2, 2.8])
with ec1:
    rider_names = plan["Fahrer"].tolist()
    sel_rider = st.selectbox("Fahrer", rider_names, key="edit_rider")
with ec2:
    idx = plan.index[plan["Fahrer"] == sel_rider][0]
    current = [r["name"] for r in races if plan.loc[idx, r["name"]] == "X"]
    race_names = [r["name"] for r in races]
    new_sel = st.multiselect(
        "Rennen für diesen Fahrer (anhaken = X)",
        options=race_names, default=current, key=f"ms_{sel_rider}",
    )
    if st.button("Übernehmen", type="primary", key="apply_x"):
        new_plan = plan.copy()
        for rname in race_names:
            new_plan.loc[idx, rname] = "X" if rname in new_sel else ""
        st.session_state.plan = new_plan
        st.session_state.msg = f"X für **{sel_rider}** aktualisiert."
        st.rerun()

st.markdown("---")
st.markdown("### Weitere Tabellenblätter")
t1, t2, t3, t4, t5, t6 = st.tabs(
    ["Rennprofil", "Fahrerwerte", "Fahreranalyse", "Eignung", "Formkurven", "Regelwerk"]
)
TABLE_H = 1100
with t1:
    st.caption("Excel-Blatt „Rennprofil“")
    st.dataframe(data["rennprofil"], use_container_width=True, height=TABLE_H, hide_index=True)
with t2:
    st.caption("Excel-Blatt „Fahrerwerte“")
    st.dataframe(data["fahrerwerte"], use_container_width=True, height=TABLE_H, hide_index=True)
with t3:
    st.caption("Excel-Blatt „Fahreranalyse“")
    st.dataframe(data["analyse"], use_container_width=True, height=TABLE_H, hide_index=True)
with t4:
    st.caption(
        "Eignungsmatrix (0–100): Fahrerwerte × Rennprofil (Berg/Sprint/Hügel/TT/Pflaster) – "
        "Farbskala Rot → Gelb → Grün"
    )
    st.markdown(build_eign_html(df_eign, races), unsafe_allow_html=True)
    st.caption("Horizontal scrollbar · aus Fahrerwerte + Rennprofil berechnet")
with t5:
    st.caption("Formkurven – Button „Formkurven berechnen“ (Teil B Regelwerk)")
    st.info("Formkurven werden laut Regelwerk in einem separaten Schritt nach der Saisonplanung berechnet.")
with t6:
    st.caption("Excel-Blatt „KI-Regelwerk“ – Auszug")
    st.text_area("Regelwerk", value=data["regelwerk"], height=500, label_visibility="collapsed")

st.markdown("---")
st.caption("PCM Season Planner 2026 · SIBOS · Daten aus PCM_Saisonplaner.xlsx")
