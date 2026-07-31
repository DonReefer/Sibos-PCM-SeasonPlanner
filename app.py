"""
SIBOS Season Planner 2026
Optik: helles Grau · Logo groß · Buttons grau umrandet
Planung: Rennname | Datum | Etappen | Belegung (x/7) als Kopfzeilen
         Rennname + Datum hochkant (90°)
         Fußzeile unter der Tabelle
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
  header[data-testid="stHeader"] {
    background: transparent !important;
    height: 0 !important;
  }
  [data-testid="stToolbar"] { display: none !important; }
  [data-testid="stDecoration"] { display: none !important; }
  #MainMenu { visibility: hidden; }
  footer { visibility: hidden; }

  .stApp {
    background: #e8eaed !important;
    color: #1f2937 !important;
  }
  .block-container {
    padding-top: 1.8rem !important;
    padding-bottom: 2rem !important;
    max-width: 100% !important;
  }

  .stButton > button {
    background: #f3f4f6 !important;
    color: #1f2937 !important;
    border: 1.5px solid #9ca3af !important;
    border-radius: 8px !important;
    font-weight: 500 !important;
    height: 2.4rem !important;
    box-shadow: none !important;
  }
  .stButton > button:hover {
    background: #e5e7eb !important;
    border-color: #6b7280 !important;
  }
  .stButton > button[kind="primary"] {
    background: #4b5563 !important;
    border-color: #374151 !important;
    color: #f9fafb !important;
    font-weight: 600 !important;
  }
  .stButton > button[kind="primary"]:hover {
    background: #374151 !important;
  }
  .stButton > button[kind="secondary"] {
    background: #e5e7eb !important;
    border-color: #9ca3af !important;
    color: #374151 !important;
    font-size: 0.88rem !important;
  }

  .stTextInput input, .stSelectbox [data-baseweb="select"] > div {
    background: #f9fafb !important;
    border: 1.5px solid #9ca3af !important;
    border-radius: 8px !important;
    color: #1f2937 !important;
  }

  .stTabs [data-baseweb="tab-list"] {
    background: #f3f4f6;
    border-radius: 10px;
    gap: 3px;
    padding: 4px;
    border: 1.5px solid #9ca3af;
  }
  .stTabs [data-baseweb="tab"] {
    color: #4b5563;
    border-radius: 8px;
    font-size: 0.88rem;
  }
  .stTabs [aria-selected="true"] {
    background: #4b5563 !important;
    color: #f9fafb !important;
  }

  div[data-testid="stDataFrame"],
  div[data-testid="stDataEditor"] {
    background: #f3f4f6 !important;
    border: 1.5px solid #9ca3af !important;
    border-radius: 8px !important;
  }

  .import-hint {
    background: #f3f4f6;
    border: 1.5px solid #9ca3af;
    border-radius: 10px;
    padding: 12px 14px;
    font-size: 0.9rem;
    color: #374151;
    margin-top: 8px;
  }

  .plan-wrap {
    overflow-x: auto;
    overflow-y: visible;
    border: 1.5px solid #9ca3af;
    border-radius: 8px;
    background: #f3f4f6;
    margin-bottom: 0.4rem;
  }
  .plan-table {
    border-collapse: collapse;
    font-size: 0.78rem;
    color: #1f2937;
    background: #f3f4f6;
    min-width: 100%;
  }
  .plan-table th, .plan-table td {
    border: 1px solid #c4c9d1;
    padding: 2px 4px;
    text-align: center;
    vertical-align: middle;
    background: #f3f4f6;
  }
  .plan-table thead th {
    background: #e5e7eb;
    font-weight: 600;
  }
  .plan-table th.vert {
    writing-mode: vertical-rl;
    transform: rotate(180deg);
    height: 130px;
    max-width: 28px;
    min-width: 22px;
    white-space: nowrap;
    font-size: 0.72rem;
    letter-spacing: 0.02em;
    padding: 6px 2px;
  }
  .plan-table th.horiz {
    min-width: 26px;
    max-width: 36px;
    font-size: 0.7rem;
    white-space: nowrap;
  }
  .plan-table th.corner {
    background: #e5e7eb;
    min-width: 120px;
    text-align: left;
    padding-left: 8px;
  }
  .plan-table td.rider {
    text-align: left;
    font-weight: 500;
    white-space: nowrap;
    min-width: 130px;
    padding-left: 8px;
    background: #eef0f3;
  }
  .plan-table td.days {
    min-width: 36px;
    font-weight: 600;
    background: #eef0f3;
  }
  .plan-table td.days.over {
    background: #fecaca;
    color: #991b1b;
  }
  .plan-table td.xcell {
    min-width: 24px;
    max-width: 30px;
    font-weight: 700;
  }
  .plan-table td.xcell.mark {
    background: #bbf7d0;
    color: #14532d;
  }
  .plan-table td.xcell.conflict {
    background: #fecaca;
    color: #991b1b;
  }
  .plan-table td.xcell.full {
    background: #d1fae5;
  }
  .plan-table th.occ-ok { background: #dcfce7; color: #166534; }
  .plan-table th.occ-full { background: #86efac; color: #14532d; font-weight: 700; }
  .plan-table th.occ-over { background: #fecaca; color: #991b1b; font-weight: 700; }

  .plan-footer {
    font-size: 0.78rem;
    color: #4b5563;
    padding: 6px 4px 2px 4px;
  }
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
            return (
                datetime(year, int(m.group(3)), int(m.group(1))),
                datetime(year, int(m.group(3)), int(m.group(2))),
            )
        except ValueError:
            return None, None
    m = re.match(r"^(\d{1,2})\.(\d{1,2})\.-(\d{1,2})\.(\d{1,2})$", s)
    if m:
        try:
            return (
                datetime(year, int(m.group(2)), int(m.group(1))),
                datetime(year, int(m.group(4)), int(m.group(3))),
            )
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
        if cnt > MAX_RIDERS:
            cls = "horiz occ-over"
        elif cnt == MAX_RIDERS:
            cls = "horiz occ-full"
        else:
            cls = "horiz occ-ok"
        cells.append(f'<th class="{cls}">{label}</th>')
    rows.append("<tr>" + "".join(cells) + "</tr>")
    thead = "<thead>" + "".join(rows) + "</thead>"

    body_rows = []
    for idx, row in plan.iterrows():
        days = int(renntage.loc[idx]) if idx in renntage.index else 0
        day_cls = "days over" if days > MAX_RENNTAGE else "days"
        cells = [
            f'<td class="rider">{esc(str(row["Fahrer"]))}</td>',
            f'<td class="{day_cls}">{days}</td>',
        ]
        for race in races:
            marked = row.get(race["name"]) == "X"
            conf = (idx, race["name"]) in conflicts
            cnt = occ[race["name"]]
            classes = ["xcell"]
            if marked and conf:
                classes.append("conflict")
            elif marked and cnt >= MAX_RIDERS:
                classes.append("full")
                classes.append("mark")
            elif marked:
                classes.append("mark")
            text = "X" if marked else ""
            cells.append(f'<td class="{" ".join(classes)}">{text}</td>')
        body_rows.append("<tr>" + "".join(cells) + "</tr>")
    tbody = "<tbody>" + "".join(body_rows) + "</tbody>"
    return f'<div class="plan-wrap"><table class="plan-table">{thead}{tbody}</table></div>'


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

c_logo, c_search, c_ki, c_imp, c_exp, c_run, c_form = st.columns(
    [2.0, 1.4, 1.2, 1.0, 1.0, 1.4, 1.4]
)

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
            data["eignung"].to_excel(w, sheet_name="Fahrerbewertung", index=False)
        st.download_button(
            "Download Excel",
            data=buf.getvalue(),
            file_name="SIBOS_Saisonplanung_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

with c_run:
    if st.button("KI-Planung starten", type="primary", use_container_width=True):
        st.session_state.msg = (
            f"**KI-Planung** mit {st.session_state.ai}: Regelwerk Teil A "
            "(alte X löschen → neue setzen → QC). Vollständige Automatik folgt."
        )
        st.toast("KI-Planung vorbereitet", icon="⚡")

with c_form:
    if st.button("Formkurven berechnen", type="secondary", use_container_width=True):
        st.session_state.msg = (
            "**Formkurven:** Separater Schritt nach der Saisonplanung (Teil B Regelwerk). "
            "Umsetzung folgt."
        )
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
        Am besten so, dass Namen und alle Zahlen gut lesbar sind.<br>
        Die App übernimmt die Namen und Werte dann automatisch ins Blatt <b>Fahrerwerte</b>.
        </div>
        """,
        unsafe_allow_html=True,
    )

if st.session_state.msg:
    st.info(st.session_state.msg)

st.markdown("---")
st.markdown("### PCM Season Planner")

st.markdown(
    build_plan_html(plan, races, renntage, occ, conflicts),
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="plan-footer">'
    "Kopfzeilen: <b>Rennen → Datum → Etappen → Belegung (aktuell/max)</b> · "
    "Grün = Maximum erreicht · Rot = überschritten · "
    "Rote X = zeitliche Überschneidung · Renntage ab 71 = über Maximum (70)"
    "</div>",
    unsafe_allow_html=True,
)

if conflicts:
    names = sorted({plan.loc[i, "Fahrer"] for i, _ in conflicts})
    st.warning("Zeitliche Überschneidung bei: " + ", ".join(names))

st.caption("X setzen / entfernen (manuell):")
edit_df = plan.copy()
edit_df.insert(1, "Renntage", renntage)
edited = st.data_editor(
    edit_df,
    use_container_width=True,
    height=1150,
    hide_index=True,
    disabled=["Fahrer", "Renntage"],
    key="plan_editor",
)

if not edited.equals(edit_df):
    new_plan = plan.copy()
    for race in races:
        if race["name"] in edited.columns:
            new_plan[race["name"]] = edited[race["name"]].apply(
                lambda v: "X" if str(v).strip().upper() in ("X", "×", "1") else ""
            )
    st.session_state.plan = new_plan
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
    st.caption("Excel-Blatt „Fahrerbewertung“ – Eignungsmatrix")
    eign = data["eignung"]
    numeric_cols = [c for c in eign.columns if c != "Fahrer"]
    has_values = any(
        eign[c].apply(lambda x: isinstance(x, (int, float)) and not pd.isna(x)).any()
        for c in numeric_cols
    )
    if has_values:
        st.dataframe(eign, use_container_width=True, height=TABLE_H, hide_index=True)
    else:
        st.info(
            "Die Eignungsmatrix ist im Excel aktuell leer (Formeln noch nicht berechnet). "
            "Vorschau der relevanten Fahrerwerte:"
        )
        st.dataframe(
            data["fahrerwerte"][["Name", "BE", "HÜG", "SP", "ZF", "KSP", "Fahrerrolle"]],
            use_container_width=True,
            height=TABLE_H,
            hide_index=True,
        )

with t5:
    st.caption("Excel-Blatt „Formkurven“ – per Button „Formkurven berechnen“ (Teil B Regelwerk)")
    st.info(
        "Formkurven werden laut KI-Regelwerk in einem **separaten Schritt nach der Saisonplanung** "
        "berechnet. Button oben rechts nutzen, sobald die Logik aktiv ist."
    )

with t6:
    st.caption("Excel-Blatt „KI-Regelwerk“ – Auszug")
    st.text_area("Regelwerk", value=data["regelwerk"], height=500, label_visibility="collapsed")

st.markdown("---")
st.caption("SIBOS Season Planner 2026 · Daten aus PCM_Saisonplaner.xlsx")
