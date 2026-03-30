"""
DSO Intake Toets — Streamlit Web App v4
State machine aanpak — één flow tegelijk, geen dubbele output
"""

import streamlit as st
import sys, os, io, builtins, tempfile
import requests as _req
from datetime import date

# ── Scripts inladen ───────────────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import importlib.util as _ilu

_echte_input = builtins.input
def _web_input(prompt=""):
    print(prompt + "1  ← automatisch gekozen")
    return "1"

def _laad(naam):
    pad = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"{naam}.py")
    spec = _ilu.spec_from_file_location(naam, pad)
    mod  = _ilu.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod

builtins.input = _web_input
_dso   = _laad("dso_bestemmingsplan")
_toets = _laad("genereer_intake_toets")
builtins.input = _echte_input

haal_data_voor_adres       = _dso.haal_data_voor_adres
haal_data_voor_coordinaten = _dso.haal_data_voor_coordinaten
genereer_intake_toets      = _toets.genereer_intake_toets

LS_BASE = "https://api.pdok.nl/bzk/locatieserver/search/v3_1"

# ── Pagina config ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="DSO Intake Toets", page_icon="🏛️", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
.stApp { background: #0d1117; color: #e6edf3; }
.dso-header {
    background: linear-gradient(135deg, #1f4e79 0%, #0d2137 100%);
    border: 1px solid #1f6feb; border-radius: 8px;
    padding: 2rem 2.5rem; margin-bottom: 1.5rem;
}
.dso-header h1 { font-family:'IBM Plex Mono',monospace; font-size:1.6rem; font-weight:600; color:#58a6ff; margin:0 0 .3rem 0; }
.dso-header p  { color:#8b949e; font-size:.9rem; margin:0; }
.stap-label    { font-family:'IBM Plex Mono',monospace; font-size:.7rem; color:#58a6ff; text-transform:uppercase; letter-spacing:2px; margin-bottom:.3rem; }
.result-card   { background:#161b22; border:1px solid #30363d; border-radius:6px; padding:1rem 1.2rem; margin:.4rem 0; }
.result-card .label  { font-size:.72rem; color:#8b949e; font-family:'IBM Plex Mono',monospace; text-transform:uppercase; letter-spacing:1px; margin-bottom:.2rem; }
.result-card .waarde { font-size:1rem; color:#e6edf3; font-weight:600; }
.result-card .waarde.auto { color:#58a6ff; }
.result-card .waarde.leeg { color:#484f58; font-style:italic; font-weight:400; }
.terminal { background:#0d1117; border:1px solid #21262d; border-radius:6px; padding:1rem 1.2rem;
    font-family:'IBM Plex Mono',monospace; font-size:.78rem; color:#3fb950; white-space:pre-wrap;
    max-height:280px; overflow-y:auto; }
.badge-auto { background:#1f3a5f; color:#58a6ff; padding:2px 10px; border-radius:20px; font-size:.75rem; margin-right:6px; }
.badge-hand { background:#3d2e00; color:#d29922; padding:2px 10px; border-radius:20px; font-size:.75rem; margin-right:6px; }
.badge-pb   { background:#3d1a1a; color:#f85149; padding:2px 10px; border-radius:20px; font-size:.75rem; }
.niet-gedig { background:#2d1f00; border:1px solid #d29922; border-radius:6px; padding:.8rem 1.2rem; color:#d29922; font-size:.85rem; margin:.5rem 0 1rem 0; }
.sectie-header { font-family:'IBM Plex Mono',monospace; font-size:.7rem; color:#58a6ff; text-transform:uppercase;
    letter-spacing:3px; border-bottom:1px solid #21262d; padding-bottom:.5rem; margin:1.5rem 0 1rem 0; }
.stDownloadButton > button { background:#238636 !important; color:white !important; border:1px solid #2ea043 !important;
    border-radius:6px !important; font-weight:600 !important; width:100% !important; font-size:1rem !important; padding:.6rem 1.5rem !important; }
.stTextInput > div > div > input { background:#161b22 !important; border:1px solid #30363d !important;
    color:#e6edf3 !important; border-radius:6px !important; font-family:'IBM Plex Mono',monospace !important; font-size:1rem !important; }
.stTextInput > div > div > input:focus { border-color:#58a6ff !important; box-shadow:0 0 0 3px rgba(88,166,255,.1) !important; }
.stNumberInput > div > div > input { background:#161b22 !important; border:1px solid #30363d !important;
    color:#e6edf3 !important; border-radius:6px !important; font-family:'IBM Plex Mono',monospace !important; }
.stButton > button { background:#1f4e79 !important; color:white !important; border:1px solid #1f6feb !important;
    border-radius:6px !important; font-weight:600 !important; width:100% !important; font-size:1rem !important; padding:.6rem 1.5rem !important; }
#MainMenu, footer, header { visibility:hidden; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="dso-header">
    <h1>🏛️ DSO Intake Toets Generator</h1>
    <p>Automatisch bestemmingsplandata ophalen en invullen in een Word-document</p>
</div>
""", unsafe_allow_html=True)
st.markdown(
    '<span class="badge-auto">Lichtblauw = automatisch (DSO)</span>'
    '<span class="badge-pb">Rood = via PowerBrowser</span>'
    '<span class="badge-hand">Geel = handmatig invullen</span>',
    unsafe_allow_html=True)
st.markdown("---")

# ── Session state initialiseren ───────────────────────────────────────────────
for k, v in {"fase": "invoer", "kandidaten": [], "gekozen": None, "data": None, "terminal_log": ""}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── Hulpfuncties ──────────────────────────────────────────────────────────────
def zoek_adressen(adres_str):
    """Geeft lijst van dicts terug met weergavenaam + rd-coördinaten."""
    params = {"q": adres_str, "fq": "type:adres", "rows": 100, "fl": "id,weergavenaam,centroide_rd"}
    r = _req.get(f"{LS_BASE}/free", params=params, timeout=10)
    r.raise_for_status()
    docs = r.json().get("response", {}).get("docs", [])
    if not docs: return []
    al = adres_str.strip().lower()
    def match(w):
        wl = w.strip().lower()
        if wl == al: return True
        d = wl.split(", ", 1)
        if len(d) == 2:
            r2 = d[1].split(" ", 1)
            p = r2[1] if len(r2) == 2 else d[1]
            if f"{d[0]}, {p}" == al or d[0] == al: return True
        return False
    gezien, res = set(), []
    for d in docs:
        n = d.get("weergavenaam", "")
        if match(n) and n not in gezien:
            gezien.add(n); res.append(d)
    return res if res else docs[:5]

def adres_naar_xy(doc):
    """Haal RD-coördinaten op uit een PDOK doc dict."""
    rd = doc.get("centroide_rd", "").replace("POINT(", "").replace(")", "").split()
    if len(rd) == 2:
        return float(rd[0]), float(rd[1])
    return None, None

def run_en_toon(fn, *args):
    """Voert fn uit, slaat terminal output op in session_state, geeft resultaat terug."""
    class Live(io.StringIO):
        def write(self, t):
            super().write(t)
            st.session_state.terminal_log = self.getvalue()
            return len(t)
    live = Live()
    old_out = sys.stdout; sys.stdout = live
    builtins.input = _web_input
    try:
        return fn(*args)
    finally:
        sys.stdout = old_out; builtins.input = _echte_input

def kaart(label, waarde):
    heeft = waarde and waarde not in ("—", "geen", "")
    return f"""<div class="result-card"><div class="label">{label}</div>
        <div class="waarde {'auto' if heeft else 'leeg'}">{waarde if heeft else '—'}</div></div>"""

def toon_resultaten(data):
    st.markdown("---")
    st.markdown('<div class="stap-label">Stap 3 — Gevonden gegevens</div>', unsafe_allow_html=True)
    if data.get("niet_gedigitaliseerd"):
        st.markdown(f"""<div class="niet-gedig">⚠️ <strong>Niet volledig gedigitaliseerd</strong> in de API.<br>
            Handmatig opzoeken via: <a href="{data.get('hyperlink','#')}" target="_blank">{data.get('hyperlink','—')}</a>
        </div>""", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(kaart("Gevonden adres",       data.get("adres_gevonden","—")), unsafe_allow_html=True)
        st.markdown(kaart("Kadastrale aanduiding", data.get("kadastrale_aanduiding","—")), unsafe_allow_html=True)
        st.markdown(kaart("Bestemmingsplan",       data.get("bestemmingsplan_naam","—")), unsafe_allow_html=True)
        st.markdown(kaart("Datum vaststelling",    data.get("bestemmingsplan_datum","—")), unsafe_allow_html=True)
    with c2:
        st.markdown(kaart("Bestemming perceel", data.get("bestemming_perceel","—")), unsafe_allow_html=True)
        st.markdown(kaart("Bestemmingstype",    data.get("bestemmingstype","—")), unsafe_allow_html=True)
        st.markdown(kaart("Functieaanduiding",  ", ".join(data.get("functieaanduidingen",[])) or "geen"), unsafe_allow_html=True)
        st.markdown(kaart("Dubbelbestemming",   ", ".join(d["naam"] for d in data.get("dubbelbestemmingen",[])) or "geen"), unsafe_allow_html=True)
    if data.get("maatvoeringen"):
        st.markdown('<div class="sectie-header">Maatvoeringen</div>', unsafe_allow_html=True)
        # Dedupliceer op naam — behoud eerste unieke waarde per naam
        gezien = {}
        for m in data["maatvoeringen"]:
            naam = m["naam"]
            if naam not in gezien:
                gezien[naam] = m
        uniek = list(gezien.values())
        cols = st.columns(3)
        for i, m in enumerate(uniek):
            with cols[i % 3]:
                st.markdown(kaart(m["naam"], f"{m['waarde']} {m.get('eenheid','')}".strip()), unsafe_allow_html=True)

def toon_download(data, label):
    st.markdown("---")
    st.markdown('<div class="stap-label">Stap 4 — Download Word-document</div>', unsafe_allow_html=True)
    with st.spinner("Word-document genereren..."):
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            tmp_pad = tmp.name
        genereer_intake_toets(data, uitvoer_pad=tmp_pad)
        with open(tmp_pad, "rb") as f: docx_bytes = f.read()
        os.unlink(tmp_pad)
    naam = f"Intake_toets_{label.replace(' ','_')[:25]}_{date.today().strftime('%Y%m%d')}.docx"
    st.download_button("📄  Download Intake Toets (.docx)", data=docx_bytes, file_name=naam,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    st.success(f"✅ Klaar! Klik op de knop om **{naam}** te downloaden.")

# ══════════════════════════════════════════════════════════════════════════════
# STATE MACHINE — één fase tegelijk
# ══════════════════════════════════════════════════════════════════════════════

# ── FASE: invoer ──────────────────────────────────────────────────────────────
st.markdown('<div class="stap-label">Stap 1 — Kies invoermethode</div>', unsafe_allow_html=True)
methode = st.radio("", ["📍 Adres", "📐 RD-coördinaten", "🧪 Testadres"],
    label_visibility="collapsed", horizontal=True)

adres_input = x_input = y_input = None
zoek_knop = False

if methode == "📍 Adres":
    c1, c2 = st.columns([4, 1])
    with c1:
        adres_input = st.text_input("Adres", placeholder="bijv. Kerkstraat 1, IJsselstein",
            label_visibility="collapsed", key="adres_veld")
    with c2:
        st.markdown("<br>", unsafe_allow_html=True)
        zoek_knop = st.button("🔍 Ophalen")

elif methode == "📐 RD-coördinaten":
    st.caption("RD-coördinaten (EPSG:28992) — te vinden in de DSO viewer of op ruimtelijkeplannen.nl")
    c1, c2, c3 = st.columns([2, 2, 1])
    with c1: x_input = st.number_input("X (RD)", value=134789, step=1, format="%d")
    with c2: y_input = st.number_input("Y (RD)", value=447145, step=1, format="%d")
    with c3:
        st.markdown("<br>", unsafe_allow_html=True)
        zoek_knop = st.button("🔍 Ophalen")

else:
    st.info("Testadres: **Graaf Walramhof 4, Nieuwegein**")
    zoek_knop = st.button("🔍 Ophalen met testadres")
    adres_input = "Graaf Walramhof 4, Nieuwegein"

# ── Nieuwe zoekopdracht → reset alles ────────────────────────────────────────
if zoek_knop:
    st.session_state.fase      = "zoeken"
    st.session_state.kandidaten = []
    st.session_state.gekozen   = None
    st.session_state.data      = None
    st.markdown("---")

    # Coördinaten: direct ophalen
    if methode == "📐 RD-coördinaten":
        st.info(f"📐 X={x_input}, Y={y_input} (RD)")
        try:
            with st.spinner("Bezig..."):
                data = run_en_toon(haal_data_voor_coordinaten, float(x_input), float(y_input))
            st.session_state.data = data
            st.session_state.fase = "resultaat"
        except Exception as e:
            st.error(f"❌ {e}"); st.stop()

    # Adres: zoek eerst kandidaten
    elif adres_input:
        with st.spinner("Adres opzoeken..."):
            try: kandidaten = zoek_adressen(adres_input.strip())
            except Exception as e: st.error(f"❌ {e}"); st.stop()

        if not kandidaten:
            st.error(f"❌ Geen adressen gevonden voor '{adres_input}'")
            st.session_state.fase = "invoer"
        elif len(kandidaten) == 1:
            # Directe match — haal coördinaten op en sla op
            gekozen = kandidaten[0].get("weergavenaam", adres_input)
            x, y = adres_naar_xy(kandidaten[0])
            st.info(f"📍 {gekozen}")
            try:
                with st.spinner("Bezig..."):
                    data = run_en_toon(haal_data_voor_coordinaten, x, y)
                data["adres_gevonden"] = gekozen
                st.session_state.data = data
                st.session_state.fase = "resultaat"
            except Exception as e:
                st.error(f"❌ {e}"); st.stop()
        else:
            # Meerdere kandidaten → keuzemenu, sla naam + coördinaten op
            st.session_state.kandidaten = [
                {"naam": d.get("weergavenaam","?"), "xy": adres_naar_xy(d)}
                for d in kandidaten
            ]
            st.session_state.fase = "keuze"
            st.rerun()
    else:
        st.warning("⚠️ Vul eerst een adres in.")
        st.session_state.fase = "invoer"

# ── FASE: keuzemenu ───────────────────────────────────────────────────────────
elif st.session_state.fase == "keuze" and st.session_state.kandidaten:
    st.markdown("---")
    st.markdown('<div class="stap-label">Meerdere adressen gevonden — kies er één</div>', unsafe_allow_html=True)
    namen = [k["naam"] for k in st.session_state.kandidaten]
    keuze_naam = st.radio("Kies het juiste adres:", namen, key="adres_keuze")
    if st.button("✓ Dit adres gebruiken"):
        # Zoek de bijbehorende coördinaten op
        gekozen_item = next(k for k in st.session_state.kandidaten if k["naam"] == keuze_naam)
        st.session_state.gekozen = gekozen_item
        st.session_state.fase    = "ophalen"
        st.rerun()

# ── FASE: ophalen na keuze ────────────────────────────────────────────────────
elif st.session_state.fase == "ophalen" and st.session_state.gekozen:
    st.markdown("---")
    gekozen = st.session_state.gekozen
    naam = gekozen["naam"]
    x, y = gekozen["xy"]
    st.info(f"📍 {naam}")
    try:
        with st.spinner("Bezig..."):
            data = run_en_toon(haal_data_voor_coordinaten, x, y)
        data["adres_gevonden"] = naam
        st.session_state.data = data
        st.session_state.fase = "resultaat"
    except Exception as e:
        st.error(f"❌ {e}"); st.session_state.fase = "invoer"; st.stop()

# ── FASE: resultaat tonen ─────────────────────────────────────────────────────
if st.session_state.fase == "resultaat" and st.session_state.data:
    data  = st.session_state.data
    label = (data.get("adres_gevonden") or data.get("adres","locatie")).split(",")[0]
    # Toon alleen de terminal log van de laatste run
    if st.session_state.get("terminal_log"):
        st.markdown('<div class="stap-label">Stap 2 — Data ophalen via DSO API</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="terminal">{st.session_state.terminal_log}</div>', unsafe_allow_html=True)
    toon_resultaten(data)
    toon_download(data, label)
    if st.button("🔄 Nieuw adres opzoeken"):
        st.session_state.fase = "invoer"
        st.session_state.data = None
        st.rerun()
