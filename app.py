"""
DSO Intake Toets — Streamlit Web App v3
========================================
Invoermethodes:
  1. Adres (met keuzemenu bij meerdere resultaten)
  2. RD-coördinaten
  3. Testadres

Gebruikt originele scripts:
  - dso_bestemmingsplan.py  → data ophalen
  - genereer_intake_toets.py → Word genereren
"""

import streamlit as st
import sys
import os
import io
import builtins
import tempfile
import requests as _req
from datetime import date

# ── Originele scripts inladen ─────────────────────────────────────────────────
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

# ── Pagina-instellingen ───────────────────────────────────────────────────────
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
.dso-header h1 {
    font-family: 'IBM Plex Mono', monospace; font-size: 1.6rem;
    font-weight: 600; color: #58a6ff; margin: 0 0 0.3rem 0;
}
.dso-header p { color: #8b949e; font-size: 0.9rem; margin: 0; }
.stap-label {
    font-family: 'IBM Plex Mono', monospace; font-size: 0.7rem;
    color: #58a6ff; text-transform: uppercase; letter-spacing: 2px; margin-bottom: 0.3rem;
}
.result-card {
    background: #161b22; border: 1px solid #30363d;
    border-radius: 6px; padding: 1rem 1.2rem; margin: 0.4rem 0;
}
.result-card .label {
    font-size: 0.72rem; color: #8b949e; font-family: 'IBM Plex Mono', monospace;
    text-transform: uppercase; letter-spacing: 1px; margin-bottom: 0.2rem;
}
.result-card .waarde      { font-size: 1rem; color: #e6edf3; font-weight: 600; }
.result-card .waarde.auto { color: #58a6ff; }
.result-card .waarde.leeg { color: #484f58; font-style: italic; font-weight: 400; }
.terminal {
    background: #0d1117; border: 1px solid #21262d; border-radius: 6px;
    padding: 1rem 1.2rem; font-family: 'IBM Plex Mono', monospace;
    font-size: 0.78rem; color: #3fb950; white-space: pre-wrap;
    max-height: 280px; overflow-y: auto;
}
.badge-auto { background:#1f3a5f; color:#58a6ff; padding:2px 10px; border-radius:20px; font-size:0.75rem; margin-right:6px; }
.badge-hand { background:#3d2e00; color:#d29922; padding:2px 10px; border-radius:20px; font-size:0.75rem; margin-right:6px; }
.badge-pb   { background:#3d1a1a; color:#f85149; padding:2px 10px; border-radius:20px; font-size:0.75rem; }
.niet-gedig {
    background: #2d1f00; border: 1px solid #d29922; border-radius: 6px;
    padding: 0.8rem 1.2rem; color: #d29922; font-size: 0.85rem; margin: 0.5rem 0 1rem 0;
}
.sectie-header {
    font-family: 'IBM Plex Mono', monospace; font-size: 0.7rem; color: #58a6ff;
    text-transform: uppercase; letter-spacing: 3px;
    border-bottom: 1px solid #21262d; padding-bottom: 0.5rem; margin: 1.5rem 0 1rem 0;
}
.stDownloadButton > button {
    background: #238636 !important; color: white !important;
    border: 1px solid #2ea043 !important; border-radius: 6px !important;
    font-family: 'IBM Plex Sans', sans-serif !important; font-weight: 600 !important;
    width: 100% !important; font-size: 1rem !important; padding: 0.6rem 1.5rem !important;
}
.stTextInput > div > div > input {
    background: #161b22 !important; border: 1px solid #30363d !important;
    color: #e6edf3 !important; border-radius: 6px !important;
    font-family: 'IBM Plex Mono', monospace !important; font-size: 1rem !important;
}
.stTextInput > div > div > input:focus {
    border-color: #58a6ff !important; box-shadow: 0 0 0 3px rgba(88,166,255,0.1) !important;
}
.stNumberInput > div > div > input {
    background: #161b22 !important; border: 1px solid #30363d !important;
    color: #e6edf3 !important; border-radius: 6px !important;
    font-family: 'IBM Plex Mono', monospace !important;
}
.stButton > button {
    background: #1f4e79 !important; color: white !important;
    border: 1px solid #1f6feb !important; border-radius: 6px !important;
    font-family: 'IBM Plex Sans', sans-serif !important; font-weight: 600 !important;
    width: 100% !important; font-size: 1rem !important; padding: 0.6rem 1.5rem !important;
}
#MainMenu, footer, header { visibility: hidden; }
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
    unsafe_allow_html=True
)
st.markdown("---")

# ── Hulpfunctie: adres zoeken zonder input() ──────────────────────────────────
def zoek_adressen(adres_str):
    params = {"q": adres_str, "fq": "type:adres", "rows": 100,
              "fl": "id,weergavenaam,centroide_rd"}
    r = _req.get(f"{LS_BASE}/free", params=params, timeout=10)
    r.raise_for_status()
    docs = r.json().get("response", {}).get("docs", [])
    if not docs:
        return []
    adres_lower = adres_str.strip().lower()
    def is_match(w):
        wl = w.strip().lower()
        if wl == adres_lower: return True
        delen = wl.split(", ", 1)
        if len(delen) == 2:
            rest = delen[1].split(" ", 1)
            plaats = rest[1] if len(rest) == 2 else delen[1]
            if f"{delen[0]}, {plaats}" == adres_lower: return True
            if delen[0] == adres_lower: return True
        return False
    gezien, exacte = set(), []
    for d in docs:
        naam = d.get("weergavenaam", "")
        if is_match(naam) and naam not in gezien:
            gezien.add(naam); exacte.append(d)
    return exacte if exacte else docs[:5]

# ── Invoer ────────────────────────────────────────────────────────────────────
st.markdown('<div class="stap-label">Stap 1 — Kies invoermethode</div>', unsafe_allow_html=True)

methode = st.radio(
    "Invoermethode",
    ["📍 Adres", "📐 RD-coördinaten", "🧪 Testadres"],
    label_visibility="collapsed", horizontal=True,
)

adres_input = None
x_input = y_input = None
zoek_knop = False

if methode == "📍 Adres":
    col1, col2 = st.columns([4, 1])
    with col1:
        adres_input = st.text_input(
            "Adres", placeholder="bijv. Kerkstraat 1, IJsselstein",
            label_visibility="collapsed", key="adres_veld"
        )
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        zoek_knop = st.button("🔍 Ophalen")

elif methode == "📐 RD-coördinaten":
    st.caption("RD-coördinaten (EPSG:28992) — te vinden in de DSO viewer of op ruimtelijkeplannen.nl")
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        x_input = st.number_input("X (RD)", value=134789, step=1, format="%d")
    with col2:
        y_input = st.number_input("Y (RD)", value=447145, step=1, format="%d")
    with col3:
        st.markdown("<br>", unsafe_allow_html=True)
        zoek_knop = st.button("🔍 Ophalen")

else:  # Testadres
    st.info("Testadres: **Graaf Walramhof 4, Nieuwegein** — Fokkesteeg-Merwestein bestemmingsplan")
    zoek_knop = st.button("🔍 Ophalen met testadres")
    adres_input = "Graaf Walramhof 4, Nieuwegein"

# ── Hulpfuncties output ───────────────────────────────────────────────────────
def run_met_capture(fn, *args, **kwargs):
    log_placeholder = st.empty()
    class LiveCapture(io.StringIO):
        def write(self, tekst):
            super().write(tekst)
            log_placeholder.markdown(
                f'<div class="terminal">{self.getvalue()}</div>',
                unsafe_allow_html=True)
            return len(tekst)
    live = LiveCapture()
    old_stdout = sys.stdout
    sys.stdout = live
    builtins.input = _web_input
    try:
        resultaat = fn(*args, **kwargs)
    finally:
        sys.stdout = old_stdout
        builtins.input = _echte_input
    return resultaat

def toon_resultaten(data):
    if data.get("niet_gedigitaliseerd"):
        st.markdown(f"""<div class="niet-gedig">
            ⚠️ <strong>Dit plan is niet volledig gedigitaliseerd</strong> in de Ruimtelijke Plannen API.<br>
            Gele velden in het Word-document moeten handmatig worden ingevuld via:<br>
            <a href="{data.get('hyperlink','#')}" target="_blank">{data.get('hyperlink','—')}</a>
        </div>""", unsafe_allow_html=True)

    def kaart(label, waarde):
        heeft = waarde and waarde not in ("—", "geen", "")
        klasse = "auto" if heeft else "leeg"
        return f"""<div class="result-card">
            <div class="label">{label}</div>
            <div class="waarde {klasse}">{waarde if heeft else "—"}</div>
        </div>"""

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(kaart("Gevonden adres",       data.get("adres_gevonden","—")), unsafe_allow_html=True)
        st.markdown(kaart("Kadastrale aanduiding", data.get("kadastrale_aanduiding","—")), unsafe_allow_html=True)
        st.markdown(kaart("Bestemmingsplan",       data.get("bestemmingsplan_naam","—")), unsafe_allow_html=True)
        st.markdown(kaart("Datum vaststelling",    data.get("bestemmingsplan_datum","—")), unsafe_allow_html=True)
    with col2:
        st.markdown(kaart("Bestemming perceel",    data.get("bestemming_perceel","—")), unsafe_allow_html=True)
        st.markdown(kaart("Bestemmingstype",       data.get("bestemmingstype","—")), unsafe_allow_html=True)
        functie = ", ".join(data.get("functieaanduidingen",[])) or "geen"
        dubbel  = ", ".join(d["naam"] for d in data.get("dubbelbestemmingen",[])) or "geen"
        st.markdown(kaart("Functieaanduiding", functie), unsafe_allow_html=True)
        st.markdown(kaart("Dubbelbestemming",  dubbel),  unsafe_allow_html=True)

    if data.get("maatvoeringen"):
        st.markdown('<div class="sectie-header">Maatvoeringen</div>', unsafe_allow_html=True)
        cols = st.columns(3)
        for i, m in enumerate(data["maatvoeringen"]):
            with cols[i % 3]:
                st.markdown(kaart(m["naam"], f"{m['waarde']} {m.get('eenheid','')}".strip()), unsafe_allow_html=True)

def toon_download(data, label):
    st.markdown("---")
    st.markdown('<div class="stap-label">Stap 4 — Download Word-document</div>', unsafe_allow_html=True)
    with st.spinner("Word-document genereren..."):
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            tmp_pad = tmp.name
        genereer_intake_toets(data, uitvoer_pad=tmp_pad)
        with open(tmp_pad, "rb") as f:
            docx_bytes = f.read()
        os.unlink(tmp_pad)
    bestandsnaam = f"Intake_toets_{label.replace(' ','_')[:25]}_{date.today().strftime('%Y%m%d')}.docx"
    st.download_button(
        label="📄  Download Intake Toets (.docx)",
        data=docx_bytes, file_name=bestandsnaam,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    st.success(f"✅ Klaar! Klik op de knop om **{bestandsnaam}** te downloaden.")

# ── Hoofdlogica ───────────────────────────────────────────────────────────────
if zoek_knop:
    # Reset vorige run
    for k in ["toon_keuze","kandidaten","gekozen_adres"]:
        st.session_state[k] = False if k=="toon_keuze" else ([] if k=="kandidaten" else None)
    st.markdown("---")

    # Coördinaten
    if methode == "📐 RD-coördinaten":
        st.markdown('<div class="stap-label">Stap 2 — Data ophalen via DSO API</div>', unsafe_allow_html=True)
        st.info(f"📐 X={x_input}, Y={y_input} (RD)")
        try:
            with st.spinner("Bezig met ophalen..."):
                data = run_met_capture(haal_data_voor_coordinaten, float(x_input), float(y_input))
        except Exception as e:
            st.error(f"❌ Fout: {e}"); st.stop()
        st.markdown("---")
        st.markdown('<div class="stap-label">Stap 3 — Gevonden gegevens</div>', unsafe_allow_html=True)
        toon_resultaten(data)
        toon_download(data, f"RD_{int(x_input)}_{int(y_input)}")

    # Adres
    elif adres_input:
        with st.spinner("Adres opzoeken..."):
            try:
                kandidaten = zoek_adressen(adres_input.strip())
            except Exception as e:
                st.error(f"❌ Adres niet gevonden: {e}"); st.stop()

        if not kandidaten:
            st.error(f"❌ Geen adressen gevonden voor '{adres_input}'")
        elif len(kandidaten) == 1:
            # Directe match
            gekozen = kandidaten[0].get("weergavenaam", adres_input)
            st.markdown('<div class="stap-label">Stap 2 — Data ophalen via DSO API</div>', unsafe_allow_html=True)
            st.info(f"📍 {gekozen}")
            try:
                with st.spinner("Bezig met ophalen..."):
                    data = run_met_capture(haal_data_voor_adres, gekozen)
            except Exception as e:
                st.error(f"❌ Fout: {e}"); st.stop()
            st.markdown("---")
            st.markdown('<div class="stap-label">Stap 3 — Gevonden gegevens</div>', unsafe_allow_html=True)
            toon_resultaten(data)
            toon_download(data, gekozen.split(",")[0])
        else:
            # Sla kandidaten op voor het keuzemenu
            st.session_state["kandidaten"] = [d.get("weergavenaam","?") for d in kandidaten]
            st.session_state["toon_keuze"] = True
            st.session_state["gekozen_adres"] = None

# ── Keuzemenu (buiten zoek_knop zodat het na herrender blijft) ────────────────
if st.session_state.get("toon_keuze") and st.session_state.get("kandidaten"):
    st.markdown("---")
    st.markdown('<div class="stap-label">Meerdere adressen gevonden — kies er één</div>', unsafe_allow_html=True)
    namen = st.session_state["kandidaten"]
    keuze = st.radio("Kies het juiste adres:", namen, key="adres_keuze")

    if st.button("✓ Dit adres gebruiken"):
        # Sla het exacte gekozen adres op en verberg keuzemenu
        st.session_state["gekozen_adres"] = keuze
        st.session_state["toon_keuze"] = False
        st.rerun()

# ── Verwerk het gekozen adres na rerun ────────────────────────────────────────
if st.session_state.get("gekozen_adres"):
    gekozen = st.session_state["gekozen_adres"]
    st.session_state["gekozen_adres"] = None  # reset zodat het niet opnieuw draait
    st.markdown("---")
    st.markdown('<div class="stap-label">Stap 2 — Data ophalen via DSO API</div>', unsafe_allow_html=True)
    st.info(f"📍 {gekozen}")
    try:
        with st.spinner("Bezig met ophalen..."):
            # Geef het exacte gevonden adres mee — dan hoeft het script niet te zoeken
            # en wordt input() nooit aangeroepen
            data = run_met_capture(haal_data_voor_adres, gekozen)
    except Exception as e:
        st.error(f"❌ Fout: {e}"); st.stop()
    st.markdown("---")
    st.markdown('<div class="stap-label">Stap 3 — Gevonden gegevens</div>', unsafe_allow_html=True)
    toon_resultaten(data)
    toon_download(data, gekozen.split(",")[0])
