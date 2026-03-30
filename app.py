"""
DSO Intake Toets — Streamlit Web App v2
========================================
Gebruikt de originele scripts rechtstreeks:
  - dso_bestemmingsplan.py  → data ophalen
  - genereer_intake_toets.py → Word genereren
"""

import streamlit as st
import sys
import os
import io
from datetime import date

# ── Originele scripts inladen ─────────────────────────────────────────────────
# Zorg dat de map van app.py in het pad zit zodat imports werken
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import importlib.util as _ilu

def _laad(naam):
    pad = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"{naam}.py")
    spec = _ilu.spec_from_file_location(naam, pad)
    mod  = _ilu.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod

_dso    = _laad("dso_bestemmingsplan")
_toets  = _laad("genereer_intake_toets")

haal_data_voor_adres = _dso.haal_data_voor_adres
genereer_intake_toets = _toets.genereer_intake_toets

# ── Pagina-instellingen ───────────────────────────────────────────────────────
st.set_page_config(
    page_title="DSO Intake Toets",
    page_icon="🏛️",
    layout="centered",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}
.stApp {
    background: #0d1117;
    color: #e6edf3;
}
.dso-header {
    background: linear-gradient(135deg, #1f4e79 0%, #0d2137 100%);
    border: 1px solid #1f6feb;
    border-radius: 8px;
    padding: 2rem 2.5rem;
    margin-bottom: 2rem;
}
.dso-header h1 {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.6rem;
    font-weight: 600;
    color: #58a6ff;
    margin: 0 0 0.3rem 0;
}
.dso-header p {
    color: #8b949e;
    font-size: 0.9rem;
    margin: 0;
}
.stap-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.7rem;
    color: #58a6ff;
    text-transform: uppercase;
    letter-spacing: 2px;
    margin-bottom: 0.3rem;
}
.result-card {
    background: #161b22;
    border: 1px solid #30363d;
    border-radius: 6px;
    padding: 1rem 1.2rem;
    margin: 0.4rem 0;
}
.result-card .label {
    font-size: 0.72rem;
    color: #8b949e;
    font-family: 'IBM Plex Mono', monospace;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-bottom: 0.2rem;
}
.result-card .waarde       { font-size: 1rem; color: #e6edf3; font-weight: 600; }
.result-card .waarde.auto  { color: #58a6ff; }
.result-card .waarde.leeg  { color: #484f58; font-style: italic; font-weight: 400; }
.terminal {
    background: #0d1117;
    border: 1px solid #21262d;
    border-radius: 6px;
    padding: 1rem 1.2rem;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.78rem;
    color: #3fb950;
    white-space: pre-wrap;
    max-height: 280px;
    overflow-y: auto;
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
    border-color: #58a6ff !important;
    box-shadow: 0 0 0 3px rgba(88,166,255,0.1) !important;
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

# ── Invoer ────────────────────────────────────────────────────────────────────
st.markdown('<div class="stap-label">Stap 1 — Voer het adres in</div>', unsafe_allow_html=True)

col1, col2 = st.columns([3, 1])
with col1:
    adres_input = st.text_input(
        label="Adres",
        placeholder="bijv. Kerkstraat 1, IJsselstein",
        label_visibility="collapsed",
        key="adres_veld"
    )
with col2:
    if st.button("Testadres"):
        st.session_state["adres_veld"] = "Prinsengracht 40A, Amsterdam"
        st.rerun()

zoek_knop = st.button("🔍  Data ophalen + Word genereren")

# ── Verwerking ────────────────────────────────────────────────────────────────
if zoek_knop and adres_input:
    st.markdown("---")
    st.markdown('<div class="stap-label">Stap 2 — Data ophalen via DSO API</div>', unsafe_allow_html=True)

    # Vang terminal output op
    log_placeholder = st.empty()

    class StreamCapture(io.StringIO):
        def write(self, tekst):
            super().write(tekst)
            log_placeholder.markdown(
                f'<div class="terminal">{self.getvalue()}</div>',
                unsafe_allow_html=True
            )
            return len(tekst)

    capture = StreamCapture()
    old_stdout = sys.stdout
    sys.stdout = capture

    try:
        with st.spinner("Bezig met ophalen..."):
            data = haal_data_voor_adres(adres_input.strip())
    except Exception as e:
        sys.stdout = old_stdout
        st.error(f"❌ Fout bij ophalen data: {e}")
        st.stop()
    finally:
        sys.stdout = old_stdout

    # ── Resultaten tonen ──────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<div class="stap-label">Stap 3 — Gevonden gegevens</div>', unsafe_allow_html=True)

    if data.get("niet_gedigitaliseerd"):
        st.markdown(f"""<div class="niet-gedig">
            ⚠️ <strong>Dit plan is niet volledig gedigitaliseerd</strong> in de Ruimtelijke Plannen API.<br>
            Gele velden in het Word-document moeten handmatig worden ingevuld via:<br>
            <a href="{data.get('hyperlink','#')}" target="_blank">{data.get('hyperlink','—')}</a>
        </div>""", unsafe_allow_html=True)

    def kaart(label, waarde):
        heeft_waarde = waarde and waarde not in ("—", "geen", "")
        klasse = "auto" if heeft_waarde else "leeg"
        w = waarde if heeft_waarde else "—"
        return f"""<div class="result-card">
            <div class="label">{label}</div>
            <div class="waarde {klasse}">{w}</div>
        </div>"""

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(kaart("Gevonden adres",       data.get("adres_gevonden", "—")), unsafe_allow_html=True)
        st.markdown(kaart("Kadastrale aanduiding", data.get("kadastrale_aanduiding", "—")), unsafe_allow_html=True)
        st.markdown(kaart("Bestemmingsplan",       data.get("bestemmingsplan_naam", "—")), unsafe_allow_html=True)
        st.markdown(kaart("Datum vaststelling",    data.get("bestemmingsplan_datum", "—")), unsafe_allow_html=True)
    with col2:
        st.markdown(kaart("Bestemming perceel",    data.get("bestemming_perceel", "—")), unsafe_allow_html=True)
        st.markdown(kaart("Bestemmingstype",       data.get("bestemmingstype", "—")), unsafe_allow_html=True)
        functie = ", ".join(data.get("functieaanduidingen", [])) or "geen"
        dubbel  = ", ".join(d["naam"] for d in data.get("dubbelbestemmingen", [])) or "geen"
        st.markdown(kaart("Functieaanduiding",     functie), unsafe_allow_html=True)
        st.markdown(kaart("Dubbelbestemming",      dubbel), unsafe_allow_html=True)

    if data.get("maatvoeringen"):
        st.markdown('<div class="sectie-header">Maatvoeringen</div>', unsafe_allow_html=True)
        cols = st.columns(3)
        for i, m in enumerate(data["maatvoeringen"]):
            with cols[i % 3]:
                waarde_str = f"{m['waarde']} {m.get('eenheid', '')}".strip()
                st.markdown(kaart(m["naam"], waarde_str), unsafe_allow_html=True)

    # ── Word genereren ────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<div class="stap-label">Stap 4 — Download Word-document</div>', unsafe_allow_html=True)

    with st.spinner("Word-document genereren..."):
        # Genereer naar tijdelijk bestand, lees terug als bytes
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            tmp_pad = tmp.name

        genereer_intake_toets(data, uitvoer_pad=tmp_pad)

        with open(tmp_pad, "rb") as f:
            docx_bytes = f.read()
        os.unlink(tmp_pad)

    adres_kort   = adres_input.split(",")[0].replace(" ", "_").replace("/", "-")[:25]
    bestandsnaam = f"Intake_toets_{adres_kort}_{date.today().strftime('%Y%m%d')}.docx"

    st.download_button(
        label="📄  Download Intake Toets (.docx)",
        data=docx_bytes,
        file_name=bestandsnaam,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    st.success(f"✅ Klaar! Klik op de knop om **{bestandsnaam}** te downloaden.")

elif zoek_knop and not adres_input:
    st.warning("⚠️ Vul eerst een adres in.")
