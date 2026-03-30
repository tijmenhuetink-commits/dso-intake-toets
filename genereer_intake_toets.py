"""
DSO Intake Toets Generator
===========================
Versie : 2.0
Datum  : 2026-03-20
Wijzigingen:
  v1.0 — eerste versie, Word-document gegenereerd vanuit DSO JSON-data
  v1.1 — importpad gecorrigeerd via sys.argv[0] (Windows-compatibel)
          vraag_invoer() lokaal gedefinieerd (stond alleen in __main__ blok DSO-script)
  v1.2 — voorbereidingsbesluit en parkeerplan automatisch ingevuld vanuit planenoverzicht
          bebouwde oppervlakte zoekt op meerdere termen
  v1.3 — lege DSO-velden tonen "geen" op lichtblauwe achtergrond
          zodat duidelijk is dat het veld wél is opgezocht maar niets opleverde
  v1.4 — PB-veldcodes verplaatst naar apart tabblad (tweede sectie in document)
          hoofdformulier is nu clean zonder [PB: ...] tekst
  v1.5 — Omgevingsplan veld toont "zie Regels op de kaart" i.p.v. "geen"
          Ozon API biedt geen publieke geo-zoekingang op XY-punt (Geoforum bevestigd)
  v1.6 — "geen" onderscheiden in twee betekenissen:
          maatvoeringen/oppervlakte: "niet opgenomen in plan" (API geeft niets terug)
          dubbelbestemming/functie/aanduiding: "geen" (correct: er is er geen)
  v1.7 — activiteiten-sectie: checkboxes ongewijzigd, PB-placeholder toegevoegd
  v1.8 — lichtrode achtergrond voor velden die straks via PB-koppeling worden gevuld
  v1.9 — aanvraagtype, omschrijving bouwplan, bijzonderheden als rood toegevoegd
          (nog te verifiëren of ZAAK.ZAAKTYPE / ZAAK.OMSCHRIJVING_KORT / ZAAK.TOELICHTING
           daadwerkelijk gevuld zijn in PB bij een zaak)
  v2.0 — niet-gedigitaliseerde plannen netjes afgehandeld:
          oranje waarschuwingsbalk met hyperlink naar ruimtelijkeplannen.nl
          gele cellen voor velden die handmatig ingevuld moeten worden
          auto() herkent NIET_BESCHIKBAAR waarden en kleurt ze geel i.p.v. blauw

Genereert een ingevuld Word-document (Intake_toets) op basis van de
JSON-output van dso_bestemmingsplan.py.

Gebruik:
  python genereer_intake_toets.py                        # vraagt adres interactief
  python genereer_intake_toets.py "Kerkstraat 1, Utrecht"
  python genereer_intake_toets.py 131653,447223          # RD-coördinaten

Output:
  Intake_toets_<adres>_<datum>.docx  op de Desktop

Benodigdheden:
  pip install requests python-docx
"""

VERSION = "2.0"

import sys
import os
import json
from datetime import date
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ── Importeer het DSO-script ──────────────────────────────────────────────────
# Zoek dso_bestemmingsplan.py op via het pad van dit script (absoluut, cross-platform)
import importlib.util as _ilu

def _laad_dso():
    """Laad dso_bestemmingsplan.py op basis van de locatie van dit script."""
    # Eigen map van dit script — werkt altijd, ook op Windows
    eigen_map = os.path.dirname(os.path.abspath(sys.argv[0]))
    kandidaten = [
        os.path.join(eigen_map, "dso_bestemmingsplan.py"),
        os.path.join(os.getcwd(), "dso_bestemmingsplan.py"),
    ]
    for pad in kandidaten:
        if os.path.isfile(pad):
            spec = _ilu.spec_from_file_location("dso_bestemmingsplan", pad)
            mod  = _ilu.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    return None

_dso = _laad_dso()
if _dso:
    haal_data_voor_adres        = _dso.haal_data_voor_adres
    haal_data_voor_coordinaten  = _dso.haal_data_voor_coordinaten
    DSO_BESCHIKBAAR = True
else:
    DSO_BESCHIKBAAR = False
    print("⚠  dso_bestemmingsplan.py niet gevonden — gebruik test-JSON")


def vraag_invoer():
    """Keuzemenu voor invoermethode (adres, coördinaten of testadres)."""
    print("=" * 55)
    print("  Kies invoermethode:")
    print("  1. Adres (bijv. Kerkstraat 1, IJsselstein)")
    print("  2. RD-coördinaten (bijv. 131653, 447223)")
    print("  3. Testadres (Prinsengracht 40A, Amsterdam)")
    print("=" * 55)
    keuze = input("  Keuze [1/2/3]: ").strip()
    if keuze == "2":
        coords = input("  X, Y (RD): ").strip()
        try:
            x_str, y_str = coords.replace(" ", "").split(",")
            return None, float(x_str), float(y_str)
        except Exception:
            print("  ⚠ Ongeldige invoer — gebruik testadres")
            return "Prinsengracht 40A, Amsterdam", None, None
    elif keuze == "3":
        return "Prinsengracht 40A, Amsterdam", None, None
    else:
        adres = input("  Adres: ").strip()
        if not adres:
            return "Prinsengracht 40A, Amsterdam", None, None
        return adres, None, None


# ── Kleuren ───────────────────────────────────────────────────────────────────
BLAUW        = RGBColor(0x1F, 0x4E, 0x79)
WIT          = RGBColor(0xFF, 0xFF, 0xFF)
LICHTBLAUW_HEX = "D6E4F0"   # automatisch gevuld (DSO-script)
LICHTROOD_HEX  = "FCE4D6"   # straks automatisch via PB-koppeling
GEEL_HEX       = "FFF2CC"   # handmatig invullen
GRIJS_HEX      = "F2F2F2"   # labelkolom
BLAUW_HEX      = "1F4E79"   # sectieheaders
DONKERBLAUW_HEX = "17375E"


# ── Hulpfuncties opmaak ───────────────────────────────────────────────────────

def set_cell_bg(cell, hex_kleur):
    """Zet achtergrondkleur van een tabelcel."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  hex_kleur)
    tcPr.append(shd)


def set_cell_borders(cell):
    """Lichte grijze rand rondom cel."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tblBorders = OxmlElement('w:tcBorders')
    for zijde in ('top', 'left', 'bottom', 'right'):
        border = OxmlElement(f'w:{zijde}')
        border.set(qn('w:val'),   'single')
        border.set(qn('w:sz'),    '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'AAAAAA')
        tblBorders.append(border)
    tcPr.append(tblBorders)


def cel_tekst(cell, tekst, vet=False, cursief=False, kleur=None, pt=9):
    """Schrijf tekst in een cel met opmaak."""
    cell.text = ""
    p   = cell.paragraphs[0]
    run = p.add_run(str(tekst))
    run.bold    = vet
    run.italic  = cursief
    run.font.size = Pt(pt)
    run.font.name = "Arial"
    if kleur:
        run.font.color.rgb = kleur
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after  = Pt(1)


def voeg_sectieheader_toe(doc, tekst):
    """Blauwe sectieheader als alinea."""
    p   = doc.add_paragraph()
    run = p.add_run(tekst)
    run.bold           = True
    run.font.size      = Pt(11)
    run.font.name      = "Arial"
    run.font.color.rgb = WIT
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(2)

    # Blauwe achtergrond via XML shading op paragraaf
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  BLAUW_HEX)
    pPr.append(shd)

    # Kleine inspringing
    ind = OxmlElement('w:ind')
    ind.set(qn('w:left'), '120')
    pPr.append(ind)
    return p


def maak_tweekoloms_tabel(doc, breedte_label=5.5, breedte_waarde=10.5):
    """Maak een tweekolomstabel zonder standaard borders."""
    tabel = doc.add_table(rows=0, cols=2)
    tabel.style = 'Table Grid'
    tabel.autofit = False
    tabel.columns[0].width = Cm(breedte_label)
    tabel.columns[1].width = Cm(breedte_waarde)
    return tabel


def voeg_rij_toe(tabel, label, waarde, kleur="wit", pb_veld=None):
    """
    Voeg een rij toe aan een tweekolomstabel.
    kleur: "blauw" = automatisch (lichtblauw), "geel" = handmatig, "wit" = neutraal
    pb_veld: optioneel PB-veldcode als tooltip-tekst (voor later)
    """
    rij  = tabel.add_row()
    cel_label  = rij.cells[0]
    cel_waarde = rij.cells[1]

    set_cell_borders(cel_label)
    set_cell_borders(cel_waarde)
    set_cell_bg(cel_label, GRIJS_HEX)

    if kleur == "blauw":
        set_cell_bg(cel_waarde, LICHTBLAUW_HEX)
        cursief = True
    elif kleur == "rood":
        set_cell_bg(cel_waarde, LICHTROOD_HEX)
        cursief = True
    elif kleur == "geel":
        set_cell_bg(cel_waarde, GEEL_HEX)
        cursief = False
    else:
        set_cell_bg(cel_waarde, "FFFFFF")
        cursief = False

    # Label
    cel_tekst(cel_label, label, vet=True, pt=9)

    # Waarde — toon alleen de echte waarde, geen PB-veldcode in de cel
    if waarde and str(waarde).strip() and str(waarde) != "—":
        cel_tekst(cel_waarde, str(waarde), cursief=cursief, pt=9)
    else:
        cel_tekst(cel_waarde, "", pt=9)

    # Celmarges
    for cel in [cel_label, cel_waarde]:
        tc   = cel._tc
        tcPr = tc.get_or_add_tcPr()
        mar  = OxmlElement('w:tcMar')
        for zijde, val in [('top','60'),('bottom','60'),('left','120'),('right','120')]:
            m = OxmlElement(f'w:{zijde}')
            m.set(qn('w:w'),    val)
            m.set(qn('w:type'), 'dxa')
            mar.append(m)
        tcPr.append(mar)


def voeg_witruimte_toe(doc, pt=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(pt)


def voeg_checkbox_toe(doc, tekst, aangevinkt=False):
    """Voeg een checkbox-regel toe."""
    p   = doc.add_paragraph()
    run = p.add_run(("☑ " if aangevinkt else "☐ ") + tekst)
    run.font.size = Pt(9)
    run.font.name = "Arial"
    p.paragraph_format.left_indent   = Cm(0.5)
    p.paragraph_format.space_before  = Pt(1)
    p.paragraph_format.space_after   = Pt(1)


# ── Activiteiten-mapping ──────────────────────────────────────────────────────
# Koppel DSO/omgevingsplan activiteiten aan de checkboxen in het formulier
# Wordt later uitgebreid zodra we weten hoe PB activiteiten opslaat

ALLE_ACTIVITEITEN = [
    "Bouwactiviteit (technisch)",
    "Bouwactiviteit (omgevingsplan)",
    "Omgevingsplanactiviteit (uitvoeren van een werk, niet zijnde een bouwwerk, of een werkzaamheid)",
    "Omgevingsplanactiviteit (slopen)",
    "Omgevingsplanactiviteit (afwijken van de regels in het omgevingsplan)",
    "Omgevingsplanactiviteit (maken of veranderen van een in- en uitrit)",
    "Omgevingsplanactiviteit (het vellen van houtopstand)",
    "Omgevingsplanactiviteit (het uitvoeren van een milieubelastende activiteit)",
    "Omgevingsplanactiviteit (aanleg/wijzigen van weg of spoorweg nabij geluidgevoelig gebouw)",
    "Omgevingsplanactiviteit (slopen in een gemeentelijk beschermd stads- of dorpsgezicht)",
    "Omgevingsplanactiviteit (slopen van gebouw in provinciaal beschermde structuur)",
    "Omgevingsplanactiviteit (het slopen van een bouwwerk in rijksbeschermd stads- of dorpsgezicht)",
    "Omgevingsplanactiviteit (het veranderen van een gemeentelijk beschermd monument)",
    "Rijksmonumentenactiviteit (Bouwwerk)",
    "Rijksmonumentenactiviteit (Archeologie)",
    "Milieubelastende activiteit (het uitvoeren van een milieubelastende activiteit)",
]


# ── Hoofdfunctie: genereer Word-document ──────────────────────────────────────

def genereer_intake_toets(data: dict, uitvoer_pad: str = None) -> str:
    """
    Genereert een ingevuld Intake Toets Word-document op basis van DSO-data.

    Parameters:
        data        — dict van haal_data_voor_adres() of haal_data_voor_coordinaten()
        uitvoer_pad — optioneel volledige bestandspad; anders Desktop

    Geeft het pad terug naar het gegenereerde bestand.
    """

    # ── Basisgegevens uit JSON ──
    adres          = data.get("adres_gevonden") or data.get("adres", "—")
    kadaster       = data.get("kadastrale_aanduiding", "—")
    bp_naam        = data.get("bestemmingsplan_naam", "—")
    bp_datum       = data.get("bestemmingsplan_datum", "—")
    hyperlink      = data.get("hyperlink", "—")
    bestemming     = data.get("bestemming_perceel", "—")
    btype          = data.get("bestemmingstype", "—")
    functie_lijst  = data.get("functieaanduidingen", [])
    dubbel_lijst   = data.get("dubbelbestemmingen", [])
    bouw_lijst     = data.get("bouwaanduidingen", [])
    maatvoeringen  = data.get("maatvoeringen", [])
    planenoverzicht = data.get("planenoverzicht", {})
    niet_gedigitaliseerd = data.get("niet_gedigitaliseerd", False)
    NIET_BESCHIKBAAR = "↗ niet beschikbaar via API — zie hyperlink plan"

    functie_str = ", ".join(functie_lijst) if functie_lijst else "geen"
    dubbel_str  = ", ".join(d["naam"] for d in dubbel_lijst) if dubbel_lijst else "geen"
    bouw_str    = ", ".join(b["naam"] for b in bouw_lijst) if bouw_lijst else "geen"

    # Bij niet-gedigitaliseerde plannen: vlak-afhankelijke velden op handmatig zetten
    if niet_gedigitaliseerd:
        bestemming  = NIET_BESCHIKBAAR
        functie_str = NIET_BESCHIKBAAR
        dubbel_str  = NIET_BESCHIKBAAR
        bouw_str    = NIET_BESCHIKBAAR

    # Maatvoeringen per naam opzoeken
    def maatvoering_waarde(zoek_naam):
        zoek = zoek_naam.lower()
        for m in maatvoeringen:
            naam = m.get("naam", "").lower()
            if zoek in naam:
                w = m.get("waarde", "—")
                e = m.get("eenheid", "")
                return f"{w} {e}".strip()
        return "—"

    bouwhoogte  = maatvoering_waarde("bouwhoogte")
    goothoogte  = maatvoering_waarde("goothoogte")
    opp_bouwperceel = maatvoering_waarde("oppervlakte")

    if niet_gedigitaliseerd:
        bouwhoogte      = NIET_BESCHIKBAAR
        goothoogte      = NIET_BESCHIKBAAR
        opp_bouwperceel = NIET_BESCHIKBAAR

    # Adresonderdelen splitsen (eenvoudig)
    adres_delen = adres.split(", ") if ", " in adres else [adres]
    straat_hnr  = adres_delen[0] if len(adres_delen) >= 1 else adres
    woonplaats  = adres_delen[-1].split(" ", 1)[-1] if len(adres_delen) >= 2 else "—"

    # Omgevingsplan uit planenoverzicht
    omgevingsplan_naam = "—"
    for p in planenoverzicht.get("omgevingsplan", []):
        omgevingsplan_naam = p.get("naam", "—")
        break

    # Voorbereidingsbesluit
    voorbereidingsbesluit_str = "—"
    vbb_lijst = planenoverzicht.get("voorbereidingsbesluit", [])
    if vbb_lijst:
        voorbereidingsbesluit_str = ", ".join(p.get("naam","—") for p in vbb_lijst)

    # Bestemmingsplan parkeren — zoek parapluplan met 'parkeer' in naam
    parkeerplan_str = "—"
    for p in planenoverzicht.get("bestemmingsplan", []):
        if "parkeer" in p.get("naam","").lower() and p.get("paraplu"):
            parkeerplan_str = p.get("naam","—")
            break

    # Bebouwde oppervlakte — probeer meerdere zoektermen
    def maatvoering_waarde_uitgebreid(zoektermen):
        for zoek in zoektermen:
            v = maatvoering_waarde(zoek)
            if v != "—":
                return v
        return "—"

    opp_bouwperceel = maatvoering_waarde_uitgebreid([
        "oppervlakte", "bebouwd", "bouwperceel", "perceeloppervlakte"
    ])

    # ── Document aanmaken ──
    doc = Document()

    # Paginamarges instellen (2 cm rondom)
    sectie = doc.sections[0]
    sectie.top_margin    = Cm(2)
    sectie.bottom_margin = Cm(2)
    sectie.left_margin   = Cm(2)
    sectie.right_margin  = Cm(2)

    # ── Stijlen instellen ──
    stijl = doc.styles['Normal']
    stijl.font.name = "Arial"
    stijl.font.size = Pt(9)

    # ── TITEL ──
    p_titel = doc.add_paragraph()
    r = p_titel.add_run("INTAKE TOETS")
    r.bold           = True
    r.font.size      = Pt(18)
    r.font.color.rgb = BLAUW
    r.font.name      = "Arial"
    p_titel.paragraph_format.space_after = Pt(2)

    # Legenda
    p_leg = doc.add_paragraph()
    for tekst, kleur_hex in [
        ("Legenda:  ", None),
        ("  Automatisch (DSO-script)  ", LICHTBLAUW_HEX),
        ("   ", None),
        ("  Straks automatisch (PB-koppeling)  ", LICHTROOD_HEX),
        ("   ", None),
        ("  Handmatig invullen  ", GEEL_HEX),
    ]:
        r = p_leg.add_run(tekst)
        r.font.size = Pt(7.5)
        r.font.name = "Arial"
        if kleur_hex:
            rPr = r._r.get_or_add_rPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'),   'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'),  kleur_hex)
            rPr.append(shd)
    p_leg.paragraph_format.space_after = Pt(4)

    voeg_witruimte_toe(doc, 2)

    # ══════════════════════════════════════════════════════
    # SECTIE 1 — ALGEMENE ZAAKGEGEVENS
    # ══════════════════════════════════════════════════════
    voeg_sectieheader_toe(doc, "1. ALGEMENE ZAAKGEGEVENS")
    t = maak_tweekoloms_tabel(doc)
    voeg_rij_toe(t, "Zaak-ID",             None,  "wit",  "ZAAK.ZAAK_IDENTIFICATIE")
    voeg_rij_toe(t, "Kenmerk",             None,  "wit",  "ZAAK.KENMERK")
    voeg_rij_toe(t, "Ontvangstdatum",      None,  "rood", "PROCES.ONTVANGST")
    voeg_rij_toe(t, "DSO Verzoeknummer",   None,  "rood", "ZAAK.DSO_VERZOEKNUMMER")
    voeg_rij_toe(t, "Casemanager",         None,  "rood", "BEHANDELAAR.SAMENGESTELDE_NAAM")
    voeg_rij_toe(t, "Afdeling",            None,  "rood", "BEHANDELAAR.AFDELING")
    voeg_rij_toe(t, "Datum adviesaanvraag RO", None, "geel")
    voeg_rij_toe(t, "Aanvraagtype",        None,  "rood", "ZAAK.ZAAKTYPE")
    voeg_rij_toe(t, "Gerelateerde zaken",  None,  "geel")
    voeg_rij_toe(t, "Omschrijving bouwplan", None, "rood", "ZAAK.OMSCHRIJVING_KORT")
    voeg_rij_toe(t, "Bijzonderheden",      None,  "rood", "ZAAK.TOELICHTING")
    voeg_witruimte_toe(doc)

    # ══════════════════════════════════════════════════════
    # SECTIE 2 — LOCATIE
    # ══════════════════════════════════════════════════════
    voeg_sectieheader_toe(doc, "2. LOCATIE")
    t = maak_tweekoloms_tabel(doc)
    voeg_rij_toe(t, "Straatnaam + huisnummer", straat_hnr,  "blauw", "ZAAK_ADRES.SAMENGESTELD_STRAAT")
    voeg_rij_toe(t, "Woonplaats",          woonplaats,  "blauw", "ZAAK_ADRES.WOONPLAATS")
    voeg_rij_toe(t, "Kadastrale aanduiding", kadaster,   "blauw", "ZAAK_OBJECTADRES.PERCEELNUMMER")
    voeg_witruimte_toe(doc)

    # ══════════════════════════════════════════════════════
    # SECTIE 3 — ACTIVITEITEN
    # ══════════════════════════════════════════════════════
    voeg_sectieheader_toe(doc, "3. ACTIVITEITEN (vink aan wat van toepassing is)")

    for act in ALLE_ACTIVITEITEN:
        voeg_checkbox_toe(doc, act, aangevinkt=False)
    voeg_witruimte_toe(doc, 2)

    # PB-placeholder voor activiteiten
    p_pb = doc.add_paragraph()
    pPr = p_pb._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), LICHTROOD_HEX)
    pPr.append(shd)
    p_pb.paragraph_format.space_before = Pt(2)
    p_pb.paragraph_format.space_after  = Pt(2)
    r1 = p_pb.add_run("Activiteiten volgens PowerBrowser:  ")
    r1.bold = True; r1.font.size = Pt(9); r1.font.name = "Arial"
    r2 = p_pb.add_run("«ZAAK.ZAAK_ACTIVITEIT_OMS»")
    r2.italic = True; r2.font.size = Pt(9); r2.font.name = "Arial"
    r2.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    r3 = p_pb.add_run("  —  wordt automatisch gevuld zodra PB-koppeling actief is")
    r3.font.size = Pt(7.5); r3.font.name = "Arial"
    r3.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA); r3.italic = True
    voeg_witruimte_toe(doc)

    # ══════════════════════════════════════════════════════
    # SECTIE 4 — VERGUNNINGSVRIJ
    # ══════════════════════════════════════════════════════
    voeg_sectieheader_toe(doc, "4. VERGUNNINGSVRIJ")
    t = doc.add_table(rows=1, cols=3)
    t.style = 'Table Grid'
    t.autofit = False
    t.columns[0].width = Cm(3)
    t.columns[1].width = Cm(3)
    t.columns[2].width = Cm(9.9)
    for cel, hdr in zip(t.rows[0].cells, ["BBL / Omgevingsplan", "Artikel", "Beschrijving"]):
        set_cell_bg(cel, BLAUW_HEX)
        set_cell_borders(cel)
        cel_tekst(cel, hdr, vet=True, kleur=WIT, pt=9)
    for _ in range(4):
        rij = t.add_row()
        for cel in rij.cells:
            set_cell_bg(cel, GEEL_HEX)
            set_cell_borders(cel)
            cel_tekst(cel, "", pt=9)
    voeg_witruimte_toe(doc)

    # ══════════════════════════════════════════════════════
    # SECTIE 5 — TOETS BBL
    # ══════════════════════════════════════════════════════
    voeg_sectieheader_toe(doc, "5. TOETS BBL")
    t = doc.add_table(rows=1, cols=3)
    t.style = 'Table Grid'
    t.autofit = False
    t.columns[0].width = Cm(3)
    t.columns[1].width = Cm(3)
    t.columns[2].width = Cm(9.9)
    for cel, hdr in zip(t.rows[0].cells, ["Afdeling", "Artikel", "Strijdigheid"]):
        set_cell_bg(cel, BLAUW_HEX)
        set_cell_borders(cel)
        cel_tekst(cel, hdr, vet=True, kleur=WIT, pt=9)
    for _ in range(4):
        rij = t.add_row()
        for cel in rij.cells:
            set_cell_bg(cel, GEEL_HEX)
            set_cell_borders(cel)
            cel_tekst(cel, "", pt=9)
    voeg_witruimte_toe(doc)

    # ══════════════════════════════════════════════════════
    # SECTIE 6 — OMGEVINGSPLANTOETS
    # ══════════════════════════════════════════════════════
    voeg_sectieheader_toe(doc, "6. OMGEVINGSPLANTOETS")
    voeg_witruimte_toe(doc, 2)

    p = doc.add_paragraph()
    r = p.add_run("Verbeelding plankaart")
    r.bold = True; r.font.size = Pt(10); r.font.name = "Arial"
    r.font.color.rgb = BLAUW
    p.paragraph_format.space_after = Pt(1)
    p_img = doc.add_paragraph()
    r = p_img.add_run("[Voeg hier schermafbeelding van de plankaart in]")
    r.font.size = Pt(8); r.font.name = "Arial"
    r.font.color.rgb = RGBColor(0x88, 0x88, 0x88); r.italic = True
    p_img.paragraph_format.space_after = Pt(6)

    # Omgevingsplan info tabel
    p = doc.add_paragraph()
    r = p.add_run("Omgevingsplan informatie")
    r.bold = True; r.font.size = Pt(10); r.font.name = "Arial"
    r.font.color.rgb = BLAUW
    p.paragraph_format.space_after = Pt(1)

    t = maak_tweekoloms_tabel(doc)
    voeg_rij_toe(t, "Hyperlink Regels op de kaart",
                 "https://omgevingswet.overheid.nl/regels-op-de-kaart/zoeken/locatie",
                 "blauw")
    def auto(waarde, leeg_tekst="geen"):
        """Geeft waarde terug met kleur:
        - geel als waarde niet beschikbaar is via API (niet gedigitaliseerd)
        - blauw als waarde gevonden is via API of leeg is
        """
        if waarde and "niet beschikbaar via API" in str(waarde):
            return waarde, "geel"
        v = waarde if waarde and waarde not in ("—", "geen", "") else leeg_tekst
        return v, "blauw"

    # Waarschuwingsbalk bij niet-gedigitaliseerde plannen
    if niet_gedigitaliseerd:
        p_warn = doc.add_paragraph()
        pPr = p_warn._p.get_or_add_pPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), 'FFD966')
        pPr.append(shd)
        r_warn = p_warn.add_run(
            f"⚠  Dit plan ({bp_naam}) is niet volledig gedigitaliseerd in de Ruimtelijke Plannen API. "
            f"De gele velden hieronder zijn niet automatisch ingevuld en moeten handmatig worden "
            f"opgezocht via: {hyperlink}"
        )
        r_warn.bold = True
        r_warn.font.size = Pt(8)
        r_warn.font.name = "Arial"
        r_warn.font.color.rgb = RGBColor(0x7F, 0x3F, 0x00)
        p_warn.paragraph_format.space_before = Pt(4)
        p_warn.paragraph_format.space_after  = Pt(4)

    # Omgevingsplan: niet opvraagbaar via publieke API op XY-punt
    # De Ozon/Geodoc koppeling is nog niet publiek beschikbaar (zie Geoforum 2024)
    omgevingsplan_weergave = omgevingsplan_naam if omgevingsplan_naam != "—"         else "↗ zie Regels op de kaart (link hierboven)"
    voeg_rij_toe(t, "Omgevingsplan", omgevingsplan_weergave, "blauw",
                 "BESTEMMINGSPLAN_UITVOER.OMSCHRIJVING")
    voeg_rij_toe(t, "Bestemmingsplan",     *auto(bp_naam),
                 "BESTEMMINGSPLAN_UITVOER.OMSCHRIJVING")
    voeg_rij_toe(t, "Datum vaststelling",  *auto(bp_datum),
                 "BESTEMMINGSPLAN_UITVOER.DATUMVASTSTELLING")
    voeg_rij_toe(t, "Hyperlink plan",      *auto(hyperlink))
    voeg_rij_toe(t, "Bestemming perceel",  *auto(bestemming),
                 "BESTEMMINGSPLAN_UITVOER.BESTEMMING_ID")
    voeg_rij_toe(t, "Dubbelbestemming",    *auto(dubbel_str))
    voeg_rij_toe(t, "(Functie)aanduiding", *auto(functie_str))
    voeg_rij_toe(t, "Bouwaanduiding",      *auto(bouw_str))
    voeg_rij_toe(t, "Voorbereidingsbesluit / ontwerpbestemmingsplan",
                 *auto(voorbereidingsbesluit_str, "geen"))
    voeg_rij_toe(t, "Bestemmingsplan parkeren", *auto(parkeerplan_str, "geen"))
    voeg_rij_toe(t, "Bebouwde oppervlakte bouwperceel", *auto(opp_bouwperceel, "niet opgenomen in plan"))
    voeg_rij_toe(t, "Maximale bouwhoogte", *auto(bouwhoogte, "niet opgenomen in plan"))
    voeg_rij_toe(t, "Maximale goothoogte", *auto(goothoogte, "niet opgenomen in plan"))

    # Alle overige maatvoeringen
    overige = [m for m in maatvoeringen
               if not any(k in m.get("naam","").lower()
                          for k in ["bouwhoogte","goothoogte","oppervlakte"])]
    for m in overige:
        w = f"{m['waarde']} {m.get('eenheid','')}".strip()
        voeg_rij_toe(t, m["naam"], w or "geen", "blauw")

    voeg_witruimte_toe(doc)

    # Regels-tabellen
    for titel, rijen in [
        ("Omgevingsregels", 3),
        ("Bestemmingsplanregels (inclusief binnenplans afwijken)", 3),
        ("Afwijkings- en/of beleidsregels", 3),
    ]:
        p = doc.add_paragraph()
        r = p.add_run(titel)
        r.bold = True; r.font.size = Pt(10); r.font.name = "Arial"
        r.font.color.rgb = BLAUW
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after  = Pt(1)

        t = doc.add_table(rows=1, cols=2)
        t.style = 'Table Grid'
        t.autofit = False
        t.columns[0].width = Cm(4)
        t.columns[1].width = Cm(11.9)
        for cel, hdr in zip(t.rows[0].cells, ["Artikel", "Strijdigheid"]):
            set_cell_bg(cel, BLAUW_HEX)
            set_cell_borders(cel)
            cel_tekst(cel, hdr, vet=True, kleur=WIT, pt=9)
        for _ in range(rijen):
            rij = t.add_row()
            for cel in rij.cells:
                set_cell_bg(cel, GEEL_HEX)
                set_cell_borders(cel)
                cel_tekst(cel, "", pt=9)
        voeg_witruimte_toe(doc, 2)

    # ══════════════════════════════════════════════════════
    # SECTIE 7 — CONCLUSIE
    # ══════════════════════════════════════════════════════
    voeg_sectieheader_toe(doc, "7. CONCLUSIE OMGEVINGSPLANTOETS")
    t = maak_tweekoloms_tabel(doc)
    voeg_rij_toe(t, "Voldoet aan bestemmingsplan?",          None, "geel")
    voeg_rij_toe(t, "Voldoet aan afwijkingsregels?",         None, "geel")
    voeg_rij_toe(t, "BOPA ja of nee?",                       None, "geel")
    voeg_rij_toe(t, "Procedure",                             None, "geel")
    voeg_rij_toe(t, "Anterieure overeenkomst / nadeelcompensatie?", None, "geel")
    voeg_rij_toe(t, "Instemming en advies benodigd?",        None, "geel")
    voeg_witruimte_toe(doc)

    # ══════════════════════════════════════════════════════
    # SECTIE 8 — TOETS OMGEVINGSPLAN (tekstblok)
    # ══════════════════════════════════════════════════════
    voeg_sectieheader_toe(doc, "8. TOETS OMGEVINGSPLAN")
    voeg_witruimte_toe(doc, 2)

    # Vooringevuld tekstblok
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  GEEL_HEX)
    pPr.append(shd)
    p.paragraph_format.space_after = Pt(4)

    def run(tekst, vet=False, cursief=False, kleur=None):
        r = p.add_run(tekst)
        r.font.size = Pt(9); r.font.name = "Arial"
        r.bold = vet; r.italic = cursief
        if kleur: r.font.color.rgb = kleur

    run("Het perceel ")
    run(straat_hnr + (f", {woonplaats}" if woonplaats != "—" else ""),
        vet=True, kleur=BLAUW)
    run(" ligt binnen het plangebied van het vigerende omgevingsplan ")
    plan_weergave = omgevingsplan_naam if omgevingsplan_naam != "—" else bp_naam
    run(f'"{plan_weergave}"', cursief=True)
    run(f", waarvan het bestemmingsplan ")
    run(f'"{bp_naam}"', cursief=True)
    run(" een onderdeel is. Het perceel heeft op basis daarvan de enkelbestemming ")
    run(f"'{bestemming}'", vet=True)
    if dubbel_str != "geen":
        run(", met de dubbelbestemming ")
        run(f"'{dubbel_str}'", vet=True)
    if bouw_str != "geen":
        run(f" en (bouw)aanduiding ")
        run(f"'{bouw_str}'", vet=True)
    run(". De aanvraag omgevingsvergunning is gesitueerd op deze bestemming.")

    voeg_witruimte_toe(doc, 2)
    p2 = doc.add_paragraph()
    r2 = p2.add_run(f"Enkelbestemming '{bestemming}' (artikel .)")
    r2.bold = True; r2.underline = True
    r2.font.size = Pt(9); r2.font.name = "Arial"
    p2.paragraph_format.space_after = Pt(2)

    p3 = doc.add_paragraph()
    pPr3 = p3._p.get_or_add_pPr()
    shd3 = OxmlElement('w:shd')
    shd3.set(qn('w:val'),'clear'); shd3.set(qn('w:color'),'auto'); shd3.set(qn('w:fill'),GEEL_HEX)
    pPr3.append(shd3)
    r3 = p3.add_run("ALLEEN DE AFWIJKINGEN MOTIVEREN — voeg hier de motivering in")
    r3.font.size = Pt(9); r3.font.name = "Arial"
    r3.font.color.rgb = RGBColor(0x88, 0x88, 0x88); r3.italic = True
    p3.paragraph_format.space_after = Pt(6)
    voeg_witruimte_toe(doc)

    # ══════════════════════════════════════════════════════
    # SECTIE 9 — BENODIGDE ADVIEZEN
    # ══════════════════════════════════════════════════════
    voeg_sectieheader_toe(doc, "9. BENODIGDE ADVIEZEN")
    t = maak_tweekoloms_tabel(doc)
    for advies in ["ROM", "Ecologie", "Bodem", "Archeologie", "", ""]:
        voeg_rij_toe(t, advies, None, "geel")
    voeg_witruimte_toe(doc)

    # ══════════════════════════════════════════════════════
    # SECTIE 10 — FLORA & FAUNA
    # ══════════════════════════════════════════════════════
    voeg_sectieheader_toe(doc, "10. FLORA & FAUNA")
    t = maak_tweekoloms_tabel(doc)
    voeg_rij_toe(t, "Hyperlink SMP-kaart",
                 "https://regelink.webgispublisher.nl/Viewer.aspx?map=SMP",
                 "blauw")
    voeg_rij_toe(t, "SMP van toepassing?",    None, "geel")
    voeg_rij_toe(t, "Beschermde diersoorten", None, "geel")
    voeg_rij_toe(t, "Beschermde planten",     None, "geel")
    voeg_rij_toe(t, "Verbeelding SMP-kaart",  "[voeg schermafbeelding in]", "geel")
    voeg_witruimte_toe(doc)

    # ══════════════════════════════════════════════════════
    # TABBLAD 2 — PB-VELDCODE REFERENTIE (voor IT)
    # ══════════════════════════════════════════════════════

    # Paginabreak voor tweede sectie
    p_break = doc.add_paragraph()
    p_break.add_run().add_break(__import__('docx.enum.text', fromlist=['WD_BREAK_TYPE']).WD_BREAK_TYPE.PAGE)

    p_ref_titel = doc.add_paragraph()
    r = p_ref_titel.add_run("BIJLAGE — PowerBrowser Veldcode Referentie")
    r.bold = True; r.font.size = Pt(14); r.font.name = "Arial"
    r.font.color.rgb = BLAUW
    p_ref_titel.paragraph_format.space_after = Pt(4)

    p_ref_sub = doc.add_paragraph()
    r = p_ref_sub.add_run(
        "Overzicht van PowerBrowser-veldcodes gekoppeld aan de velden in dit formulier. "
        "Aan te leveren aan IT/PB-beheerder voor configuratie van het sjabloon."
    )
    r.font.size = Pt(9); r.font.name = "Arial"
    r.font.color.rgb = RGBColor(0x55, 0x55, 0x55); r.italic = True
    p_ref_sub.paragraph_format.space_after = Pt(8)

    # Referentietabel: formulierveld | PB-veldcode | bron
    PB_VELDEN = [
        # Sectie, Formulierveld, PB-veldcode, Bron
        ("Algemene zaakgegevens", "Zaak-ID",              "ZAAK.ZAAK_IDENTIFICATIE",              "PowerBrowser"),
        ("Algemene zaakgegevens", "Kenmerk",              "ZAAK.KENMERK",                         "PowerBrowser"),
        ("Algemene zaakgegevens", "Ontvangstdatum",       "PROCES.ONTVANGST",                     "PowerBrowser"),
        ("Algemene zaakgegevens", "DSO Verzoeknummer",    "ZAAK.DSO_VERZOEKNUMMER",               "PowerBrowser"),
        ("Algemene zaakgegevens", "Casemanager",          "BEHANDELAAR.SAMENGESTELDE_NAAM",       "PowerBrowser"),
        ("Algemene zaakgegevens", "Afdeling",             "BEHANDELAAR.AFDELING",                 "PowerBrowser"),
        ("Locatie",               "Straatnaam + huisnummer", "ZAAK_ADRES.SAMENGESTELD_STRAAT",    "PowerBrowser"),
        ("Locatie",               "Woonplaats",           "ZAAK_ADRES.WOONPLAATS",                "PowerBrowser"),
        ("Locatie",               "Kadastrale aanduiding","ZAAK_OBJECTADRES.PERCEELNUMMER",       "PowerBrowser"),
        ("Activiteiten",          "Activiteiten (lijst)", "ZAAK.ZAAK_ACTIVITEIT_OMS",             "PowerBrowser"),
        ("Activiteiten",          "Subproducten",         "SUBPRODUCT_ROW.SUBPRODUCT1-5",         "PowerBrowser"),
        ("Zaakgegevens ⚠",        "Aanvraagtype",         "ZAAK.ZAAKTYPE",                        "⚠ verifiëren"),
        ("Zaakgegevens ⚠",        "Omschrijving bouwplan","ZAAK.OMSCHRIJVING_KORT",               "⚠ verifiëren"),
        ("Zaakgegevens ⚠",        "Bijzonderheden",       "ZAAK.TOELICHTING",                     "⚠ verifiëren"),
        ("Zaakgegevens ⚠ checken", "Aanvraagtype",        "ZAAK.ZAAKTYPE",                        "PB — nog te verifiëren"),
        ("Zaakgegevens ⚠ checken", "Omschrijving bouwplan","ZAAK.OMSCHRIJVING_KORT",              "PB — nog te verifiëren"),
        ("Zaakgegevens ⚠ checken", "Bijzonderheden",      "ZAAK.TOELICHTING",                     "PB — nog te verifiëren"),
        ("Omgevingsplan info",    "Omgevingsplan",        "BESTEMMINGSPLAN_UITVOER.OMSCHRIJVING", "DSO API → PB"),
        ("Omgevingsplan info",    "Bestemmingsplan",      "BESTEMMINGSPLAN_UITVOER.OMSCHRIJVING", "DSO API → PB"),
        ("Omgevingsplan info",    "Datum vaststelling",   "BESTEMMINGSPLAN_UITVOER.DATUMVASTSTELLING", "DSO API → PB"),
        ("Omgevingsplan info",    "Datum onherroepelijk", "BESTEMMINGSPLAN_UITVOER.DATUMONHERROEPELIJK", "DSO API → PB"),
        ("Omgevingsplan info",    "Bestemming perceel",   "BESTEMMINGSPLAN_UITVOER.BESTEMMING_ID","DSO API → PB"),
        ("Omgevingsplan info",    "Status plan",          "BESTEMMINGSPLAN_UITVOER.STATUSBESTEMMINGSPLAN_ID", "DSO API → PB"),
    ]

    t_ref = doc.add_table(rows=1, cols=4)
    t_ref.style = 'Table Grid'
    t_ref.autofit = False
    t_ref.columns[0].width = Cm(3.5)
    t_ref.columns[1].width = Cm(5.5)
    t_ref.columns[2].width = Cm(7.5)
    t_ref.columns[3].width = Cm(3.4)
    for cel, hdr in zip(t_ref.rows[0].cells,
                        ["Sectie", "Formulierveld", "PB-veldcode", "Bron"]):
        set_cell_bg(cel, BLAUW_HEX)
        set_cell_borders(cel)
        cel_tekst(cel, hdr, vet=True, kleur=WIT, pt=9)

    vorige_sectie = ""
    for sectie, veld, code, bron in PB_VELDEN:
        rij = t_ref.add_row()
        cellen = rij.cells

        # Sectielabel alleen tonen als het verandert
        set_cell_borders(cellen[0])
        if sectie != vorige_sectie:
            set_cell_bg(cellen[0], GRIJS_HEX)
            cel_tekst(cellen[0], sectie, vet=True, pt=8)
            vorige_sectie = sectie
        else:
            set_cell_bg(cellen[0], "FFFFFF")
            cel_tekst(cellen[0], "", pt=8)

        set_cell_bg(cellen[1], "FFFFFF")
        set_cell_borders(cellen[1])
        cel_tekst(cellen[1], veld, pt=9)

        set_cell_bg(cellen[2], LICHTBLAUW_HEX)
        set_cell_borders(cellen[2])
        cel_tekst(cellen[2], code, cursief=True, pt=9)

        bron_kleur = GEEL_HEX if "DSO" in bron else ("FCE4D6" if "verifiëren" in bron else "FFFFFF")
        set_cell_bg(cellen[3], bron_kleur)
        set_cell_borders(cellen[3])
        cel_tekst(cellen[3], bron, pt=8)

    voeg_witruimte_toe(doc, 6)
    p_leg2 = doc.add_paragraph()
    for tekst, kleur_hex in [
        ("Lichtblauw", LICHTBLAUW_HEX), ("  = PB-veldcode    ", None),
        ("Geel", GEEL_HEX), ("  = nog te koppelen via DSO API    ", None),
    ]:
        r = p_leg2.add_run(tekst)
        r.font.size = Pt(8); r.font.name = "Arial"
        if kleur_hex:
            rPr = r._r.get_or_add_rPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), kleur_hex)
            rPr.append(shd)

    # ── Voettekst legenda ──
    p_voet = doc.add_paragraph()
    r_voet = p_voet.add_run(
        f"Gegenereerd op {date.today().strftime('%d-%m-%Y')} via DSO Intake Toets Generator v{VERSION}  |  "
        "Lichtblauw = automatisch (DSO-script)  |  Geel = handmatig invullen  |  "
        "«PB-veld» = later automatisch via PowerBrowser"
    )
    r_voet.font.size = Pt(7)
    r_voet.font.name = "Arial"
    r_voet.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
    p_voet.paragraph_format.space_before = Pt(8)

    # ── Opslaan ──
    if not uitvoer_pad:
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        adres_kort = adres.split(",")[0].replace(" ", "_").replace("/", "-")[:30]
        bestandsnaam = f"Intake_toets_{adres_kort}_{date.today().strftime('%Y%m%d')}.docx"
        uitvoer_pad = os.path.join(desktop, bestandsnaam)

    doc.save(uitvoer_pad)
    print(f"\n  ✓ Word-document opgeslagen: {uitvoer_pad}")
    return uitvoer_pad


# ── Standalone uitvoering ──────────────────────────────────────────────────────

if __name__ == "__main__":

    if not DSO_BESCHIKBAAR:
        # Test met dummy-data als DSO-script niet beschikbaar is
        print("⚠  Testmodus: DSO-script niet beschikbaar, gebruik voorbeelddata")
        test_data = {
            "adres": "Prinsengracht 40A, Amsterdam",
            "adres_gevonden": "Prinsengracht 40A, 1015 DV Amsterdam",
            "kadastrale_aanduiding": "ASD04-G-3456",
            "bestemmingsplan_naam": "Bestemmingsplan Grachtengordel West",
            "bestemmingsplan_datum": "2015-07-02",
            "hyperlink": "https://www.ruimtelijkeplannen.nl/viewer/viewer?planidn=NL.IMRO.0363.A1502BPSTD-VG01",
            "bestemming_perceel": "Gemengd - 1",
            "bestemmingstype": "enkelbestemming",
            "functieaanduidingen": ["wonen", "kantoor"],
            "dubbelbestemmingen": [{"naam": "Waarde - Cultuurhistorie", "artikelnummer": "15"}],
            "bouwaanduidingen": [],
            "maatvoeringen": [
                {"naam": "maximale bouwhoogte", "waarde": "15", "eenheid": "m"},
                {"naam": "maximale goothoogte", "waarde": "11", "eenheid": "m"},
            ],
            "planenoverzicht": {"omgevingsplan": [], "bestemmingsplan": [], "voorbereidingsbesluit": [], "beheersverordening": [], "inpassingsplan": []},
        }
        uitvoer = genereer_intake_toets(test_data)

    else:
        # Echte DSO-data ophalen
        if len(sys.argv) > 1:
            arg = " ".join(sys.argv[1:])
            if "," in arg and arg.replace(",","").replace(".","").replace(" ","").replace("-","").isdigit():
                parts = arg.replace(" ","").split(",")
                data = haal_data_voor_coordinaten(float(parts[0]), float(parts[1]))
            else:
                data = haal_data_voor_adres(arg)
        else:
            invoer_adres, invoer_x, invoer_y = vraag_invoer()
            if invoer_x is not None:
                data = haal_data_voor_coordinaten(invoer_x, invoer_y)
            else:
                data = haal_data_voor_adres(invoer_adres)

        uitvoer = genereer_intake_toets(data)
