"""
DSO Bestemmingsplan Data Ophaler
================================
Versie : 3.3
Datum  : 2026-03-20
Wijzigingen:
  v0.1 — eerste versie
  v0.2 — Accept header gewijzigd naar application/hal+json
  v0.3 — body veldnaam gewijzigd naar "geometrie", debug-output toegevoegd
  v0.4 — body veldnaam gecorrigeerd naar "_geo" (gevonden via testscript)
  v0.5 — parapluplan-filtering toegevoegd, fallback voor bestemmingsvlakken
  v0.6 — detaildata (vlakken, aanduidingen, maatvoeringen) via IHR base URL
  v0.7 — planId in URL-pad gezet per Geoforum documentatie (correct endpoint)
  v0.8 — functieaanduidingen/maatvoeringen als top-level GET met ?bestemmingsvlak=
          gebiedsaanduidingen als top-level POST met ?planId= (uit OAS3 YAML spec)
  v0.9 — alle detail-endpoints onder /plannen/{planId}/ gezet op basis van Swagger
  v1.0 — datum uit planstatusInfo, type/hoofdgroep uit vlak, isParapluplan officieel
          maatvoering waarde uit omvang[] gelezen
  v1.1 — functieaanduidingen via GET, maatvoeringen via POST _geo
          omvang[] volledig uitgelezen met aparte naam per maat
  v1.2 — is_parapluplan() krijgt nu altijd een dict (was string op regel 199)
  v1.3 — haal_maatvoeringen() aanroep x/y parameters toegevoegd
  v1.4 — functieaanduidingen + maatvoeringen via _links van bestemmingsvlak
          JSON opgeslagen op Desktop met vast pad
  v1.5 — functieaanduidingen correct "geen" als _links leeg is
          maatvoeringen via bouwvlak _links (preciezer), fallback via _geo
  v1.6 — adresinvoer via prompt als geen argument opgegeven
  v1.7 — enkelbestemming wordt altijd verkozen boven dubbelbestemming
  v1.8 — dubbelbestemmingen correct via punt-zoek (niet perceel-geometrie)
          bouwaanduidingen toegevoegd via POST _zoek met punt
  v1.9 — planenoverzicht toegevoegd: bestemmingsplannen, omgevingsplannen,
          voorbereidingsbesluiten, beheersverordeningen en inpassingsplannen
  v2.0 — uitgebreide paraplu-detectie (datacenters, darkstores, etc.)
          oudste moederplan geselecteerd als basis (niet het nieuwste)
  v2.1 — RD-coördinaten als invoer toegevoegd (speld op de kaart)
          keuzemenu bij opstarten: adres / coördinaten / testadres
  v2.2 — reverse geocoding toegevoegd: coördinaten → dichtstbijzijnd adres
  v2.3 — nieuwste moederplan gekozen ipv oudste
          adresinvoer toont keuze bij meerdere resultaten
          stap 4 print altijd een label, maatvoeringen correct ingesprongen
  v2.4 — keuzemenu bij adres altijd tonen bij meerdere resultaten
          fallback naar ouder basisplan als vigerend plan geen vlakken heeft
  v2.5 — keuzemenu toont unieke adressen, 20 resultaten ophalen
  v2.6 — 100 resultaten ophalen, filter op exacte adresstart, duplicaten verwijderd
  v2.7 — filter op adres + komma zodat huisnummer 10,11,12 niet meer meekomen
          hint getoond als adres er niet bij staat: voeg plaatsnaam toe
  v2.8 — kadastrale aanduiding via Locatieserver lookup (gekoppeld_perceel)
  v2.9 — keuzemenu logica vereenvoudigd: exacte match = direct door
          Grote Straat 1 toont keuze, Prinsengracht 40A Amsterdam niet
  v3.0 — kadastrale aanduiding ook bij coördinateninvoer via reverse geocoding adres-ID
  v3.2 — niet-gedigitaliseerde plannen netjes afgehandeld:
          melding "niet beschikbaar via API" + hyperlink in resultaat en samenvatting
          vlag "niet_gedigitaliseerd: True" in JSON-output voor Word-generator
  v3.3 — paraplu-keywords uitgebreid: terrasregels, terrassen, detailhandel, TAM-omgevingsplan
          keuzemenu bij geen exacte adresMatch (was: stilletjes docs[0] kiezen)

Haalt automatisch bestemmingsplandata op voor een opgegeven adres.

Vult de volgende formuliervelden automatisch in:
  - Naam vigerend bestemmingsplan
  - Bestemming perceel
  - Dubbelbestemming
  - Functieaanduiding
  - Bebouwde oppervlakte bouwperceel (maatvoeringen)
  - Maximale bouwhoogte / goothoogte
  - Hyperlink ruimtelijkeplannen.nl
  - Kadastraal perceel (sectie + nummer)

Stap 1 : Adres → RD-coördinaten (PDOK Locatieserver, gratis, geen key)
Stap 2 : Coördinaten → Kadastraal perceel (PDOK Locatieserver, gratis)
Stap 3 : Coördinaten → Vigerend bestemmingsplan (Ruimtelijke Plannen API v4)
Stap 4 : Plan-ID → Bestemmingsvlak op locatie
Stap 5 : Bestemmingsvlak → Functieaanduidingen, dubbelbestemmingen, maatvoeringen

Benodigdheden:
  pip install requests
"""

import requests
import json
import sys

VERSION = "3.3"

# ─────────────────────────────────────────────
# CONFIGURATIE — pas hier je API-key aan
# ─────────────────────────────────────────────
RP_API_KEY = "085ebb90bd31d7ce9a6c3ebfb40745e5"   # Ruimtelijke Plannen API
RP_BASE    = "https://ruimte.omgevingswet.overheid.nl/ruimtelijke-plannen/api/opvragen/v4"
LS_BASE    = "https://api.pdok.nl/bzk/locatieserver/search/v3_1"


# ─────────────────────────────────────────────
# HULPFUNCTIES
# ─────────────────────────────────────────────

def rp_headers(met_body=False):
    """Geeft de standaard headers terug voor de Ruimtelijke Plannen API."""
    headers = {
        "X-Api-Key": RP_API_KEY,
        "Accept": "application/hal+json",   # ← API vereist hal+json, niet plain json
        "Content-Crs": "epsg:28992",
        "Accept-Crs": "epsg:28992",
    }
    if met_body:
        headers["Content-Type"] = "application/json"  # ← alleen bij POST met body
    return headers


def stap(nr, omschrijving):
    print(f"\n{'─'*55}")
    print(f"  Stap {nr}: {omschrijving}")
    print(f"{'─'*55}")


# ─────────────────────────────────────────────
# STAP 1 — Adres omzetten naar RD-coördinaten
# ─────────────────────────────────────────────

def adres_naar_rd(adres: str) -> dict:
    """
    Gebruikt de PDOK Locatieserver om een adres om te zetten naar
    RD-coördinaten (X/Y in EPSG:28992) én kadastrale gegevens.

    Bij meerdere resultaten wordt een keuzemenu getoond.
    Tip: voeg plaatsnaam toe als het adres er niet bij staat.

    Geeft een dict terug met:
      x, y          — RD-coördinaten
      weergavenaam  — volledig adres zoals gevonden
      kadastrale_aanduiding — bijv. "IJsselstein A 1234"
    """
    url = f"{LS_BASE}/free"
    params = {
        "q": adres,
        "fq": "type:adres",
        "rows": 100,
        "fl": "id,weergavenaam,centroide_rd,kadastrale_aanduiding",
    }

    resp = requests.get(url, params=params, timeout=10)
    resp.raise_for_status()
    data = resp.json()

    docs = data.get("response", {}).get("docs", [])
    if not docs:
        raise ValueError(f"Adres niet gevonden: '{adres}'")

    adres_lower = adres.strip().lower()

    def is_exacte_match(weergave):
        w = weergave.strip().lower()
        if w == adres_lower:
            return True
        # Splits "Straat HNR, POSTCODE Plaats" -> vergelijk met "Straat HNR, Plaats"
        delen = w.split(", ", 1)
        if len(delen) == 2:
            straat_hnr = delen[0]
            rest_delen = delen[1].split(" ", 1)
            plaatsnaam = rest_delen[1] if len(rest_delen) == 2 else delen[1]
            if f"{straat_hnr}, {plaatsnaam}" == adres_lower:
                return True
            if straat_hnr == adres_lower:
                return True
        return False

    exacte = []
    gezien = set()
    for d in docs:
        naam = d.get("weergavenaam", "")
        if is_exacte_match(naam) and naam not in gezien:
            gezien.add(naam)
            exacte.append(d)

    if len(exacte) == 1:
        doc = exacte[0]
    elif len(exacte) == 0:
        # Geen exacte match — toon de beste suggesties en vraag om bevestiging
        suggesties = docs[:5]
        print(f"  ⚠ Geen exacte match voor '{adres}'.")
        print(f"  Let op: onderstaande suggesties kunnen sterk afwijken.")
        print(f"  Controleer of de straatnaam correct is gespeld.")
        print(f"  Beste suggesties:")
        for i, d in enumerate(suggesties, 1):
            print(f"    {i}. {d.get('weergavenaam', '?')}")
        print(f"  (Of typ het adres opnieuw met plaatsnaam voor een betere match)")
        keuze = input(f"  Kies [1-{len(suggesties)}] of typ nieuw adres: ").strip()
        try:
            idx = int(keuze) - 1
            if not 0 <= idx < len(suggesties):
                idx = 0
            doc = suggesties[idx]
        except ValueError:
            if keuze:
                print(f"  Zoek opnieuw naar: '{keuze}'...")
                return adres_naar_rd(keuze)
            else:
                doc = suggesties[0]
    else:
        print(f"  {len(exacte)} adressen gevonden voor '{adres}':")
        for i, d in enumerate(exacte[:10], 1):
            print(f"    {i}. {d.get('weergavenaam', '?')}")
        print(f"  (Staat uw adres er niet bij? Typ het opnieuw met plaatsnaam)")
        keuze = input(f"  Kies [1-{min(len(exacte), 10)}] of typ nieuw adres: ").strip()
        try:
            idx = int(keuze) - 1
            if not 0 <= idx < len(exacte):
                idx = 0
            doc = exacte[idx]
        except ValueError:
            if keuze:
                print(f"  Zoek opnieuw naar: '{keuze}'...")
                return adres_naar_rd(keuze)
            else:
                doc = exacte[0]

    print(f"  Gevonden adres : {doc.get('weergavenaam', '—')}")

    # centroide_rd heeft formaat "POINT(x y)"
    rd_str = doc.get("centroide_rd", "")
    coords = rd_str.replace("POINT(", "").replace(")", "").split()
    if len(coords) != 2:
        raise ValueError(f"Kon RD-coördinaten niet lezen uit: '{rd_str}'")

    x, y = float(coords[0]), float(coords[1])
    print(f"  RD-coördinaten : X={x:.2f}, Y={y:.2f}")

    # Kadastrale gegevens via lookup op adres-ID
    kadastrale = "—"
    adres_id = doc.get("id", "")
    if adres_id:
        r_lookup = requests.get(
            f"{LS_BASE}/lookup",
            params={"id": adres_id, "fl": "gekoppeld_perceel"},
            timeout=10
        )
        if r_lookup.ok:
            lookup_docs = r_lookup.json().get("response", {}).get("docs", [])
            if lookup_docs:
                percelen = lookup_docs[0].get("gekoppeld_perceel", [])
                if percelen:
                    # Formatteer: "ASD08-L-2091" → "Amsterdam L 2091"
                    # Of toon de ruwe code als fallback
                    kadastrale = ", ".join(percelen)

    print(f"  Kadastrale aand.: {kadastrale}")

    return {
        "x": x,
        "y": y,
        "weergavenaam": doc.get("weergavenaam", adres),
        "kadastrale_aanduiding": kadastrale,
    }


# ─────────────────────────────────────────────
# STAP 2 — Vigerend bestemmingsplan ophalen
# ─────────────────────────────────────────────

def is_parapluplan(plan: dict) -> bool:
    """Geeft True als het plan officieel een parapluplan is (via isParapluplan veld),
    of als de naam typische paraplu-keywords bevat als fallback."""
    if plan.get("isParapluplan") is True:
        return True
    naam_lower = plan.get("naam", "").lower()
    keywords = [
        # Expliciete paraplu-aanduidingen
        "paraplu",
        # Thematische plannen die geen volledige bestemmingen bevatten
        "mantelzorg", "parkeer", "kruimelgeval", "wooneenheid",
        "bed and breakfast", "datacenter", "darkstore", "flitsbezorg",
        "kelders", "kelderbouw", "baliefunctie", "winkeldiversiteit",
        "grondwater", "staalslak", "drijvende bouwwerk", "hyperscale",
        "terrasregel", "terrassen", "detailhandel", "reclame",
        "TAM-omgevingsplan", "tam-omgevingsplan",
        # Procedurele plannen
        "voorbereidingsbesluit", "herziening",
    ]
    return any(kw in naam_lower for kw in keywords)


def haal_bestemmingsplan(x: float, y: float) -> dict | None:
    """
    Zoekt het vigerende bestemmingsplan op een RD-locatie via de
    Ruimtelijke Plannen API v4 (POST /plannen/_zoek).

    Haalt meerdere plannen op en filtert parapluplannen eruit zodat
    het echte moederbestemmingsplan wordt teruggegeven. Parapluplannen
    worden wél apart gerapporteerd.

    Geeft de plangegevens terug als dict, of None als er niets gevonden is.
    """
    url = f"{RP_BASE}/plannen/_zoek"
    params = {
        "planType": "bestemmingsplan",
        "planStatus": "vigerend",
        "page": 0,
        "pageSize": 10,   # meer ophalen zodat we kunnen filteren
    }
    body = {
        "_geo": {
            "intersects": {
                "type": "Point",
                "coordinates": [x, y]
            }
        }
    }

    resp = requests.post(url, headers=rp_headers(met_body=True), params=params,
                         json=body, timeout=15)
    if not resp.ok:
        print(f"  DEBUG response: {resp.text[:500]}")
    resp.raise_for_status()
    data = resp.json()

    plannen = data.get("_embedded", {}).get("plannen", [])
    if not plannen:
        print("  ⚠ Geen vigerend bestemmingsplan gevonden op deze locatie.")
        return None

    # Toon alle gevonden plannen
    print(f"  {len(plannen)} plan(nen) gevonden op deze locatie:")
    for p in plannen:
        tag = " [paraplu]" if is_parapluplan(p) else ""
        print(f"    • {p.get('naam','?')} ({p.get('datum','?')}){tag}")

    # Filter: houd alleen niet-parapluplannen over
    moederplannen = [p for p in plannen if not is_parapluplan(p)]
    parapluplannen = [p for p in plannen if is_parapluplan(p)]

    if not moederplannen:
        print("  ⚠ Alleen parapluplannen gevonden — gebruik eerste plan toch.")
        moederplannen = plannen  # fallback

    # Sorteer moederplannen op datum — oudste eerst is het moederplan
    # (parapluplannen zijn vaak recenter dan het moederplan)
    def plan_datum_sort(p):
        return (p.get("planstatusInfo") or {}).get("datum", "0000-00-00")

    moederplannen_gesorteerd = sorted(moederplannen, key=plan_datum_sort, reverse=True)
    plan = moederplannen_gesorteerd[0]  # nieuwste vigerende moederplan
    if len(moederplannen_gesorteerd) > 1:
        print(f"  ℹ Meerdere moederplannen — nieuwste gekozen als basis:")
        for mp in moederplannen_gesorteerd:
            datum = (mp.get("planstatusInfo") or {}).get("datum", "?")
            print(f"    • {mp.get('naam')} ({datum})")
    plan_id    = plan.get("id", "—")
    plan_naam  = plan.get("naam", "—")
    plan_datum = plan.get("planstatusInfo", {}).get("datum", "—")

    print(f"\n  ✓ Geselecteerd  : {plan_naam}")
    print(f"  Plan-ID         : {plan_id}")
    print(f"  Vastgesteld     : {plan_datum}")

    # Hyperlink naar ruimtelijkeplannen.nl
    hyperlink = f"https://www.ruimtelijkeplannen.nl/viewer/viewer?planidn={plan_id}"
    print(f"  Hyperlink       : {hyperlink}")

    # Parapluplannen apart bijhouden voor rapportage
    paraplu_namen = [p.get("naam", "?") for p in parapluplannen]
    if paraplu_namen:
        print(f"  Parapluplannen  : {', '.join(paraplu_namen)}")

    # Haal ook alle andere plantypen op voor het overzicht
    alle_plannen = data.get("_embedded", {}).get("plannen", [])

    # Relevante plantypen voor vergunningverlening
    relevante_types = {
        "bestemmingsplan"     : [],
        "omgevingsplan"       : [],
        "voorbereidingsbesluit": [],
        "beheersverordening"  : [],
        "inpassingsplan"      : [],
    }

    # Haal ook andere plantypen op via aparte zoekopdracht
    for plantype in ["omgevingsplan", "voorbereidingsbesluit",
                     "beheersverordening", "inpassingsplan"]:
        r_extra = requests.post(
            f"{RP_BASE}/plannen/_zoek",
            headers=rp_headers(met_body=True),
            params={"planType": plantype, "planStatus": "vigerend",
                    "page": 0, "pageSize": 5},
            json=body, timeout=10
        )
        if r_extra.ok:
            extra = (r_extra.json().get("_embedded") or {}).get("plannen", [])
            for p in extra:
                relevante_types[plantype].append({
                    "naam" : p.get("naam", "—"),
                    "datum": (p.get("planstatusInfo") or {}).get("datum", "—"),
                    "id"   : p.get("id", "—"),
                })

    # Voeg bestemmingsplannen toe
    for p in plannen:
        relevante_types["bestemmingsplan"].append({
            "naam"    : p.get("naam", "—"),
            "datum"   : (p.get("planstatusInfo") or {}).get("datum", "—"),
            "id"      : p.get("id", "—"),
            "paraplu" : is_parapluplan(p),
        })

    # Print relevant overzicht
    print(f"\n  Plannen-overzicht op deze locatie:")
    for plantype, plannen_lijst in relevante_types.items():
        if plannen_lijst:
            for p in plannen_lijst:
                paraplu_tag = " [paraplu]" if p.get("paraplu") else ""
                print(f"    [{plantype}] {p['naam']} ({p['datum']}){paraplu_tag}")

    # Bewaar alle moederplannen voor fallback bij lege vlakken
    alle_moeder = [
        {
            "id"   : p.get("id","—"),
            "naam" : p.get("naam","—"),
            "datum": (p.get("planstatusInfo") or {}).get("datum","—"),
        }
        for p in moederplannen_gesorteerd
    ]

    return {
        "id"              : plan_id,
        "naam"            : plan_naam,
        "datum"           : plan_datum,
        "hyperlink"       : hyperlink,
        "parapluplannen"  : paraplu_namen,
        "planenoverzicht" : relevante_types,
        "alle_moederplannen": alle_moeder,
    }


# ─────────────────────────────────────────────
# STAP 3 — Bestemmingsvlak ophalen op locatie
# ─────────────────────────────────────────────

def haal_bestemmingsvlak(plan_id: str, x: float, y: float) -> dict | None:
    """
    Haalt het bestemmingsvlak op dat de opgegeven locatie bevat.
    Gebruikt de IHR (data.informatiehuisruimte.nl) API met _geo als
    query parameter voor de ruimtelijke zoekopdracht.
    """
    # planId zit in het PAD (niet als query param)
    url = f"{RP_BASE}/plannen/{plan_id}/bestemmingsvlakken/_zoek"
    params = {"pageSize": 10}
    body = {
        "_geo": {
            "intersects": {
                "type": "Point",
                "coordinates": [x, y]
            }
        }
    }

    resp = requests.post(url, headers=rp_headers(met_body=True), params=params,
                         json=body, timeout=15)
    if not resp.ok:
        print(f"  DEBUG {resp.status_code}: {resp.text[:300]}")
    resp.raise_for_status()
    data = resp.json()

    vlakken = (data.get("_embedded") or {}).get("bestemmingsvlakken", [])
    if not vlakken:
        print("  ⚠ Geen bestemmingsvlak gevonden in dit plan.")
        print("  ℹ Het plan is mogelijk niet gedigitaliseerd in de Ruimtelijke Plannen API.")
        print("  ℹ Raadpleeg ruimtelijkeplannen.nl voor de plankaart.")
        return None

    # Splits in enkelvlakken en dubbelbestemmingen
    enkelvlakken  = [v for v in vlakken if v.get("type") == "enkelbestemming"]
    dubbelVlakken = [v for v in vlakken if v.get("type") == "dubbelbestemming"]

    print(f"  {len(vlakken)} bestemmingsvlak(ken) gevonden op deze locatie:")
    for v in vlakken:
        print(f"    • [{v.get('type','?')}] {v.get('naam','?')} (art. {v.get('artikelnummer','—')})")

    # Kies de enkelstemming als primaire bestemming
    if enkelvlakken:
        vlak = enkelvlakken[0]
    else:
        vlak = vlakken[0]
        print(f"  ℹ Geen enkelbestemming — eerste vlak gebruikt")

    # Sla dubbelbestemmingen op met alle beschikbare info
    dubbel_namen = []
    for dv in dubbelVlakken:
        info = {
            "naam"         : dv.get("naam", "—"),
            "artikelnummer": dv.get("artikelnummer", "—"),
            "id"           : dv.get("id", "—"),
        }
        # Voeg tekstlink toe als beschikbaar
        dv_links = dv.get("_links", {})
        teksten  = dv_links.get("teksten", [])
        if teksten:
            href = teksten[0].get("href") if isinstance(teksten[0], dict) else teksten[0]
            info["tekst_url"] = href
        dubbel_namen.append(info)
    vlak_id     = vlak.get("id", "—")
    naam        = vlak.get("naam", "—")
    btype       = vlak.get("type", "—")
    bhoofdgroep = vlak.get("bestemmingshoofdgroep", "—")

    print(f"  ✓ Geselecteerd : {naam}")
    print(f"  Type           : {btype} / {bhoofdgroep}")
    print(f"  Vlak-ID        : {vlak_id}")
    if dubbel_namen:
        print(f"  Dubbelbestemm. : {len(dubbel_namen)}x gevonden:")
        for d in dubbel_namen:
            print(f"    • {d['naam']} (art. {d['artikelnummer']})")

    links = vlak.get("_links", {})
    return {
        "id"             : vlak_id,
        "naam"           : naam,
        "type"           : btype,
        "hoofdgroep"     : bhoofdgroep,
        "links"          : links,
        "dubbelbestemmingen": dubbel_namen,
    }


# ─────────────────────────────────────────────
# STAP 4 — Functieaanduidingen ophalen
# ─────────────────────────────────────────────

def haal_functieaanduidingen(plan_id: str, vlak_id: str, x: float, y: float,
                             vlak_links: dict = None) -> list[str]:
    """
    Haalt de functieaanduidingen op via de _links van het bestemmingsvlak.
    Als _links leeg is, heeft dit vlak geen functieaanduidingen — dat is normaal.
    """
    # _links is een lijst van hrefs, of leeg als er geen aanduidingen zijn
    fa_links = (vlak_links or {}).get("functieaanduidingen", [])

    if not fa_links:
        print("  Functieaand.   : geen")
        return []

    # Volg elke link en verzamel de namen
    resultaat = []
    seen = set()
    for link in fa_links if isinstance(fa_links, list) else [fa_links]:
        href = link.get("href") if isinstance(link, dict) else None
        if not href:
            continue
        resp = requests.get(href, headers=rp_headers(), timeout=15)
        if resp.ok:
            item = resp.json()
            naam = item.get("naam", "—")
            if naam not in seen:
                seen.add(naam)
                resultaat.append(naam)
    # Early return hier, print gebeurt na
    return resultaat
    resp.raise_for_status()
    data = resp.json()

    if resultaat:
        print(f"  Functieaand.   : {', '.join(resultaat)}")
    return resultaat


# ─────────────────────────────────────────────
# STAP 5 — Dubbelbestemmingen ophalen
# ─────────────────────────────────────────────

def haal_dubbelbestemmingen(plan_id: str, x: float, y: float) -> list[str]:
    """
    Haalt dubbelbestemmingen op (opgeslagen als gebiedsaanduidingen
    van het type 'dubbelbestemming') op de locatie binnen het plan.
    """
    url = f"{RP_BASE}/plannen/{plan_id}/gebiedsaanduidingen/_zoek"
    params = {"pageSize": 50}
    body = {
        "_geo": {
            "intersects": {
                "type": "Point",
                "coordinates": [x, y]
            }
        }
    }
    resp = requests.post(url, headers=rp_headers(met_body=True), params=params,
                         json=body, timeout=15)
    if not resp.ok:
        print(f"  DEBUG gebiedsaanduidingen: {resp.status_code} {resp.text[:200]}")
    resp.raise_for_status()
    data = resp.json()

    items = (data.get("_embedded") or {}).get("gebiedsaanduidingen", [])

    # Filter op type dubbelbestemming
    dubbel = [
        item.get("naam", "—")
        for item in items
        if "dubbelbestemming" in item.get("type", "").lower()
        or "dubbelbestemming" in item.get("naam", "").lower()
    ]

    if dubbel:
        print(f"  Dubbelbestemm. : {', '.join(dubbel)}")
    else:
        print("  Dubbelbestemm. : geen")

    # Geef alle gebiedsaanduidingen ook mee voor volledigheid
    alle = [item.get("naam", "—") for item in items]
    return dubbel, alle


# ─────────────────────────────────────────────
# STAP 6 — Maatvoeringen ophalen
# ─────────────────────────────────────────────

def haal_maatvoeringen(plan_id: str, vlak_id: str, x: float, y: float,
                       vlak_links: dict = None) -> list[dict]:
    """
    Haalt maatvoeringen op. Strategie:
    1. Via bouwvlakken uit _links van het bestemmingsvlak (meest precies)
    2. Fallback: POST _zoek met _geo op het punt
    """
    resultaat = []

    # Strategie 1: haal maatvoeringen via de bouwvlakken van dit bestemmingsvlak
    bouwvlak_links = (vlak_links or {}).get("bouwvlakken", [])
    if bouwvlak_links:
        for bv_link in bouwvlak_links:
            bv_href = bv_link.get("href") if isinstance(bv_link, dict) else None
            if not bv_href:
                continue
            # Haal maatvoeringen op voor dit bouwvlak
            maatv_url = bv_href.replace("/bouwvlakken/", "/bouwvlakken/").rstrip("/")
            # Gebruik het bouwvlak-ID om maatvoeringen te zoeken
            bv_id = bv_href.rstrip("/").split("/")[-1]
            r = requests.get(
                f"{RP_BASE}/plannen/{plan_id}/maatvoeringen",
                headers=rp_headers(),
                params={"bouwvlak": bv_id, "pageSize": 20},
                timeout=15)
            if r.ok:
                items = (r.json().get("_embedded") or {}).get("maatvoeringen", [])
                if items:
                    for item in items:
                        for o in item.get("omvang", []):
                            naam = o.get("naam", "—")
                            waarde = o.get("waarde", "—")
                            resultaat.append({"naam": naam, "waarde": waarde, "eenheid": "m"})
                    if resultaat:
                        return resultaat  # gevonden via bouwvlak

    # Strategie 2: POST _zoek met _geo op de locatie
    url = f"{RP_BASE}/plannen/{plan_id}/maatvoeringen/_zoek"
    body = {"_geo": {"intersects": {"type": "Point", "coordinates": [x, y]}}}
    resp = requests.post(url, headers=rp_headers(met_body=True),
                         json=body, params={"pageSize": 20}, timeout=15)
    resp.raise_for_status()
    data = resp.json()
    items = (data.get("_embedded") or {}).get("maatvoeringen", [])

    for item in items:
        naam   = item.get("naam", "—")
        omvang = item.get("omvang", [])
        # omvang is een lijst van {naam, waarde} — elke maat apart uitprinten
        if omvang:
            for o in omvang:
                onaam   = o.get("naam", naam)
                owaarde = o.get("waarde", "—")
                print(f"  Maatvoering    : {onaam} = {owaarde} m")
                resultaat.append({"naam": onaam, "waarde": owaarde, "eenheid": "m"})
        else:
            waarde = item.get("waarde", "—")
            print(f"  Maatvoering    : {naam} = {waarde}")
            resultaat.append({"naam": naam, "waarde": waarde, "eenheid": ""})

    if not resultaat:
        print("  Maatvoeringen  : geen gevonden voor dit vlak")

    return resultaat


# ─────────────────────────────────────────────
# HOOFDFUNCTIE — alles samenvoegen
# ─────────────────────────────────────────────

def haal_data_voor_coordinaten(x: float, y: float) -> dict:
    """
    Directe invoer via RD-coördinaten (speld op de kaart).
    Slaat stap 1 (geocodering) over en gaat direct naar stap 2.
    """
    print(f"\n{'═'*55}")
    print(f"  DSO Bestemmingsplan Data Ophaler  v{VERSION}")
    print(f"  Coördinaten: X={x:.2f}, Y={y:.2f} (RD)")
    print(f"{'═'*55}")

    # Reverse geocoding: coördinaten → dichtstbijzijnd adres + kadastrale aanduiding
    adres_gevonden = f"RD: X={x:.2f}, Y={y:.2f}"
    kadastrale_aanduiding = "—"
    try:
        r_rev = requests.get(
            f"{LS_BASE}/reverse",
            params={"X": x, "Y": y, "type": "adres", "rows": 1},
            timeout=10
        )
        if r_rev.ok:
            docs = r_rev.json().get("response", {}).get("docs", [])
            if docs:
                adres_gevonden = docs[0].get("weergavenaam", adres_gevonden)
                print(f"  Dichtstbijzijnd adres: {adres_gevonden}")
                # Kadastrale aanduiding via lookup op adres-ID
                adres_id = docs[0].get("id", "")
                if adres_id:
                    r_lookup = requests.get(
                        f"{LS_BASE}/lookup",
                        params={"id": adres_id, "fl": "gekoppeld_perceel"},
                        timeout=10
                    )
                    if r_lookup.ok:
                        lookup_docs = r_lookup.json().get("response", {}).get("docs", [])
                        if lookup_docs:
                            percelen = lookup_docs[0].get("gekoppeld_perceel", [])
                            if percelen:
                                kadastrale_aanduiding = ", ".join(percelen)
                print(f"  Kadastrale aand.: {kadastrale_aanduiding}")
    except Exception:
        pass  # geen adres gevonden, coördinaten gebruiken

    resultaat = {
        "adres": f"RD: {x:.2f}, {y:.2f}",
        "adres_gevonden": adres_gevonden,
        "kadastrale_aanduiding": kadastrale_aanduiding,
        "bestemmingsplan_naam": "—",
        "bestemmingsplan_datum": "—",
        "hyperlink": "—",
        "bestemming_perceel": "—",
        "bestemmingstype": "—",
        "functieaanduidingen": [],
        "dubbelbestemmingen": [],
        "bouwaanduidingen": [],
        "maatvoeringen": [],
    }

    # Stap 2 t/m 6 direct uitvoeren met de opgegeven coördinaten
    stap(2, "Vigerend bestemmingsplan ophalen")
    plan = haal_bestemmingsplan(x, y)
    if not plan:
        print("\n  ✗ Geen bestemmingsplan gevonden op deze locatie.")
        return resultaat

    resultaat["bestemmingsplan_naam"]  = plan["naam"]
    resultaat["bestemmingsplan_datum"] = plan["datum"]
    resultaat["hyperlink"]             = plan["hyperlink"]
    resultaat["planenoverzicht"]       = plan.get("planenoverzicht", {})

    stap(3, "Bestemmingsvlak ophalen")
    vlak = haal_bestemmingsvlak(plan["id"], x, y)

    # Fallback: probeer andere moederplannen als dit plan geen vlakken heeft
    if not vlak and "alle_moederplannen" in plan:
        for ouder_plan in plan["alle_moederplannen"][1:]:
            print(f"  ↩ Probeer ouder plan: {ouder_plan['naam']} ({ouder_plan['datum']})")
            vlak = haal_bestemmingsvlak(ouder_plan["id"], x, y)
            if vlak:
                print(f"  ✓ Bestemmingsvlak gevonden in: {ouder_plan['naam']}")
                resultaat["bestemmingsplan_naam"] = ouder_plan["naam"]
                resultaat["bestemmingsplan_datum"] = ouder_plan["datum"]
                resultaat["hyperlink"] = f"https://www.ruimtelijkeplannen.nl/viewer/viewer?planidn={ouder_plan['id']}"
                break

    if not vlak:
        print(f"\n  ⚠ Bestemmingsvlakken niet beschikbaar via API.")
        print(f"  ℹ Raadpleeg het plan handmatig via:")
        print(f"    {resultaat['hyperlink']}")
        resultaat["bestemming_perceel"] = "⚠ Niet beschikbaar via API — zie hyperlink"
        resultaat["niet_gedigitaliseerd"] = True
        return resultaat
    resultaat["bestemming_perceel"] = vlak["naam"]
    resultaat["bestemmingstype"]    = vlak["type"]

    stap(4, "Functieaanduidingen ophalen")
    resultaat["functieaanduidingen"] = haal_functieaanduidingen(
        plan["id"], vlak["id"], x, y, vlak.get("links", {}))

    stap(5, "Dubbelbestemmingen & bouwaanduidingen ophalen")
    dubbel_uit_vlak = vlak.get("dubbelbestemmingen", [])
    if dubbel_uit_vlak:
        for d in dubbel_uit_vlak:
            print(f"  Dubbelbestemm. : {d['naam']} (art. {d['artikelnummer']})")
    else:
        print("  Dubbelbestemm. : geen")
    resultaat["dubbelbestemmingen"] = dubbel_uit_vlak

    bouw_url = f"{RP_BASE}/plannen/{plan['id']}/bouwaanduidingen/_zoek"
    bouw_body = {"_geo": {"intersects": {"type": "Point", "coordinates": [x, y]}}}
    br = requests.post(bouw_url, headers=rp_headers(met_body=True),
                       json=bouw_body, params={"pageSize": 20}, timeout=15)
    bouwaanduidingen = []
    if br.ok:
        for ba in (br.json().get("_embedded") or {}).get("bouwaanduidingen", []):
            info = {"naam": ba.get("naam","—"), "artikelnummer": ba.get("artikelnummer","—")}
            bouwaanduidingen.append(info)
            print(f"  Bouwaanduiding : {info['naam']} (art. {info['artikelnummer']})")
    if not bouwaanduidingen:
        print("  Bouwaanduidingen: geen")
    resultaat["bouwaanduidingen"] = bouwaanduidingen

    stap(6, "Maatvoeringen ophalen")
    resultaat["maatvoeringen"] = haal_maatvoeringen(
        plan["id"], vlak["id"], x, y, vlak.get("links", {}))

    # Samenvatting
    _print_samenvatting(resultaat)
    return resultaat


def _print_samenvatting(resultaat: dict):
    """Print de samenvatting van de opgehaalde data."""
    print(f"\n{'═'*55}")
    print("  SAMENVATTING — formuliervelden")
    print(f"{'═'*55}")
    print(f"  Adres                    : {resultaat.get('adres_gevonden', resultaat.get('adres','—'))}")
    print(f"  Kadastraal perceel       : {resultaat.get('kadastrale_aanduiding','—')}")
    print(f"  Naam bestemmingsplan     : {resultaat.get('bestemmingsplan_naam','—')}")
    print(f"  Datum vastgesteld        : {resultaat.get('bestemmingsplan_datum','—')}")
    print(f"  Bestemming perceel       : {resultaat.get('bestemming_perceel','—')}")
    print(f"  Bestemmingstype          : {resultaat.get('bestemmingstype','—')}")
    if resultaat.get("niet_gedigitaliseerd"):
        print(f"  ⚠ Dit plan is niet gedigitaliseerd in de Ruimtelijke Plannen API.")
        print(f"  ℹ Bestemming, dubbelbestemming en maatvoeringen handmatig opzoeken via:")
        print(f"    {resultaat.get('hyperlink','—')}")
    funct = ', '.join(resultaat.get('functieaanduidingen',[])) or 'geen'
    print(f"  Functieaanduiding        : {funct}")
    dubbel_lijst = resultaat.get("dubbelbestemmingen", [])
    if dubbel_lijst:
        for d in dubbel_lijst:
            url = f" → {d['tekst_url']}" if d.get("tekst_url") else ""
            print(f"  Dubbelbestemming         : {d['naam']} (art. {d['artikelnummer']}){url}")
    else:
        print(f"  Dubbelbestemming         : geen")
    bouw_lijst = resultaat.get("bouwaanduidingen", [])
    if bouw_lijst:
        for b in bouw_lijst:
            print(f"  Bouwaanduiding           : {b['naam']} (art. {b['artikelnummer']})")
    else:
        print(f"  Bouwaanduiding           : geen")
    for m in resultaat.get('maatvoeringen', []):
        eenheid = m.get('eenheid', '')
        waarde_str = f"{m['waarde']} {eenheid}".strip()
        print(f"  {m['naam']:<30}: {waarde_str}")
    print(f"  Hyperlink                : {resultaat.get('hyperlink','—')}")
    overzicht = resultaat.get("planenoverzicht", {})
    if any(v for v in overzicht.values()):
        print(f"\n  Plannen op deze locatie:")
        for plantype, plannen_lijst in overzicht.items():
            for p in plannen_lijst:
                paraplu_tag = " [paraplu]" if p.get("paraplu") else ""
                print(f"    [{plantype}] {p['naam']} ({p['datum']}){paraplu_tag}")
    print(f"{'═'*55}\n")


def haal_data_voor_adres(adres: str) -> dict:
    """
    Haalt alle bestemmingsplandata op voor een adres en
    geeft een gestructureerde dict terug klaar voor gebruik.
    """
    print(f"\n{'═'*55}")
    print(f"  DSO Bestemmingsplan Data Ophaler  v{VERSION}")
    print(f"  Adres: {adres}")
    print(f"{'═'*55}")

    resultaat = {
        "adres": adres,
        "kadastrale_aanduiding": "—",
        "bestemmingsplan_naam": "—",
        "bestemmingsplan_datum": "—",
        "hyperlink": "—",
        "bestemming_perceel": "—",
        "bestemmingstype": "—",
        "functieaanduidingen": [],
        "dubbelbestemmingen": [],
        "alle_gebiedsaanduidingen": [],
        "maatvoeringen": [],
    }

    # Stap 1: adres → coördinaten + kadaster
    stap(1, "Adres omzetten naar RD-coördinaten")
    locatie = adres_naar_rd(adres)
    x, y = locatie["x"], locatie["y"]
    resultaat["adres_gevonden"]         = locatie["weergavenaam"]
    resultaat["kadastrale_aanduiding"]  = locatie["kadastrale_aanduiding"]

    # Stap 2: coördinaten → bestemmingsplan
    stap(2, "Vigerend bestemmingsplan ophalen")
    plan = haal_bestemmingsplan(x, y)
    if not plan:
        print("\n  ✗ Geen bestemmingsplan gevonden. Script stopt hier.")
        return resultaat

    resultaat["bestemmingsplan_naam"]  = plan["naam"]
    resultaat["bestemmingsplan_datum"] = plan["datum"]
    resultaat["hyperlink"]             = plan["hyperlink"]

    # Stap 3: bestemmingsvlak ophalen
    stap(3, "Bestemmingsvlak ophalen")
    vlak = haal_bestemmingsvlak(plan["id"], x, y)

    # Fallback: probeer andere moederplannen als dit plan geen vlakken heeft
    if not vlak and "alle_moederplannen" in plan:
        for ouder_plan in plan["alle_moederplannen"][1:]:
            print(f"  ↩ Probeer ouder plan: {ouder_plan['naam']} ({ouder_plan['datum']})")
            vlak = haal_bestemmingsvlak(ouder_plan["id"], x, y)
            if vlak:
                print(f"  ✓ Bestemmingsvlak gevonden in: {ouder_plan['naam']}")
                resultaat["bestemmingsplan_naam"] = ouder_plan["naam"]
                resultaat["bestemmingsplan_datum"] = ouder_plan["datum"]
                resultaat["hyperlink"] = f"https://www.ruimtelijkeplannen.nl/viewer/viewer?planidn={ouder_plan['id']}"
                break

    if not vlak:
        print(f"\n  ⚠ Bestemmingsvlakken niet beschikbaar via API.")
        print(f"  ℹ Raadpleeg het plan handmatig via:")
        print(f"    {resultaat['hyperlink']}")
        resultaat["bestemming_perceel"] = "⚠ Niet beschikbaar via API — zie hyperlink"
        resultaat["niet_gedigitaliseerd"] = True
        return resultaat

    resultaat["bestemming_perceel"] = vlak["naam"]
    resultaat["bestemmingstype"]    = vlak["type"]

    # Stap 4: functieaanduidingen
    stap(4, "Functieaanduidingen ophalen")
    resultaat["functieaanduidingen"] = haal_functieaanduidingen(
        plan["id"], vlak["id"], x, y, vlak.get("links", {}))

    # Stap 5: dubbelbestemmingen
    stap(5, "Dubbelbestemmingen & bouwaanduidingen ophalen")
    # Dubbelbestemmingen zijn al gevonden in stap 3
    dubbel_uit_vlak = vlak.get("dubbelbestemmingen", [])
    if dubbel_uit_vlak:
        for d in dubbel_uit_vlak:
            print(f"  Dubbelbestemm. : {d['naam']} (art. {d['artikelnummer']})")
            if d.get("tekst_url"):
                print(f"    Tekst: {d['tekst_url']}")
    else:
        print("  Dubbelbestemm. : geen")
    resultaat["dubbelbestemmingen"] = dubbel_uit_vlak

    # Bouwaanduidingen via POST _zoek met punt
    bouw_url = f"{RP_BASE}/plannen/{plan['id']}/bouwaanduidingen/_zoek"
    bouw_body = {"_geo": {"intersects": {"type": "Point", "coordinates": [x, y]}}}
    br = requests.post(bouw_url, headers=rp_headers(met_body=True),
                       json=bouw_body, params={"pageSize": 20}, timeout=15)
    bouwaanduidingen = []
    if br.ok:
        ba_items = (br.json().get("_embedded") or {}).get("bouwaanduidingen", [])
        for ba in ba_items:
            info = {
                "naam"         : ba.get("naam", "—"),
                "artikelnummer": ba.get("artikelnummer", "—"),
            }
            bouwaanduidingen.append(info)
            print(f"  Bouwaanduiding : {info['naam']} (art. {info['artikelnummer']})")
    if not bouwaanduidingen:
        print("  Bouwaanduidingen: geen")
    resultaat["bouwaanduidingen"] = bouwaanduidingen

    # Stap 6: maatvoeringen
    stap(6, "Maatvoeringen ophalen")
    resultaat["maatvoeringen"] = haal_maatvoeringen(plan["id"], vlak["id"], x, y)

    _print_samenvatting(resultaat)

    return resultaat


# ─────────────────────────────────────────────
# UITVOEREN
# ─────────────────────────────────────────────

if __name__ == "__main__":

    # ── Invoer: adres of RD-coördinaten ──
    def vraag_invoer():
        print("=" * 55)
        print("  Kies invoermethode:")
        print("  1. Adres (bijv. Kerkstraat 1, IJsselstein)")
        print("  2. RD-coördinaten (bijv. 131653, 447223)")
        print("  3. Testadres (Prinsengracht 40A, Amsterdam)")
        print("=" * 55)
        keuze = input("  Keuze [1/2/3]: ").strip()

        if keuze == "2":
            print("  Tip: coördinaten vind je in de DSO viewer")
            print("  (klik op de kaart, lees X en Y af linksboven)")
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

    if len(sys.argv) > 1:
        # Command-line argument: adres of "X,Y"
        arg = " ".join(sys.argv[1:])
        if "," in arg and arg.replace(",","").replace(".","").replace(" ","").replace("-","").isdigit():
            parts = arg.replace(" ","").split(",")
            invoer_adres, invoer_x, invoer_y = None, float(parts[0]), float(parts[1])
        else:
            invoer_adres, invoer_x, invoer_y = arg, None, None
    else:
        invoer_adres, invoer_x, invoer_y = vraag_invoer()

    try:
        if invoer_x is not None:
            data = haal_data_voor_coordinaten(invoer_x, invoer_y)
        else:
            data = haal_data_voor_adres(invoer_adres)

        # Optioneel: sla de ruwe data op als JSON
        import os
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        json_pad = os.path.join(desktop, "dso_resultaat.json")
        with open(json_pad, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"  JSON opgeslagen in: {json_pad}")

    except requests.HTTPError as e:
        print(f"\n  ✗ API-fout: {e.response.status_code} — {e.response.text[:300]}")
    except Exception as e:
        print(f"\n  ✗ Fout: {e}")
