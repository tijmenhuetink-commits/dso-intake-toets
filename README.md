# DSO Intake Toets Generator

Automatisch bestemmingsplandata ophalen en invullen in een Word-document (Intake Toets), via de Ruimtelijke Plannen API en PDOK Locatieserver.

## Gebruik

1. Vul een adres in
2. Klik op "Data ophalen"
3. Download het gegenereerde Word-document

## Wat wordt automatisch ingevuld?

| Veld | Bron |
|---|---|
| Adres (gevonden) | PDOK Locatieserver |
| Kadastrale aanduiding | PDOK Locatieserver |
| Naam vigerend bestemmingsplan | Ruimtelijke Plannen API v4 |
| Datum vaststelling | Ruimtelijke Plannen API v4 |
| Bestemming perceel | Ruimtelijke Plannen API v4 |
| Functieaanduiding | Ruimtelijke Plannen API v4 |
| Dubbelbestemming | Ruimtelijke Plannen API v4 |
| Bouwaanduiding | Ruimtelijke Plannen API v4 |
| Maatvoeringen | Ruimtelijke Plannen API v4 |
| Hyperlink ruimtelijkeplannen.nl | Automatisch gegenereerd |

## Lokaal draaien

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deployen op Streamlit Community Cloud

1. Zet deze repository op GitHub
2. Ga naar [share.streamlit.io](https://share.streamlit.io)
3. Koppel je GitHub repository
4. Klik Deploy

---
*Ontwikkeld voor gemeentelijke vergunningverlening — DSO Automatisering project*
