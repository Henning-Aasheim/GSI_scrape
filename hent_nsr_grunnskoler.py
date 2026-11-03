"""Hent aktive og nedlagte grunnskoler fra Nasjonalt skoleregister (NSR) API v4.

Installer:
    pip install requests pandas openpyxl

Kjor:
    python hent_nsr_grunnskoler.py

Utdata:
    nsr_grunnskoler.csv
    nsr_grunnskoler.xlsx
    nsr_grunnskoler_raw.jsonl
"""

from __future__ import annotations

import json
import time
from pathlib import Path

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://data-nsr.udir.no/v4"
ANTALL_PER_SIDE = 1000
PAUSE_MELLOM_DETALJKALL = 0.05
TA_MED_NEDLAGTE = True

OUT_CSV = Path("nsr_grunnskoler.csv")
OUT_XLSX = Path("nsr_grunnskoler.xlsx")
OUT_JSONL = Path("nsr_grunnskoler_raw.jsonl")


def make_session() -> requests.Session:
    retry = Retry(
        total=6,
        connect=6,
        read=6,
        backoff_factor=1,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET"}),
        respect_retry_after_header=True,
    )
    session = requests.Session()
    session.headers.update({"Accept": "application/json", "User-Agent": "NSR-datauttrekk/1.0"})
    session.mount("https://", HTTPAdapter(max_retries=retry))
    return session


def get_json(session: requests.Session, url: str, params: dict | None = None) -> dict:
    response = session.get(url, params=params, timeout=60)
    response.raise_for_status()
    return response.json()


def hent_alle_enheter(session: requests.Session) -> list[dict]:
    """Henter den paginerte oversikten fra GET /v4/enheter."""
    alle = []
    side = 1

    while True:
        data = get_json(
            session,
            f"{BASE_URL}/enheter",
            params={"sidenummer": side, "antallperside": ANTALL_PER_SIDE},
        )
        enheter = data.get("EnhetListe") or []
        alle.extend(enheter)

        antall_sider = data.get("AntallSider") or side
        print(f"Oversikt: side {side}/{antall_sider}, totalt hentet {len(alle)}")

        if side >= antall_sider or not enheter:
            break
        side += 1

    return alle


def er_relevant_grunnskole(enhet: dict) -> bool:
    """Beholder grunnskoler; eventuelt bare aktive grunnskoler."""
    if not enhet.get("ErSkole", False):
        return False
    if not enhet.get("ErGrunnskole", False):
        return False
    if not TA_MED_NEDLAGTE and not enhet.get("ErAktiv", False):
        return False
    return True


def hent_detaljer(session: requests.Session, organisasjonsnummer: str) -> dict:
    return get_json(session, f"{BASE_URL}/enhet/{organisasjonsnummer}")


def get_nested(obj: dict, *keys, default=None):
    current = obj
    for key in keys:
        if not isinstance(current, dict):
            return default
        current = current.get(key)
        if current is None:
            return default
    return current


def list_to_text(values: list[dict] | None, *keys) -> str:
    result = []
    for item in values or []:
        value = get_nested(item, *keys)
        if value not in (None, ""):
            result.append(str(value))
    return " | ".join(result)


def finn_eier(enhet: dict) -> tuple[str | None, str | None]:
    """Finner eierrelasjonen. Foretrekker relasjonstype med id 1 eller navnet Eierstruktur."""
    relasjoner = enhet.get("ForeldreRelasjoner") or []
    for relasjon in relasjoner:
        rel_id = str(get_nested(relasjon, "Relasjonstype", "Id", default=""))
        rel_navn = str(get_nested(relasjon, "Relasjonstype", "Navn", default=""))
        if rel_id == "1" or rel_navn.casefold() == "eierstruktur":
            return (
                get_nested(relasjon, "Enhet", "Organisasjonsnummer"),
                get_nested(relasjon, "Enhet", "Navn"),
            )
    return None, None


def eierform(enhet: dict) -> str:
    privat = enhet.get("ErPrivatskole")
    offentlig = enhet.get("ErOffentligSkole")
    if privat is True and offentlig is not True:
        return "Privat"
    if offentlig is True and privat is not True:
        return "Offentlig"
    if privat is True and offentlig is True:
        return "Konflikt: både privat og offentlig"
    return "Uavklart"


def flat_ut(enhet: dict) -> dict:
    eier_orgnr, eier_navn = finn_eier(enhet)
    godkjenning = enhet.get("PrivatskoleGodkjenning") or {}

    return {
        "Organisasjonsnummer": enhet.get("Organisasjonsnummer"),
        "Navn": enhet.get("Navn"),
        "FulltNavn": enhet.get("FulltNavn"),
        "Karakteristikk": enhet.get("Karakteristikk"),
        "ErAktiv": enhet.get("ErAktiv"),
        "ErSkole": enhet.get("ErSkole"),
        "ErGrunnskole": enhet.get("ErGrunnskole"),
        "ErPrivatskole": enhet.get("ErPrivatskole"),
        "ErOffentligSkole": enhet.get("ErOffentligSkole"),
        "Eierform": eierform(enhet),
        "ErSpesialskole": enhet.get("ErSpesialskole"),
        "ErGrunnopplaering": enhet.get("ErGrunnopplaering"),
        "Fylkesnummer": get_nested(enhet, "Fylke", "Fylkesnummer"),
        "Fylke": get_nested(enhet, "Fylke", "Navn"),
        "Kommunenummer": get_nested(enhet, "Kommune", "Kommunenummer"),
        "Kommune": get_nested(enhet, "Kommune", "Navn"),
        "MaalformId": get_nested(enhet, "Maalform", "Id"),
        "Maalform": get_nested(enhet, "Maalform", "Navn"),
        "Elevtall": enhet.get("Elevtall"),
        "AntallAnsatte": enhet.get("AntallAnsatte"),
        "SkoletrinnGSFra": enhet.get("SkoletrinnGSFra"),
        "SkoletrinnGSTil": enhet.get("SkoletrinnGSTil"),
        "EierOrganisasjonsnummer": eier_orgnr,
        "EierNavn": eier_navn,
        "OrganisasjonsformId": get_nested(enhet, "Organisasjonsform", "Id"),
        "Organisasjonsform": get_nested(enhet, "Organisasjonsform", "Navn"),
        "Naeringskoder": list_to_text(enhet.get("Naeringskoder"), "Kode"),
        "NaeringskodeNavn": list_to_text(enhet.get("Naeringskoder"), "Navn"),
        "Skolekategorier": list_to_text(enhet.get("Skolekategorier"), "Navn"),
        "PrivatskoleGodkjenningslov": godkjenning.get("Godkjenningslov"),
        "PrivatskoleGodkjentStatus": godkjenning.get("GodkjentStatus"),
        "PrivatskoleGodkjentStatusTekst": godkjenning.get("GodkjentStatusTekst"),
        "PrivatskoleGodkjenningAktiv": godkjenning.get("ErAktiv"),
        "OppstartsEllerStiftelsesdato": enhet.get("OppstartsEllerStiftelsesdato"),
        "UtgaattDato": enhet.get("UtgaattDato"),
        "UtgaattypeId": get_nested(enhet, "Utgaattype", "Id"),
        "Utgaattype": get_nested(enhet, "Utgaattype", "Navn"),
        "DatoEndret": enhet.get("DatoEndret"),
        "Internettadresse": enhet.get("Internettadresse"),
        "Beliggenhetsadresse": get_nested(enhet, "Beliggenhetsadresse", "Adresse"),
        "Postnummer": get_nested(enhet, "Beliggenhetsadresse", "Postnummer"),
        "Poststed": get_nested(enhet, "Beliggenhetsadresse", "Poststed"),
        "Lengdegrad": get_nested(enhet, "Koordinat", "Lengdegrad"),
        "Breddegrad": get_nested(enhet, "Koordinat", "Breddegrad"),
    }


def main() -> None:
    session = make_session()
    oversikt = hent_alle_enheter(session)
    grunnskoler = [e for e in oversikt if er_relevant_grunnskole(e)]
    print(f"Relevante grunnskoler: {len(grunnskoler)}")

    detaljer = []
    with OUT_JSONL.open("w", encoding="utf-8") as raw_file:
        for nr, enkel in enumerate(grunnskoler, start=1):
            orgnr = enkel.get("Organisasjonsnummer")
            if not orgnr:
                print(f"Hopper over enhet uten organisasjonsnummer: {enkel.get('Navn')}")
                continue

            try:
                full = hent_detaljer(session, orgnr)
            except requests.HTTPError as exc:
                print(f"FEIL {orgnr} {enkel.get('Navn')}: {exc}")
                continue

            # Ekstra kontroll fordi fullmodellen er autoritativ for detaljfeltene.
            if not full.get("ErGrunnskole", False):
                continue
            if not TA_MED_NEDLAGTE and not full.get("ErAktiv", False):
                continue

            raw_file.write(json.dumps(full, ensure_ascii=False) + "\n")
            detaljer.append(flat_ut(full))

            if nr % 100 == 0 or nr == len(grunnskoler):
                print(f"Detaljer: {nr}/{len(grunnskoler)}")
            time.sleep(PAUSE_MELLOM_DETALJKALL)

    df = pd.DataFrame(detaljer)
    df = df.sort_values(["ErAktiv", "Fylke", "Kommune", "Navn"], ascending=[False, True, True, True])
    df.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="NSR grunnskoler", index=False)
        pd.DataFrame(
            {
                "Kontrollpunkt": [
                    "Antall enheter",
                    "Aktive",
                    "Nedlagte/inaktive",
                    "Private",
                    "Offentlige",
                    "Uavklarte/konflikt",
                    "Mangler målform",
                ],
                "Antall": [
                    len(df),
                    int((df["ErAktiv"] == True).sum()),
                    int((df["ErAktiv"] == False).sum()),
                    int((df["Eierform"] == "Privat").sum()),
                    int((df["Eierform"] == "Offentlig").sum()),
                    int((~df["Eierform"].isin(["Privat", "Offentlig"])).sum()),
                    int(df["Maalform"].isna().sum()),
                ],
            }
        ).to_excel(writer, sheet_name="Kontroll", index=False)

    print(f"Lagret: {OUT_CSV.resolve()}")
    print(f"Lagret: {OUT_XLSX.resolve()}")
    print(f"Lagret: {OUT_JSONL.resolve()}")


if __name__ == "__main__":
    main()
