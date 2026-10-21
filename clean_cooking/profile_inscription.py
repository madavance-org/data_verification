"""Profil ponctuel du datagrid Inscription : valeurs des listes, formats (masqués) et comptages.
N'affiche aucune donnée personnelle : les champs sensibles sont réduits à un masque
(chiffre -> 9, lettre -> A) et seuls des comptages sont imprimés. N'écrit rien nulle part."""

import csv
import io
import os
import re
from collections import Counter, defaultdict

from common.mwater_client import download_datagrid_raw, mwater_login

DATAGRID_INSCRIPTION = "ac3ae2ce1f7540f586ae1ab5673419a8"

LISTES = [
    "Status", "Etat de paiment", "Mode d'utilisation", "Modèle/type de foyer amélioré commandé",
    "Quantité de foyer améliorés efficaces non fonctionnels à remettre",
    "Telephone number (if available) of the household (Not Applicable)",
    "Annexe G : Reçu de paiement (Asked)", "Type", "Approval Level", "Number of Rejections",
]
MASQUES = [
    "Household ID code > Unique ID", "Code village > Unique ID", "ID",
    "Numéro de carte d'identité nationale (CIN) du représentant du ménage",
    "Telephone number (if available) of the household",
    "Drafted On", "Submitted On", "Photo de la maison", "Annexe A : Fiche d'inscription des membres",
    "Annexe G : Reçu de paiement", "Adresse du ménage > FOKONTANY", "Name of household representative",
]
OBLIGATOIRES = [
    "Household ID code > Unique ID", "Code village > Unique ID", "Name of household representative",
    "Telephone number (if available) of the household", "Adresse du ménage > REGION",
    "Adresse du ménage > DISTRICT", "Adresse du ménage > COMMUNE", "Adresse du ménage > FOKONTANY",
    "Modèle/type de foyer amélioré commandé", "Mode d'utilisation", "Etat de paiment",
    "Photo de la maison", "Photo carte d'identité nationale", "Annexe A : Fiche d'inscription des membres",
    "Annexe B : Accord Carbone", "Annexe D : Formulaire de consentement pour la publication de données et de photos",
    "Annexe G : Reçu de paiement", "Numéro de carte d'identité nationale (CIN) du représentant du ménage",
]


def masque(v):
    v = (v or "").strip()
    if not v:
        return "<vide>"
    return re.sub(r"[A-Za-zÀ-ÿ]", "A", re.sub(r"\d", "9", v))[:40]


def main():
    client_id = mwater_login(os.environ["MWATER_USERNAME"].strip(), os.environ["MWATER_PASSWORD"].strip())
    text = download_datagrid_raw(DATAGRID_INSCRIPTION, client_id, timeout=900).decode("utf-8-sig")
    header = next(csv.reader(io.StringIO(text)))
    doublons = [h for h, n in Counter(header).items() if n > 1]
    rows = list(csv.DictReader(io.StringIO(text)))
    print(f"Lignes : {len(rows)} / En-têtes bruts : {len(header)} / En-têtes en double : {doublons}")

    print("\n--- Valeurs des listes (top 8)")
    for c in LISTES:
        print(f"{c!r}:", Counter((r.get(c) or "").strip() for r in rows).most_common(8))

    print("\n--- Masques de format (top 6)")
    for c in MASQUES:
        print(f"{c!r}:", Counter(masque(r.get(c)) for r in rows).most_common(6))

    print("\n--- Colonnes obligatoires vides (par Status)")
    for c in OBLIGATOIRES:
        vides = Counter((r.get("Status") or "") for r in rows if not (r.get(c) or "").strip())
        print(f"{c!r}: {dict(vides)}")

    print("\n--- Rejets")
    rej = [r for r in rows if (r.get("Rejection message") or "").strip()]
    print("Messages de rejet non vides :", len(rej), "dont Status Rejected :",
          sum(1 for r in rej if r.get("Status") == "Rejected"))
    print("Contenant 'doublon' :", sum(1 for r in rej if "doublon" in r["Rejection message"].lower()))

    print("\n--- Unicité Household ID (hors Draft)")
    ids = Counter((r.get("Household ID code > Unique ID") or "").strip() for r in rows
                  if r.get("Status") != "Draft" and (r.get("Household ID code > Unique ID") or "").strip())
    dup = {k: n for k, n in ids.items() if n > 1}
    print("ménages uniques :", len(ids), "/ ID en double :", len(dup), "/ lignes concernées :", sum(dup.values()))

    print("\n--- Cohérence village / adresse (REGION, DISTRICT, COMMUNE, FOKONTANY)")
    adresse_villages = defaultdict(Counter)
    for r in rows:
        v = (r.get("Code village > Unique ID") or "").strip()
        a = tuple((r.get(f"Adresse du ménage > {k}") or "").strip() for k in ("REGION", "DISTRICT", "COMMUNE", "FOKONTANY"))
        if v and all(a):
            adresse_villages[a][v] += 1
    multi = {a: c for a, c in adresse_villages.items() if len(c) > 1}
    print("adresses :", len(adresse_villages), "/ adresses avec >1 village :", len(multi),
          "/ lignes minoritaires :", sum(sum(c.values()) - max(c.values()) for c in multi.values()))
    villages_adresses = defaultdict(set)
    for a, c in adresse_villages.items():
        for v in c:
            villages_adresses[v].add(a)
    print("villages avec >1 adresse :", sum(1 for s in villages_adresses.values() if len(s) > 1), "/", len(villages_adresses))


if __name__ == "__main__":
    main()
