"""Profil ponctuel du datagrid Distribution (et croisement avec Inscription) : valeurs des listes,
formats masqués, comptages. N'affiche aucune donnée personnelle (champs réduits à un masque,
chiffre -> 9, lettre -> A) et n'écrit rien nulle part."""

import re
from collections import Counter, defaultdict

from common.dates import parse_dt
from common.mwater_client import download_datagrid_with_retry, mwater_login

import os

DATAGRID_DISTRIBUTION = "ba07c16edc57417093d9735c93a97e52"
DATAGRID_INSCRIPTION = "ac3ae2ce1f7540f586ae1ab5673419a8"

C_HOUSEHOLD = "Code ID du ménage > Unique ID"
C_MODELE = "Modèle/type de foyer amélioré reçu"
C_QR = ["QR Code 1", "QR Code 2", "QR Code 3"]
C_TYPE = "Type"
C_QTE = "Quantité de foyer améliorés du projet reçus"
C_RETOURNES = "Foyers non efficient retournés"
C_MONTANT = "Montant payé"
C_DATE = "Date de distribution"
C_FICHE = "Fiche de distribution"
C_CONVENTION = "Photo de convention d'utilisation carbone"


def masque(v):
    v = (v or "").strip()
    if not v:
        return "<vide>"
    return re.sub(r"[A-Za-zÀ-ÿ]", "A", re.sub(r"\d", "9", v))[:40]


def top(rows, col, n=8, fn=lambda v: (v or "").strip()):
    return Counter(fn(r.get(col)) for r in rows).most_common(n)


def main():
    client_id = mwater_login(os.environ["MWATER_USERNAME"].strip(), os.environ["MWATER_PASSWORD"].strip())
    rows = download_datagrid_with_retry(DATAGRID_DISTRIBUTION, client_id, attempts=8, wait_seconds=60)
    print(f"Distribution : {len(rows)} lignes")

    print("\n--- Valeurs des listes")
    for c in ("Status", C_MODELE, C_TYPE, C_QTE, C_RETOURNES, "Approval Level", "Number of Rejections"):
        print(f"{c!r}:", top(rows, c))

    print("\n--- Masques de format")
    for c in [C_HOUSEHOLD, *C_QR, C_MONTANT, C_DATE, "Distribution Datetime", "Drafted On", "Submitted On",
              C_FICHE, C_CONVENTION]:
        print(f"{c!r}:", top(rows, c, 6, masque))
    montants = Counter()
    for r in rows:
        m = (r.get(C_MONTANT) or "").strip()
        montants["vide" if not m else ("negatif" if m.startswith("-") else ("0" if m in ("0", "0.0") else "positif"))] += 1
    print("Montant (catégories) :", dict(montants))

    print("\n--- Type x nombre de QR renseignés")
    croise = Counter()
    for r in rows:
        nb = sum(1 for c in C_QR if (r.get(c) or "").strip())
        remplis = "".join(str(i + 1) for i, c in enumerate(C_QR) if (r.get(c) or "").strip()) or "aucun"
        croise[((r.get(C_TYPE) or "").strip(), nb, remplis)] += 1
    for k, n in sorted(croise.items(), key=lambda kv: -kv[1])[:12]:
        print(k, n)

    print("\n--- Colonnes obligatoires vides (par Status)")
    for c in (C_HOUSEHOLD, C_MODELE, C_MONTANT, C_DATE, C_FICHE, C_CONVENTION, C_RETOURNES, C_QTE):
        print(f"{c!r}:", dict(Counter((r.get('Status') or '') for r in rows if not (r.get(c) or '').strip())))

    print("\n--- Unicité / rejets")
    codes = Counter((r.get("Response Code") or "") for r in rows)
    print("Response Code en double :", sum(1 for n in codes.values() if n > 1))
    rej = [r for r in rows if (r.get("Rejection message") or "").strip()]
    print("Messages de rejet :", len(rej), "dont Rejected :", sum(1 for r in rej if r.get("Status") == "Rejected"),
          "dont 'doublon' :", sum(1 for r in rej if "doublon" in r["Rejection message"].lower()))

    print("\n--- Promptitude")
    drafted_after = date_after = 0
    for r in rows:
        d, s, dd = parse_dt(r.get("Drafted On")), parse_dt(r.get("Submitted On")), parse_dt(r.get(C_DATE))
        drafted_after += bool(d and s and s < d)
        date_after += bool(dd and s and dd.date() > s.date())
    print("Submitted < Drafted :", drafted_after, "/ Date de distribution > Submitted On :", date_after)

    print("\n--- Cohérence avec l'Inscription")
    insc = download_datagrid_with_retry(DATAGRID_INSCRIPTION, client_id, attempts=8, wait_seconds=60)
    final_ids = {(r.get("Household ID code > Unique ID") or "").strip() for r in insc if r.get("Status") == "Final"}
    drafted_by_hh = {}
    for r in insc:
        h = (r.get("Household ID code > Unique ID") or "").strip()
        d = parse_dt(r.get("Drafted On"))
        if h and r.get("Status") == "Final" and d:
            drafted_by_hh[h] = min(d, drafted_by_hh.get(h, d))
    sans_insc = avant = 0
    for r in rows:
        h = (r.get(C_HOUSEHOLD) or "").strip()
        if not h:
            continue
        if h not in final_ids:
            sans_insc += 1
        else:
            dd = parse_dt(r.get(C_DATE))
            if dd and h in drafted_by_hh and dd.date() < drafted_by_hh[h].date():
                avant += 1
    print("Distributions sans inscription Final :", sans_insc,
          "/ date de distribution < date de début de saisie de l'inscription :", avant)
    deja = Counter((r.get(C_HOUSEHOLD) or "").strip() for r in rows if (r.get(C_HOUSEHOLD) or "").strip())
    print("Ménages avec plusieurs distributions :", sum(1 for n in deja.values() if n > 1))


if __name__ == "__main__":
    main()
