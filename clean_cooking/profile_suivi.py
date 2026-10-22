"""Profil ponctuel du datagrid Suivi post distribution (anomalies). Lit le CSV par position (le
datagrid contient des colonnes en double, dont plusieurs "Status"). N'affiche aucune donnée
personnelle : les colonnes à faible cardinalité montrent leurs valeurs (listes de choix), les
autres sont réduites à un masque (chiffre -> 9, lettre -> A). N'écrit rien nulle part."""

import csv
import io
import os
import re
from collections import Counter, defaultdict

from common.mwater_client import download_datagrid_raw_with_retry, mwater_login

DATAGRID_SUIVI = "e4a38892c65840ef8d1caf59b0fe6198"


def masque(v):
    v = (v or "").strip()
    return re.sub(r"[A-Za-zÀ-ÿ]", "A", re.sub(r"\d", "9", v))[:34]


def main():
    client_id = mwater_login(os.environ["MWATER_USERNAME"].strip(), os.environ["MWATER_PASSWORD"].strip())
    text = download_datagrid_raw_with_retry(DATAGRID_SUIVI, client_id, attempts=8, wait_seconds=60).decode("utf-8-sig")
    table = list(csv.reader(io.StringIO(text)))
    header, rows = table[0], table[1:]
    print(f"Lignes : {len(rows)} / En-têtes : {len(header)}")
    largeurs = Counter(len(r) for r in rows)
    print("Largeur des lignes :", dict(largeurs))

    positions = defaultdict(list)
    for i, h in enumerate(header):
        positions[h].append(i)
    print("\n--- En-têtes en double (nom : positions 1-indexées)")
    for h, pos in positions.items():
        if len(pos) > 1:
            print(f"{h!r}: {[p + 1 for p in pos]}")

    print("\n--- Colonne 3 (Status de réponse)")
    print(Counter(r[2] for r in rows))

    print("\n--- Profil par colonne : n° | nom | remplies | distinctes | valeurs ou masques")
    for i, h in enumerate(header):
        valeurs = [r[i].strip() for r in rows if i < len(r) and r[i].strip()]
        distinctes = Counter(valeurs)
        if not valeurs:
            resume = "vide"
        elif len(distinctes) <= 10 and not h.startswith(("Photo", "Drafted", "Submitted", "Last Modified")):
            resume = str(dict(distinctes.most_common(10)))[:260]
        else:
            resume = "masques " + str(Counter(masque(v) for v in valeurs).most_common(3))[:200]
        print(f"{i + 1:>3} | {h[:60]} | {len(valeurs)} | {len(distinctes)} | {resume}")


if __name__ == "__main__":
    main()
