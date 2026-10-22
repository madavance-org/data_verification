"""Comptage ponctuel des photos du datagrid Réhabilitation (formulaire combiné), par colonne photo
et par type de travaux. Lecture seule : n'affiche que des comptages (aucune adresse de photo,
aucune donnée personnelle) et ne télécharge aucune photo."""

import csv
import io
import os
import re
from collections import Counter, defaultdict

from common.mwater_client import download_datagrid_raw_with_retry, mwater_login

DATAGRID_REHAB = "0638d32971704164b9ac22d549d4818e"
IMAGE_RE = re.compile(r"https://api\.mwater\.co/v3/images/([0-9a-f]{32})")


def main():
    client_id = mwater_login(os.environ["MWATER_USERNAME"].strip(), os.environ["MWATER_PASSWORD"].strip())
    texte = download_datagrid_raw_with_retry(DATAGRID_REHAB, client_id, attempts=8, wait_seconds=60).decode("utf-8-sig")
    table = list(csv.reader(io.StringIO(texte)))
    header, rows = table[0], table[1:]
    print(f"Lignes : {len(rows)} / En-têtes : {len(header)}")
    positions = defaultdict(list)
    for i, h in enumerate(header):
        positions[h].append(i + 1)
    print("En-têtes en double :", {h: p for h, p in positions.items() if len(p) > 1})

    i_type = header.index("Type de travaux") if "Type de travaux" in header else None
    i_status = header.index("Status")
    print("Types de travaux x Status :")
    croise = Counter((r[i_type] if i_type is not None else "?", r[i_status]) for r in rows)
    for k, n in sorted(croise.items(), key=lambda kv: -kv[1]):
        print(f"   {n:>5} | {k}")

    # Colonnes contenant des adresses de photos
    colonnes_photo = []
    for i, h in enumerate(header):
        n = sum(len(IMAGE_RE.findall(r[i])) for r in rows if i < len(r))
        if n:
            colonnes_photo.append((i, h, n))
    print(f"\nColonnes photo : {len(colonnes_photo)}")
    total_toutes = total_pr = 0
    ids_toutes, ids_pr = set(), set()
    print("n° | colonne | photos (tous types) | photos (Première réhabilitation) | lignes avec photo (Première réhab.)")
    for i, h, n in colonnes_photo:
        n_pr = lignes_pr = 0
        for r in rows:
            if i >= len(r):
                continue
            ids = IMAGE_RE.findall(r[i])
            ids_toutes.update(ids)
            if i_type is not None and r[i_type].strip() == "Première réhabilitation":
                n_pr += len(ids)
                lignes_pr += 1 if ids else 0
                ids_pr.update(ids)
        total_toutes += n
        total_pr += n_pr
        print(f"{i + 1:>3} | {h[:55]} | {n} | {n_pr} | {lignes_pr}")
    print(f"\nTOTAL photos, tous types : {total_toutes} ({len(ids_toutes)} identifiants distincts)")
    print(f"TOTAL photos, Première réhabilitation : {total_pr} ({len(ids_pr)} identifiants distincts)")


if __name__ == "__main__":
    main()
