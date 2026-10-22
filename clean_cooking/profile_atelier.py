"""Profil ponctuel du datagrid Atelier de réparation et croisement avec la Distribution. Lecture
par position (détection des colonnes en double). N'affiche aucune donnée personnelle : colonnes à
faible cardinalité = valeurs (listes de choix), autres = masques (chiffre -> 9, lettre -> A).
N'écrit rien nulle part."""

import csv
import io
import os
import re
from collections import Counter, defaultdict

from common.dates import parse_dt
from common.mwater_client import (
    download_datagrid_raw_with_retry,
    download_datagrid_with_retry,
    mwater_login,
)

DATAGRID_ATELIER = "e129499f02c141a3a45e65bfb8b3f7d7"
DATAGRID_DISTRIBUTION = "ba07c16edc57417093d9735c93a97e52"


def masque(v):
    v = (v or "").strip()
    return re.sub(r"[A-Za-zÀ-ÿ]", "A", re.sub(r"\d", "9", v))[:34]


def main():
    client_id = mwater_login(os.environ["MWATER_USERNAME"].strip(), os.environ["MWATER_PASSWORD"].strip())
    texte = download_datagrid_raw_with_retry(DATAGRID_ATELIER, client_id, attempts=8, wait_seconds=60).decode("utf-8-sig")
    table = list(csv.reader(io.StringIO(texte)))
    header, rows = table[0], table[1:]
    print(f"Lignes : {len(rows)} / En-têtes : {len(header)} / Largeurs : {dict(Counter(len(r) for r in rows))}")
    positions = defaultdict(list)
    for i, h in enumerate(header):
        positions[h].append(i + 1)
    print("En-têtes en double :", {h: p for h, p in positions.items() if len(p) > 1})
    print("Status (col 3) :", Counter(r[2] for r in rows))

    print("\n--- Profil par colonne : n° | nom | remplies | distinctes | valeurs ou masques")
    for i, h in enumerate(header):
        valeurs = [r[i].strip() for r in rows if i < len(r) and r[i].strip()]
        distinctes = Counter(valeurs)
        if not valeurs:
            resume = "vide"
        elif len(distinctes) <= 10 and not h.startswith(("Drafted", "Submitted", "Last Modified")):
            resume = str(dict(distinctes.most_common(10)))[:240]
        else:
            resume = "masques " + str(Counter(masque(v) for v in valeurs).most_common(3))[:180]
        print(f"{i + 1:>3} | {h[:58]} | {len(valeurs)} | {len(distinctes)} | {resume}")

    # --- Croisement avec la Distribution (par nom d'en-tête, première occurrence)
    idx = {h: pos[0] - 1 for h, pos in positions.items()}

    def val(r, nom):
        i = idx.get(nom)
        return r[i].strip() if i is not None and i < len(r) else ""

    print("\n--- Distribution : QR -> types / dates")
    dist = download_datagrid_with_retry(DATAGRID_DISTRIBUTION, client_id, attempts=8, wait_seconds=60)
    qr_types, qr_dates, qr_menages = defaultdict(set), defaultdict(list), defaultdict(set)
    menages_final = set()
    for d in dist:
        if (d.get("Status") or "").strip() != "Final":
            continue
        menage = (d.get("Code ID du ménage > Unique ID") or "").strip()
        menages_final.add(menage)
        for c in ("QR Code 1", "QR Code 2", "QR Code 3"):
            q = (d.get(c) or "").strip().upper()
            if q:
                qr_types[q].add((d.get("Type") or "").strip())
                dt = parse_dt(d.get("Date de distribution"))
                if dt:
                    qr_dates[q].append(dt.date())
                qr_menages[q].add(menage)
    print("QR distribués :", len(qr_types), "/ QR avec plusieurs distributions :",
          sum(1 for v in qr_dates.values() if len(v) > 1))

    print("\n--- Atelier : croisements par stove")
    croise_type, croise_redistribue = Counter(), Counter()
    sans_dist = date_avant = menage_diff = menage_sans_dist = total = 0
    types_atelier = Counter()
    for r in rows:
        if r[2] != "Final":
            continue
        menage = val(r, "Household ID > Unique ID")
        if menage and menage not in menages_final:
            menage_sans_dist += 1
        for n, (qr_n, com_n, na_n, type_n, retour_n, redistribue_n) in {
            1: ("QR code - stove 1", "QR code - stove 1 (Comments)", "QR code - stove 1 (Not Applicable)",
                "Stoves' type - stove 1", "Return date - stove 1", "Redistributed stoves' type - stove 1"),
            2: ("QR code - stove 2", "QR code - stove 2 (Comments)", "QR code - stove 2 (Not Applicable)",
                "Stoves' type - stove 2", "Return date - stove 2", "Redistributed stoves' type - stove 2"),
        }.items():
            q = val(r, qr_n) or (val(r, com_n) if val(r, na_n).lower() == "true" else "")
            if not q:
                continue
            total += 1
            q = q.upper().rstrip()
            t = val(r, type_n)
            types_atelier[t] += 1
            if q not in qr_types:
                sans_dist += 1
                continue
            croise_type[(t, tuple(sorted(qr_types[q])))] += 1
            if val(r, redistribue_n):
                croise_redistribue[(t, val(r, redistribue_n))] += 1
            dr = parse_dt(val(r, retour_n))
            if dr and qr_dates[q] and dr.date() < min(qr_dates[q]):
                date_avant += 1
            if menage and menage not in qr_menages[q]:
                menage_diff += 1
    print("Stoves (Final) avec QR :", total, "/ QR absent de la Distribution :", sans_dist,
          "/ retour avant la 1re distribution :", date_avant, "/ QR distribué à un autre ménage :", menage_diff)
    print("Réparations (Final) dont le ménage n'a aucune distribution Final :", menage_sans_dist)
    print("Types Atelier :", dict(types_atelier))
    print("Type retourné x types distribués pour ce QR :")
    for k, n in croise_type.most_common(12):
        print("  ", n, k)
    print("Type retourné x type redistribué :")
    for k, n in croise_redistribue.most_common(12):
        print("  ", n, k)


if __name__ == "__main__":
    main()
