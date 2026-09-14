#!/usr/bin/env python3
"""
Diagnostic : pour les réponses "Appel maintenance préventive" dont le point d'eau référencé ne
résout à aucune entité mWater active (colonne "Water Point ID > Unique ID" vide dans le
datagrid), détermine si le code répondu correspond en réalité à une entité qui a depuis été
fusionnée (déduplication, propriété `_merged_entities` côté mWater) dans une autre — plutôt que
supprimée ou jamais renseignée.

Contexte : découverte du 30/07/2026 sur la réponse Rindra_Madavance-DV7526 / point d'eau
742895319, fusionné dans 742895333 le 15/01/2026 (avant même la soumission de la réponse, qui
référençait donc un code déjà obsolète). Ce script généralise cette investigation à toutes les
réponses "Appel maintenance préventive" concernées, en la rendant rejouable dans le temps (les
fusions continuent de se produire) plutôt que figée à la liste trouvée ce jour-là.

Usage (depuis la racine du repo) :
    python -m premiere_rehabilitation.find_merged_water_points [--csv sortie.csv]

Variables d'environnement requises : MWATER_USERNAME, MWATER_PASSWORD.
"""

from __future__ import annotations

import argparse
import csv
import os
import sys

from common.mwater_client import (
    download_datagrid,
    extract_water_point_code,
    fetch_raw_responses_by_code,
    mwater_login,
)
from common.water_point_merges import resolve_water_point_codes

DATAGRID_APPEL = "5f443b8a8c144502a304c8c5c24d4f82"
FORM_APPEL = "c08b3fe26d0f42c084074701f29eb75e"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--csv", help="Chemin d'un fichier CSV où écrire le rapport (en plus de l'affichage console).")
    args = parser.parse_args()

    username = os.environ.get("MWATER_USERNAME", "").strip()
    password = os.environ.get("MWATER_PASSWORD", "").strip()
    if not username or not password:
        sys.exit("MWATER_USERNAME et MWATER_PASSWORD doivent être définis dans l'environnement.")

    print("Authentification mWater...")
    client_id = mwater_login(username, password)

    print("Téléchargement du datagrid Appel maintenance préventive...")
    appel_rows = download_datagrid(DATAGRID_APPEL, client_id)
    print(f"  {len(appel_rows)} réponses")

    candidates = [
        r.get("Response Code", "")
        for r in appel_rows
        if r.get("Response Code") and not (r.get("Water Point ID > Unique ID") or "").strip()
    ]
    candidates = sorted(set(c for c in candidates if c))
    print(f"  {len(candidates)} réponses avec Water Point ID vide dans le datagrid")

    if not candidates:
        print("Rien à investiguer.")
        return 0

    print("Récupération des réponses brutes pour retrouver le code réellement saisi...")
    raw_by_code = fetch_raw_responses_by_code(client_id, FORM_APPEL, candidates)

    rows = []
    referenced_codes = set()
    for response_code, raw_list in raw_by_code.items():
        for raw in raw_list:
            wp_code = extract_water_point_code(raw)
            rows.append({"response_code": response_code, "response_id": raw.get("_id"), "wp_code": wp_code})
            if wp_code:
                referenced_codes.add(wp_code)

    print(f"  {len(referenced_codes)} codes de point d'eau uniques référencés, à vérifier...")

    print("Vérification directe et recherche de fusions (_merged_entities) dans mWater...")
    direct, merged = resolve_water_point_codes(client_id, referenced_codes)

    report = []
    for row in rows:
        wp_code = row["wp_code"]
        if wp_code is None:
            resolution, new_code = "Non répondu (dontknow)", ""
        elif wp_code in direct:
            resolution, new_code = "Valide directement", wp_code
        elif wp_code in merged:
            resolution, new_code = "Fusionné", merged[wp_code].get("code", "")
        else:
            resolution, new_code = "Introuvable", ""
        report.append({
            "Response Code": row["response_code"],
            "Response ID": row["response_id"],
            "Code référencé": wp_code or "",
            "Résolution": resolution,
            "Nouveau code": new_code,
        })

    print("\n" + "-" * 100)
    for r in report:
        print(f"{r['Response Code']:30} | {r['Code référencé']:12} | {r['Résolution']:22} | {r['Nouveau code']}")
    print("-" * 100)

    n_merged = sum(1 for r in report if r["Résolution"] == "Fusionné")
    n_missing = sum(1 for r in report if r["Résolution"] == "Introuvable")
    n_valid = sum(1 for r in report if r["Résolution"] == "Valide directement")
    n_dontknow = sum(1 for r in report if r["Résolution"] == "Non répondu (dontknow)")
    print(f"\nRésumé : {n_merged} fusionnés, {n_valid} valides directement, "
          f"{n_missing} introuvables, {n_dontknow} non répondus. Total : {len(report)}.")

    if args.csv:
        with open(args.csv, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=list(report[0].keys()))
            writer.writeheader()
            writer.writerows(report)
        print(f"\nRapport écrit dans {args.csv}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
