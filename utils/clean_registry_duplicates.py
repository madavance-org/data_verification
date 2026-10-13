#!/usr/bin/env python3
"""
Nettoyage des doublons du registre mWater
("MadAvance || Vérification de données || Formulaire || Actif").

Avant la correction du 15/09/2026, les scripts d'Appel et de Réparation recréaient une
ligne du registre à chaque exécution pour les anomalies sans point d'eau. Ce script garde,
pour chaque anomalie (déploiement + Response Code + Dimension + Sous-dimension +
Description), la ligne la plus récente et supprime les autres.

Par défaut : SIMULATION (aucune écriture, affiche seulement les chiffres). Ajouter
--apply pour supprimer réellement.

La suppression d'une réponse via l'API n'a jamais été testée dans ce repo. En mode
--apply, le script essaie d'abord UNE seule suppression : si mWater la refuse, il bascule
sur le rejet (statut "rejected", via le même POST complet que l'écriture du registre) ; si
le rejet est aussi refusé, il s'arrête sans rien faire d'autre. Il s'arrête aussi après 3
échecs consécutifs.

Usage (depuis la racine du repo) :
    python -m utils.clean_registry_duplicates [--apply]

Variables d'environnement requises : MWATER_USERNAME, MWATER_PASSWORD.
"""

import argparse
import json
import os
import sys
from collections import defaultdict

import requests

from common.mwater_client import MWATER_API_BASE, mwater_login

FORM_LOG_VERIFICATION = "1febfeabe8054be6979350b81d1652fe"
# Questions du registre qui identifient une anomalie : Response Code, Dimension,
# Sous-dimension, Description.
KEY_QUESTIONS = (
    "cd92c27826674193ac4b35bf5198a167",
    "40e7a60a2c6b41589313e72f816cdcb5",
    "1ba451d29b91404f8c9d16a3e8d08cd4",
    "f9adfa4bb56c4e16afc4fd1628839512",
)


Q_POINT_EAU = "663f2f64c91f48a69bdf45909498109c"


def existing_response_key(item):
    """Clé de l'anomalie, point d'eau inclus : deux points d'eau différents peuvent avoir
    la même anomalie (même Response Code, même description) sans être des doublons."""
    data = item.get("data", {})
    site = (data.get(Q_POINT_EAU) or {}).get("value")
    wp_code = site.get("code") if isinstance(site, dict) else (site or "")
    return tuple((data.get(q) or {}).get("value") for q in KEY_QUESTIONS) + (wp_code or "",)


MAX_CONSECUTIVE_FAILURES = 3


def fetch_registry_responses(client_id):
    resp = requests.get(
        f"{MWATER_API_BASE}/responses",
        params={"client": client_id, "filter": json.dumps({"form": FORM_LOG_VERIFICATION}),
                "limit": 10000},
        timeout=300,
    )
    resp.raise_for_status()
    items = resp.json()
    if len(items) >= 10000:
        sys.exit("Plus de 10000 réponses : la liste est probablement tronquée, arrêt par sécurité.")
    return items


def sort_key(item):
    return (item.get("submittedOn") or "", (item.get("modified") or {}).get("on") or "",
            item["_id"])


def find_duplicates(items):
    """Renvoie (à_supprimer, nb_groupes) : toutes les lignes sauf la plus récente de chaque
    groupe (déploiement + clé de l'anomalie). Les lignes déjà rejetées sont ignorées."""
    groups = defaultdict(list)
    for item in items:
        if (item.get("status") or "").lower() == "rejected":
            continue
        key = existing_response_key(item)
        if not key[0]:
            continue
        groups[(item.get("deployment"),) + key].append(item)

    to_remove, nb_groups = [], 0
    for group in groups.values():
        if len(group) > 1:
            nb_groups += 1
            group.sort(key=sort_key)
            to_remove.extend(group[:-1])
    return to_remove, nb_groups


def delete_response(client_id, response_id):
    resp = requests.delete(f"{MWATER_API_BASE}/responses/{response_id}",
                           params={"client": client_id}, timeout=60)
    return resp.ok, f"HTTP {resp.status_code} {resp.text[:300]}"


def reject_response(client_id, item):
    document = {
        "_id": item["_id"],
        "form": item["form"],
        "formRev": item["formRev"],
        "deployment": item["deployment"],
        "user": item["user"],
        "status": "rejected",
        "approvals": item.get("approvals", []),
        "startedOn": item["startedOn"],
        "submittedOn": item["submittedOn"],
        "data": item["data"],
    }
    resp = requests.post(f"{MWATER_API_BASE}/responses", params={"client": client_id},
                         json=document, timeout=60)
    return resp.ok, f"HTTP {resp.status_code} {resp.text[:300]}"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--apply", action="store_true",
                        help="Supprime réellement (sinon simulation).")
    args = parser.parse_args()

    username = os.environ.get("MWATER_USERNAME", "").strip()
    password = os.environ.get("MWATER_PASSWORD", "").strip()
    if not username or not password:
        sys.exit("MWATER_USERNAME et MWATER_PASSWORD doivent être définis.")

    client_id = mwater_login(username, password)
    items = fetch_registry_responses(client_id)
    to_remove, nb_groups = find_duplicates(items)

    by_deployment = defaultdict(int)
    for item in to_remove:
        by_deployment[item.get("deployment")] += 1
    print(f"{len(items)} réponses dans le registre ; {nb_groups} anomalies en double ; "
          f"{len(to_remove)} lignes en trop.")
    for deployment, n in sorted(by_deployment.items(), key=lambda kv: -kv[1]):
        print(f"  déploiement {deployment} : {n} lignes en trop")

    if not args.apply:
        print("SIMULATION : rien n'a été modifié. Relancer avec --apply pour nettoyer.")
        return 0
    if not to_remove:
        print("Rien à nettoyer.")
        return 0

    mode = "delete"
    first = to_remove[0]
    ok, detail = delete_response(client_id, first["_id"])
    if ok:
        print("Suppression acceptée par mWater (essai sur 1 ligne).")
        done = 1
        to_process = to_remove[1:]
    else:
        print(f"Suppression refusée ({detail}). Essai du rejet sur la même ligne...")
        ok, detail = reject_response(client_id, first)
        if not ok:
            print(f"Rejet refusé aussi ({detail}). Arrêt, rien d'autre n'a été modifié.")
            return 1
        print("Rejet accepté par mWater (essai sur 1 ligne).")
        mode, done = "reject", 1
        to_process = to_remove[1:]

    failures = 0
    for item in to_process:
        ok, detail = (delete_response(client_id, item["_id"]) if mode == "delete"
                      else reject_response(client_id, item))
        if ok:
            done += 1
            failures = 0
        else:
            failures += 1
            print(f"  échec sur {item['_id']} : {detail}", file=sys.stderr)
            if failures >= MAX_CONSECUTIVE_FAILURES:
                print(f"{failures} échecs consécutifs, arrêt.", file=sys.stderr)
                break

    print(f"Terminé ({mode}) : {done}/{len(to_remove)} lignes traitées.")
    remaining, _ = find_duplicates(fetch_registry_responses(client_id))
    print(f"Doublons restants après nettoyage : {len(remaining)}")
    return 0 if done == len(to_remove) else 1


if __name__ == "__main__":
    sys.exit(main())
