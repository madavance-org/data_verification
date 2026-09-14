#!/usr/bin/env python3
"""
Télécharge tel quel le datagrid mWater "Clean Water || Water Point" (ID
5971727de1bd44e2bc2dff09948c2068), qui expose déjà les colonnes "Unique ID" (code actuel d'un
point d'eau) et "Previous mWater IDs" (anciens codes fusionnés dedans, vide si aucun) pour
l'ensemble des points d'eau — indépendamment de toute activité/formulaire.

Simplification du 30/07/2026 : ce datagrid, déjà configuré dans le portail mWater, rend inutile
l'export complet de l'organisation (~270 Mo, plusieurs minutes) utilisé dans une première
version de ce script — un simple GET suffit.

Usage (depuis la racine du repo) :
    python -m utils.list_merged_water_points --csv sortie.csv [--upload]

--upload dépose le CSV dans le dossier SharePoint MWATER_MERGES_FOLDER_LINK (variable
d'environnement), via Microsoft Graph (AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET).

Variables d'environnement requises : MWATER_USERNAME, MWATER_PASSWORD ; en plus si --upload :
AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, MWATER_MERGES_FOLDER_LINK.
"""

from __future__ import annotations

import argparse
import os
import sys

from common.mwater_client import download_datagrid_raw, mwater_login
from common.sharepoint import graph_token, resolve_share_link, upload_to_sharepoint

DATAGRID_WATER_POINT = "5971727de1bd44e2bc2dff09948c2068"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--csv", default="Clean Water - Water Point.csv", help="Chemin du fichier CSV de sortie.")
    parser.add_argument("--upload", action="store_true", help="Dépose le CSV sur SharePoint (MWATER_MERGES_FOLDER_LINK).")
    args = parser.parse_args()

    username = os.environ.get("MWATER_USERNAME", "").strip()
    password = os.environ.get("MWATER_PASSWORD", "").strip()
    if not username or not password:
        sys.exit("MWATER_USERNAME et MWATER_PASSWORD doivent être définis dans l'environnement.")

    print("Authentification mWater...")
    client_id = mwater_login(username, password)

    print("Téléchargement du datagrid...")
    content = download_datagrid_raw(DATAGRID_WATER_POINT, client_id)
    with open(args.csv, "wb") as f:
        f.write(content)
    print(f"  Écrit dans {args.csv} ({len(content)} octets)")

    if args.upload:
        folder_link = os.environ.get("MWATER_MERGES_FOLDER_LINK", "").strip()
        tenant_id = os.environ.get("AZURE_TENANT_ID", "").strip()
        azure_client_id = os.environ.get("AZURE_CLIENT_ID", "").strip()
        azure_client_secret = os.environ.get("AZURE_CLIENT_SECRET", "").strip()
        if not all([folder_link, tenant_id, azure_client_id, azure_client_secret]):
            sys.exit("--upload requiert MWATER_MERGES_FOLDER_LINK, AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET.")
        print("Authentification Microsoft Graph...")
        token = graph_token(tenant_id, azure_client_id, azure_client_secret)
        print("Résolution du dossier SharePoint...")
        drive_id, folder_item_id = resolve_share_link(token, folder_link)
        print("Dépôt du fichier...")
        item = upload_to_sharepoint(
            token, drive_id, folder_item_id, args.csv, os.path.basename(args.csv),
            content_type="text/csv",
        )
        print(f"Déposé : {item.get('webUrl', '')}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
