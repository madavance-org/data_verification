#!/usr/bin/env python3
"""
Télécharge tel quel le datagrid mWater "Clean Water || Water Point" (ID
5971727de1bd44e2bc2dff09948c2068), qui expose déjà les colonnes "Unique ID" (code actuel d'un
point d'eau) et "Previous mWater IDs" (anciens codes fusionnés dedans, vide si aucun) pour
l'ensemble des points d'eau — indépendamment de toute activité/formulaire.

Simplification du 30/07/2026 : ce datagrid, déjà configuré dans le portail mWater, rend inutile
l'export complet de l'organisation (~270 Mo, plusieurs minutes) utilisé dans une première
version de ce script — un simple GET suffit.

Usage :
    python list_merged_water_points.py --csv sortie.csv [--upload]

--upload dépose le CSV dans le dossier SharePoint MWATER_MERGES_FOLDER_LINK (variable
d'environnement), via Microsoft Graph (AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET).

Variables d'environnement requises : MWATER_USERNAME, MWATER_PASSWORD ; en plus si --upload :
AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, MWATER_MERGES_FOLDER_LINK.
"""

from __future__ import annotations

import argparse
import base64
import os
import sys

import requests

MWATER_API_BASE = "https://api.mwater.co/v3"
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
DATAGRID_WATER_POINT = "5971727de1bd44e2bc2dff09948c2068"


def raise_for_status_verbose(response: requests.Response) -> None:
    if not response.ok:
        print(f"HTTP {response.status_code} sur {response.url}", file=sys.stderr)
        print(response.text[:2000], file=sys.stderr)
        response.raise_for_status()


def mwater_login(username: str, password: str) -> str:
    resp = requests.post(
        f"{MWATER_API_BASE}/clients", json={"username": username, "password": password}, timeout=30,
    )
    raise_for_status_verbose(resp)
    client_id = resp.json().get("client")
    if not client_id:
        raise RuntimeError("Authentification mWater : champ 'client' absent de la réponse")
    return client_id


def download_datagrid_raw(datagrid_id: str, client_id: str) -> bytes:
    """Télécharge le datagrid tel quel, sans parsing ni transformation."""
    resp = requests.get(
        f"{MWATER_API_BASE}/datagrids/{datagrid_id}/download",
        params={"client": client_id, "share": "", "extraFilters": "[]", "format": "csv"},
        timeout=120,
    )
    raise_for_status_verbose(resp)
    return resp.content


def graph_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    resp = requests.post(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    raise_for_status_verbose(resp)
    return resp.json()["access_token"]


def resolve_share_link(token: str, share_url: str) -> tuple[str, str]:
    b64 = base64.b64encode(share_url.encode("utf-8")).decode("utf-8")
    encoded = "u!" + b64.replace("+", "-").replace("/", "_").rstrip("=")
    resp = requests.get(
        f"{GRAPH_API_BASE}/shares/{encoded}/driveItem",
        headers={"Authorization": f"Bearer {token}"},
        timeout=30,
    )
    raise_for_status_verbose(resp)
    data = resp.json()
    return data["parentReference"]["driveId"], data["id"]


def upload_csv(token: str, drive_id: str, folder_item_id: str, file_path: str, file_name: str) -> str:
    with open(file_path, "rb") as f:
        content = f.read()
    resp = requests.put(
        f"{GRAPH_API_BASE}/drives/{drive_id}/items/{folder_item_id}:/{file_name}:/content",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "text/csv"},
        data=content,
        timeout=120,
    )
    raise_for_status_verbose(resp)
    return resp.json().get("webUrl", "")


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
        web_url = upload_csv(token, drive_id, folder_item_id, args.csv, os.path.basename(args.csv))
        print(f"Déposé : {web_url}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
