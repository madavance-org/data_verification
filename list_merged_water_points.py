#!/usr/bin/env python3
"""
Liste tous les points d'eau MadAvance qui ont été fusionnés (déduplication) dans un autre,
indépendamment de toute activité/formulaire — contrairement à find_merged_water_points.py, qui
part des réponses "Appel maintenance préventive" pour trouver les fusions qui les concernent,
ce script part directement des entités "Water point" elles-mêmes.

Principe : déclenche un export complet de l'organisation mWater (même mécanisme que
mWater_backup/backup_mwater.py), en extrait "Entity-Water point.csv", et ne garde que les deux
colonnes utiles : "Unique ID" (code actuel du point d'eau) et "Previous mWater IDs" (anciens
codes fusionnés dedans, séparés par "; " — vide si le point d'eau n'a jamais absorbé de
doublon). Seules les lignes ayant au moins un ancien code sont conservées.

Usage :
    python list_merged_water_points.py --csv sortie.csv [--upload]

--upload dépose le CSV dans le dossier SharePoint MWATER_MERGES_FOLDER_LINK (variable
d'environnement), via Microsoft Graph (AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET).

Variables d'environnement requises : MWATER_USERNAME, MWATER_PASSWORD, MWATER_ORG_ID (optionnel,
par défaut l'organisation MadAvance) ; en plus si --upload : AZURE_TENANT_ID, AZURE_CLIENT_ID,
AZURE_CLIENT_SECRET, MWATER_MERGES_FOLDER_LINK.
"""

from __future__ import annotations

import argparse
import base64
import csv
import io
import os
import sys
import time
import zipfile

import requests

MWATER_API_BASE = "https://api.mwater.co/v3"
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
MWATER_ORG_ID_DEFAULT = "aaaf0a14e4ce44eaa7a2bcfd1c74aa56"
WATER_POINT_CSV_NAME = "Entity-Water point.csv"

POLL_MAX_ATTEMPTS = int(os.environ.get("POLL_MAX_ATTEMPTS", "60"))
POLL_DELAY_SECONDS = int(os.environ.get("POLL_DELAY_SECONDS", "30"))


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


def run_export_and_download(client_id: str, org_id: str) -> bytes:
    """Même principe que mWater_backup/backup_mwater.py : crée un job d'export, attend sa fin,
    télécharge le zip résultant. Repris ici en version simplifiée (pas de retry sur échec
    explicite du job, ce script n'étant pas destiné à tourner sans supervision)."""
    print("Création du job d'export mWater...")
    resp = requests.post(
        f"{MWATER_API_BASE}/organizations/{org_id}/export-job", params={"client": client_id}, timeout=30,
    )
    raise_for_status_verbose(resp)
    job_id = resp.json()["jobId"]
    print(f"  Job créé : {job_id}")

    print(f"Attente de la fin du job (jusqu'à {POLL_MAX_ATTEMPTS * POLL_DELAY_SECONDS // 60} min)...")
    for attempt in range(1, POLL_MAX_ATTEMPTS + 1):
        resp = requests.get(f"{MWATER_API_BASE}/jobs/{job_id}", timeout=30)
        raise_for_status_verbose(resp)
        status = resp.json().get("status")
        print(f"  Poll {attempt}/{POLL_MAX_ATTEMPTS} — status={status}")
        if status == "success":
            break
        if status in ("error", "failed", "failure"):
            raise RuntimeError(f"Le job d'export mWater a échoué : {resp.json()}")
        time.sleep(POLL_DELAY_SECONDS)
    else:
        raise RuntimeError("Timeout : le job d'export mWater n'est pas terminé à temps.")

    print("Téléchargement de l'export...")
    url = f"{MWATER_API_BASE}/organizations/{org_id}/export-jobs/{job_id}/file"
    with requests.get(url, params={"client": client_id}, stream=True, timeout=300) as resp:
        raise_for_status_verbose(resp)
        return resp.content


def extract_merged_water_points(zip_bytes: bytes) -> list[dict]:
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        with zf.open(WATER_POINT_CSV_NAME) as f:
            text = f.read().decode("utf-8-sig")
    reader = csv.DictReader(io.StringIO(text))
    rows = []
    for row in reader:
        previous_ids = (row.get("Previous mWater IDs") or "").strip()
        if not previous_ids:
            continue
        rows.append({
            "Unique ID": (row.get("Unique ID") or "").strip(),
            "Previous mWater IDs": previous_ids,
        })
    return rows


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
    parser.add_argument("--csv", default="merged_water_points.csv", help="Chemin du fichier CSV de sortie.")
    parser.add_argument("--upload", action="store_true", help="Dépose le CSV sur SharePoint (MWATER_MERGES_FOLDER_LINK).")
    args = parser.parse_args()

    username = os.environ.get("MWATER_USERNAME", "").strip()
    password = os.environ.get("MWATER_PASSWORD", "").strip()
    org_id = os.environ.get("MWATER_ORG_ID", MWATER_ORG_ID_DEFAULT).strip()
    if not username or not password:
        sys.exit("MWATER_USERNAME et MWATER_PASSWORD doivent être définis dans l'environnement.")

    print("Authentification mWater...")
    client_id = mwater_login(username, password)

    zip_bytes = run_export_and_download(client_id, org_id)
    print(f"  Export téléchargé : {len(zip_bytes) / 1_000_000:.1f} Mo")

    print(f"Extraction de {WATER_POINT_CSV_NAME}...")
    rows = extract_merged_water_points(zip_bytes)
    print(f"  {len(rows)} points d'eau avec au moins une fusion")

    with open(args.csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["Unique ID", "Previous mWater IDs"])
        writer.writeheader()
        writer.writerows(rows)
    print(f"Écrit dans {args.csv}")

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
