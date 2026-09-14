"""Authentification et lecture mWater, partagées par tous les scripts de ce repo.

Regroupe les fonctions qui étaient dupliquées à l'identique (ou presque) dans chaque
script d'activité avant l'introduction de ce module partagé (voir README, section
"Module partagé common/"). Un bug corrigé ici est corrigé pour tous les scripts qui
l'importent — c'est le but : avant, il fallait répercuter le correctif dans chaque copie
séparément.
"""

import csv
import io
import json

import requests

from common.http_utils import raise_for_status_verbose

MWATER_API_BASE = "https://api.mwater.co/v3"


def mwater_login(username, password):
    resp = requests.post(
        f"{MWATER_API_BASE}/clients",
        json={"username": username, "password": password},
        timeout=30,
    )
    raise_for_status_verbose(resp)
    client_id = resp.json().get("client")
    if not client_id:
        raise RuntimeError("Authentification mWater : champ 'client' absent de la réponse")
    return client_id


def download_datagrid(datagrid_id, client_id):
    """Télécharge un datagrid mWater et le retourne comme liste de dict (une par ligne)."""
    resp = requests.get(
        f"{MWATER_API_BASE}/datagrids/{datagrid_id}/download",
        params={"client": client_id, "share": "", "extraFilters": "[]", "format": "csv"},
        timeout=120,
    )
    raise_for_status_verbose(resp)
    # utf-8-sig pour gérer le BOM renvoyé par mWater
    text = resp.content.decode("utf-8-sig")
    reader = csv.DictReader(io.StringIO(text))
    return list(reader)


def download_datagrid_raw(datagrid_id, client_id):
    """Télécharge le datagrid tel quel (bytes), sans parsing ni transformation."""
    resp = requests.get(
        f"{MWATER_API_BASE}/datagrids/{datagrid_id}/download",
        params={"client": client_id, "share": "", "extraFilters": "[]", "format": "csv"},
        timeout=120,
    )
    raise_for_status_verbose(resp)
    return resp.content


def fetch_raw_responses_by_code(client_id, form_id, codes, chunk_size=200):
    """Récupère les réponses brutes mWater (endpoint /responses, avant jointure du datagrid)
    pour une liste de 'Response Code'. Sert à distinguer un champ jamais répondu d'un champ
    répondu dont l'entité liée (ex. Water Point) a été supprimée/fusionnée : dans ce cas, le
    datagrid exporte le champ vide alors que la réponse contient bien une référence à une
    entité, visible dans le tableau `entities` de la réponse brute.

    Renvoie {code: [réponses brutes]} — en liste, car un même Response Code peut correspondre à
    plusieurs réponses distinctes (voir Rindra_Madavance-DV7526, où 29 réponses partagent le
    même code) ; un dict à valeur unique perdrait silencieusement toutes les réponses sauf la
    dernière traitée.
    """
    results = {}
    codes = [c for c in codes if c]
    for i in range(0, len(codes), chunk_size):
        chunk = codes[i:i + chunk_size]
        resp = requests.get(
            f"{MWATER_API_BASE}/responses",
            params={
                "client": client_id,
                "filter": json.dumps({"form": form_id, "code": {"$in": chunk}}),
            },
            timeout=60,
        )
        raise_for_status_verbose(resp)
        for item in resp.json():
            results.setdefault(item.get("code"), []).append(item)
    return results


def extract_water_point_code(raw_response, entity_type="water_point"):
    """Extrait le code de l'entité référencée dans une réponse brute (tableau `entities`),
    indépendamment de l'ID de question. None si aucune référence de ce type."""
    if not raw_response:
        return None
    for entity in raw_response.get("entities", []):
        if entity.get("entityType") == entity_type:
            return entity.get("value")
    return None


def fetch_entities_by_code(client_id, entity_type, codes, chunk_size=100):
    """Renvoie {code: entité} pour les entités qui existent encore directement sous ce code."""
    results = {}
    codes = sorted(set(c for c in codes if c))
    for i in range(0, len(codes), chunk_size):
        chunk = codes[i:i + chunk_size]
        resp = requests.get(
            f"{MWATER_API_BASE}/entities/{entity_type}",
            params={"client": client_id, "filter": json.dumps({"code": {"$in": chunk}})},
            timeout=60,
        )
        raise_for_status_verbose(resp)
        for item in resp.json():
            results[item.get("code")] = item
    return results


def fetch_merge_targets(client_id, entity_type, codes, chunk_size=100):
    """Renvoie {ancien_code: entité_cible} pour les codes qui ont été fusionnés dans une autre
    entité (propriété `_merged_entities` de l'entité cible)."""
    results = {}
    codes = sorted(set(c for c in codes if c))
    for i in range(0, len(codes), chunk_size):
        chunk = codes[i:i + chunk_size]
        resp = requests.get(
            f"{MWATER_API_BASE}/entities/{entity_type}",
            params={"client": client_id, "filter": json.dumps({"_merged_entities": {"$in": chunk}})},
            timeout=60,
        )
        raise_for_status_verbose(resp)
        for item in resp.json():
            for old_code in item.get("_merged_entities", []):
                if old_code in chunk:
                    results[old_code] = item
    return results


def find_water_point_entity_id(water_point_id, client_id, cache=None):
    """Résout un Water Point ID (le 'code' mWater, ex. '1124056539') vers l'_id d'entité
    attendu par un champ Site de formulaire. `cache`, si fourni (dict), évite de refaire la
    requête pour un même point d'eau partagé par plusieurs anomalies dans une même exécution.

    Endpoint et paramètre confirmés par un test déjà fait côté MadAvance (repo mWater_backup,
    diagnostics/test_entities_endpoint.py et test_entities_org_filter.py) : le type d'entité
    fait partie du chemin (/v3/entities/water_point, pas /v3/entities?type=...), et le filtre
    se passe via le paramètre 'selector' (pas 'filter', utilisé ailleurs pour /v3/responses —
    deux conventions différentes chez mWater)."""
    if cache is not None and water_point_id in cache:
        return cache[water_point_id]
    resp = requests.get(
        f"{MWATER_API_BASE}/entities/water_point",
        params={"client": client_id, "selector": json.dumps({"code": water_point_id})},
        timeout=30,
    )
    raise_for_status_verbose(resp)
    items = resp.json()
    entity_id = items[0]["_id"] if items else None
    if cache is not None:
        cache[water_point_id] = entity_id
    return entity_id
