"""Écriture des anomalies dans le formulaire registre mWater
("MadAvance || Vérification de données || Formulaire || Actif").

Une même ligne du registre = une anomalie, identifiée par (Response Code, Dimension,
Sous-dimension, Description) lus directement dans les données déjà stockées. Une
exécution met à jour la ligne existante au lieu d'en créer une nouvelle (voir la
correction du 15/09/2026 dans verify_maintenance_preventive.py : l'ancienne clé, basée
sur le Water Point ID relu via `entities`, ne retrouvait jamais les lignes sans point
d'eau et a produit des milliers de doublons).

Chaque activité a son propre déploiement et son énumérateur autorisé, passés en
paramètres de insert_log_in_mwater().
"""

import json
import sys
import uuid
from datetime import datetime

import requests

from common.http_utils import raise_for_status_verbose
from common.mwater_client import MWATER_API_BASE, find_water_point_entity_id

FORM_LOG_VERIFICATION = "1febfeabe8054be6979350b81d1652fe"

Q_POINT_EAU = "663f2f64c91f48a69bdf45909498109c"
Q_DIMENSION = "40e7a60a2c6b41589313e72f816cdcb5"
Q_SOUS_DIMENSION = "1ba451d29b91404f8c9d16a3e8d08cd4"
Q_RESPONSE_CODE = "cd92c27826674193ac4b35bf5198a167"
Q_DESCRIPTION = "f9adfa4bb56c4e16afc4fd1628839512"
Q_CHAMP_CONCERNE = "1c562ea5696e4a88a0ddd1d5cd06aa5f"
Q_DETAILS = "01e0435d0c3543939c573e9292423eaf"
Q_STATUT = "de575ee83a0141119161cc0c96b496f2"
Q_PREMIERE_DETECTION = "5c20dbea032a4cfebcc1d4b7db62f455"
Q_DERNIERE_DETECTION = "ff5e64c0168649f2b72643a732e67349"
Q_DATE_RESOLUTION = "c93d563571ac4f17b2a2278c6ae841be"

# "Supprimé" n'a pas de choix correspondant dans le formulaire : ces lignes sont ignorées.
STATUT_CHOICE_IDS = {
    "Nouveau": "A2dsjGf",
    "Toujours ouvert": "VD3AM6A",
    "Résolu": "3znmNE5",
}


def parse_log_date(value):
    """Convertit une date du log ('JJ/MM/AAAA', datetime ou date) en 'AAAA-MM-JJ'.
    Retourne None si vide/invalide."""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if hasattr(value, "isoformat") and not isinstance(value, str):
        return value.isoformat()
    value = (value or "").strip()
    if not value:
        return None
    try:
        return datetime.strptime(value, "%d/%m/%Y").strftime("%Y-%m-%d")
    except ValueError:
        return None


def is_valid_water_point_id(water_point_id):
    return str(water_point_id or "").strip().isdigit()


def build_payload(row, client_id):
    """Données de la réponse registre pour une ligne du log, ou None si le statut n'est
    pas pris en charge. Le champ Site n'est renseigné que si le point d'eau existe
    encore dans mWater ; sinon la valeur d'origine est conservée dans Détails."""
    choice_id = STATUT_CHOICE_IDS.get(row.get("Statut"))
    if not choice_id:
        print(f"  [registre mWater] ignoré (statut '{row.get('Statut')}' non pris en "
              f"charge) : {row.get('Response Code')}", file=sys.stderr)
        return None

    water_point_id = str(row.get("Water Point ID") or "").strip()
    entity_id = None
    if is_valid_water_point_id(water_point_id):
        entity_id = find_water_point_entity_id(water_point_id, client_id)
    details = row.get("Détails") or ""
    if not entity_id and water_point_id:
        note = "[Water Point ID saisi (invalide) : " + water_point_id + "]"
        details = (details + "\n" + note) if details else note

    data = {
        Q_SOUS_DIMENSION: {"value": row.get("Sous-dimension", "")},
        Q_RESPONSE_CODE: {"value": row.get("Response Code", "")},
        Q_DETAILS: {"value": details},
        Q_STATUT: {"value": choice_id},
        Q_PREMIERE_DETECTION: {"value": parse_log_date(row.get("Première détection"))},
        Q_DERNIERE_DETECTION: {"value": parse_log_date(row.get("Dernière détection"))},
    }
    if entity_id:
        data[Q_POINT_EAU] = {"value": {"code": water_point_id}}
    if row.get("Dimension"):
        data[Q_DIMENSION] = {"value": row["Dimension"]}
    if row.get("Description"):
        data[Q_DESCRIPTION] = {"value": row["Description"]}
    if row.get("Champ concerné"):
        data[Q_CHAMP_CONCERNE] = {"value": row["Champ concerné"]}
    date_resolution = parse_log_date(row.get("Date de résolution"))
    if date_resolution:
        data[Q_DATE_RESOLUTION] = {"value": date_resolution}
    return {k: v for k, v in data.items() if v.get("value") not in (None, "")}


def existing_response_key(item):
    data = item.get("data", {})
    return tuple((data.get(q) or {}).get("value") for q in
                 (Q_RESPONSE_CODE, Q_DIMENSION, Q_SOUS_DIMENSION, Q_DESCRIPTION))


def row_key(row):
    return (row.get("Response Code"), row.get("Dimension"), row.get("Sous-dimension"),
            row.get("Description"))


def fetch_existing_responses(client_id, deployment_id):
    """{clé: _id} des réponses déjà présentes dans le registre pour ce déploiement
    (chaque activité ne touche que ses propres lignes)."""
    resp = requests.get(
        f"{MWATER_API_BASE}/responses",
        params={
            "client": client_id,
            "filter": json.dumps({"form": FORM_LOG_VERIFICATION, "deployment": deployment_id}),
            "limit": 10000,
        },
        timeout=120,
    )
    raise_for_status_verbose(resp)
    by_key = {}
    for item in resp.json():
        key = existing_response_key(item)
        if key[0]:
            by_key[key] = item["_id"]
    return by_key


def fetch_form_rev(client_id):
    resp = requests.get(
        f"{MWATER_API_BASE}/forms/{FORM_LOG_VERIFICATION}",
        params={"client": client_id},
        timeout=30,
    )
    raise_for_status_verbose(resp)
    return resp.json()["_rev"]


def insert_log_in_mwater(merged_rows, client_id, deployment_id, enumerator_id):
    """Crée ou met à jour une réponse registre par ligne du log. Sauté (avec message)
    si le déploiement ou l'énumérateur n'est pas renseigné. Non bloquant : l'appelant
    doit l'entourer d'un try/except, le fichier Excel restant la source fiable."""
    if not deployment_id or not enumerator_id:
        print("  [registre mWater] sauté : déploiement ou énumérateur pas encore "
              "renseigné pour cette activité. Le fichier Excel/SharePoint reste la "
              "source fiable.")
        return

    print("Écriture dans le registre mWater...")
    form_rev = fetch_form_rev(client_id)
    existing = fetch_existing_responses(client_id, deployment_id)

    created = updated = ignored = 0
    for row in merged_rows:
        payload = build_payload(row, client_id)
        if payload is None:
            ignored += 1
            continue
        response_id = existing.get(row_key(row))
        now = datetime.utcnow().isoformat() + "Z"
        document = {
            "_id": response_id or uuid.uuid4().hex,
            "form": FORM_LOG_VERIFICATION,
            "formRev": form_rev,
            "deployment": deployment_id,
            "user": enumerator_id,
            "status": "final",
            "approvals": [],
            "startedOn": now,
            "submittedOn": now,
            "data": payload,
        }
        resp = requests.post(
            f"{MWATER_API_BASE}/responses",
            params={"client": client_id},
            json=document,
            timeout=30,
        )
        raise_for_status_verbose(resp)
        if response_id:
            updated += 1
        else:
            created += 1

    print(f"  registre mWater : {created} créées, {updated} mises à jour, {ignored} ignorées")
