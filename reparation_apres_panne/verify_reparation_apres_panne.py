#!/usr/bin/env python3
"""
Vérification de données — Réparation après panne
====================================================

Applique les six dimensions du "Manuel de vérification de données" MadAvance
(Complétude, Promptitude, Validité, Unicité, Cohérence, Fiabilité) à
l'activité "Réparation après panne", sur deux sources mWater :

  - Le formulaire actif "Clean Water || Réparation après panne || Survey ||
    Actif" (datagrid DATAGRID_REPARATION).
  - L'ancien formulaire combiné (inactif) "Première réhabilitation /
    Entretien préventif / Réparation après panne / Analyse de l'eau"
    (datagrid DATAGRID_REHAB), filtré sur le champ "Type de travaux" pour
    n'en retenir que les réponses de réparation/entretien préventif — ce
    même datagrid sert aussi de source de vérité pour la dimension
    Fiabilité (réhabilitations réussies), sur son propre sous-ensemble
    "Première réhabilitation".

Contrairement à verify_maintenance_preventive.py (Appel maintenance
préventive) :
  - Pas de dimension Promptitude sur le délai de réparation (pas de délai
    de référence défini pour l'instant) — seule la chronologie Drafted
    On / Submitted On est vérifiée.
  - La règle de Cohérence déjà couverte côté Appel (signal code présent
    dans Appel avec pompe = 'No', date du signal code <= date de
    complétion) n'est PAS réimplémentée ici, pour éviter le doublon de
    calcul et de log — seule la cohérence interne à ce formulaire (champs
    conditionnels selon 'Is the repair possible ?') est vérifiée.
  - Fichier de log dédié (data_verification_reparation_log.xlsx),
    distinct de celui d'Appel maintenance préventive.

Le script télécharge les données via l'API mWater, applique les règles,
fusionne avec le log existant sur SharePoint, réenregistre ce même
fichier, et envoie un email de confirmation.
"""

import os
import sys
from collections import Counter, defaultdict
from datetime import datetime

from common.dates import parse_dt
from common.mwater_client import (
    download_datagrid,
    extract_water_point_code,
    fetch_raw_responses_by_code,
    mwater_login,
)
from common.sharepoint import download_existing_log, graph_token, resolve_share_link, send_html_email, upload_to_sharepoint
from common.signal_code import SIGNAL_CODE_RE, propose_correction

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# IDs des datagrids/formulaires mWater (préconfigurés dans le portail, pas des secrets)
DATAGRID_REPARATION = "d9f1c36a2d6340429658b6628fe81b88"
DATAGRID_REHAB = "0638d32971704164b9ac22d549d4818e"

FORM_REPARATION_ACTIF = "958b4763788348d699e7d8c5821f92ee"

# Choix "Type de travaux" de l'ancien formulaire combiné (inactif) à inclure dans le
# périmètre de ce script — voir handoff : id 7abnNjk = "Breakdown repair" / "Réparation
# après panne", id 4K8SB4u = "Regular preventive maintenance" / "Entretien préventif
# régulier". Valeurs FR ci-dessous PAS ENCORE VALIDÉES contre un export réel du datagrid
# (par analogie avec "Première réhabilitation", déjà utilisé pour la Fiabilité côté Appel)
# — à vérifier avant la première exécution en production.
LEGACY_TYPES_TRAVAUX_REPARATION = {"Réparation après panne", "Entretien préventif régulier"}

# Fichier de log dédié à cette activité (distinct de celui d'Appel maintenance préventive
# — décision du 14/09/2026 : pas de fichier partagé entre activités).
LOG_FILE_NAME = "data_verification_reparation_log.xlsx"

LOG_HEADERS = [
    "Dimension", "Sous-dimension", "Response Code", "Water Point ID",
    "Description", "Champ concerné", "Détails", "Statut", "Première détection",
    "Dernière détection", "Date de résolution",
]

# ---------------------------------------------------------------------------
# mWater : insertion du log de vérification dans le formulaire dédié (même formulaire
# que verify_maintenance_preventive.py, déploiement différent — voir sa docstring pour
# le détail du mécanisme et des champs obligatoires).
# ---------------------------------------------------------------------------

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

STATUT_CHOICE_IDS = {
    "Nouveau": "A2dsjGf",
    "Toujours ouvert": "VD3AM6A",
    "Résolu": "3znmNE5",
}

# Déploiement Réparation après panne du formulaire de log (donné par Lanja le 14/09/2026).
DEPLOYMENT_LOG_VERIFICATION = "8b541444ac4e465bb93d0aad0d4dff68"

# Même énumérateur que pour Appel maintenance préventive (voir
# ENUMERATEUR_LOG_VERIFICATION dans verify_maintenance_preventive.py) : Lanja est aussi
# la personne autorisée à soumettre sur ce déploiement.
ENUMERATEUR_LOG_VERIFICATION = "0336889caebc433d8ab0bba5dc919bed"


def parse_log_date(value):
    """Convertit une date du log en 'AAAA-MM-JJ' pour l'API mWater. Retourne None si
    vide/invalide (voir verify_maintenance_preventive.py pour le détail des deux formes
    d'entrée possibles)."""
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


def build_mwater_log_response(row):
    """Construit le payload de réponse mWater pour une ligne du log (voir
    verify_maintenance_preventive.py pour le détail des champs obligatoires). Pas de
    résolution d'entité Site ici : ce script n'a pas encore été testé en écriture mWater,
    on ne réintroduit pas cette complexité tant que DEPLOYMENT_LOG_VERIFICATION est
    inconnu."""
    statut = row.get("Statut")
    choice_id = STATUT_CHOICE_IDS.get(statut)
    if not choice_id:
        print(f"  [mWater log] ignoré (statut '{statut}' non pris en charge par le "
              f"formulaire) : {row.get('Response Code')}", file=sys.stderr)
        return None

    data = {
        Q_SOUS_DIMENSION: {"value": row.get("Sous-dimension", "")},
        Q_RESPONSE_CODE: {"value": row.get("Response Code", "")},
        Q_DETAILS: {"value": row.get("Détails", "")},
        Q_STATUT: {"value": choice_id},
        Q_PREMIERE_DETECTION: {"value": parse_log_date(row.get("Première détection"))},
        Q_DERNIERE_DETECTION: {"value": parse_log_date(row.get("Dernière détection"))},
    }
    water_point_id = str(row.get("Water Point ID") or "").strip()
    if water_point_id.isdigit():
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


def fetch_form_rev(form_id, client_id):
    import requests
    from common.http_utils import raise_for_status_verbose
    from common.mwater_client import MWATER_API_BASE
    resp = requests.get(
        f"{MWATER_API_BASE}/forms/{form_id}",
        params={"client": client_id},
        timeout=30,
    )
    raise_for_status_verbose(resp)
    return resp.json()["_rev"]


def fetch_existing_log_responses(client_id):
    """Récupère les réponses déjà présentes dans le formulaire de log mWater, indexées
    par Response Code (voir verify_maintenance_preventive.py pour la limite connue sur
    les collisions de clé)."""
    import json
    import requests
    from common.http_utils import raise_for_status_verbose
    from common.mwater_client import MWATER_API_BASE
    resp = requests.get(
        f"{MWATER_API_BASE}/responses",
        params={"client": client_id, "filter": json.dumps({"form": FORM_LOG_VERIFICATION})},
        timeout=60,
    )
    raise_for_status_verbose(resp)
    by_rc = {}
    for item in resp.json():
        rc = (item.get("data", {}).get(Q_RESPONSE_CODE) or {}).get("value")
        if rc:
            by_rc[rc] = item["_id"]
    return by_rc


def inserer_log_dans_mwater(merged_rows, client_id):
    """Pousse le log de vérification dans le formulaire mWater dédié (voir
    verify_maintenance_preventive.py pour le détail du mécanisme upsert). Non bloquant
    par construction de l'appelant (voir main())."""
    if not DEPLOYMENT_LOG_VERIFICATION or not ENUMERATEUR_LOG_VERIFICATION:
        print("  [mWater log] sauté : DEPLOYMENT_LOG_VERIFICATION / "
              "ENUMERATEUR_LOG_VERIFICATION pas encore renseignés pour cette activité "
              "(voir commentaire en tête de script). Le fichier Excel/SharePoint reste "
              "la source fiable.")
        return

    import uuid
    import requests
    from common.http_utils import raise_for_status_verbose
    from common.mwater_client import MWATER_API_BASE

    print("Insertion du log dans mWater (formulaire de log dédié)...")
    form_rev = fetch_form_rev(FORM_LOG_VERIFICATION, client_id)
    existing_by_rc = fetch_existing_log_responses(client_id)

    created, updated, ignored = 0, 0, 0
    for row in merged_rows:
        payload = build_mwater_log_response(row)
        if payload is None:
            ignored += 1
            continue
        response_id = existing_by_rc.get(row.get("Response Code"))
        now = datetime.utcnow().isoformat() + "Z"
        document = {
            "_id": response_id or uuid.uuid4().hex,
            "form": FORM_LOG_VERIFICATION,
            "formRev": form_rev,
            "deployment": DEPLOYMENT_LOG_VERIFICATION,
            "user": ENUMERATEUR_LOG_VERIFICATION,
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

    print(f"  mWater : {created} créées, {updated} mises à jour, {ignored} ignorées")


# ---------------------------------------------------------------------------
# Résolution des Water Point ID vides (site fusionné/supprimé mais réponse valide)
# ---------------------------------------------------------------------------

def enrich_missing_water_point_ids(reparation_rows, client_id):
    """Pour les lignes 'Final' dont 'Water Point ID > Unique ID' ressort vide dans le
    datagrid, retrouve le code brut réellement répondu (via l'API brute) et le réinjecte
    dans la ligne — même mécanisme que verify_maintenance_preventive.py (voir sa
    docstring pour le détail). Ne s'applique qu'aux réponses du formulaire actif : les
    lignes de l'ancien formulaire combiné n'ont pas de Response Code résolvable sur
    FORM_REPARATION_ACTIF."""
    candidates = [
        r for r in reparation_rows
        if r.get("Status") == "Final"
        and not (r.get("Water Point ID > Unique ID") or "").strip()
        and r.get("_from_legacy_form") is not True
    ]
    if not candidates:
        return 0

    codes = sorted({r.get("Response Code", "") for r in candidates if r.get("Response Code")})
    raw_by_code = fetch_raw_responses_by_code(client_id, FORM_REPARATION_ACTIF, codes)

    already_resolved_by_rc = defaultdict(set)
    for r in reparation_rows:
        rc = r.get("Response Code", "")
        wp = (r.get("Water Point ID > Unique ID") or "").strip()
        if rc and wp:
            already_resolved_by_rc[rc].add(wp)

    unassigned_by_rc = {}
    for rc, raws in raw_by_code.items():
        already = already_resolved_by_rc.get(rc, set())
        unassigned_by_rc[rc] = [raw for raw in raws if extract_water_point_code(raw) not in already]

    count = 0
    for r in candidates:
        rc = r.get("Response Code", "")
        pool = unassigned_by_rc.get(rc) or []
        if not pool:
            continue
        raw = pool.pop(0)
        wp_code = extract_water_point_code(raw)
        if wp_code:
            r["Water Point ID > Unique ID"] = wp_code
            r["_wp_enriched"] = True
            already_resolved_by_rc[rc].add(wp_code)
            count += 1
    return count


# ---------------------------------------------------------------------------
# Ancien formulaire combiné : réponses de réparation/entretien préventif à inclure
# ---------------------------------------------------------------------------

def filter_legacy_reparation_rows(rehab_rows):
    """Filtre le datagrid Première réhabilitation (mixte) sur les types de travaux
    équivalents à une réparation, saisis avant la mise en service du formulaire actif —
    voir LEGACY_TYPES_TRAVAUX_REPARATION.

    ATTENTION : l'ancien formulaire n'a pas le même schéma de questions que le nouveau
    (pas de 'Is the repair possible ?', 'Signal code', 'Completion date of the work'
    confirmés sous ces mêmes noms de colonne). Seul le Water Point ID est renormalisé
    ci-dessous (colonne différente sur ce datagrid) ; les dimensions qui dépendent des
    champs spécifiques au nouveau formulaire ne détecteront simplement rien sur ces
    lignes historiques tant que ce mapping n'est pas complété et validé contre un export
    réel — à traiter avant la mise en production si ces réponses historiques doivent
    être couvertes plus finement que la seule Complétude/Fiabilité."""
    legacy_wp_field = "De quel point d'eau s'agit-il? > Unique ID"
    rows = []
    for r in rehab_rows:
        if (r.get("Type de travaux") or "").strip() not in LEGACY_TYPES_TRAVAUX_REPARATION:
            continue
        row = dict(r)
        row["Water Point ID > Unique ID"] = row.get(legacy_wp_field, "")
        row["_from_legacy_form"] = True
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# Anomalie : structure commune
# ---------------------------------------------------------------------------

class Anomaly:
    def __init__(self, dimension, subdimension, response_code, water_point_id, description,
                 details="", champ_concerne=""):
        self.dimension = dimension
        self.subdimension = subdimension
        self.response_code = response_code
        self.water_point_id = water_point_id
        self.description = description
        self.details = details
        self.champ_concerne = champ_concerne


# ---------------------------------------------------------------------------
# Dimension : Complétude
# ---------------------------------------------------------------------------

def check_completude(reparation_rows):
    """Pas d'option 'Don't Know' sur ce formulaire (contrairement à Appel maintenance
    préventive) : un Water Point ID vide sur une réponse finalisée est donc soit un site
    supprimé/inaccessible (si `_wp_enriched`), soit une vraie anomalie."""
    anomalies = []
    for r in reparation_rows:
        status = r.get("Status", "")
        rc = r.get("Response Code", "")
        wp = r.get("Water Point ID > Unique ID", "")

        if status == "Draft":
            anomalies.append(Anomaly(
                "Complétude", "Brouillon non soumis", rc, wp,
                "Brouillon jamais soumis",
                f"Drafted On: {r.get('Drafted On', '')}",
            ))
        elif status == "Final" and not wp.strip():
            anomalies.append(Anomaly(
                "Complétude", "Water Point ID manquant", rc, "",
                "Réponse finalisée sans Water Point ID renseigné",
            ))
        elif status == "Final" and r.get("_wp_enriched"):
            anomalies.append(Anomaly(
                "Complétude", "Site associé supprimé/inaccessible", rc, wp,
                "Water Point ID répondu mais site associé supprimé ou inaccessible",
                f"Code référencé : {wp} — non résolu dans l'export mWater malgré une réponse valide (à vérifier dans le portail)",
            ))
    return anomalies


# ---------------------------------------------------------------------------
# Dimension : Promptitude
# ---------------------------------------------------------------------------

def check_promptitude(reparation_rows):
    """Contrairement au handoff initial : pas de vérification du délai de réparation
    (Completion date of the work vs Date of breakdown) — décision du 14/09/2026 de ne
    pas intégrer cette sous-règle, faute de délai de référence défini. Seule la
    chronologie de saisie est vérifiée, comme pour les autres activités de ce repo."""
    anomalies = []
    for r in reparation_rows:
        drafted = parse_dt(r.get("Drafted On"))
        submitted = parse_dt(r.get("Submitted On"))
        if drafted and submitted and submitted < drafted:
            anomalies.append(Anomaly(
                "Promptitude", "Chronologie Drafted/Submitted",
                r.get("Response Code", ""), r.get("Water Point ID > Unique ID", ""),
                "Submitted On antérieur à Drafted On",
                f"Drafted On: {drafted} / Submitted On: {submitted}",
            ))
    return anomalies


# ---------------------------------------------------------------------------
# Dimension : Validité
# ---------------------------------------------------------------------------

def check_validite(reparation_rows):
    anomalies = []
    for r in reparation_rows:
        rc = r.get("Response Code", "")
        wp = r.get("Water Point ID > Unique ID", "")

        signal_code = (r.get("Signal code") or "").strip()
        if signal_code and not SIGNAL_CODE_RE.match(signal_code):
            corrected, note = propose_correction(signal_code)
            description = "Signal code ne respecte pas le format attendu"
            if note:
                description += f" ({note})"
            details = f"Signal code correct proposé : {corrected}" if corrected else ""
            anomalies.append(Anomaly(
                "Validité", "Format signal code", rc, wp, description,
                details=details, champ_concerne=signal_code,
            ))

        if wp and not wp.strip().isdigit():
            anomalies.append(Anomaly(
                "Validité", "Format Water Point ID", rc, wp,
                "Water Point ID non numérique",
            ))

        repair_possible = (r.get("Is the repair possible ?") or "").strip()
        if repair_possible and repair_possible not in ("Oui", "Non"):
            anomalies.append(Anomaly(
                "Validité", "Valeur 'Is the repair possible ?'", rc, wp,
                "Hors des valeurs attendues (Oui/Non)",
                champ_concerne=repair_possible,
            ))

        repair_successful = (r.get("Is the repair succesfull ?") or "").strip()
        if repair_successful and repair_successful not in ("Oui", "Non"):
            anomalies.append(Anomaly(
                "Validité", "Valeur 'Is the repair succesfull ?'", rc, wp,
                "Hors des valeurs attendues (Oui/Non)",
                champ_concerne=repair_successful,
            ))
    return anomalies


# ---------------------------------------------------------------------------
# Dimension : Unicité
# ---------------------------------------------------------------------------

def check_unicite(reparation_rows):
    """Comme pour Appel maintenance préventive : un rejet 'doublon' ne compte que si la
    réponse est toujours au statut Rejected (pas si elle est repassée Final depuis)."""
    anomalies = []
    non_draft = [r for r in reparation_rows if r.get("Status") != "Draft"]
    counts = Counter(
        (r.get("Signal code") or "").strip()
        for r in non_draft
        if (r.get("Signal code") or "").strip()
    )
    duplicated_codes = {code for code, n in counts.items() if n > 1}

    for r in reparation_rows:
        signal_code = (r.get("Signal code") or "").strip()
        status = r.get("Status")
        rejection = (r.get("Rejection message") or "").strip()
        if status != "Draft" and signal_code in duplicated_codes:
            anomalies.append(Anomaly(
                "Unicité", "Doublon signal code",
                r.get("Response Code", ""), r.get("Water Point ID > Unique ID", ""),
                "Signal code dupliqué",
                champ_concerne=signal_code,
            ))
        elif status == "Rejected" and "doublon" in rejection.lower():
            anomalies.append(Anomaly(
                "Unicité", "Doublon signal code",
                r.get("Response Code", ""), r.get("Water Point ID > Unique ID", ""),
                "Rejet mWater pour doublon de signal code",
                details="" if signal_code else "Signal code vide",
                champ_concerne=signal_code,
            ))
    return anomalies


# ---------------------------------------------------------------------------
# Dimension : Cohérence
# ---------------------------------------------------------------------------

def check_coherence(reparation_rows):
    """Ne réimplémente PAS la règle déjà couverte côté Appel maintenance préventive
    (signal code présent dans Appel avec pompe 'No', date du signal code <= date de
    complétion) — voir docstring en tête de fichier. Vérifie seulement ce qui est
    spécifique à ce formulaire : cohérence entre 'Is the repair possible ?' et la
    présence/absence des champs conditionnels qui en dépendent (section 'Repair is
    possible' du formulaire, native au design, absente de la réponse quand
    'Is the repair possible ? = Non' — pas une anomalie dans ce cas-là)."""
    anomalies = []
    for r in reparation_rows:
        if r.get("Status") == "Draft":
            continue
        rc = r.get("Response Code", "")
        wp = r.get("Water Point ID > Unique ID", "")
        repair_possible = (r.get("Is the repair possible ?") or "").strip()
        repair_successful = (r.get("Is the repair succesfull ?") or "").strip()
        completion_date = (r.get("Completion date of the work") or "").strip()

        if repair_possible == "Non":
            if repair_successful or completion_date:
                anomalies.append(Anomaly(
                    "Cohérence", "Champs conditionnels incohérents", rc, wp,
                    "'Is the repair possible ?' = Non mais des champs de la section "
                    "'Repair is possible' sont renseignés",
                    details=f"Is the repair succesfull ? = '{repair_successful}' / "
                            f"Completion date of the work = '{completion_date}'",
                ))
        elif repair_possible == "Oui":
            if not repair_successful:
                anomalies.append(Anomaly(
                    "Cohérence", "Champs conditionnels incohérents", rc, wp,
                    "'Is the repair possible ?' = Oui mais 'Is the repair succesfull ?' "
                    "n'est pas renseigné",
                ))
            if not completion_date:
                anomalies.append(Anomaly(
                    "Cohérence", "Champs conditionnels incohérents", rc, wp,
                    "'Is the repair possible ?' = Oui mais 'Completion date of the work' "
                    "n'est pas renseigné",
                ))
    return anomalies


# ---------------------------------------------------------------------------
# Dimension : Fiabilité
# ---------------------------------------------------------------------------

def check_fiabilite(reparation_rows, rehab_rows):
    """Même mécanisme que verify_maintenance_preventive.py, appliqué aux rejets propres
    à Réparation après panne plutôt qu'à ceux d'Appel."""
    anomalies = []
    wp_field = "De quel point d'eau s'agit-il? > Unique ID"

    rehab_only = [r for r in rehab_rows if (r.get("Type de travaux") or "").strip() == "Première réhabilitation"]
    success_wp_ids = {
        r.get(wp_field, "").strip()
        for r in rehab_only
        if r.get("Status") == "Final" and r.get(wp_field, "").strip()
    }

    all_wp_context = defaultdict(list)
    for r in rehab_rows:
        wp = r.get(wp_field, "").strip()
        if wp:
            all_wp_context[wp].append((r.get("Type de travaux", ""), r.get("Status", "")))

    for r in reparation_rows:
        if r.get("_from_legacy_form"):
            continue  # ces lignes viennent déjà du datagrid Réhabilitation lui-même
        rejection = (r.get("Rejection message") or "").strip()
        status = r.get("Status")
        if status != "Rejected" or "id incorrect" not in rejection.lower():
            continue

        rc = r.get("Response Code", "")
        wp = r.get("Water Point ID > Unique ID", "").strip()

        if wp in success_wp_ids:
            anomalies.append(Anomaly(
                "Fiabilité", "Rejet ID incorrect", rc, wp,
                "Rejet 'ID incorrect' potentiellement injustifié",
                "Ce Water Point ID existe dans la liste des Premières réhabilitations réussies (Status Final)",
            ))
        elif wp in all_wp_context:
            context = "; ".join(f"{t or 'type inconnu'} ({s})" for t, s in all_wp_context[wp])
            anomalies.append(Anomaly(
                "Fiabilité", "Rejet ID incorrect", rc, wp,
                "Rejet 'ID incorrect' à réexaminer — WP ID trouvé ailleurs dans le datagrid Réhabilitation",
                f"Trouvé sous : {context}",
            ))
    return anomalies


# ---------------------------------------------------------------------------
# Rapport Excel — log unique avec suivi Nouveau / Toujours ouvert / Résolu / Supprimé
# ---------------------------------------------------------------------------

def anomaly_key(a):
    """Identifiant stable d'une anomalie à travers les exécutions successives — voir
    verify_maintenance_preventive.py pour la raison d'inclure Water Point ID (plusieurs
    points d'eau peuvent partager la même Dimension/Sous-dimension/Response Code/
    Description sur une réponse en masse)."""
    return (a.dimension, a.subdimension, a.response_code, a.description, a.water_point_id)


def build_submitted_on_lookup(reparation_rows):
    lookup = {}
    for r in reparation_rows:
        rc = r.get("Response Code")
        dt = parse_dt(r.get("Submitted On"))
        if rc and dt:
            lookup[rc] = dt.strftime("%d/%m/%Y")
    return lookup


def merge_with_log(current_anomalies, existing_rows, today_str, submitted_on_lookup=None,
                    existing_response_codes=None):
    """Fusionne les anomalies détectées aujourd'hui avec le log existant — même logique à
    4 statuts que verify_maintenance_preventive.py (voir sa docstring pour le détail de
    Résolu vs Supprimé)."""
    submitted_on_lookup = submitted_on_lookup or {}
    existing_response_codes = existing_response_codes or set()
    existing_open = {}
    already_closed = []
    for row in existing_rows:
        key = (row.get("Dimension"), row.get("Sous-dimension"), row.get("Response Code"),
               row.get("Description"), row.get("Water Point ID"))
        if row.get("Statut") in ("Résolu", "Supprimé"):
            already_closed.append(row)
        else:
            existing_open[key] = row

    merged = []
    seen_keys = set()
    new_count = 0

    for a in current_anomalies:
        key = anomaly_key(a)
        seen_keys.add(key)
        existing = existing_open.get(key)
        if not existing:
            new_count += 1
        merged.append({
            "Dimension": a.dimension,
            "Sous-dimension": a.subdimension,
            "Response Code": a.response_code,
            "Water Point ID": a.water_point_id,
            "Description": a.description,
            "Champ concerné": a.champ_concerne,
            "Détails": a.details,
            "Statut": "Toujours ouvert" if existing else "Nouveau",
            "Première détection": existing.get("Première détection") if existing else today_str,
            "Dernière détection": today_str,
            "Date de résolution": "",
        })

    resolved_count = 0
    deleted_count = 0
    for key, row in existing_open.items():
        if key not in seen_keys:
            row = dict(row)
            if row.get("Response Code") not in existing_response_codes:
                row["Statut"] = "Supprimé"
                row["Date de résolution"] = row.get("Première détection", "")
                deleted_count += 1
            else:
                row["Statut"] = "Résolu"
                row["Date de résolution"] = submitted_on_lookup.get(row.get("Response Code"), today_str)
                resolved_count += 1
            merged.append(row)

    merged.extend(already_closed)
    return merged, new_count, resolved_count, deleted_count


def build_log_report(merged_rows, output_path):
    from openpyxl import Workbook
    from openpyxl.styles import Font

    wb = Workbook()
    ws = wb.active
    ws.title = "Anomalies"
    ws.append(LOG_HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in merged_rows:
        ws.append([row.get(h, "") for h in LOG_HEADERS])

    open_rows = [r for r in merged_rows if r.get("Statut") not in ("Résolu", "Supprimé")]
    summary = wb.create_sheet("Résumé")
    summary.append(["Dimension", "Nombre d'anomalies ouvertes"])
    for cell in summary[1]:
        cell.font = Font(bold=True)
    counts = Counter(r.get("Dimension") for r in open_rows)
    for dimension in ["Complétude", "Promptitude", "Validité", "Unicité", "Cohérence", "Fiabilité"]:
        summary.append([dimension, counts.get(dimension, 0)])
    summary.append(["Total ouvert", len(open_rows)])

    wb.save(output_path)


def send_confirmation_email(token, sender, recipients, file_name, open_count, new_count,
                             resolved_count, deleted_count, counts_by_dimension):
    lines = "".join(f"<li>{dim} : {n}</li>" for dim, n in counts_by_dimension.items())
    body_html = (
        f"<p>Vérification de données \"Réparation après panne\" exécutée.</p>"
        f"<p>Anomalies actuellement ouvertes : <b>{open_count}</b> "
        f"(dont {new_count} nouvelles) — {resolved_count} résolue(s) et {deleted_count} "
        f"supprimée(s) (réponse mWater disparue) depuis la dernière exécution.</p>"
        f"<ul>{lines}</ul>"
        f"<p>Log mis à jour sur SharePoint : <b>{file_name}</b></p>"
    )
    send_html_email(token, sender, recipients, f"Vérification de données — {file_name}", body_html)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    mwater_username = os.environ["MWATER_USERNAME"].strip()
    mwater_password = os.environ["MWATER_PASSWORD"].strip()

    azure_tenant_id = os.environ["AZURE_TENANT_ID"].strip()
    azure_client_id = os.environ["AZURE_CLIENT_ID"].strip()
    azure_client_secret = os.environ["AZURE_CLIENT_SECRET"].strip()

    sharepoint_folder_link = os.environ["SHAREPOINT_FOLDER_LINK"].strip()

    email_sender = os.environ["EMAIL_SENDER"].strip()
    email_recipients = os.environ["EMAIL_RECIPIENTS"].strip()

    print("Authentification mWater...")
    client_id = mwater_login(mwater_username, mwater_password)

    print("Téléchargement des datagrids...")
    reparation_rows = download_datagrid(DATAGRID_REPARATION, client_id)
    rehab_rows = download_datagrid(DATAGRID_REHAB, client_id)
    print(f"  Réparation après panne : {len(reparation_rows)} / Réhabilitation (mixte) : {len(rehab_rows)}")

    legacy_rows = filter_legacy_reparation_rows(rehab_rows)
    print(f"  {len(legacy_rows)} réponses historiques de l'ancien formulaire combiné incluses")
    reparation_rows = reparation_rows + legacy_rows

    print("Résolution des Water Point ID vides (site fusionné/supprimé mais réponse valide)...")
    enriched_count = enrich_missing_water_point_ids(reparation_rows, client_id)
    print(f"  {enriched_count or 0} lignes enrichies avec leur code brut")

    print("Application des règles de vérification (6 dimensions)...")
    anomalies = []
    anomalies += check_completude(reparation_rows)
    anomalies += check_promptitude(reparation_rows)
    anomalies += check_validite(reparation_rows)
    anomalies += check_unicite(reparation_rows)
    anomalies += check_coherence(reparation_rows)
    anomalies += check_fiabilite(reparation_rows, rehab_rows)
    print(f"  {len(anomalies)} anomalies détectées")

    print("Authentification Microsoft Graph...")
    token = graph_token(azure_tenant_id, azure_client_id, azure_client_secret)

    print("Résolution du lien SharePoint...")
    sharepoint_drive_id, sharepoint_folder_item_id = resolve_share_link(token, sharepoint_folder_link)

    print("Téléchargement du log existant sur SharePoint...")
    existing_rows = download_existing_log(token, sharepoint_drive_id, sharepoint_folder_item_id, LOG_FILE_NAME)
    print(f"  {len(existing_rows)} lignes déjà présentes dans le log")

    today_str = datetime.now().strftime("%d/%m/%Y")
    submitted_on_lookup = build_submitted_on_lookup(reparation_rows)
    existing_response_codes = {r.get("Response Code", "") for r in reparation_rows}
    merged_rows, new_count, resolved_count, deleted_count = merge_with_log(
        anomalies, existing_rows, today_str,
        submitted_on_lookup=submitted_on_lookup,
        existing_response_codes=existing_response_codes)
    open_rows = [r for r in merged_rows if r.get("Statut") not in ("Résolu", "Supprimé")]
    print(f"  {len(open_rows)} anomalies ouvertes ({new_count} nouvelles, {resolved_count} résolues, "
          f"{deleted_count} supprimées)")

    try:
        inserer_log_dans_mwater(merged_rows, client_id)
    except Exception as exc:
        print(f"AVERTISSEMENT : échec de l'insertion du log dans mWater ({exc}) - "
              f"le fichier Excel reste la source fiable pour cette exécution.",
              file=sys.stderr)

    output_path = f"/tmp/{LOG_FILE_NAME}"
    build_log_report(merged_rows, output_path)
    print(f"Rapport généré : {output_path}")

    print("Dépôt du log mis à jour sur SharePoint...")
    upload_to_sharepoint(token, sharepoint_drive_id, sharepoint_folder_item_id, output_path, LOG_FILE_NAME)

    print("Envoi de l'email de confirmation...")
    counts_by_dimension = Counter(r.get("Dimension") for r in open_rows)
    send_confirmation_email(token, email_sender, email_recipients, LOG_FILE_NAME,
                             len(open_rows), new_count, resolved_count, deleted_count, counts_by_dimension)

    print("Terminé.")


if __name__ == "__main__":
    main()
