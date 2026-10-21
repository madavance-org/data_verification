#!/usr/bin/env python3
"""
Vérification de données — Clean Cooking : Distribution
=======================================================

Applique les six dimensions du "Manuel de vérification — Clean Cooking : Distribution"
(Complétude, Promptitude, Validité, Unicité, Cohérence, Fiabilité) au datagrid mWater
"Formulaire de distribution" (formulaire f1fcf0c427ea46d2a702757bece97b39), en le croisant avec le
datagrid Inscription pour la Cohérence. Met à jour un log Excel cumulatif
(data_verification_clean_cooking_distribution_log.xlsx : Nouveau / Toujours ouvert / Résolu /
Supprimé) déposé sur SharePoint, et envoie un email de confirmation.

Mode test : avec DRY_RUN=true, le script affiche le décompte des anomalies sans toucher à
SharePoint ni envoyer d'email.

Écarts assumés avec le guide (à valider) :
- Cohérence, stock par QR code : le guide s'appuie sur le journal des foyers (DuckDB,
  recap_activite_foyers.journal), qui n'existe pas côté mWater. Cette règle n'est PAS appliquée ici ;
  elle demande de définir les mouvements qui retirent un QR code (retours, remplacements, vols,
  ventes, transmissions) à partir des autres formulaires.
- Cohérence, date de distribution vs inscription : la "date d'inscription" est approchée par la
  date de début de saisie (Drafted On) de l'inscription Final du ménage.
"""

import os
import re
from collections import Counter

from clean_cooking.anomaly_log import (
    Anomaly,
    STATUTS_CLOS,
    build_log_report,
    build_submitted_on_lookup,
    has_entity,
    is_digits,
    merge_with_log,
    print_summary,
    send_confirmation_email,
    today,
    val,
)
from common.dates import parse_dt
from common.mwater_client import download_datagrid_with_retry, fetch_raw_responses_by_code, mwater_login
from common.sharepoint import (
    download_existing_log,
    folder_link_from_env,
    graph_token,
    resolve_share_link,
    upload_to_sharepoint,
)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

DATAGRID_DISTRIBUTION = "ba07c16edc57417093d9735c93a97e52"
FORM_DISTRIBUTION = "f1fcf0c427ea46d2a702757bece97b39"
DATAGRID_INSCRIPTION = "ac3ae2ce1f7540f586ae1ab5673419a8"

LOG_FILE_NAME = "data_verification_clean_cooking_distribution_log.xlsx"

# Noms de colonnes du datagrid Distribution (orthographe exacte du portail)
C_STATUS = "Status"
C_CODE = "Response Code"
C_HOUSEHOLD = "Code ID du ménage > Unique ID"
C_MODELE = "Modèle/type de foyer amélioré reçu"
C_TYPE = "Type"
C_QR = ["QR Code 1", "QR Code 2", "QR Code 3"]
C_QUANTITE = "Quantité de foyer améliorés du projet reçus"
C_RETOURNES = "Foyers non efficient retournés"
C_MONTANT = "Montant payé"
C_DATE = "Date de distribution"
C_FICHE = "Fiche de distribution"
C_CONVENTION = "Photo de convention d'utilisation carbone"
C_DRAFTED = "Drafted On"
C_SUBMITTED = "Submitted On"
C_REJECTION = "Rejection message"

# Colonnes du datagrid Inscription utilisées pour la Cohérence
I_STATUS = "Status"
I_HOUSEHOLD = "Household ID code > Unique ID"
I_DRAFTED = "Drafted On"

# QR codes attendus selon le type de foyer (guide : 2 pour le Type 1, 1 pour les Types 2 et 3)
QR_ATTENDUS = {
    "Type 1": ["QR Code 1", "QR Code 2"],
    "Type 2": ["QR Code 1"],
    "Type 3": ["QR Code 3"],
}
QR_RE = re.compile(r"^[A-Z0-9]{6}$")
NOMBRE_RE = re.compile(r"^[0-9]+([.,][0-9]+)?$")


def new(dimension, subdimension, r, description, details=""):
    return Anomaly(dimension, subdimension, val(r, C_CODE), val(r, C_HOUSEHOLD), description, details)


def type_du_modele(r):
    """Type de foyer (Type 1/2/3) déduit du libellé du modèle reçu ; None si absent."""
    m = re.search(r"Type\s*([123])", val(r, C_MODELE))
    return f"Type {m.group(1)}" if m else None


# ---------------------------------------------------------------------------
# Entité ménage supprimée/inaccessible (distinction Complétude)
# ---------------------------------------------------------------------------

def resolve_missing_households(rows, client_id):
    codes = sorted({val(r, C_CODE) for r in rows if not val(r, C_HOUSEHOLD)})
    if not codes:
        return {}
    raw = fetch_raw_responses_by_code(client_id, FORM_DISTRIBUTION, codes)
    return {c: has_entity(raw.get(c), "household") for c in codes}


# ---------------------------------------------------------------------------
# Dimensions
# ---------------------------------------------------------------------------

def check_completude(rows, household_entities):
    out = []
    for r in rows:
        status = val(r, C_STATUS)
        if status != "Final":
            out.append(new("Complétude", "Réponse non finalisée", r,
                           "Distribution non finalisée", f"Status : {status or 'vide'}"))

        if not val(r, C_HOUSEHOLD):
            if household_entities.get(val(r, C_CODE)):
                out.append(new("Complétude", "Ménage supprimé/inaccessible", r,
                               "Identifiant du ménage répondu mais site associé supprimé ou inaccessible",
                               "Non résolu dans l'export mWater malgré une réponse valide (à vérifier dans le portail)"))
            else:
                out.append(new("Complétude", "Identifiant du ménage manquant", r, "Identifiant du ménage vide"))

        modele = val(r, C_MODELE)
        if not modele:
            out.append(new("Complétude", "Modèle de foyer manquant", r, "Modèle de foyer reçu non renseigné"))
        else:
            type_foyer = type_du_modele(r) or val(r, C_TYPE)
            attendus = QR_ATTENDUS.get(type_foyer, [])
            for col in attendus:
                if not val(r, col):
                    out.append(new("Complétude", "QR code manquant", r,
                                   f"{col} manquant pour un foyer {type_foyer}",
                                   f"QR codes attendus : {len(attendus)}"))
            # Le guide exige un nombre de QR codes égal à la quantité attendue : un QR en trop
            # (ex. QR Code 1 renseigné pour un Type 3) est aussi un écart.
            if attendus:
                for col in C_QR:
                    if col not in attendus and val(r, col):
                        out.append(new("Complétude", "QR code en trop", r,
                                       f"{col} renseigné alors qu'il n'est pas attendu pour un foyer {type_foyer}",
                                       f"QR codes attendus : {', '.join(attendus)}"))

        # Montant payé et foyers retournés : toujours obligatoires (0 accepté)
        if not val(r, C_MONTANT):
            out.append(new("Complétude", "Montant payé manquant", r,
                           "Montant payé non renseigné (0 si la distribution est gratuite)"))
        if not val(r, C_DATE):
            out.append(new("Complétude", "Date de distribution manquante", r, "Date de distribution non renseignée"))
        if not val(r, C_RETOURNES):
            out.append(new("Complétude", "Foyers retournés manquant", r,
                           "Nombre de foyers non efficients retournés non renseigné (0 si aucun)"))
        for label, col in (("Fiche de distribution", C_FICHE), ("Photo de la convention d'utilisation carbone", C_CONVENTION)):
            if not val(r, col):
                out.append(new("Complétude", "Pièce justificative manquante", r, f"Pièce manquante : {label}"))
    return out


def check_promptitude(rows):
    out = []
    for r in rows:
        drafted, submitted, distribution = (parse_dt(val(r, C_DRAFTED)), parse_dt(val(r, C_SUBMITTED)),
                                            parse_dt(val(r, C_DATE)))
        if drafted and submitted and submitted < drafted:
            out.append(new("Promptitude", "Chronologie Drafted/Submitted", r,
                           "Submitted On antérieur à Drafted On", f"Drafted On : {drafted} / Submitted On : {submitted}"))
        if distribution and submitted and distribution.date() > submitted.date():
            out.append(new("Promptitude", "Date de distribution postérieure à la soumission", r,
                           "Date de distribution postérieure à Submitted On",
                           f"Date de distribution : {distribution.date()} / Submitted On : {submitted.date()}"))
    return out


def check_validite(rows):
    out = []
    for r in rows:
        household = val(r, C_HOUSEHOLD)
        if household and not is_digits(household):
            out.append(new("Validité", "Format identifiant ménage", r, "Identifiant du ménage non numérique",
                           f"Valeur : '{household}'"))
        for col in C_QR:
            qr = val(r, col)
            if qr and not QR_RE.match(qr):
                out.append(new("Validité", "Format QR code", r,
                               f"{col} invalide (6 caractères alphanumériques en majuscules attendus)",
                               f"Valeur : '{qr}'"))
        modele, type_col = val(r, C_MODELE), val(r, C_TYPE)
        if modele:
            deduit = type_du_modele(r)
            if deduit is None:
                out.append(new("Validité", "Modèle de foyer", r,
                               "Modèle de foyer reçu hors des valeurs du formulaire (Type 1, 2 ou 3)",
                               f"Valeur : '{modele}'"))
            elif type_col and type_col != deduit:
                out.append(new("Validité", "Type / modèle", r, "Type de foyer ne correspondant pas au modèle",
                               f"Modèle : '{modele}' / Type : '{type_col}'"))
        quantite = val(r, C_QUANTITE)
        if quantite and quantite not in ("1", "2", "1.0", "2.0"):
            out.append(new("Validité", "Quantité de foyers reçus", r,
                           "Quantité de foyers améliorés reçus invalide (1 ou 2 attendu)", f"Valeur : '{quantite}'"))
        retournes = val(r, C_RETOURNES)
        if retournes and not re.fullmatch(r"[0-9]+(\.0+)?", retournes):
            out.append(new("Validité", "Foyers retournés", r,
                           "Nombre de foyers non efficients retournés invalide (entier ≥ 0 attendu)",
                           f"Valeur : '{retournes}'"))
        montant = val(r, C_MONTANT)
        if montant and not NOMBRE_RE.match(montant):
            out.append(new("Validité", "Montant payé", r, "Montant payé invalide (nombre ≥ 0 attendu)",
                           f"Valeur : '{montant}'"))
    return out


def check_unicite(rows):
    out = []
    counts = Counter(val(r, C_CODE) for r in rows if val(r, C_CODE))
    for r in rows:
        code = val(r, C_CODE)
        if code and counts[code] > 1:
            out.append(new("Unicité", "Doublon Response Code", r, "Response Code dupliqué",
                           f"{counts[code]} réponses partagent ce code"))
        elif val(r, C_STATUS) == "Rejected" and "doublon" in val(r, C_REJECTION).lower():
            out.append(new("Unicité", "Rejet pour doublon", r, "Rejet mWater pour doublon", val(r, C_REJECTION)))
    return out


def check_coherence(rows, inscription_rows):
    final_ids, drafted_par_menage = set(), {}
    for r in inscription_rows:
        household = val(r, I_HOUSEHOLD)
        if household and val(r, I_STATUS) == "Final":
            final_ids.add(household)
            drafted = parse_dt(val(r, I_DRAFTED))
            if drafted and (household not in drafted_par_menage or drafted < drafted_par_menage[household]):
                drafted_par_menage[household] = drafted
    out = []
    for r in rows:
        household = val(r, C_HOUSEHOLD)
        if not household:
            continue
        if household not in final_ids:
            out.append(new("Cohérence", "Inscription du ménage", r,
                           "Ménage distribué sans inscription au statut Final"))
            continue
        distribution, inscription = parse_dt(val(r, C_DATE)), drafted_par_menage.get(household)
        if distribution and inscription and distribution.date() < inscription.date():
            out.append(new("Cohérence", "Date distribution vs inscription", r,
                           "Date de distribution antérieure à la date d'inscription du ménage",
                           f"Distribution : {distribution.date()} / Inscription (début de saisie) : {inscription.date()}"))
    return out


def check_fiabilite(rows):
    out = []
    for r in rows:
        # Un rejet n'est une anomalie que si la réponse est toujours au statut Rejected.
        if val(r, C_STATUS) == "Rejected":
            out.append(new("Fiabilité", "Rejet mWater non corrigé", r, "Réponse rejetée non corrigée",
                           val(r, C_REJECTION) or "Aucun message de rejet"))
    return out


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    dry_run = os.environ.get("DRY_RUN", "").strip().lower() in ("1", "true", "yes")

    print("Authentification mWater...")
    client_id = mwater_login(os.environ["MWATER_USERNAME"].strip(), os.environ["MWATER_PASSWORD"].strip())

    print("Téléchargement du datagrid Distribution...")
    # ~90 000 lignes : mWater répond parfois 524 (export trop long à générer) -> plus de tentatives
    rows = download_datagrid_with_retry(DATAGRID_DISTRIBUTION, client_id, attempts=8, wait_seconds=60)
    print(f"  {len(rows)} lignes")
    print("Téléchargement du datagrid Inscription (Cohérence)...")
    inscription_rows = download_datagrid_with_retry(DATAGRID_INSCRIPTION, client_id, attempts=8, wait_seconds=60)
    print(f"  {len(inscription_rows)} lignes")

    print("Résolution des ID ménage vides (site supprimé ou jamais répondu)...")
    household_entities = resolve_missing_households(rows, client_id)
    print(f"  {len(household_entities)} réponses vérifiées sur l'API brute")

    print("Application des règles de vérification (6 dimensions)...")
    anomalies = []
    anomalies += check_completude(rows, household_entities)
    anomalies += check_promptitude(rows)
    anomalies += check_validite(rows)
    anomalies += check_unicite(rows)
    anomalies += check_coherence(rows, inscription_rows)
    anomalies += check_fiabilite(rows)
    print(f"  {len(anomalies)} anomalies détectées")
    print_summary(anomalies)

    if dry_run:
        print("DRY_RUN : aucun accès SharePoint, aucun email, aucun log écrit.")
        return

    azure_tenant_id = os.environ["AZURE_TENANT_ID"].strip()
    azure_client_id = os.environ["AZURE_CLIENT_ID"].strip()
    azure_client_secret = os.environ["AZURE_CLIENT_SECRET"].strip()
    sharepoint_folder_link = folder_link_from_env("SHAREPOINT_FOLDER_LINK_DISTRIBUTION")
    email_sender = os.environ["EMAIL_SENDER"].strip()
    email_recipients = os.environ["EMAIL_RECIPIENTS"].strip()

    print("Authentification Microsoft Graph...")
    token = graph_token(azure_tenant_id, azure_client_id, azure_client_secret)
    print("Résolution du lien SharePoint...")
    drive_id, folder_item_id = resolve_share_link(token, sharepoint_folder_link)

    print("Téléchargement du log existant sur SharePoint...")
    existing_rows = download_existing_log(token, drive_id, folder_item_id, LOG_FILE_NAME)
    print(f"  {len(existing_rows)} lignes déjà présentes dans le log")

    merged_rows, new_count, resolved_count, deleted_count = merge_with_log(
        anomalies, existing_rows, today(),
        submitted_on_lookup=build_submitted_on_lookup(rows, C_CODE, C_SUBMITTED),
        existing_response_codes={val(r, C_CODE) for r in rows})
    open_rows = [r for r in merged_rows if r.get("Statut") not in STATUTS_CLOS]
    print(f"  {len(open_rows)} anomalies ouvertes ({new_count} nouvelles, {resolved_count} résolues, "
          f"{deleted_count} supprimées)")

    output_path = f"/tmp/{LOG_FILE_NAME}"
    build_log_report(merged_rows, output_path)
    print("Dépôt du log mis à jour sur SharePoint...")
    upload_to_sharepoint(token, drive_id, folder_item_id, output_path, LOG_FILE_NAME)

    print("Envoi de l'email de confirmation...")
    send_confirmation_email(token, email_sender, email_recipients, "Clean Cooking — Distribution",
                            LOG_FILE_NAME, len(open_rows), new_count, resolved_count, deleted_count,
                            Counter(r.get("Dimension") for r in open_rows))
    print("Terminé.")


if __name__ == "__main__":
    main()
