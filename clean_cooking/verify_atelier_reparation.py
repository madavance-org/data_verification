#!/usr/bin/env python3
"""
Vérification de données — Clean Cooking : Atelier de réparation
================================================================

Applique les six dimensions du "Manuel de vérification — Clean Cooking : Atelier de réparation"
au datagrid mWater "Clean Cooking II Réparation des foyers II MERV" (formulaire
d32769553b26450c90d228f555724f5f), en croisant avec le datagrid Distribution pour la Cohérence.
Met à jour un log Excel cumulatif (Nouveau / Toujours ouvert / Résolu / Supprimé) déposé sur
SharePoint et envoie un email de confirmation. Avec DRY_RUN=true : décompte uniquement, sans
SharePoint ni email.

Écarts assumés avec le guide (à valider) :
- Brouillons : une réponse au statut Draft n'est signalée que comme "non finalisée" ; ses champs
  vides ne sont pas listés un par un (un brouillon n'est pas censé être complet).
- Types : "Ecoa CHAR" est rapproché des types Distribution 1 et 3 (foyers à charbon), "Ecoa WOOD" du
  type 2 ; tout autre type (ex. "FabStove") est signalé comme ne correspondant à aucun type distribué.
- Cohérence : un QR distribué plusieurs fois (retour, réparation) est comparé à sa PREMIÈRE
  distribution Final pour la date, et à l'ensemble de ses types pour le type.
- Non appliqué : "le QR code retourné doit appartenir au ménage de la réparation" (point resté
  ouvert dans le guide ; 44 stoves concernés à l'analyse).
- Unicité du numéro de signal : appliquée telle que décrite dans le guide, mais la règle y est
  marquée "à confirmer" (un même signal peut-il couvrir plusieurs ménages ?).
"""

import csv
import io
import os
import re
from collections import Counter, defaultdict

from clean_cooking.anomaly_log import (
    Anomaly,
    STATUTS_CLOS,
    build_log_report,
    build_submitted_on_lookup,
    merge_with_log,
    print_summary,
    send_confirmation_email,
    today,
)
from common.dates import parse_dt
from common.mwater_client import download_datagrid_raw_with_retry, download_datagrid_with_retry, mwater_login
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

DATAGRID_ATELIER = "e129499f02c141a3a45e65bfb8b3f7d7"
DATAGRID_DISTRIBUTION = "ba07c16edc57417093d9735c93a97e52"

LOG_FILE_NAME = "data_verification_clean_cooking_atelier_reparation_log.xlsx"

QR_RE = re.compile(r"^[A-Z0-9]{6}$")
MOTIFS_DOUBLON = ("doublon",)

# Types Atelier -> types Distribution compatibles
TYPES_COMPATIBLES = {"Ecoa CHAR": {"Type 1", "Type 3"}, "Ecoa WOOD": {"Type 2"}}

# Noms de colonnes du datagrid (orthographe exacte du portail)
C_STATUS, C_CODE, C_DRAFTED, C_SUBMITTED, C_REJECTION = "Status", "Response Code", "Drafted On", "Submitted On", "Rejection message"
C_PANNE = "Date of inactvity (failure)"
C_SIGNAL_NB = "Signal number: Reference - Number"
C_SIGNAL_DEP = "Signal number: Reference - Department"
C_MENAGE = "Household ID > Unique ID"
C_PHONE, C_PHONE_NA = "Numéro de téléphone", "Numéro de téléphone (Not Applicable)"
C_NOMBRE = "How many stove is being returned ?"
QR_REMPLACEMENT = "Replaced stoves' QR code"


def champ(base, n):
    return f"{base} - stove {n}"


# Champs de l'atelier, exigés si "réparation possible" = Yes : (libellé, nom de colonne de base)
CHAMPS_ATELIER = [
    ("Date d'entrée en atelier", "Date of entry"),
    ("Nom du réceptionnaire", "Receiver Name"),
    ("Poste du réceptionnaire", "Receiver Position"),
    ("Raison d'entrée", "Reason to enter the workshop"),
    ("Date de début de réparation", "Repair start date"),
    ("Date de fin de réparation", "Repair end date"),
    ("Statut de réparation", "Status of repair"),
    ("Date de redistribution", "Redistribution date"),
    ("Type de stove redistribué", "Redistributed stoves' type"),
    ("Photo de redistribution", "Redistributed stoves' picture"),
]

# Ordre chronologique attendu pour chaque stove : (libellé, colonne de base)
CHRONOLOGIE = [
    ("panne", None), ("retour", "Return date"), ("entrée en atelier", "Date of entry"),
    ("début de réparation", "Repair start date"), ("fin de réparation", "Repair end date"),
    ("redistribution", "Redistribution date"),
]


def val(r, nom):
    return (r.get(nom) or "").strip()


def is_true(r, nom):
    return val(r, nom).lower() == "true"


def date_de(r, nom):
    d = parse_dt(val(r, nom))
    return d.date() if d else None


def new(dimension, subdimension, r, description, details=""):
    return Anomaly(dimension, subdimension, val(r, C_CODE), val(r, C_MENAGE), description, details)


def nombre_de_stoves(r):
    return 2 if val(r, C_NOMBRE) == "2" else 1


def qr_retourne(r, n):
    """QR du stove : champ QR, ou commentaire lorsque « non applicable » est coché."""
    base = champ("QR code", n)
    if val(r, base):
        return val(r, base).rstrip(" \t")
    if is_true(r, f"{base} (Not Applicable)") and val(r, f"{base} (Comments)"):
        return val(r, f"{base} (Comments)").rstrip(" \t")
    return ""


def qr_remplacement(r, n):
    base = champ(QR_REMPLACEMENT, n)
    if val(r, base):
        return val(r, base).rstrip(" \t")
    if is_true(r, f"{base} (Not Applicable)") and val(r, f"{base} (Comments)"):
        return val(r, f"{base} (Comments)").rstrip(" \t")
    return ""


# ---------------------------------------------------------------------------
# Complétude
# ---------------------------------------------------------------------------

def check_completude(rows):
    out = []
    for r in rows:
        statut = val(r, C_STATUS)
        if statut != "Final":
            out.append(new("Complétude", "Réponse non finalisée", r, "Réparation non finalisée",
                           f"Status : {statut or 'vide'}"))
        if statut == "Draft":
            continue  # un brouillon n'est pas censé être complet (voir docstring)

        def exiger(sous_dimension, libelle, ok):
            if not ok:
                out.append(new("Complétude", sous_dimension, r, f"{libelle} non renseigné"))

        g = "Champ général"
        exiger(g, "Référence du signal", bool(val(r, C_SIGNAL_NB)))
        exiger(g, "Département du signal", bool(val(r, C_SIGNAL_DEP)))
        exiger(g, "Ménage", bool(val(r, C_MENAGE)))
        exiger(g, "Date de la panne", bool(val(r, C_PANNE)))
        exiger(g, "Nombre de foyers retournés", bool(val(r, C_NOMBRE)))
        exiger("Téléphone", "Téléphone (« non applicable » non coché)",
               bool(val(r, C_PHONE)) or is_true(r, C_PHONE_NA))

        for n in range(1, nombre_de_stoves(r) + 1):
            s = f"Champ du stove {n}"
            exiger(s, f"Date de retour (stove {n})", bool(val(r, champ("Return date", n))))
            exiger(s, f"Type de stove (stove {n})", bool(val(r, champ("Stoves' type", n))))
            exiger("QR code manquant", f"QR code du stove {n}", bool(qr_retourne(r, n)))
            exiger(s, f"Photo (stove {n})", bool(val(r, champ("Picture", n))))
            exiger(s, f"Types d'anomalies (stove {n})", bool(val(r, champ("Types of anomalies", n))))

            if val(r, champ("Is the reparation possible?", n)) == "Yes":
                a = f"Champ d'atelier du stove {n}"
                for libelle, base in CHAMPS_ATELIER:
                    exiger(a, f"{libelle} (stove {n})", bool(val(r, champ(base, n))))

            # QR de remplacement : vérifié uniquement lorsque la question est posée
            if is_true(r, f"{champ(QR_REMPLACEMENT, n)} (Asked)"):
                exiger("QR code de remplacement manquant", f"QR code de remplacement (stove {n})",
                       bool(qr_remplacement(r, n)))
    return out


# ---------------------------------------------------------------------------
# Promptitude
# ---------------------------------------------------------------------------

def check_promptitude(rows):
    out = []
    for r in rows:
        drafted, submitted = parse_dt(val(r, C_DRAFTED)), parse_dt(val(r, C_SUBMITTED))
        if drafted and submitted and submitted < drafted:
            out.append(new("Promptitude", "Chronologie Drafted/Submitted", r, "Submitted On antérieur à Drafted On",
                           f"Drafted On : {drafted} / Submitted On : {submitted}"))
        soumission = submitted.date() if submitted else None
        for n in range(1, nombre_de_stoves(r) + 1):
            dates = []
            for libelle, base in CHRONOLOGIE:
                d = date_de(r, C_PANNE) if base is None else date_de(r, champ(base, n))
                if d:
                    dates.append((libelle, d))
            for (l1, d1), (l2, d2) in zip(dates, dates[1:]):
                if d2 < d1:
                    out.append(new("Promptitude", "Ordre des dates", r,
                                   f"Date de {l2} antérieure à la date de {l1} (stove {n})",
                                   f"{l1} : {d1} / {l2} : {d2}"))
            if soumission:
                for libelle, d in dates:
                    if d > soumission:
                        out.append(new("Promptitude", "Date postérieure à la soumission", r,
                                       f"Date de {libelle} postérieure à Submitted On (stove {n})",
                                       f"{libelle} : {d} / Submitted On : {soumission}"))
    return out


# ---------------------------------------------------------------------------
# Validité
# ---------------------------------------------------------------------------

def check_validite(rows):
    out = []
    for r in rows:
        if val(r, C_STATUS) == "Draft":
            continue
        for n in (1, 2):
            for libelle, valeur in ((f"QR code retourné (stove {n})", qr_retourne(r, n)),
                                    (f"QR code de remplacement (stove {n})", qr_remplacement(r, n))):
                if valeur and not QR_RE.match(valeur):
                    out.append(new("Validité", "Format QR code", r,
                                   f"{libelle} invalide (6 caractères alphanumériques en majuscules attendus)",
                                   f"Valeur : '{valeur}'"))
    return out


# ---------------------------------------------------------------------------
# Unicité
# ---------------------------------------------------------------------------

def check_unicite(rows):
    out = []
    codes = Counter(val(r, C_CODE) for r in rows if val(r, C_CODE))
    signaux = Counter((val(r, C_SIGNAL_NB), val(r, C_SIGNAL_DEP)) for r in rows
                      if val(r, C_SIGNAL_NB) and val(r, C_SIGNAL_DEP) and val(r, C_STATUS) != "Draft")
    for r in rows:
        code = val(r, C_CODE)
        if code and codes[code] > 1:
            out.append(new("Unicité", "Doublon Response Code", r, "Response Code dupliqué",
                           f"{codes[code]} réponses partagent ce code"))
        cle = (val(r, C_SIGNAL_NB), val(r, C_SIGNAL_DEP))
        if all(cle) and val(r, C_STATUS) != "Draft" and signaux[cle] > 1:
            out.append(new("Unicité", "Doublon numéro de signal", r,
                           "Numéro de signal utilisé pour plusieurs réponses",
                           f"Signal {cle[0]} / {cle[1]} présent dans {signaux[cle]} réponses"))
        if val(r, C_STATUS) == "Rejected" and any(m in val(r, C_REJECTION).lower() for m in MOTIFS_DOUBLON):
            out.append(new("Unicité", "Rejet pour doublon", r, "Rejet mWater pour doublon", val(r, C_REJECTION)))
    return out


# ---------------------------------------------------------------------------
# Cohérence (croisement avec la Distribution, statut Final)
# ---------------------------------------------------------------------------

def index_distribution(distribution_rows):
    types, dates, menages = defaultdict(set), defaultdict(list), set()
    for d in distribution_rows:
        if (d.get("Status") or "").strip() != "Final":
            continue
        menage = (d.get("Code ID du ménage > Unique ID") or "").strip()
        if menage:
            menages.add(menage)
        date_dist = parse_dt(d.get("Date de distribution"))
        for c in ("QR Code 1", "QR Code 2", "QR Code 3"):
            q = (d.get(c) or "").strip().upper()
            if q:
                types[q].add((d.get("Type") or "").strip())
                if date_dist:
                    dates[q].append(date_dist.date())
    return types, dates, menages


def check_coherence(rows, distribution):
    types, dates, menages = distribution
    out = []
    for r in rows:
        if val(r, C_STATUS) == "Draft":
            continue
        menage = val(r, C_MENAGE)
        if menage and menage not in menages:
            out.append(new("Cohérence", "Distribution du ménage", r,
                           "Ménage sans distribution au statut Final"))
        for n in range(1, nombre_de_stoves(r) + 1):
            qr = qr_retourne(r, n).upper()
            type_retourne = val(r, champ("Stoves' type", n))
            if qr and QR_RE.match(qr):  # un QR de format invalide est déjà signalé en Validité
                if qr not in types:
                    out.append(new("Cohérence", "QR code non distribué", r,
                                   f"QR code retourné (stove {n}) absent de la distribution (statut Final)",
                                   f"QR code : {qr}"))
                else:
                    retour = date_de(r, champ("Return date", n))
                    if retour and dates[qr] and retour < min(dates[qr]):
                        out.append(new("Cohérence", "Date de retour vs distribution", r,
                                       f"Date de retour (stove {n}) antérieure à la distribution du QR code",
                                       f"Retour : {retour} / Première distribution : {min(dates[qr])} / QR : {qr}"))
                    if type_retourne and not (TYPES_COMPATIBLES.get(type_retourne, set()) & types[qr]):
                        out.append(new("Cohérence", "Type retourné vs distribué", r,
                                       f"Type de stove retourné (stove {n}) différent du type distribué pour ce QR code",
                                       f"Retourné : {type_retourne} / Distribué : {', '.join(sorted(types[qr]))} / QR : {qr}"))
            redistribue = val(r, champ("Redistributed stoves' type", n))
            if redistribue and type_retourne and redistribue != type_retourne:
                out.append(new("Cohérence", "Type redistribué vs retourné", r,
                               f"Type de stove redistribué (stove {n}) différent du type retourné",
                               f"Retourné : {type_retourne} / Redistribué : {redistribue}"))
    return out


# ---------------------------------------------------------------------------
# Fiabilité
# ---------------------------------------------------------------------------

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

    print("Téléchargement du datagrid Atelier de réparation...")
    texte = download_datagrid_raw_with_retry(DATAGRID_ATELIER, client_id, attempts=8, wait_seconds=60).decode("utf-8-sig")
    lecteur = csv.reader(io.StringIO(texte))
    header = next(lecteur)
    doublons = [h for h, n in Counter(header).items() if n > 1]
    if doublons:
        raise RuntimeError(f"Colonnes en double dans le datagrid Atelier ({doublons}) : la lecture par nom "
                           "n'est plus fiable, passer à une lecture par position.")
    rows = [dict(zip(header, ligne)) for ligne in lecteur if any(c.strip() for c in ligne)]
    print(f"  {len(rows)} lignes")

    print("Téléchargement du datagrid Distribution (Cohérence)...")
    distribution_rows = download_datagrid_with_retry(DATAGRID_DISTRIBUTION, client_id, attempts=8, wait_seconds=60)
    distribution = index_distribution(distribution_rows)
    print(f"  {len(distribution_rows)} lignes, {len(distribution[0])} QR codes distribués (Final)")

    print("Application des règles de vérification (6 dimensions)...")
    anomalies = []
    anomalies += check_completude(rows)
    anomalies += check_promptitude(rows)
    anomalies += check_validite(rows)
    anomalies += check_unicite(rows)
    anomalies += check_coherence(rows, distribution)
    anomalies += check_fiabilite(rows)
    print(f"  {len(anomalies)} anomalies détectées")
    print_summary(anomalies)

    if dry_run:
        print("  Motifs les plus fréquents :")
        for (dimension, description), n in Counter((a.dimension, a.description) for a in anomalies).most_common(25):
            print(f"    {n:>4} | {dimension} | {description}")
        print("DRY_RUN : aucun accès SharePoint, aucun email, aucun log écrit.")
        return

    azure_tenant_id = os.environ["AZURE_TENANT_ID"].strip()
    azure_client_id = os.environ["AZURE_CLIENT_ID"].strip()
    azure_client_secret = os.environ["AZURE_CLIENT_SECRET"].strip()
    sharepoint_folder_link = folder_link_from_env("SHAREPOINT_FOLDER_LINK_ATELIER_REPARATION")
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
    send_confirmation_email(token, email_sender, email_recipients, "Clean Cooking — Atelier de réparation",
                            LOG_FILE_NAME, len(open_rows), new_count, resolved_count, deleted_count,
                            Counter(r.get("Dimension") for r in open_rows))
    print("Terminé.")


if __name__ == "__main__":
    main()
