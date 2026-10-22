#!/usr/bin/env python3
"""
Vérification de données — Clean Cooking : Suivi post distribution (anomalies post distribution)
================================================================================================

Applique les six dimensions du "Manuel de vérification — Clean Cooking : Anomalies post
distribution" au datagrid mWater "raw_data_clean_cooking_anomalies" (formulaire
a9a5419c9ae34356ae9722a6ed85ef58), en croisant les QR codes cités avec le datagrid Distribution.
Met à jour un log Excel cumulatif (Nouveau / Toujours ouvert / Résolu / Supprimé) déposé sur
SharePoint, et envoie un email de confirmation. Avec DRY_RUN=true : décompte uniquement, sans
SharePoint ni email.

Lecture PAR POSITION : le datagrid a 279 colonnes dont beaucoup en double (8 colonnes "Status",
des QR Code, des Photo du PV...). Lire par nom écraserait la colonne de statut de la réponse
(colonne 3) par un statut de bloc ("Enregistré"/"Traité"). Les numéros ci-dessous sont les
positions 1-indexées du datagrid (profil du 22/10/2026) ; si le datagrid est modifié dans le
portail, le script s'arrête (contrôle des en-têtes de repère) plutôt que de lire de mauvaises colonnes.

Règles de bloc : un bloc (déménagement, retour, booklet, changement de foyer, vol/perte, vente,
transmission) n'est vérifié que si le formulaire l'a posé (colonne "(Asked)" du bloc = True), ce qui
reproduit "les champs d'un bloc ne sont exigés que si le type d'anomalie correspondant est coché".

Écarts assumés avec le guide (à valider) :
- Cohérence : seuls les QR "retourné / ancien foyer / volé-perdu / vendu / transmis" sont rapprochés
  de la Distribution (statut Final) ; les nouveaux QR d'un changement de foyer et les foyers
  retrouvés sont exclus. Un QR de format invalide est signalé en Validité seulement.
- Téléphone : mWater l'exporte sans le 0 initial (9 chiffres) ; il est normalisé avant le contrôle
  de la plage 0300000000-0389999999.
- Les anciennes valeurs de liste "???" ne sont pas des anomalies (aucun contrôle de valeur de liste).
"""

import csv
import io
import os
import re
from collections import Counter

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
from common.mwater_client import (
    download_datagrid_raw_with_retry,
    download_datagrid_with_retry,
    mwater_login,
)
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

DATAGRID_SUIVI = "e4a38892c65840ef8d1caf59b0fe6198"
DATAGRID_DISTRIBUTION = "ba07c16edc57417093d9735c93a97e52"

LOG_FILE_NAME = "data_verification_clean_cooking_suivi_post_distribution_log.xlsx"

NB_COLONNES = 279
# En-têtes de repère (position 1-indexée -> début de l'en-tête attendu) : garde-fou contre un
# changement de structure du datagrid dans le portail.
REPERES = {
    3: "Status", 4: "Response Code", 5: "Drafted On", 6: "Submitted On", 13: "Rejection message",
    17: "NUMERO DE SIGNALEMENT", 20: "Date de l'activité", 21: "ID MENAGE", 23: "NUMERO DE TELEPHONE",
    34: "Date du déménagement", 53: "Date du retour", 96: "Date de changement",
    114: "Date du changement de Foyer", 163: "Date de l'acte de vol", 190: "Date de la vente",
    238: "Date du retrait", 240: "Date de transmission",
}

QR_RE = re.compile(r"^[A-Z0-9]{6}$")
MOTIFS_DOUBLON = ("doublon", "doublant", "triplant", "existe déjà")

# Colonnes (1-indexées) : généralités
C_STATUS, C_CODE, C_DRAFTED, C_SUBMITTED, C_REJECTION = 3, 4, 5, 6, 13
C_SIGNAL_NB, C_SIGNAL_DEP, C_DATE_ACTIVITE, C_MENAGE, C_PHONE, C_PHONE_NA = 17, 18, 20, 21, 23, 25
C_ADMIN = (26, 27, 28, 29)
C_TRAITEMENT, C_TYPES = 30, 32


def col(r, n):
    return r[n - 1].strip() if n - 1 < len(r) else ""


def filled(r, n):
    return bool(col(r, n))


def is_true(r, n):
    return col(r, n).lower() == "true"


def date_de(r, n):
    d = parse_dt(col(r, n))
    return d.date() if d else None


def new(dimension, subdimension, r, description, details=""):
    return Anomaly(dimension, subdimension, col(r, C_CODE), col(r, C_MENAGE), description, details)


# Emplacements de QR code : (bloc, libellé, candidats [(colonne, colonne "non applicable" ou None)],
# rapproché de la Distribution ?). Un commentaire ne compte que si "non applicable" est coché.
QR_SLOTS = [
    ("foyer retourné", "foyer retourné 1", [(65, None), (66, 68)], True),
    ("foyer retourné", "foyer retourné 2", [(70, None), (71, 73)], True),
    ("foyer retourné", "foyer retourné 3", [(75, None), (76, 77)], True),
    ("ancien foyer", "ancien foyer changé 1", [(125, None), (126, 127), (129, None)], True),
    ("ancien foyer", "ancien foyer changé 2", [(131, None), (132, 133), (135, None)], True),
    ("nouveau foyer", "nouveau foyer 1", [(144, None), (145, 146)], False),
    ("nouveau foyer", "nouveau foyer 2", [(148, None), (149, 150)], False),
    ("foyer volé/perdu", "foyer volé/perdu 1", [(169, None)], True),
    ("foyer volé/perdu", "foyer volé/perdu 2", [(171, None)], True),
    ("foyer vendu", "foyer vendu 1", [(198, None)], True),
    ("foyer vendu", "foyer vendu 2", [(200, None)], True),
    ("foyer retrouvé", "foyer retrouvé 1", [(210, None)], False),
    ("foyer retrouvé", "foyer retrouvé 2", [(213, None)], False),
    ("foyer transmis", "foyer transmis 1", [(244, None)], True),
    ("foyer transmis", "foyer transmis 2", [(246, None)], True),
]


def slot_value(r, candidats):
    for c, na in candidats:
        if filled(r, c) and (na is None or is_true(r, na)):
            return col(r, c).rstrip(" \t")
    return ""


def verifier_en_tetes(header):
    if len(header) != NB_COLONNES:
        raise RuntimeError(f"Le datagrid Suivi post distribution a {len(header)} colonnes "
                           f"({NB_COLONNES} attendues) : vérifier sa configuration dans le portail.")
    for position, debut in REPERES.items():
        if not header[position - 1].strip().startswith(debut):
            raise RuntimeError(f"Colonne {position} : « {header[position - 1]} » au lieu de « {debut}… » : "
                               "la structure du datagrid a changé.")


# ---------------------------------------------------------------------------
# Complétude
# ---------------------------------------------------------------------------

def check_completude(rows):
    out = []
    for r in rows:
        statut = col(r, C_STATUS)
        if statut != "Final":
            out.append(new("Complétude", "Réponse non finalisée", r, "Réponse non finalisée",
                           f"Status : {statut or 'vide'}"))

        def exiger(sous_dimension, libelle, ok):
            if not ok:
                out.append(new("Complétude", sous_dimension, r, f"{libelle} non renseigné"))

        g = "Informations générales"
        exiger(g, "Numéro de signalement (nombre)", filled(r, C_SIGNAL_NB))
        exiger(g, "Numéro de signalement (département)", filled(r, C_SIGNAL_DEP))
        exiger(g, "Date de l'activité", filled(r, C_DATE_ACTIVITE))
        exiger(g, "Identifiant du ménage", filled(r, C_MENAGE))
        exiger(g, "Informations administratives", all(filled(r, c) for c in C_ADMIN))
        exiger(g, "Lieu de traitement", filled(r, C_TRAITEMENT))
        exiger(g, "Type d'anomalie", filled(r, C_TYPES))
        exiger(g, "Téléphone (« non applicable » non coché)", filled(r, C_PHONE) or is_true(r, C_PHONE_NA))

        def ok_ou_na(c, na):
            return filled(r, c) or is_true(r, na)

        def qr(libelles_slots):
            return all(slot_value(r, QR_SLOTS[i][2]) for i in libelles_slots)

        # --- Déménagement
        if is_true(r, 35):
            b = "Bloc déménagement"
            exiger(b, "Date du déménagement", filled(r, 34))
            exiger(b, "Nouvelle adresse", all(filled(r, c) for c in (36, 37, 38, 39)))
            exiger(b, "Adresse précédente", all(filled(r, c) for c in (42, 43, 44, 45)))
            exiger(b, "Observation sur le déménagement", ok_ou_na(47, 48))
            exiger(b, "Statut", filled(r, 50))

        # --- Foyer retourné
        if is_true(r, 54):
            b = "Bloc foyer retourné"
            exiger(b, "Date du retour", filled(r, 53))
            exiger(b, "Motif du retour", filled(r, 55))
            if col(r, 55).startswith("Autre"):
                exiger(b, "Précision du motif « Autre »", filled(r, 56))
            exiger(b, "Nombre de foyers retournés", filled(r, 61))
            exiger(b, "Type de foyer retourné", filled(r, 63))
            for asked_col, slot in ((69, 0), (74, 1), (78, 2)):
                if is_true(r, asked_col):
                    exiger(b, f"QR code du foyer retourné ({QR_SLOTS[slot][1]})", qr([slot]))
            exiger(b, "Photo du foyer retourné", ok_ou_na(79, 80))
            if is_true(r, 83):
                exiger(b, "Photo du second foyer retourné", filled(r, 82))
            exiger(b, "Photo du PV", ok_ou_na(84, 85))
            exiger(b, "Observations sur le foyer retourné", ok_ou_na(87, 88))
            exiger(b, "Statut", filled(r, 90))
            exiger(b, "Observations sur le statut", ok_ou_na(92, 93))

        # --- Changement de booklet
        if is_true(r, 97):
            b = "Bloc changement de booklet"
            exiger(b, "Date de changement", filled(r, 96))
            exiger(b, "Motif du changement", filled(r, 98))
            if col(r, 98).startswith(("Autre", "Other")):
                exiger(b, "Précision du motif « Autre »", filled(r, 99) or filled(r, 101))
            if col(r, 98) == "Carnet du foyer détérioré":
                exiger(b, "Photo de l'ancien booklet", ok_ou_na(103, 104))
            exiger(b, "Photo du nouveau booklet", ok_ou_na(106, 107))
            exiger(b, "Statut", filled(r, 109))
            exiger(b, "Observations sur le statut", ok_ou_na(111, 112))

        # --- Changement de foyer (ICS)
        if is_true(r, 115):
            b = "Bloc changement de foyer"
            exiger(b, "Date du changement", filled(r, 114))
            exiger(b, "Raison du changement", filled(r, 116) or filled(r, 118))
            exiger(b, "Nombre de foyers à changer", filled(r, 121))
            exiger(b, "Type de l'ancien foyer", filled(r, 123))
            exiger(b, "QR code de l'ancien foyer", qr([3]))
            if col(r, 121) == "2":
                exiger(b, "QR code du second ancien foyer", qr([4]))
            exiger(b, "Photo de l'ancien foyer", ok_ou_na(137, 138))
            exiger(b, "Type du nouveau foyer", filled(r, 140))
            exiger(b, "Nombre de nouveaux foyers", filled(r, 142))
            exiger(b, "QR code du nouveau foyer", qr([5]))
            if col(r, 142) == "2":
                exiger(b, "QR code du second nouveau foyer", qr([6]))
            exiger(b, "Photo du nouveau foyer", ok_ou_na(152, 153))
            exiger(b, "Observations sur le foyer remplacé", ok_ou_na(155, 156))
            exiger(b, "Statut", filled(r, 158))
            exiger(b, "Observations sur le statut", ok_ou_na(160, 161))

        # --- Vol ou perte
        if is_true(r, 164):
            b = "Bloc vol ou perte"
            exiger(b, "Date de l'acte", filled(r, 163))
            exiger(b, "Nombre de foyers volés/perdus", filled(r, 165))
            exiger(b, "Type de foyer", filled(r, 167))
            exiger(b, "QR code du foyer volé/perdu", qr([7]))
            if col(r, 165) == "2":
                exiger(b, "QR code du second foyer volé/perdu", qr([8]))
            exiger(b, "Description des circonstances", filled(r, 173))
            exiger(b, "Date de déclaration à MadAvance", filled(r, 175))
            exiger(b, "Foyer déjà retourné au ménage ?", filled(r, 181))
            if col(r, 181) == "Oui":
                exiger(b, "Date de retour au ménage", filled(r, 183))
            exiger(b, "Photos des pièces justificatives", ok_ou_na(185, 186))
            exiger(b, "Statut", filled(r, 188))

        # --- Vente
        if is_true(r, 191):
            b = "Bloc vente"
            exiger(b, "Date de la vente", filled(r, 190))
            exiger(b, "Responsable de la vente", filled(r, 192))
            exiger(b, "ID du ménage", filled(r, 194))
            exiger(b, "Nombre de foyers vendus", filled(r, 196))
            exiger(b, "QR code du foyer vendu", qr([9]))
            if col(r, 196) == "2":
                exiger(b, "QR code du second foyer vendu", qr([10]))
            exiger(b, "Le ménage a-t-il retrouvé le foyer ?", filled(r, 206))
            if col(r, 206) == "Oui":
                exiger(b, "Nombre de foyers retrouvés", filled(r, 208))
                exiger(b, "QR code du foyer retrouvé", qr([11]))
                if col(r, 208) == "2":
                    exiger(b, "QR code du second foyer retrouvé", qr([12]))
                exiger(b, "Photo du foyer retrouvé", ok_ou_na(216, 217))
            exiger(b, "Décision à l'égard du ménage", filled(r, 219) or filled(r, 222))
            if col(r, 219).startswith("Autre"):
                exiger(b, "Précision de la décision « Autre »", filled(r, 220))
            exiger(b, "Photo du PV", ok_ou_na(224, 225))
            exiger(b, "Date de fin de la dernière utilisation", filled(r, 227))
            exiger(b, "Date de fin d'adhésion du ménage", filled(r, 229))
            exiger(b, "Statut", filled(r, 231))
            exiger(b, "Observations sur le statut", ok_ou_na(233, 234))

        # --- Transmission vers un autre ménage
        if is_true(r, 239):
            b = "Bloc transmission"
            exiger(b, "ID du ménage récipiendaire", filled(r, 236))
            exiger(b, "Date du retrait", filled(r, 238))
            exiger(b, "Date de transmission", filled(r, 240))
            exiger(b, "Nombre de foyers", filled(r, 242))
            exiger(b, "QR code du foyer transmis", qr([13]))
            if col(r, 242) == "2":
                exiger(b, "QR code du second foyer transmis", qr([14]))
            exiger(b, "ID du ménage titulaire", filled(r, 248))
            exiger(b, "Nom du titulaire", filled(r, 250))
            exiger(b, "Informations administratives du titulaire", all(filled(r, c) for c in (252, 253, 254, 255)))
            exiger(b, "Photos du carnet précédent", ok_ou_na(257, 258))
            exiger(b, "Nom du ménage actuel", filled(r, 260))
            exiger(b, "Informations administratives du ménage actuel", all(filled(r, c) for c in (262, 263, 264, 265)))
            exiger(b, "Photos du carnet actuel", ok_ou_na(267, 268))
            exiger(b, "Motif de transmission", filled(r, 270))
            if col(r, 270).startswith("Autre"):
                exiger(b, "Précision du motif « Autre »", filled(r, 271))
            exiger(b, "Observations sur le ménage", ok_ou_na(275, 276))
    return out


# ---------------------------------------------------------------------------
# Promptitude
# ---------------------------------------------------------------------------

# (colonne de date, libellé) des événements de chaque bloc
DATES_EVENEMENTS = [
    (34, "déménagement"), (53, "retour du foyer"), (96, "changement de booklet"),
    (114, "changement de foyer"), (163, "vol ou perte"), (190, "vente"), (240, "transmission"),
]


def check_promptitude(rows):
    out = []
    for r in rows:
        drafted, submitted = parse_dt(col(r, C_DRAFTED)), parse_dt(col(r, C_SUBMITTED))
        activite = date_de(r, C_DATE_ACTIVITE)
        soumission = submitted.date() if submitted else None

        def anomalie(sous, description, details):
            out.append(new("Promptitude", sous, r, description, details))

        if drafted and submitted and submitted < drafted:
            anomalie("Chronologie Drafted/Submitted", "Submitted On antérieur à Drafted On",
                     f"Drafted On : {drafted} / Submitted On : {submitted}")
        if activite and soumission and activite > soumission:
            anomalie("Date de l'activité", "Date de l'activité postérieure à Submitted On",
                     f"Date de l'activité : {activite} / Submitted On : {soumission}")
        for c, libelle in DATES_EVENEMENTS:
            d = date_de(r, c)
            if d is None:
                continue
            if activite and d > activite:
                anomalie("Date d'événement", f"Date de {libelle} postérieure à la date de l'activité",
                         f"Date de {libelle} : {d} / Date de l'activité : {activite}")
            if soumission and d > soumission:
                anomalie("Date d'événement", f"Date de {libelle} postérieure à Submitted On",
                         f"Date de {libelle} : {d} / Submitted On : {soumission}")

        acte, declaration, retour_menage = date_de(r, 163), date_de(r, 175), date_de(r, 183)
        if acte and declaration and acte > declaration:
            anomalie("Vol ou perte", "Date de l'acte postérieure à la déclaration à MadAvance",
                     f"Acte : {acte} / Déclaration : {declaration}")
        if acte and retour_menage and retour_menage < acte:
            anomalie("Vol ou perte", "Date de retour du foyer au ménage antérieure à la date de l'acte",
                     f"Acte : {acte} / Retour au ménage : {retour_menage}")
        vente, fin_adhesion = date_de(r, 190), date_de(r, 229)
        if vente and fin_adhesion and fin_adhesion < vente:
            anomalie("Vente", "Date de fin d'adhésion antérieure à la date de la vente",
                     f"Vente : {vente} / Fin d'adhésion : {fin_adhesion}")
        retrait, transmission = date_de(r, 238), date_de(r, 240)
        if retrait and transmission and retrait > transmission:
            anomalie("Transmission", "Date du retrait postérieure à la date de transmission",
                     f"Retrait : {retrait} / Transmission : {transmission}")
    return out


# ---------------------------------------------------------------------------
# Validité
# ---------------------------------------------------------------------------

def telephone_valide(valeur):
    """mWater exporte le numéro sans le 0 initial : 9 chiffres -> on le rétablit. Plage attendue
    0300000000-0389999999."""
    if not valeur.isdigit():
        return False
    if len(valeur) == 9:
        valeur = "0" + valeur
    return len(valeur) == 10 and 300000000 <= int(valeur) <= 389999999


def check_validite(rows):
    out = []
    for r in rows:
        for _bloc, libelle, candidats, _rapproche in QR_SLOTS:
            valeur = slot_value(r, candidats)
            if valeur and not QR_RE.match(valeur):
                out.append(new("Validité", "Format QR code", r,
                               f"QR code invalide ({libelle} : 6 caractères alphanumériques en majuscules attendus)",
                               f"Valeur : '{valeur}'"))
        phone = col(r, C_PHONE)
        if phone and not telephone_valide(phone):
            out.append(new("Validité", "Numéro de téléphone", r,
                           "Numéro de téléphone hors de la plage 0300000000-0389999999", f"Valeur : '{phone}'"))
    return out


# ---------------------------------------------------------------------------
# Unicité
# ---------------------------------------------------------------------------

def check_unicite(rows):
    out = []
    codes = Counter(col(r, C_CODE) for r in rows if col(r, C_CODE))
    signalements = Counter((col(r, C_SIGNAL_NB), col(r, C_SIGNAL_DEP)) for r in rows
                           if col(r, C_SIGNAL_NB) and col(r, C_SIGNAL_DEP))
    for r in rows:
        code = col(r, C_CODE)
        if code and codes[code] > 1:
            out.append(new("Unicité", "Doublon Response Code", r, "Response Code dupliqué",
                           f"{codes[code]} réponses partagent ce code"))
        cle = (col(r, C_SIGNAL_NB), col(r, C_SIGNAL_DEP))
        if all(cle) and signalements[cle] > 1:
            out.append(new("Unicité", "Doublon numéro de signalement", r,
                           "Numéro de signalement utilisé pour plusieurs réponses",
                           f"Signalement {cle[0]} / {cle[1]} présent dans {signalements[cle]} réponses"))
        rejet = col(r, C_REJECTION).lower()
        if col(r, C_STATUS) == "Rejected" and any(m in rejet for m in MOTIFS_DOUBLON):
            out.append(new("Unicité", "Rejet pour doublon", r, "Rejet mWater pour doublon", col(r, C_REJECTION)))
    return out


# ---------------------------------------------------------------------------
# Cohérence : QR cités présents dans la Distribution
# ---------------------------------------------------------------------------

def qr_distribues(distribution_rows):
    qrs = set()
    for r in distribution_rows:
        if (r.get("Status") or "").strip() != "Final":
            continue
        for c in ("QR Code 1", "QR Code 2", "QR Code 3"):
            v = (r.get(c) or "").strip().upper()
            if v:
                qrs.add(v)
    return qrs


def check_coherence(rows, distribues):
    out = []
    for r in rows:
        for _bloc, libelle, candidats, rapproche in QR_SLOTS:
            if not rapproche:
                continue
            valeur = slot_value(r, candidats)
            # Un QR de format invalide est déjà signalé en Validité.
            if valeur and QR_RE.match(valeur) and valeur not in distribues:
                out.append(new("Cohérence", "QR code absent de la distribution", r,
                               f"QR code cité ({libelle}) absent de la distribution (statut Final)",
                               f"QR code : {valeur}"))
    return out


# ---------------------------------------------------------------------------
# Fiabilité
# ---------------------------------------------------------------------------

def check_fiabilite(rows):
    out = []
    for r in rows:
        # Un rejet n'est une anomalie que si la réponse est toujours au statut Rejected.
        if col(r, C_STATUS) == "Rejected":
            out.append(new("Fiabilité", "Rejet mWater non corrigé", r, "Réponse rejetée non corrigée",
                           col(r, C_REJECTION) or "Aucun message de rejet"))
    return out


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    dry_run = os.environ.get("DRY_RUN", "").strip().lower() in ("1", "true", "yes")

    print("Authentification mWater...")
    client_id = mwater_login(os.environ["MWATER_USERNAME"].strip(), os.environ["MWATER_PASSWORD"].strip())

    print("Téléchargement du datagrid Suivi post distribution...")
    texte = download_datagrid_raw_with_retry(DATAGRID_SUIVI, client_id, attempts=8, wait_seconds=60).decode("utf-8-sig")
    table = list(csv.reader(io.StringIO(texte)))
    verifier_en_tetes(table[0])
    rows = [r for r in table[1:] if any(c.strip() for c in r)]
    print(f"  {len(rows)} lignes")

    print("Téléchargement du datagrid Distribution (Cohérence des QR codes)...")
    distribution_rows = download_datagrid_with_retry(DATAGRID_DISTRIBUTION, client_id, attempts=8, wait_seconds=60)
    distribues = qr_distribues(distribution_rows)
    print(f"  {len(distribution_rows)} lignes, {len(distribues)} QR codes distribués (Final)")

    print("Application des règles de vérification (6 dimensions)...")
    anomalies = []
    anomalies += check_completude(rows)
    anomalies += check_promptitude(rows)
    anomalies += check_validite(rows)
    anomalies += check_unicite(rows)
    anomalies += check_coherence(rows, distribues)
    anomalies += check_fiabilite(rows)
    print(f"  {len(anomalies)} anomalies détectées")
    print_summary(anomalies)

    if dry_run:
        print("DRY_RUN : aucun accès SharePoint, aucun email, aucun log écrit.")
        return

    azure_tenant_id = os.environ["AZURE_TENANT_ID"].strip()
    azure_client_id = os.environ["AZURE_CLIENT_ID"].strip()
    azure_client_secret = os.environ["AZURE_CLIENT_SECRET"].strip()
    sharepoint_folder_link = folder_link_from_env("SHAREPOINT_FOLDER_LINK_SUIVI_POST_DISTRIBUTION")
    email_sender = os.environ["EMAIL_SENDER"].strip()
    email_recipients = os.environ["EMAIL_RECIPIENTS"].strip()

    print("Authentification Microsoft Graph...")
    token = graph_token(azure_tenant_id, azure_client_id, azure_client_secret)
    print("Résolution du lien SharePoint...")
    drive_id, folder_item_id = resolve_share_link(token, sharepoint_folder_link)

    print("Téléchargement du log existant sur SharePoint...")
    existing_rows = download_existing_log(token, drive_id, folder_item_id, LOG_FILE_NAME)
    print(f"  {len(existing_rows)} lignes déjà présentes dans le log")

    lignes_dict = [{"Response Code": col(r, C_CODE), "Submitted On": col(r, C_SUBMITTED)} for r in rows]
    merged_rows, new_count, resolved_count, deleted_count = merge_with_log(
        anomalies, existing_rows, today(),
        submitted_on_lookup=build_submitted_on_lookup(lignes_dict),
        existing_response_codes={col(r, C_CODE) for r in rows})
    open_rows = [r for r in merged_rows if r.get("Statut") not in STATUTS_CLOS]
    print(f"  {len(open_rows)} anomalies ouvertes ({new_count} nouvelles, {resolved_count} résolues, "
          f"{deleted_count} supprimées)")

    output_path = f"/tmp/{LOG_FILE_NAME}"
    build_log_report(merged_rows, output_path)
    print("Dépôt du log mis à jour sur SharePoint...")
    upload_to_sharepoint(token, drive_id, folder_item_id, output_path, LOG_FILE_NAME)

    print("Envoi de l'email de confirmation...")
    send_confirmation_email(token, email_sender, email_recipients, "Clean Cooking — Suivi post distribution",
                            LOG_FILE_NAME, len(open_rows), new_count, resolved_count, deleted_count,
                            Counter(r.get("Dimension") for r in open_rows))
    print("Terminé.")


if __name__ == "__main__":
    main()
