#!/usr/bin/env python3
"""
Archivage des photos — Clean Water : Première réhabilitation
============================================================

Extrait de mWater les photos des réhabilitations (datagrid « Réhabilitation », lignes
« Type de travaux = Première réhabilitation » au statut Final) et les dépose dans SharePoint, dans
le dossier `photos` de l'activité, classées par localisation puis par point d'eau :

    région / district / commune / fokontany / village / ID du point d'eau / fichier

Nom du fichier : activité _ ID du point d'eau _ custom ID _ champ photo _ date des travaux _ numéro.jpg
(le « custom ID » est l'`alt_id` de la fiche du point d'eau dans mWater ; le numéro distingue les photos
d'un même champ et garantit qu'un fichier n'en écrase jamais un autre).

Localisation : lue sur la fiche du point d'eau (niveaux administratifs 1 à 5), à défaut dans les colonnes
Région / District / Commune / Fokontany / Village du datagrid. Une photo sans point d'eau, ou dont un des
cinq niveaux est introuvable, est rangée dans `non_classé/<ID du point d'eau ou code de la réponse>`.

Incrémental et reprenable : un registre `photos_archivees.json` (dans le dossier `photos`) liste les
photos déjà archivées ; une photo n'est jamais retéléchargée ni redéposée. Le script s'arrête de lui-même
après MAX_RUNTIME_MINUTES (défaut 300, sous la limite de 6 h d'un job GitHub) et reprend au passage suivant.

Variables d'environnement :
    DRY_RUN=true            plan et décomptes uniquement : rien n'est téléchargé ni déposé
    MAX_PHOTOS=n            limite le nombre de photos traitées (premiers essais)
    MAX_RUNTIME_MINUTES=n   durée maximale de traitement (défaut 300)
    PHOTO_WORKERS=n         téléchargements/dépôts en parallèle (défaut 8)
    SHAREPOINT_FOLDER_LINK_PREMIERE_REHABILITATION_PHOTOS   lien du dossier `photos` (obligatoire hors DRY_RUN)

Aucune photo ni donnée personnelle n'est affichée dans les journaux ni stockée sur GitHub.
"""

import csv
import io
import os
import time
from collections import Counter, namedtuple
from concurrent.futures import ThreadPoolExecutor

from common.mwater_client import download_datagrid_raw_with_retry, fetch_entities_by_code, mwater_login
from common.photos import GraphClient, PhotoArchive, clean_part, download_image, image_ids
from common.sharepoint import graph_token, resolve_share_link

DATAGRID_REHAB = "0638d32971704164b9ac22d549d4818e"

PREFIXE_FICHIER = "Premiere-rehabilitation"
TYPE_TRAVAUX = "Première réhabilitation"
NON_CLASSE = "non_classé"

C_TYPE, C_STATUS, C_CODE = "Type de travaux", "Status", "Response Code"
C_WP = "De quel point d'eau s'agit-il? > Unique ID"
C_DATE_TRAVAUX = "Date de début des travaux (ou échantillonnage)"
C_SUBMITTED = "Submitted On"
NIVEAUX_GRILLE = ["Région", "District", "Commune", "Fokontany", "Village"]
NIVEAUX_FICHE = ["admin_div1", "admin_div2", "admin_div3", "admin_div4", "admin_div5"]

# Libellés courts des champs photo dans les noms de fichiers : (morceau du titre du champ, libellé)
LIBELLES_CHAMPS = [
    ("problème", "avant-travaux"),
    ("activité de réhabilitation", "pendant-travaux"),
    ("travaux réalisés", "apres-travaux"),
    ("non fonctionnement", "preuve-panne"),
    ("représentant local", "representant-local"),
    ("eau coulant", "eau-pompe-reparee"),
]


def libelle_champ(titre):
    minuscules = titre.lower()
    for morceau, libelle in LIBELLES_CHAMPS:
        if morceau in minuscules:
            return libelle
    return clean_part(titre, 45, "-")  # champ inconnu : titre nettoyé


Tache = namedtuple("Tache", "image_id parties nom_fichier champ")


def lire_datagrid(client_id):
    texte = download_datagrid_raw_with_retry(DATAGRID_REHAB, client_id, timeout=600, attempts=8,
                                             wait_seconds=60).decode("utf-8-sig")
    table = list(csv.reader(io.StringIO(texte)))
    return table[0], table[1:]


def premiere_position(header, nom):
    return header.index(nom) if nom in header else None


def valeur(ligne, position):
    return ligne[position].strip() if position is not None and position < len(ligne) else ""


def construire_taches(header, lignes, entites, registre):
    """Une tâche par photo à archiver : (identifiant, dossier, préfixe du nom). Les noms sont
    numérotés en tenant compte des chemins déjà pris (registre) pour ne jamais en écraser un."""
    i_type, i_status = premiere_position(header, C_TYPE), premiere_position(header, C_STATUS)
    i_code, i_wp = premiere_position(header, C_CODE), premiere_position(header, C_WP)
    i_date, i_sub = premiere_position(header, C_DATE_TRAVAUX), premiere_position(header, C_SUBMITTED)
    i_grille = [premiere_position(header, n) for n in NIVEAUX_GRILLE]

    retenues = [l for l in lignes
                if valeur(l, i_type) == TYPE_TRAVAUX and valeur(l, i_status) == "Final"]

    # Colonnes contenant des photos pour ces lignes (par position : le datagrid a des colonnes en double)
    colonnes_photo = [i for i in range(len(header)) if any(image_ids(valeur(l, i)) for l in retenues)]

    retenues.sort(key=lambda l: ((valeur(l, i_date) or valeur(l, i_sub))[:10], valeur(l, i_code)))
    utilises = set(registre.values())
    taches, stats = [], Counter()
    for l in retenues:
        wp, code_reponse = valeur(l, i_wp), valeur(l, i_code)
        date = (valeur(l, i_date) or valeur(l, i_sub))[:10] or "sans-date"
        fiche = entites.get(wp) or {}
        niveaux = [str(fiche.get(nf) or valeur(l, ig) or "").strip()
                   for nf, ig in zip(NIVEAUX_FICHE, i_grille)]
        if wp and all(niveaux):
            parties = [clean_part(n) for n in niveaux] + [clean_part(wp)]
            stats["classées"] += 1
        else:
            parties = [NON_CLASSE, clean_part(wp or code_reponse)]
            stats["non classées"] += 1
        custom_id = clean_part(fiche.get("alt_id") or "sans-id", 30, "-")
        for i in colonnes_photo:
            champ = libelle_champ(header[i])
            for image_id in image_ids(valeur(l, i)):
                if image_id in registre:
                    stats["déjà archivées"] += 1
                    continue
                prefixe = "_".join([PREFIXE_FICHIER, clean_part(wp or code_reponse, 30, "-"), custom_id, champ, date])
                dossier = "/".join(parties)
                n = 1
                while f"{dossier}/{prefixe}_{n}.jpg" in utilises:
                    n += 1
                utilises.add(f"{dossier}/{prefixe}_{n}.jpg")
                taches.append(Tache(image_id, parties, f"{prefixe}_{n}.jpg", champ))
    return taches, stats, len(retenues), len(colonnes_photo)


def main():
    dry_run = os.environ.get("DRY_RUN", "").strip().lower() in ("1", "true", "yes")
    max_photos = int(os.environ.get("MAX_PHOTOS") or 0)
    max_minutes = float(os.environ.get("MAX_RUNTIME_MINUTES") or 300)
    workers = int(os.environ.get("PHOTO_WORKERS") or 8)
    debut = time.time()

    print("Authentification mWater...")
    client_id = mwater_login(os.environ["MWATER_USERNAME"].strip(), os.environ["MWATER_PASSWORD"].strip())
    print("Téléchargement du datagrid Réhabilitation...")
    header, lignes = lire_datagrid(client_id)
    print(f"  {len(lignes)} lignes")

    archive, registre = None, {}
    if not dry_run:
        tenant, cid, secret = (os.environ[k].strip() for k in ("AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET"))
        lien = os.environ["SHAREPOINT_FOLDER_LINK_PREMIERE_REHABILITATION_PHOTOS"].strip()
        drive_id, racine_id = resolve_share_link(graph_token(tenant, cid, secret), lien)
        archive = PhotoArchive(GraphClient(tenant, cid, secret), drive_id, racine_id)
        registre = archive.load_registry()
        print(f"  {len(registre)} photos déjà archivées (registre)")

    i_type, i_status, i_wp = (premiere_position(header, n) for n in (C_TYPE, C_STATUS, C_WP))
    codes = {valeur(l, i_wp) for l in lignes if valeur(l, i_type) == TYPE_TRAVAUX and valeur(l, i_status) == "Final"}
    print(f"Lecture des fiches de {len(codes - {''})} points d'eau...")
    entites = fetch_entities_by_code(client_id, "water_point", codes - {""})
    print(f"  {len(entites)} fiches trouvées")

    taches, stats, n_lignes, n_colonnes = construire_taches(header, lignes, entites, registre)
    print(f"{n_lignes} réhabilitations Final, {n_colonnes} colonnes photo, {len(taches)} photos à archiver "
          f"({dict(stats)})")
    if max_photos:
        taches = taches[:max_photos]
        print(f"  limité à {len(taches)} photos (MAX_PHOTOS)")

    if dry_run:
        print("Photos par champ :", dict(Counter(t.champ for t in taches).most_common()))
        print("Dossiers distincts :", len({tuple(t.parties) for t in taches}),
              "| profondeur max :", max((len(t.parties) for t in taches), default=0),
              "| nom de fichier le plus long :", max((len(t.nom_fichier) for t in taches), default=0),
              "| chemin le plus long :", max((len("/".join(t.parties)) + len(t.nom_fichier) + 1 for t in taches), default=0))
        for t in taches[:3]:
            print("  exemple :", "/".join(t.parties) + "/" + t.nom_fichier)
        print("DRY_RUN : rien n'a été téléchargé ni déposé.")
        return

    reussies = echecs = 0
    taille = workers * 25
    for debut_lot in range(0, len(taches), taille):
        if (time.time() - debut) / 60 > max_minutes:
            print(f"Durée maximale atteinte ({max_minutes:.0f} min) : arrêt propre, reprise au prochain passage.")
            break
        lot = taches[debut_lot:debut_lot + taille]

        def traiter(t):
            contenu = download_image(t.image_id)
            archive.upload(archive.folder(t.parties), t.nom_fichier, contenu)
            return t

        with ThreadPoolExecutor(max_workers=workers) as pool:
            futures = [(t, pool.submit(traiter, t)) for t in lot]
            for t, futur in futures:
                try:
                    futur.result()
                    registre[t.image_id] = "/".join(t.parties) + "/" + t.nom_fichier
                    reussies += 1
                except Exception as exc:  # une photo en échec sera retentée au prochain passage
                    echecs += 1
                    print(f"  échec photo {t.image_id[:8]}… : {type(exc).__name__}")
        archive.save_registry(registre)
        print(f"  {reussies} photos archivées, {echecs} échecs ({(time.time() - debut) / 60:.0f} min)")

    print(f"Terminé : {reussies} photos archivées, {echecs} échecs, {len(registre)} photos au registre.")


if __name__ == "__main__":
    main()
