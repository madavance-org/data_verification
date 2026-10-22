#!/usr/bin/env python3
"""
Archivage des photos — Clean Cooking : Inscription
===================================================

Extrait de mWater les photos des inscriptions (datagrid Inscription, réponses au statut Final) et les dépose
dans SharePoint, dans le dossier `photos` de l'activité, classées par localisation puis par ménage :

    région / district / commune / fokontany / village / nom du représentant du ménage / fichier

Nom du fichier : Inscription _ nom du représentant _ champ photo _ date de saisie _ numéro.jpg
Le numéro distingue les photos d'un même champ et garantit qu'un fichier n'en écrase jamais un autre.

Champs photo : maison, carte d'identité, Annexe A (fiche d'inscription), Annexe B (accord carbone),
Annexe D (consentement), Annexe G (reçu de paiement).

Règles de classement :
- Région, district, commune, fokontany : colonnes « Adresse du ménage » du datagrid. S'il en manque un, ou
  si le nom du représentant est vide, la photo va dans `non_classé/<nom ou code de la réponse>`.
- Village : nom lu sur la fiche du village (entité mWater « community ») à partir du code village. Sans code
  village (~7 % des inscriptions) ou sans nom : sous-dossier `village-inconnu`, les autres niveaux étant conservés.
- Homonymes : le dossier du ménage porte le nom seul ; si deux ménages (ID différents) ont le même nom dans le
  même village, l'ID du ménage est ajouté aux deux (`nom-ID`) pour ne pas mélanger leurs photos.

Incrémental et reprenable (registre `photos_archivees.json`, comme pour la Première réhabilitation) ; arrêt
propre après MAX_RUNTIME_MINUTES. Variables : DRY_RUN, MAX_PHOTOS, MAX_RUNTIME_MINUTES, PHOTO_WORKERS et
SHAREPOINT_FOLDER_LINK_INSCRIPTION_PHOTOS (obligatoire hors DRY_RUN).

Aucune photo ni donnée personnelle (noms) n'est affichée dans les journaux ni stockée sur GitHub.
"""

import csv
import io
import os
import time
from collections import Counter, defaultdict, namedtuple

from common.mwater_client import download_datagrid_raw_with_retry, fetch_entities_by_code, mwater_login
from common.photos import GraphClient, PhotoArchive, archiver_taches, clean_part, image_ids
from common.sharepoint import graph_token, resolve_share_link

DATAGRID_INSCRIPTION = "ac3ae2ce1f7540f586ae1ab5673419a8"

PREFIXE_FICHIER = "Inscription"
NON_CLASSE = "non_classé"
VILLAGE_INCONNU = "village-inconnu"

C_STATUS, C_CODE, C_DRAFTED = "Status", "Response Code", "Drafted On"
C_MENAGE = "Household ID code > Unique ID"
C_VILLAGE = "Code village > Unique ID"
C_NOM = "Name of household representative"
C_ADRESSE = [f"Adresse du ménage > {n}" for n in ("REGION", "DISTRICT", "COMMUNE", "FOKONTANY")]

CHAMPS_PHOTO = {
    "Photo de la maison": "maison",
    "Photo carte d'identité nationale": "carte-identite",
    "Annexe A : Fiche d'inscription des membres": "fiche-inscription",
    "Annexe B : Accord Carbone": "accord-carbone",
    "Annexe D : Formulaire de consentement pour la publication de données et de photos": "consentement",
    "Annexe G : Reçu de paiement": "recu-paiement",
}

Tache = namedtuple("Tache", "image_id parties nom_fichier champ")


def lire_datagrid(client_id):
    texte = download_datagrid_raw_with_retry(DATAGRID_INSCRIPTION, client_id, timeout=900, attempts=8,
                                             wait_seconds=60).decode("utf-8-sig")
    lecteur = csv.reader(io.StringIO(texte))
    header = next(lecteur)
    doublons = [h for h, n in Counter(header).items() if n > 1]
    if doublons:
        raise RuntimeError(f"Colonnes en double dans le datagrid Inscription ({doublons}) : lecture par nom non fiable.")
    return [dict(zip(header, ligne)) for ligne in lecteur if any(c.strip() for c in ligne)]


def val(ligne, colonne):
    return (ligne.get(colonne) or "").strip()


def construire_taches(lignes, noms_villages, registre):
    retenues = [r for r in lignes if val(r, C_STATUS) == "Final"]
    retenues.sort(key=lambda r: (val(r, C_DRAFTED)[:10], val(r, C_CODE)))

    # Homonymes : même nom dans un même village (même localisation) avec des ID de ménage différents
    ids_par_nom = defaultdict(set)
    for r in retenues:
        cle = (tuple(val(r, c) for c in C_ADRESSE), val(r, C_VILLAGE), val(r, C_NOM).lower())
        ids_par_nom[cle].add(val(r, C_MENAGE))

    utilises = set(registre.values())
    taches, stats = [], Counter()
    for r in retenues:
        menage, code_reponse = val(r, C_MENAGE), val(r, C_CODE)
        nom = val(r, C_NOM)
        cle = (tuple(val(r, c) for c in C_ADRESSE), val(r, C_VILLAGE), nom.lower())
        homonyme = len(ids_par_nom[cle]) > 1 and bool(menage)
        dossier_menage = clean_part(nom or code_reponse, 60)
        fichier_menage = clean_part(nom or code_reponse, 40, "-")
        if homonyme:
            dossier_menage = f"{dossier_menage}-{clean_part(menage, 15)}"
            fichier_menage = f"{fichier_menage}-{clean_part(menage, 15)}"
            stats["ménages homonymes"] += 1

        niveaux = [val(r, c) for c in C_ADRESSE]
        if nom and all(niveaux):
            village = noms_villages.get(val(r, C_VILLAGE)) or ""
            parties = [clean_part(n) for n in niveaux] + [clean_part(village) if village else VILLAGE_INCONNU, dossier_menage]
            stats["classées"] += 1
            stats["village inconnu"] += 0 if village else 1
        else:
            parties = [NON_CLASSE, dossier_menage]
            stats["non classées"] += 1

        date = val(r, C_DRAFTED)[:10] or "sans-date"
        for colonne, libelle in CHAMPS_PHOTO.items():
            for image_id in image_ids(val(r, colonne)):
                if image_id in registre:
                    stats["photos déjà archivées"] += 1
                    continue
                prefixe = "_".join([PREFIXE_FICHIER, fichier_menage, libelle, date])
                dossier = "/".join(parties)
                n = 1
                while f"{dossier}/{prefixe}_{n}.jpg" in utilises:
                    n += 1
                utilises.add(f"{dossier}/{prefixe}_{n}.jpg")
                taches.append(Tache(image_id, parties, f"{prefixe}_{n}.jpg", libelle))
    return taches, stats, len(retenues)


def main():
    dry_run = os.environ.get("DRY_RUN", "").strip().lower() in ("1", "true", "yes")
    max_photos = int(os.environ.get("MAX_PHOTOS") or 0)
    max_minutes = float(os.environ.get("MAX_RUNTIME_MINUTES") or 300)
    workers = int(os.environ.get("PHOTO_WORKERS") or 8)
    debut = time.time()

    print("Authentification mWater...")
    client_id = mwater_login(os.environ["MWATER_USERNAME"].strip(), os.environ["MWATER_PASSWORD"].strip())
    print("Téléchargement du datagrid Inscription...")
    lignes = lire_datagrid(client_id)
    print(f"  {len(lignes)} lignes")

    archive, registre = None, {}
    if not dry_run:
        tenant, cid, secret = (os.environ[k].strip() for k in ("AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET"))
        lien = os.environ["SHAREPOINT_FOLDER_LINK_INSCRIPTION_PHOTOS"].strip()
        drive_id, racine_id = resolve_share_link(graph_token(tenant, cid, secret), lien)
        archive = PhotoArchive(GraphClient(tenant, cid, secret), drive_id, racine_id)
        registre = archive.load_registry()
        print(f"  {len(registre)} photos déjà archivées (registre)")

    codes = {val(r, C_VILLAGE) for r in lignes if val(r, C_STATUS) == "Final"} - {""}
    print(f"Lecture des fiches de {len(codes)} villages...")
    villages = fetch_entities_by_code(client_id, "community", codes)
    noms_villages = {code: (e.get("name") or "") for code, e in villages.items()}
    print(f"  {len(villages)} fiches trouvées")

    taches, stats, n_lignes = construire_taches(lignes, noms_villages, registre)
    print(f"{n_lignes} inscriptions Final, {len(taches)} photos à archiver ({dict(stats)})")
    if max_photos:
        taches = taches[:max_photos]
        print(f"  limité à {len(taches)} photos (MAX_PHOTOS)")

    if dry_run:
        print("Photos par champ :", dict(Counter(t.champ for t in taches).most_common()))
        print("Dossiers distincts :", len({tuple(t.parties) for t in taches}),
              "| profondeur max :", max((len(t.parties) for t in taches), default=0),
              "| nom de fichier le plus long :", max((len(t.nom_fichier) for t in taches), default=0),
              "| chemin le plus long :", max((len("/".join(t.parties)) + len(t.nom_fichier) + 1 for t in taches), default=0))
        print("DRY_RUN : rien n'a été téléchargé ni déposé.")
        return

    reussies, echecs = archiver_taches(archive, registre, taches, workers, max_minutes, debut)
    print(f"Terminé : {reussies} photos archivées, {echecs} échecs, {len(registre)} photos au registre.")


if __name__ == "__main__":
    main()
