"""Archivage de photos mWater dans SharePoint, partagé par les scripts d'export de photos.

Briques : lecture des identifiants de photos d'une cellule de datagrid, téléchargement d'une photo,
nettoyage des noms de dossiers/fichiers pour SharePoint, client Microsoft Graph avec renouvellement
du jeton et nouvelles tentatives, création de l'arborescence de dossiers, envoi de fichiers (petits ou
volumineux) et registre des photos déjà archivées (reprise après interruption, aucun doublon).

Tout est adressé PAR IDENTIFIANT (pas de syntaxe chemin `items/{id}:/{nom}:`) : voir la note de
common/sharepoint.py sur les erreurs 400 observées en authentification applicative.
"""

import json
import re
import threading
import time

import requests

from common.http_utils import raise_for_status_verbose
from common.sharepoint import GRAPH_API_BASE, graph_token

IMAGE_URL = "https://api.mwater.co/v3/images/{}"
IMAGE_RE = re.compile(r"https://api\.mwater\.co/v3/images/([0-9a-f]{32})")

SIMPLE_UPLOAD_MAX = 100_000_000      # en dessous : envoi simple (Graph accepte jusqu'à 250 Mo) ; au-delà : session d'envoi
CHUNK_SIZE = 320 * 1024 * 30         # ~9,4 Mo, multiple de 320 Kio comme l'exige Graph
TRANSIENT_STATUS = {429, 500, 502, 503, 504}
TOKEN_LIFETIME_SECONDS = 40 * 60     # renouvellement du jeton bien avant son expiration (~60 min)
MAX_RETRIES = 6
REGISTRY_NAME = "photos_archivees.json"

_INVALID = re.compile(r'[\\/:*?"<>|#%~&{}\x00-\x1f]')


def image_ids(cell):
    """Identifiants de photos d'une cellule de datagrid (adresses séparées), dans l'ordre, sans doublon."""
    vus, ordre = set(), []
    for identifiant in IMAGE_RE.findall(cell or ""):
        if identifiant not in vus:
            vus.add(identifiant)
            ordre.append(identifiant)
    return ordre


def clean_part(valeur, max_len=60, separateur=None):
    """Nom sûr pour SharePoint : sans caractères interdits, sans espaces ni points en bout.
    Avec `separateur`, les espaces internes sont remplacés par ce séparateur (noms de fichiers)."""
    texte = _INVALID.sub(" ", str(valeur or "")).strip()
    texte = re.sub(r"\s+", separateur if separateur else " ", texte).strip(" .")
    return texte[:max_len].strip(" .") or "sans-nom"


def download_image(image_id, attempts=5):
    """Télécharge une photo (adresse publique, sans authentification). Réessaie sur erreur transitoire."""
    derniere = None
    for tentative in range(1, attempts + 1):
        try:
            resp = requests.get(IMAGE_URL.format(image_id), timeout=120)
            if resp.status_code == 200 and resp.content:
                return resp.content
            derniere = f"HTTP {resp.status_code}"
            if resp.status_code in (400, 403, 404):
                break  # définitif : inutile de réessayer
        except requests.exceptions.RequestException as exc:
            derniere = type(exc).__name__
        time.sleep(min(2 ** tentative, 30))
    raise RuntimeError(f"photo {image_id} non téléchargée ({derniere})")


class GraphClient:
    """Requêtes Microsoft Graph avec jeton renouvelé automatiquement et nouvelles tentatives."""

    def __init__(self, tenant_id, client_id, client_secret):
        self._identifiants = (tenant_id, client_id, client_secret)
        self._verrou = threading.Lock()
        self._jeton, self._obtenu = None, 0.0

    def _token(self, force=False):
        with self._verrou:
            if force or self._jeton is None or time.time() - self._obtenu > TOKEN_LIFETIME_SECONDS:
                self._jeton = graph_token(*self._identifiants)
                self._obtenu = time.time()
            return self._jeton

    def request(self, method, url, authentifie=True, **kwargs):
        base_headers = dict(kwargs.pop("headers", {}) or {})
        delai = kwargs.pop("timeout", 120)
        tentative, jeton_force = 0, False
        while True:
            headers = dict(base_headers)
            if authentifie:
                headers["Authorization"] = f"Bearer {self._token(force=jeton_force)}"
            try:
                resp = requests.request(method, url, headers=headers, timeout=delai, **kwargs)
            except requests.exceptions.RequestException:
                if tentative >= MAX_RETRIES:
                    raise
                time.sleep(min(2 ** tentative, 60))
                tentative += 1
                continue
            if resp.status_code == 401 and authentifie and not jeton_force:
                jeton_force = True  # jeton expiré : on en redemande un et on rejoue
                continue
            if resp.status_code in TRANSIENT_STATUS and tentative < MAX_RETRIES:
                try:
                    attente = float(resp.headers.get("Retry-After", ""))
                except ValueError:
                    attente = min(2 ** tentative, 60)
                time.sleep(attente)
                tentative += 1
                continue
            return resp


class PhotoArchive:
    """Arborescence de dossiers SharePoint sous un dossier racine, et registre des photos archivées."""

    def __init__(self, graph, drive_id, root_id):
        self.graph, self.drive_id, self.root_id = graph, drive_id, root_id
        self._dossiers = {(): root_id}
        self._verrou = threading.Lock()

    # --- dossiers -----------------------------------------------------------------------------
    def _enfants(self, parent_id):
        url = f"{GRAPH_API_BASE}/drives/{self.drive_id}/items/{parent_id}/children?$top=200&$select=id,name,folder"
        while url:
            resp = self.graph.request("GET", url)
            raise_for_status_verbose(resp)
            donnees = resp.json()
            yield from donnees.get("value", [])
            url = donnees.get("@odata.nextLink")

    def folder(self, parties):
        """Identifiant du dossier racine/parties[0]/parties[1]/..., créé si besoin (mis en cache)."""
        parties = tuple(parties)
        with self._verrou:
            courant = ()
            for nom in parties:
                suivant = courant + (nom,)
                if suivant not in self._dossiers:
                    self._dossiers[suivant] = self._trouver_ou_creer(self._dossiers[courant], nom)
                courant = suivant
            return self._dossiers[courant]

    def _trouver_ou_creer(self, parent_id, nom):
        url = f"{GRAPH_API_BASE}/drives/{self.drive_id}/items/{parent_id}/children"
        resp = self.graph.request("POST", url, json={"name": nom, "folder": {},
                                                     "@microsoft.graph.conflictBehavior": "fail"})
        if resp.status_code in (200, 201):
            return resp.json()["id"]
        if resp.status_code == 409:  # existe déjà
            for enfant in self._enfants(parent_id):
                if enfant.get("name") == nom and enfant.get("folder") is not None:
                    return enfant["id"]
        raise_for_status_verbose(resp)
        raise RuntimeError(f"dossier '{nom}' introuvable après création")

    # --- fichiers -----------------------------------------------------------------------------
    def upload(self, folder_id, filename, contenu):
        """Dépose (ou remplace) un fichier dans un dossier, petit ou volumineux."""
        url = f"{GRAPH_API_BASE}/drives/{self.drive_id}/items/{folder_id}/children"
        resp = self.graph.request("POST", url, json={"name": filename, "file": {},
                                                     "@microsoft.graph.conflictBehavior": "replace"})
        raise_for_status_verbose(resp)
        item_id = resp.json()["id"]
        base = f"{GRAPH_API_BASE}/drives/{self.drive_id}/items/{item_id}"
        if len(contenu) <= SIMPLE_UPLOAD_MAX:
            resp = self.graph.request("PUT", f"{base}/content", data=contenu,
                                      headers={"Content-Type": "application/octet-stream"})
            raise_for_status_verbose(resp)
            return
        resp = self.graph.request("POST", f"{base}/createUploadSession",
                                  json={"item": {"@microsoft.graph.conflictBehavior": "replace"}})
        raise_for_status_verbose(resp)
        upload_url, total, debut = resp.json()["uploadUrl"], len(contenu), 0
        while debut < total:
            fin = min(debut + CHUNK_SIZE, total) - 1
            morceau = contenu[debut:fin + 1]
            resp = self.graph.request("PUT", upload_url, authentifie=False, data=morceau, headers={
                "Content-Length": str(len(morceau)), "Content-Range": f"bytes {debut}-{fin}/{total}"})
            raise_for_status_verbose(resp)
            debut = fin + 1

    # --- registre des photos déjà archivées ---------------------------------------------------
    def load_registry(self):
        """{identifiant de photo: chemin relatif}. Vide au premier passage."""
        for enfant in self._enfants(self.root_id):
            if enfant.get("name") == REGISTRY_NAME and enfant.get("folder") is None:
                resp = self.graph.request("GET", f"{GRAPH_API_BASE}/drives/{self.drive_id}/items/{enfant['id']}/content")
                raise_for_status_verbose(resp)
                try:
                    return dict(resp.json().get("photos", {}))
                except ValueError:
                    return {}
        return {}

    def save_registry(self, registre):
        contenu = json.dumps({"photos": registre}, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        self.upload(self.root_id, REGISTRY_NAME, contenu)


def archiver_taches(archive, registre, taches, workers, max_minutes, debut):
    """Télécharge puis dépose les photos par lots (en parallèle), en enregistrant le registre après
    chaque lot. S'arrête proprement après `max_minutes`. Une photo en échec n'est pas inscrite au
    registre : elle sera retentée au passage suivant. Retourne (réussies, échecs)."""
    from concurrent.futures import ThreadPoolExecutor

    reussies = echecs = 0
    taille = workers * 25

    def traiter(tache):
        contenu = download_image(tache.image_id)
        archive.upload(archive.folder(tache.parties), tache.nom_fichier, contenu)

    for debut_lot in range(0, len(taches), taille):
        if (time.time() - debut) / 60 > max_minutes:
            print(f"Durée maximale atteinte ({max_minutes:.0f} min) : arrêt propre, reprise au prochain passage.")
            break
        lot = taches[debut_lot:debut_lot + taille]
        with ThreadPoolExecutor(max_workers=workers) as pool:
            futures = [(t, pool.submit(traiter, t)) for t in lot]
            for t, futur in futures:
                try:
                    futur.result()
                    registre[t.image_id] = "/".join(t.parties) + "/" + t.nom_fichier
                    reussies += 1
                except Exception as exc:
                    echecs += 1
                    print(f"  échec photo {t.image_id[:8]}… : {type(exc).__name__}")
        archive.save_registry(registre)
        print(f"  {reussies} photos archivées, {echecs} échecs ({(time.time() - debut) / 60:.0f} min)")
    return reussies, echecs
