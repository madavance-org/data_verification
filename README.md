# data_verification

Vérification automatisée des données mWater pour l'activité **Appel maintenance préventive**, sur la base des six dimensions du Manuel de vérification de données MadAvance (Complétude, Promptitude, Validité, Unicité, Cohérence, Fiabilité).

Le script `verify_maintenance_preventive.py` :

1. S'authentifie sur l'API mWater et télécharge 4 datagrids déjà configurés dans le portail (Appel maintenance préventive, Maintenance préventive, Réparation après panne, Première réhabilitation).
2. Applique les règles de vérification (voir le détail dans le manuel ClickUp lié).
3. Télécharge le log existant sur SharePoint (`data_verification_call_log.xlsx`) et le fusionne avec les anomalies détectées aujourd'hui.
4. Réenregistre ce même fichier (nom fixe, pas de date dans le nom) sur SharePoint via Microsoft Graph.
5. Envoie un email de confirmation avec le décompte d'anomalies ouvertes par dimension, plus le nombre de nouvelles et de résolues.

Exécution automatique hebdomadaire via GitHub Actions (lundi 06:30 UTC), ou manuelle via `workflow_dispatch`. Ne modifie jamais les données mWater — uniquement de la détection/reporting. La correction reste une étape manuelle séparée.

### Le fichier de log (`data_verification_call_log.xlsx`)

Contrairement à un rapport horodaté recréé à chaque exécution, c'est **le même fichier** qui est mis à jour à chaque run — il sert d'historique cumulé. Chaque ligne du log a un identifiant stable (Dimension + Sous-dimension + Response Code + Description) qui permet de reconnaître une anomalie d'une exécution à l'autre, et un `Statut` :

- **Nouveau** : détectée pour la première fois.
- **Toujours ouvert** : déjà vue lors d'une exécution précédente, toujours présente.
- **Résolu** : présente avant, plus détectée aujourd'hui — considérée comme corrigée. La ligne reste dans le fichier (historique), mais n'est plus mise à jour.

Colonnes : `Dimension`, `Sous-dimension`, `Response Code`, `Water Point ID`, `Description`, `Détails`, `Statut`, `Première détection`, `Dernière détection`, `Date de résolution`. L'onglet "Résumé" ne compte que les anomalies actuellement ouvertes (Nouveau + Toujours ouvert), pas les résolues.

Si le fichier n'existe pas encore sur SharePoint (première exécution), le script part d'un log vide et toutes les anomalies détectées sont marquées "Nouveau".

## Secrets GitHub Actions requis (Settings > Secrets and variables > Actions)

| Secret | Description |
| --- | --- |
| `MWATER_USERNAME` / `MWATER_PASSWORD` | Identifiants mWater |
| `AZURE_TENANT_ID` | Même App Registration que le repo `mWater_backup` ("BackupOffice365") |
| `AZURE_CLIENT_ID` | idem |
| `AZURE_CLIENT_SECRET` | idem — généré dans Entra ID, valeur connue uniquement dans GitHub Secrets |
| `SHAREPOINT_FOLDER_LINK` | Lien de partage SharePoint du dossier cible, collé tel quel (ex. `https://madavancengo.sharepoint.com/:f:/s/ITMadAvance/...`). **Ne pas** essayer de le convertir en `driveId`/`itemId` à la main — le script le résout lui-même à chaque exécution via `/shares/{shareId}/driveItem`. |
| `EMAIL_SENDER` | Boîte d'envoi (ex. `it@madavance.org`) |
| `EMAIL_RECIPIENTS` | Destinataire(s), séparés par des virgules |

Ces secrets ne sont pas partagés automatiquement entre repos GitHub : même si `AZURE_TENANT_ID`/`AZURE_CLIENT_ID`/`AZURE_CLIENT_SECRET` existent déjà dans `mWater_backup`, il faut les recopier dans les secrets de **ce** repo.

### `SHAREPOINT_FOLDER_LINK`

Historique : on utilisait auparavant deux secrets séparés (`SHAREPOINT_DRIVE_ID` / `SHAREPOINT_FOLDER_ITEM_ID`), à résoudre manuellement à partir du lien de partage. Source de bugs (mauvaise valeur collée par erreur, un lien au lieu d'un ID). Le script résout maintenant directement le lien de partage au démarrage — un seul secret à renseigner, moins d'erreur possible.

## Datagrids mWater utilisés (IDs codés en dur dans le script, pas des secrets)

| Activité | Formulaire | Datagrid |
| --- | --- | --- |
| Appel maintenance préventive | `c08b3fe26d0f42c084074701f29eb75e` | `5f443b8a8c144502a304c8c5c24d4f82` |
| Maintenance préventive | `de26d89a5c8a4452b42158c622be20d0` | `2cfe0ba7ac264d119cfc8964b5f3cebc` |
| Réparation après panne | `958b4763788348d699e7d8c5821f92ee` | `d9f1c36a2d6340429658b6628fe81b88` |
| Première réhabilitation (datagrid mixte, filtré dans le script) | `86cf66efdd3749dd8a121314bab3675a` | `0638d32971704164b9ac22d549d4818e` |

## Règles appliquées

- **Complétude** : `Status = Draft` (brouillon jamais soumis) ou `Status = Final` sans Water Point ID renseigné.
- **Promptitude** : `Drafted On <= Submitted On`.
- **Validité** : `Signal reference` doit respecter `{DEPLOYMENT}_{JJMMAAAA}_{E|S}{N}` ; Water Point ID doit être numérique.
- **Unicité** : doublons de `Signal reference`, ou rejet mWater contenant "doublon" dans `Rejection message`.
- **Cohérence** : préfixe de déploiement cohérent ; si pompe "Partially" → code présent dans Maintenance préventive ; si "No" → présent dans Réparation après panne ; date embarquée dans le code ≤ `Completion date of the work` du formulaire correspondant.
- **Fiabilité** : rejets contenant "ID incorrect" croisés avec les Water Point ID des Premières réhabilitations réussies (`Type de travaux = "Première réhabilitation"` et `Status = Final`) dans le datagrid Réhabilitation ; si le WP ID est trouvé ailleurs dans ce même datagrid (autre type/statut), le contexte est indiqué pour réexamen.

## Test en local

```bash
pip install -r requirements.txt
export MWATER_USERNAME=... MWATER_PASSWORD=...
export AZURE_TENANT_ID=... AZURE_CLIENT_ID=... AZURE_CLIENT_SECRET=...
export SHAREPOINT_DRIVE_ID=... SHAREPOINT_FOLDER_ITEM_ID=...
export EMAIL_SENDER=... EMAIL_RECIPIENTS=...
python verify_maintenance_preventive.py
```

## Test de validation (logique seule, sans upload)

Testé le 21/07/2026 sur un export réel des 4 datagrids (3073 réponses Appel, 732 Maintenance, 171 Réparation, 1688 Réhabilitation), après plusieurs corrections faites avec Lanja en cours de route (distinction site supprimé/Don't Know en Complétude, exclusion des rejets déjà corrigés en Unicité, etc.) : 299 anomalies détectées (Cohérence 186, Unicité 50, Promptitude 29, Complétude 33, Fiabilité 1). **À valider avec Lanja avant mise en production** : passer en revue un échantillon des anomalies de Cohérence "absent du formulaire" pour confirmer qu'il ne s'agit pas de faux positifs (ex. suivi encore en attente plutôt qu'anomalie réelle).

La logique de log (Nouveau / Toujours ouvert / Résolu) a été testée par simulation de deux exécutions successives : les statuts et dates de première/dernière détection et de résolution se comportent comme attendu. Pas encore testée en conditions réelles sur deux exécutions GitHub Actions successives.
