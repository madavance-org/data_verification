"""Résolution des fusions de points d'eau mWater (déduplication).

Extrait de premiere_rehabilitation/find_merged_water_points.py pour être réutilisable par
n'importe quel script d'activité qui référence un Water Point ID : un même point d'eau peut
avoir été fusionné (dédupliqué) dans une autre entité après la saisie d'une réponse, ce qui le
fait ressortir vide dans le datagrid sans que ce soit une vraie anomalie de Complétude.
"""

from common.mwater_client import fetch_entities_by_code, fetch_merge_targets


def resolve_water_point_codes(client_id, codes, entity_type="water_point"):
    """Pour une liste de codes de point d'eau, renvoie (direct, merged) :
    - direct : {code: entité} pour les codes qui existent encore tels quels
    - merged : {ancien_code: entité_cible} pour les codes fusionnés dans une autre entité

    Un code absent des deux dicts a été supprimé, ou n'a jamais existé."""
    codes = sorted(set(c for c in codes if c))
    direct = fetch_entities_by_code(client_id, entity_type, codes)
    still_missing = [c for c in codes if c not in direct]
    merged = fetch_merge_targets(client_id, entity_type, still_missing)
    return direct, merged
