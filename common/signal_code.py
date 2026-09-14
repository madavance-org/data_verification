"""Format et correction du Signal code, partagés par les scripts qui le valident.

Format attendu : {DEPLOYMENT}_{JJMMAAAA}_{E|S}{N} (ex. MAR_21072026_E14).
"""

import re
from datetime import datetime

SIGNAL_CODE_RE = re.compile(r"^([A-Z]+)_(\d{2})(\d{2})(\d{4})_([ES])(\d+)$")

# Format "propre" mais avec séparateurs/ordre incorrects : ex. "MAR-08052026-E1",
# "MAR _09072026_E2", "MAR_26/02/2026_S14", "FTU_28042026E1"
LOOSE_RE = re.compile(
    r"^\s*([A-Za-z]+)[\s_\-]*(\d{2})[\s/\-]?(\d{2})[\s/\-]?(\d{4})"
    r"[\s_\-]*([EeSs])[\s_\-]*(\d+)\.?\s*$"
)

# Même chose mais avec E/S et numéro inversés : ex. "MAR_11022025_1E"
SWAPPED_RE = re.compile(
    r"^\s*([A-Za-z]+)[\s_\-]*(\d{2})[\s/\-]?(\d{2})[\s/\-]?(\d{4})"
    r"[\s_\-]*(\d+)[\s_\-]*([EeSs])\.?\s*$"
)


def propose_correction(raw):
    """Tente de reconstruire un Signal code valide à partir d'une valeur mal formatée.
    Retourne (valeur_corrigee_ou_None, note). None si le code est trop ambigu/dégradé
    pour une correction automatique fiable (nécessite une vérification manuelle)."""

    m = LOOSE_RE.match(raw)
    if not m:
        m = SWAPPED_RE.match(raw)
        if m:
            prefix, dd, mm, yyyy, num, es = m.groups()
        else:
            return None, "Format non reconnu automatiquement"
    else:
        prefix, dd, mm, yyyy, es, num = m.groups()

    # Sanité de la date avant de proposer une correction
    try:
        datetime(int(yyyy), int(mm), int(dd))
    except ValueError:
        return None, "Date invalide dans le code (jour/mois incohérents)"

    corrected = f"{prefix.upper()}_{dd}{mm}{yyyy}_{es.upper()}{num}"
    if corrected == raw:
        return None, "Format non reconnu automatiquement"
    return corrected, ""
