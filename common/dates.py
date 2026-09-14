"""Parsing de dates/horodatages mWater, partagé par tous les scripts de ce repo."""

from datetime import datetime


def parse_dt(value):
    """Parse un horodatage mWater 'AAAA-MM-JJ HH:MM:SS' ou une date 'AAAA-MM-JJ'.
    Retourne None si vide/invalide."""
    value = (value or "").strip()
    if not value:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None
