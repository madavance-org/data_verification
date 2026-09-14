"""Aide HTTP partagée par tous les scripts de ce repo."""

import sys


def raise_for_status_verbose(response):
    """Comme response.raise_for_status(), mais loggue d'abord le corps complet de la
    réponse (tronqué) sur stderr — indispensable pour diagnostiquer une erreur mWater/Graph
    depuis les logs GitHub Actions, où le message par défaut de requests n'inclut pas le corps."""
    if not response.ok:
        print(f"HTTP {response.status_code} sur {response.url}", file=sys.stderr)
        print(response.text[:2000], file=sys.stderr)
        response.raise_for_status()
