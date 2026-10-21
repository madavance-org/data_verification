"""Inspection ponctuelle des datagrids Clean Cooking : en-têtes de colonnes, nombre de lignes
et répartition des Status. N'affiche aucune donnée de ligne et n'écrit rien nulle part."""

import csv
import io
import os
from collections import Counter

from common.mwater_client import download_datagrid_raw, mwater_login

DATAGRIDS = {
    "Inscription (présumé)": "ac3ae2ce1f7540f586ae1ab5673419a8",
    "Distribution (présumé)": "ba07c16edc57417093d9735c93a97e52",
    "Suivi post distribution (présumé)": "e4a38892c65840ef8d1caf59b0fe6198",
    "Atelier de réparation (présumé)": "e129499f02c141a3a45e65bfb8b3f7d7",
}


def main():
    client_id = mwater_login(os.environ["MWATER_USERNAME"].strip(), os.environ["MWATER_PASSWORD"].strip())
    for label, datagrid_id in DATAGRIDS.items():
        raw = download_datagrid_raw(datagrid_id, client_id)
        rows = list(csv.DictReader(io.StringIO(raw.decode("utf-8-sig"))))
        headers = list(rows[0].keys()) if rows else []
        print(f"\n===== {label} — {datagrid_id}")
        print(f"Lignes : {len(rows)} / Colonnes : {len(headers)}")
        if "Status" in headers:
            print("Status :", dict(Counter(r.get("Status", "") for r in rows)))
        for i, h in enumerate(headers, 1):
            print(f"  {i:>3}. {h}")


if __name__ == "__main__":
    main()
