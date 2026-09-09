#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Journalise un échec de génération dans state/snapshot.json (journal) et
régénère docs/runs.json + les pages du site, pour que l'historique des
générations affiche aussi les erreurs.

Usage (workflow, étape `if: failure()`) :
  python outils/journaliser_erreur.py generation.log
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from genere_rncp import (JOURNAL_MAX, charger_journal, copier_gabarits,  # noqa: E402
                         ecrire_runs_json, entree_journal, log)


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("log", nargs="?", default="")
    p.add_argument("--state", default="state/snapshot.json")
    p.add_argument("--outdir", default="docs")
    p.add_argument("--site", default="site/index.html")
    args = p.parse_args()

    try:
        from zoneinfo import ZoneInfo
        maintenant = datetime.now(ZoneInfo("Europe/Paris"))
    except Exception:  # noqa: BLE001
        maintenant = datetime.now(timezone.utc)

    message = "Échec de la génération (voir le run GitHub Actions)."
    if args.log and Path(args.log).exists():
        lignes = [l.strip() for l in Path(args.log).read_text(encoding="utf-8",
                                                               errors="replace").splitlines()
                  if l.strip() and "fiches lues" not in l]
        if lignes:
            message = " | ".join(lignes[-3:])[:600]

    chemin = Path(args.state)
    journal = charger_journal(chemin)
    journal.append(entree_journal("erreur", maintenant, message=message))
    journal = journal[-JOURNAL_MAX:]

    contenu = {"version": 2, "codes": {}, "historique": []}
    if chemin.exists():
        try:
            lu = json.loads(chemin.read_text(encoding="utf-8"))
            if lu.get("version") == 2:
                contenu = lu
        except ValueError:
            pass
    contenu["journal"] = journal
    chemin.parent.mkdir(parents=True, exist_ok=True)
    chemin.write_text(json.dumps(contenu, ensure_ascii=False, indent=1), encoding="utf-8")

    outdir = Path(args.outdir)
    ecrire_runs_json(outdir, journal, maintenant)
    copier_gabarits(outdir, Path(args.site))
    log(f"Échec journalisé : {message}")


if __name__ == "__main__":
    main()
