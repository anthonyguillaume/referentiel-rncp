"""Tests de la règle de sélection des fiches suivies.

    python -m unittest discover tests
"""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from genere_rncp import selectionner  # noqa: E402


def fiche(actif=True, formacodes=("21752",), successeurs=()):
    return {"intitule": "x", "type": "CAP", "niveau": "Niveau 3", "actif": actif,
            "date_fin": "", "successeurs": list(successeurs), "formacodes": list(formacodes)}


EXACTS, PREFIXES = {"21752"}, set()


class SelectionTest(unittest.TestCase):
    def test_active_du_perimetre(self):
        rep = {"RNCP1": fiche(), "RNCP2": fiche(formacodes=["99999"])}
        self.assertEqual(set(selectionner(rep, {}, EXACTS, PREFIXES)), {"RNCP1"})

    def test_inactive_jamais_suivie_exclue(self):
        rep = {"RNCP1": fiche(actif=False)}
        self.assertEqual(set(selectionner(rep, {}, EXACTS, PREFIXES)), set())

    def test_deja_suivie_reste_suivie(self):
        rep = {"RNCP1": fiche(actif=False, formacodes=["99999"])}
        self.assertEqual(set(selectionner(rep, {"RNCP1": {}}, EXACTS, PREFIXES)), {"RNCP1"})

    def test_successeur_actif_hors_perimetre_entre(self):
        # Cas RNCP37245 → RNCP42124 (signalement du 16/09/2026) : l'ancienne fiche est
        # suivie via la référence, la remplaçante porte un formacode hors périmètre.
        rep = {"RNCP37245": fiche(actif=False, formacodes=["21759"], successeurs=["RNCP42124"]),
               "RNCP42124": fiche(formacodes=["21759"])}
        sel = selectionner(rep, {"RNCP37245": {}}, EXACTS, PREFIXES)
        self.assertEqual(set(sel), {"RNCP37245", "RNCP42124"})

    def test_successeur_transitif(self):
        rep = {"RNCP1": fiche(actif=False, successeurs=["RNCP2"]),
               "RNCP2": fiche(actif=False, formacodes=["99999"], successeurs=["RNCP3"]),
               "RNCP3": fiche(formacodes=["99999"])}
        sel = selectionner(rep, {"RNCP1": {}}, EXACTS, PREFIXES)
        # RNCP2, inactive et jamais suivie, n'entre pas mais la chaîne est remontée
        self.assertEqual(set(sel), {"RNCP1", "RNCP3"})

    def test_successeur_absent_ou_cyclique(self):
        rep = {"RNCP1": fiche(successeurs=["RNCP2", "RNCP404"]),
               "RNCP2": fiche(formacodes=["99999"], successeurs=["RNCP1"])}
        self.assertEqual(set(selectionner(rep, {}, EXACTS, PREFIXES)), {"RNCP1", "RNCP2"})

    def test_successeur_d_une_fiche_non_suivie_ignore(self):
        rep = {"RNCP1": fiche(formacodes=["99999"], successeurs=["RNCP2"]),
               "RNCP2": fiche(formacodes=["99999"])}
        self.assertEqual(set(selectionner(rep, {}, EXACTS, PREFIXES)), set())


if __name__ == "__main__":
    unittest.main()
