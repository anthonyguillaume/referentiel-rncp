#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Génération du référentiel des diplômes RNCP depuis l'export open data officiel
de France compétences (data.gouv.fr), sur un périmètre de formacodes.

À chaque exécution :
  - télécharge le dernier export RNCP et le filtre sur config/formacodes.txt
    (fiches actives ; une fiche suivie qui se désactive reste dans le fichier) ;
  - compare avec la génération précédente (state/snapshot.json) ;
  - produit docs/diplomes_rncp.xlsx : lignes VERTES (ajoutées), ORANGES
    (modifiées, avec le détail avant → après), ROUGES (fiches désactivées ou
    disparues), onglets Synthèse, Changements et un onglet par domaine ;
  - produit le site statique docs/index.html + docs/data.json (recherche,
    filtres, lien de téléchargement de l'Excel) pour GitHub Pages ;
  - produit summary.md / summary.html (résumé du run), tient le journal des
    générations (docs/runs.json, page runs.html) et pose les sorties
    GitHub Actions (changes / sujet / resume).

Usage local :
  python genere_rncp.py                        # télécharge le dernier export
  python genere_rncp.py --xml export.xml       # export déjà téléchargé
  python genere_rncp.py --debug-fiche RNCP35803
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from datetime import date, datetime, timezone
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

# --------------------------------------------------------------------------- #
# Constantes
# --------------------------------------------------------------------------- #

DATASET_SLUG = (
    "repertoire-national-des-certifications-professionnelles-et-repertoire-specifique"
)
API_V2_RESSOURCES = f"https://www.data.gouv.fr/api/2/datasets/{DATASET_SLUG}/resources/"
API_V1_DATASET = f"https://www.data.gouv.fr/api/1/datasets/{DATASET_SLUG}/"
HTTP_HEADERS = {"User-Agent": "controle-rncp/2.0 (veille diplomes RNCP)"}
URL_FICHE = "https://www.francecompetences.fr/recherche/rncp/{num}/"

RE_RNCP = re.compile(r"RNCP\s*0*(\d+)", re.IGNORECASE)
RE_DATE_NOM = re.compile(r"(20\d{2})-(\d{2})-(\d{2})")
RE_VERSION = re.compile(r"v(\d+)[-_.](\d+)", re.IGNORECASE)

EVOLUTIONS = ("AJOUTÉ", "MODIFIÉ", "SUPPRIMÉ")
FILLS = {
    "AJOUTÉ": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
    "MODIFIÉ": PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
    "SUPPRIMÉ": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
}
FONTS = {
    "AJOUTÉ": Font(color="006100", bold=True),
    "MODIFIÉ": Font(color="9C6500", bold=True),
    "SUPPRIMÉ": Font(color="9C0006", bold=True),
}
FILL_ENTETE = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
FONT_INACTIF = Font(color="9C0006", bold=True)
FONT_LIEN = Font(color="0563C1", underline="single")

ENTETES = ["Code RNCP", "Intitulé", "Type", "Niveau", "Fiche active", "Échéance",
           "Remplacée par", "Formacodes", "Domaines",
           "Fiche France compétences", "Évolution", "Détail de l'évolution",
           "Dernier changement", "Suivi depuis"]

CHAMPS_COMPARES = [
    ("intitule", "Intitulé"),
    ("type", "Type"),
    ("niveau", "Niveau"),
    ("actif", "Fiche active"),
    ("date_fin", "Échéance"),
    ("successeurs", "Remplacée par"),
    ("formacodes", "Formacodes"),
]


# --------------------------------------------------------------------------- #
# Utilitaires
# --------------------------------------------------------------------------- #

def canon(valeur) -> str | None:
    if valeur is None:
        return None
    if isinstance(valeur, (int, float)) and not isinstance(valeur, bool):
        n = int(valeur)
        return f"RNCP{n}" if n > 0 else None
    m = RE_RNCP.search(str(valeur))
    return f"RNCP{int(m.group(1))}" if m else None


def parse_date_souple(s: str | None) -> date | None:
    if not s:
        return None
    s = s.strip()[:10]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def fr_date(iso: str | None) -> str:
    d = parse_date_souple(iso)
    return d.strftime("%d/%m/%Y") if d else ""


def strip_ns(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def log(msg: str) -> None:
    print(msg, flush=True)


# --------------------------------------------------------------------------- #
# 1. Récupération du dernier export RNCP (data.gouv.fr)
# --------------------------------------------------------------------------- #

def lister_ressources(session: requests.Session) -> list[dict]:
    ressources: list[dict] = []
    try:
        for page in (1, 2):
            r = session.get(API_V2_RESSOURCES, params={"page": page, "page_size": 50},
                            headers=HTTP_HEADERS, timeout=30)
            r.raise_for_status()
            data = r.json().get("data", [])
            ressources.extend(data)
            if len(data) < 50:
                break
    except Exception as exc:  # noqa: BLE001
        log(f"  API v2 indisponible ({exc}), tentative via l'API v1…")
    if not ressources:
        r = session.get(API_V1_DATASET, headers=HTTP_HEADERS, timeout=60)
        r.raise_for_status()
        ressources = r.json().get("resources", [])
    if not ressources:
        raise SystemExit("Impossible de lister les ressources data.gouv.fr")
    return ressources


def candidats_export_rncp(ressources: list[dict]) -> list[dict]:
    candidats = []
    for res in ressources:
        titre = f"{res.get('title') or ''} {res.get('url') or ''}".lower()
        est_zip = (res.get("format") or "").lower() == "zip" or titre.rstrip().endswith(".zip")
        if not est_zip or "rncp" not in titre or "csv" in titre:
            continue
        m = RE_DATE_NOM.search(titre)
        d = date(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else date.min
        v = RE_VERSION.search(titre)
        version = (int(v.group(1)), int(v.group(2))) if v else (0, 0)
        candidats.append({"res": res, "date": d, "version": version})
    candidats.sort(key=lambda c: (c["date"], c["version"]), reverse=True)
    return [c["res"] for c in candidats]


def telecharger_et_extraire(session: requests.Session, tmpdir: Path) -> tuple[Path, str]:
    candidats = candidats_export_rncp(lister_ressources(session))
    if not candidats:
        raise SystemExit("Aucun export RNCP (.zip) trouvé — nommage des ressources modifié ?")
    derniere_erreur = None
    for res in candidats[:3]:
        nom = res.get("title") or res.get("url", "").rsplit("/", 1)[-1]
        log(f"  Téléchargement : {nom}")
        try:
            zip_path = tmpdir / "export_rncp.zip"
            with session.get(res["url"], headers=HTTP_HEADERS, timeout=(15, 600),
                             stream=True) as r:
                r.raise_for_status()
                with open(zip_path, "wb") as f:
                    shutil.copyfileobj(r.raw, f, length=1024 * 1024)
            log(f"  Reçu : {zip_path.stat().st_size / 1e6:.1f} Mo")
            with zipfile.ZipFile(zip_path) as z:
                membres = [i for i in z.infolist() if i.filename.lower().endswith(".xml")]
                if not membres:
                    raise ValueError("aucun .xml dans l'archive")
                membre = max(membres, key=lambda i: i.file_size)
                xml_path = tmpdir / "export_rncp.xml"
                with z.open(membre) as src, open(xml_path, "wb") as dst:
                    shutil.copyfileobj(src, dst, length=1024 * 1024)
            log(f"  XML extrait : {membre.filename} ({membre.file_size / 1e6:.0f} Mo)")
            return xml_path, nom
        except Exception as exc:  # noqa: BLE001
            derniere_erreur = exc
            log(f"  Échec ({exc}), ressource suivante…")
    raise SystemExit(f"Téléchargement impossible : {derniere_erreur}")


# --------------------------------------------------------------------------- #
# 2. Analyse du XML (streaming, tolérant aux variations de schéma)
# --------------------------------------------------------------------------- #

def _texte(elem: ET.Element, tag: str) -> str | None:
    for enfant in elem:
        if strip_ns(enfant.tag) == tag:
            t = "".join(enfant.itertext()).strip()
            return t or None
    return None


def _codes_lies(elem: ET.Element, motif: str) -> list[str]:
    codes: set[str] = set()
    for node in elem.iter():
        if motif in strip_ns(node.tag).upper():
            for m in RE_RNCP.finditer("".join(node.itertext())):
                codes.add(f"RNCP{int(m.group(1))}")
    return sorted(codes, key=lambda c: int(c[4:]))


def _formacodes(elem: ET.Element) -> list[str]:
    resultat, vus = [], set()
    for node in elem.iter():
        if strip_ns(node.tag).upper() == "FORMACODE":
            code = (node.findtext("CODE") or "").strip()
            if not code:
                m = re.search(r"\b(\d{5})\b", "".join(node.itertext()))
                code = m.group(1) if m else ""
            if code and code not in vus:
                vus.add(code)
                resultat.append(code)
    return resultat


def _niveau(elem: ET.Element) -> str:
    for node in elem.iter():
        if "NOMENCLATURE_EUROPE" in strip_ns(node.tag).upper():
            m = re.search(r"(\d)", (node.findtext("NIVEAU") or "")
                          + "".join(node.itertext())[:40])
            if m:
                return f"Niveau {m.group(1)}"
    return ""


def parser_export(xml_path: Path, debug_fiche: str | None = None) -> dict[str, dict]:
    repertoire: dict[str, dict] = {}
    it = ET.iterparse(str(xml_path), events=("start", "end"))
    _, racine = next(it)
    n = 0
    for event, elem in it:
        if event != "end" or strip_ns(elem.tag) != "FICHE":
            continue
        numero = canon(_texte(elem, "NUMERO_FICHE"))
        if numero:
            if debug_fiche and numero == canon(debug_fiche):
                print(ET.tostring(elem, encoding="unicode")[:30000])
                raise SystemExit(0)
            actif_txt = (_texte(elem, "ACTIF") or "").strip().lower()
            abrege = None
            for enfant in elem:
                if strip_ns(enfant.tag) == "ABREGE":
                    abrege = (enfant.findtext("CODE") or "").strip() or None
            d = parse_date_souple(_texte(elem, "DATE_FIN_ENREGISTREMENT"))
            repertoire[numero] = {
                "intitule": _texte(elem, "INTITULE") or "",
                "type": abrege or "",
                "niveau": _niveau(elem),
                "actif": actif_txt.startswith("o"),
                "date_fin": d.isoformat() if d else "",
                "successeurs": _codes_lies(elem, "NOUVELLE_CERTIFICATION"),
                "formacodes": _formacodes(elem),
            }
            n += 1
            if n % 5000 == 0:
                log(f"  … {n} fiches lues")
        elem.clear()
        racine.clear()
    if debug_fiche:
        raise SystemExit(f"Fiche {debug_fiche} introuvable dans l'export.")
    log(f"  {n} fiches RNCP chargées "
        f"({sum(1 for f in repertoire.values() if f['actif'])} actives)")
    return repertoire


# --------------------------------------------------------------------------- #
# 3. Périmètre
# --------------------------------------------------------------------------- #

def lire_perimetre(chemin: Path) -> tuple[set[str], set[str]]:
    """-> (formacodes exacts à 5 chiffres, préfixes de domaine à 3 chiffres)"""
    exacts, prefixes = set(), set()
    if chemin.exists():
        for ligne in chemin.read_text(encoding="utf-8").splitlines():
            v = ligne.split("#", 1)[0].strip()
            if not v.isdigit():
                continue
            (exacts if len(v) >= 4 else prefixes).add(v)
    return exacts, prefixes


def dans_perimetre(formacodes: list[str], exacts: set[str], prefixes: set[str]) -> bool:
    if not exacts and not prefixes:
        return True
    for fc in formacodes:
        if fc in exacts or fc[:3] in prefixes:
            return True
    return False


# --------------------------------------------------------------------------- #
# 4. Diff avec la génération précédente
# --------------------------------------------------------------------------- #

def _aff(champ: str, valeur) -> str:
    if champ == "actif":
        return "Oui" if valeur else "Non"
    if champ == "date_fin":
        return fr_date(valeur)
    if isinstance(valeur, list):
        return ", ".join(valeur)
    return str(valeur or "")


def comparer(ancien: dict, nouveau: dict, jour: str) -> tuple[dict, list[dict]]:
    """-> ({code: (évolution, détail)}, événements du journal)"""
    diff: dict[str, tuple[str, str]] = {}
    events: list[dict] = []

    def ev(code, type_, avant, apres):
        events.append({"date": jour, "code": code, "type": type_,
                       "avant": str(avant), "apres": str(apres)})

    anciens, nouveaux = set(ancien), set(nouveau)
    for c in sorted(nouveaux - anciens, key=lambda x: int(x[4:])):
        diff[c] = ("AJOUTÉ", "Nouvelle fiche dans le périmètre")
        ev(c, "Fiche ajoutée", "", nouveau[c]["intitule"])
    for c in sorted(anciens - nouveaux, key=lambda x: int(x[4:])):
        diff[c] = ("SUPPRIMÉ", "Fiche disparue de l'export ou sortie du périmètre")
        ev(c, "Fiche disparue", ancien[c].get("intitule", ""), "")
    for c in sorted(anciens & nouveaux, key=lambda x: int(x[4:])):
        o, n = ancien[c], nouveau[c]
        morceaux = []
        for champ, libelle in CHAMPS_COMPARES:
            if o.get(champ) != n.get(champ):
                morceaux.append(f"{libelle} : {_aff(champ, o.get(champ))} → "
                                f"{_aff(champ, n.get(champ))}")
                ev(c, f"{libelle} modifié", _aff(champ, o.get(champ)),
                   _aff(champ, n.get(champ)))
        if morceaux:
            diff[c] = ("MODIFIÉ", " ; ".join(morceaux))
    return diff, events


# --------------------------------------------------------------------------- #
# 5. Classeur Excel
# --------------------------------------------------------------------------- #

def _ligne_valeurs(code: str, f: dict, evolution: str, detail: str) -> list:
    return [
        code, f.get("intitule", ""), f.get("type", ""), f.get("niveau", ""),
        "Oui" if f.get("actif") else "Non", fr_date(f.get("date_fin")),
        ", ".join(f.get("successeurs", [])),
        ", ".join(f.get("formacodes", [])),
        ", ".join(sorted({fc[:3] for fc in f.get("formacodes", [])})),
        URL_FICHE.format(num=code[4:]),
        evolution or "—", detail,
        fr_date(f.get("dernier_changement")), fr_date(f.get("premiere_vue")),
    ]


def ecrire_classeur(chemin: Path, lignes: list[tuple[str, dict, str, str]],
                    historique: list[dict], n_nouveaux: int, meta: dict) -> None:
    wb = Workbook()

    # ---- Onglet Diplômes ----
    ws = wb.active
    ws.title = "Diplômes"
    for i, t in enumerate(ENTETES, start=1):
        c = ws.cell(row=1, column=i, value=t)
        c.font = Font(bold=True)
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for col, largeur in zip("ABCDEFGHIJKLMN",
                            (12, 62, 9, 10, 11, 12, 22, 34, 14, 44, 12, 60, 14, 12)):
        ws.column_dimensions[col].width = largeur
    r = 1
    for code, f, evolution, detail in lignes:
        r += 1
        for i, v in enumerate(_ligne_valeurs(code, f, evolution, detail), start=1):
            ws.cell(row=r, column=i, value=v)
        lien = ws.cell(row=r, column=10)
        lien.hyperlink = lien.value
        lien.font = FONT_LIEN
        if evolution in FILLS:
            fill = FILLS[evolution]
            for col in range(1, len(ENTETES) + 1):
                ws.cell(row=r, column=col).fill = fill
            ws.cell(row=r, column=11).font = FONTS[evolution]
        if not f.get("actif"):
            ws.cell(row=r, column=5).font = FONT_INACTIF
    ws.freeze_panes = "A2"
    if r > 1:
        ref = f"A1:{get_column_letter(len(ENTETES))}{r}"
        table = Table(displayName="Diplomes", ref=ref)
        table.tableColumns = [TableColumn(id=i + 1, name=t)
                              for i, t in enumerate(ENTETES)]
        table.tableStyleInfo = TableStyleInfo(name="TableStyleLight9",
                                              showRowStripes=True)
        table.autoFilter = AutoFilter(ref=ref)
        ws.add_table(table)

    # ---- Synthèse ----
    syn = wb.create_sheet("Synthèse", 0)
    syn["A1"] = f"Diplômes RNCP — génération du {meta['date_generation']}"
    syn["A1"].font = Font(bold=True, size=14)
    syn["A2"] = f"Export France compétences : {meta['source_export']}"
    syn["A3"] = (f"Périmètre : {meta['perimetre']} — "
                 f"{meta['nb_lignes']} diplômes dans le fichier")
    if meta.get("note"):
        syn["A4"] = meta["note"]
        syn["A4"].font = Font(bold=True, color="9C0006")
    gras = Font(bold=True)
    l = 6
    syn.cell(row=l, column=1, value="Évolution depuis la dernière génération").font = gras
    for evo in EVOLUTIONS:
        l += 1
        c = syn.cell(row=l, column=1, value=evo)
        c.fill = FILLS[evo]
        c.font = FONTS[evo]
        syn.cell(row=l, column=2, value=meta["compteurs"].get(evo, 0))
    l += 1
    syn.cell(row=l, column=1, value="Sans changement")
    syn.cell(row=l, column=2, value=meta["compteurs"].get("—", 0))
    l += 2
    syn.cell(row=l, column=1, value="Lignes ayant évolué :").font = gras
    l += 1
    entetes_syn = ["Code RNCP", "Intitulé", "Évolution", "Détail", "Fiche active",
                   "Échéance", "Remplacée par"]
    for i, t in enumerate(entetes_syn, start=1):
        c = syn.cell(row=l, column=i, value=t)
        c.font = gras
        c.fill = FILL_ENTETE
    debut_actions = l
    ordre = {"SUPPRIMÉ": 0, "MODIFIÉ": 1, "AJOUTÉ": 2}
    for code, f, evolution, detail in sorted(
            (x for x in lignes if x[2] in FILLS),
            key=lambda x: (ordre[x[2]], int(x[0][4:]))):
        l += 1
        vals = [code, f.get("intitule", ""), evolution, detail,
                "Oui" if f.get("actif") else "Non", fr_date(f.get("date_fin")),
                ", ".join(f.get("successeurs", []))]
        for i, v in enumerate(vals, start=1):
            syn.cell(row=l, column=i, value=v)
        syn.cell(row=l, column=3).fill = FILLS[evolution]
        syn.cell(row=l, column=3).font = FONTS[evolution]
    for col, largeur in zip("ABCDEFG", (12, 62, 12, 70, 11, 12, 22)):
        syn.column_dimensions[col].width = largeur
    syn.freeze_panes = syn.cell(row=debut_actions + 1, column=1)
    syn.auto_filter.ref = f"A{debut_actions}:G{max(l, debut_actions)}"

    # ---- Changements (journal) ----
    chg = wb.create_sheet("Changements", 1)
    titre = chg.cell(row=1, column=1,
                     value="Journal des changements — lignes surlignées : génération courante")
    titre.font = Font(bold=True, color="1F4E79")
    for i, t in enumerate(["Date", "Code RNCP", "Changement", "Avant", "Après"], start=1):
        c = chg.cell(row=2, column=i, value=t)
        c.font = gras
        c.fill = FILL_ENTETE
    fill_nouveau = PatternFill(start_color="BDD7EE", end_color="BDD7EE",
                               fill_type="solid")
    for j, e in enumerate(reversed(historique), start=3):
        for i, v in enumerate([e["date"], e["code"], e["type"], e["avant"],
                               e["apres"]], start=1):
            cel = chg.cell(row=j, column=i, value=v)
            if j - 3 < n_nouveaux:
                cel.fill = fill_nouveau
    for col, largeur in zip("ABCDE", (12, 13, 26, 55, 55)):
        chg.column_dimensions[col].width = largeur
    chg.freeze_panes = "A3"
    chg.auto_filter.ref = f"A2:E{max(3, 2 + len(historique))}"

    # ---- Un onglet par domaine (formacode simplifié) ----
    par_domaine: dict[str, list] = {}
    for code, f, evolution, detail in lignes:
        for dom in sorted({fc[:3] for fc in f.get("formacodes", [])}):
            par_domaine.setdefault(dom, []).append((code, f, evolution))
    for dom in sorted(par_domaine):
        feuille = wb.create_sheet(f"Formacode {dom}")
        for i, t in enumerate(["Code RNCP", "Intitulé", "Fiche active", "Évolution"],
                              start=1):
            c = feuille.cell(row=1, column=i, value=t)
            c.font = gras
            c.fill = FILL_ENTETE
        rr = 1
        for code, f, evolution in par_domaine[dom]:
            rr += 1
            feuille.cell(row=rr, column=1, value=code)
            feuille.cell(row=rr, column=2, value=f.get("intitule", ""))
            feuille.cell(row=rr, column=3, value="Oui" if f.get("actif") else "Non")
            feuille.cell(row=rr, column=4, value=evolution or "—")
            if evolution in FILLS:
                for col in range(1, 5):
                    feuille.cell(row=rr, column=col).fill = FILLS[evolution]
        for col, largeur in zip("ABCD", (12, 70, 11, 12)):
            feuille.column_dimensions[col].width = largeur
        feuille.freeze_panes = "A2"
        feuille.auto_filter.ref = f"A1:D{rr}"

    wb.save(chemin)


# --------------------------------------------------------------------------- #
# 6. Site statique + résumés
# --------------------------------------------------------------------------- #

def ecrire_site(outdir: Path, gabarit: Path, lignes, meta: dict) -> None:
    rows = []
    for code, f, evolution, detail in lignes:
        rows.append({
            "code": code, "intitule": f.get("intitule", ""),
            "type": f.get("type", ""), "niveau": f.get("niveau", ""),
            "actif": bool(f.get("actif")), "echeance": fr_date(f.get("date_fin")),
            "successeurs": f.get("successeurs", []),
            "formacodes": f.get("formacodes", []),
            "evolution": evolution or "", "detail": detail,
            "dernier_changement": fr_date(f.get("dernier_changement")),
            "url": URL_FICHE.format(num=code[4:]),
        })
    data = {"meta": {"genere_le": meta["date_generation"],
                     "export": meta["source_export"],
                     "perimetre": meta["perimetre"],
                     "compteurs": meta["compteurs"],
                     "note": meta.get("note", "")},
            "rows": rows}
    (outdir / "data.json").write_text(json.dumps(data, ensure_ascii=False),
                                      encoding="utf-8")
    copier_gabarits(outdir, gabarit)


def copier_gabarits(outdir: Path, gabarit: Path) -> None:
    """Copie index.html et les autres pages (runs.html…) du dossier du gabarit."""
    if gabarit.parent.is_dir():
        for page in gabarit.parent.glob("*.html"):
            shutil.copyfile(page, outdir / page.name)


# --------------------------------------------------------------------------- #
# 6 bis. Journal des générations (runs.json, page runs.html)
# --------------------------------------------------------------------------- #

JOURNAL_MAX = 1500


def url_run_actions() -> str:
    serveur = os.environ.get("GITHUB_SERVER_URL", "")
    depot = os.environ.get("GITHUB_REPOSITORY", "")
    run_id = os.environ.get("GITHUB_RUN_ID", "")
    return f"{serveur}/{depot}/actions/runs/{run_id}" if serveur and depot and run_id else ""


def entree_journal(statut: str, maintenant: datetime, **champs) -> dict:
    e = {"date": maintenant.strftime("%Y-%m-%dT%H:%M"), "statut": statut, "export": "",
         "diplomes": 0, "ajoutes": 0, "modifies": 0, "supprimes": 0, "message": "",
         "run_url": url_run_actions()}
    e.update(champs)
    return e


def ecrire_runs_json(outdir: Path, journal: list[dict], maintenant: datetime) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    (outdir / "runs.json").write_text(
        json.dumps({"genere_le": maintenant.strftime("%d/%m/%Y %H:%M"),
                    "runs": list(reversed(journal))}, ensure_ascii=False),
        encoding="utf-8")


def charger_journal(chemin_state: Path) -> list[dict]:
    if chemin_state.exists():
        try:
            contenu = json.loads(chemin_state.read_text(encoding="utf-8"))
        except ValueError:
            return []
        if contenu.get("version") == 2:
            return contenu.get("journal", [])
    return []


def construire_resumes(meta: dict, events: list[dict], lignes,
                       premiere: bool) -> tuple[str, str, str]:
    cpt = meta["compteurs"]
    bilan = (f"{cpt.get('AJOUTÉ', 0)} ajouté(s), {cpt.get('MODIFIÉ', 0)} modifié(s), "
             f"{cpt.get('SUPPRIMÉ', 0)} supprimé(s)")
    sujet = (f"Initialisation — {meta['nb_lignes']} diplômes suivis" if premiere
             else f"{bilan}")
    site = os.environ.get("SITE_URL", "").strip()

    md = [f"# Diplômes RNCP — génération du {meta['date_generation']}", "",
          f"Export France compétences : {meta['source_export']}  ",
          f"Périmètre : {meta['perimetre']}  ",
          f"Diplômes dans le fichier : {meta['nb_lignes']}  ",
          f"Bilan : **{bilan}**", ""]
    if site:
        md.append(f"Consulter et rechercher : {site}")
        md.append("")
    if premiere:
        md.append("Première génération : le fichier joint constitue l'état de référence.")
    else:
        evolues = [x for x in lignes if x[2] in FILLS]
        if evolues:
            md.append("| Code | Évolution | Intitulé | Détail |")
            md.append("|---|---|---|---|")
            ordre = {"SUPPRIMÉ": 0, "MODIFIÉ": 1, "AJOUTÉ": 2}
            for code, f, evolution, detail in sorted(
                    evolues, key=lambda x: (ordre[x[2]], int(x[0][4:])))[:80]:
                md.append(f"| {code} | {evolution} | {f.get('intitule','')} "
                          f"| {detail} |")
            if len(evolues) > 80:
                md.append(f"| … | | {len(evolues)-80} autres lignes | voir fichier |")
    texte_md = "\n".join(md) + "\n"

    def chip(evo):
        f = FILLS[evo].start_color.rgb[-6:]
        c = FONTS[evo].color.rgb[-6:]
        return (f'<span style="background:#{f};color:#{c};padding:1px 8px;'
                f'border-radius:9px;font-weight:600">{evo}</span>')

    h = [f"<h2>Diplômes RNCP — génération du {meta['date_generation']}</h2>",
         f"<p>Export : {meta['source_export']}<br>Périmètre : {meta['perimetre']}<br>"
         f"Diplômes : {meta['nb_lignes']}<br>"
         + " · ".join(f"{chip(e)} {cpt.get(e, 0)}" for e in EVOLUTIONS) + "</p>"]
    if site:
        h.append(f'<p><a href="{site}">Consulter et rechercher en ligne</a></p>')
    if premiere:
        h.append("<p>Première génération : le fichier joint constitue l'état "
                 "de référence.</p>")
    else:
        evolues = [x for x in lignes if x[2] in FILLS]
        if evolues:
            h.append('<table border="1" cellspacing="0" cellpadding="5" '
                     'style="border-collapse:collapse;font-family:Arial;'
                     'font-size:13px">')
            h.append("<tr style='background:#DDEBF7'><th>Code</th><th>Évolution</th>"
                     "<th>Intitulé</th><th>Détail</th></tr>")
            ordre = {"SUPPRIMÉ": 0, "MODIFIÉ": 1, "AJOUTÉ": 2}
            for code, f, evolution, detail in sorted(
                    evolues, key=lambda x: (ordre[x[2]], int(x[0][4:])))[:80]:
                h.append(f"<tr><td>{code}</td><td>{chip(evolution)}</td>"
                         f"<td>{f.get('intitule','')}</td><td>{detail}</td></tr>")
            h.append("</table>")
    h.append("<p><small>Détail complet : classeur joint et site (onglets Synthèse "
             "et Changements).</small></p>")
    return texte_md, "\n".join(h) + "\n", sujet


def ecrire_github_output(changes: bool, sujet: str, resume: str) -> None:
    chemin = os.environ.get("GITHUB_OUTPUT")
    if not chemin:
        return
    sujet = re.sub(r"[\r\n]+", " ", sujet)
    resume = re.sub(r"[\r\n\"'`]+", " ", resume)
    with open(chemin, "a", encoding="utf-8") as f:
        f.write(f"changes={'true' if changes else 'false'}\n")
        f.write(f"sujet={sujet}\n")
        f.write(f"resume={resume}\n")


# --------------------------------------------------------------------------- #
# Programme principal
# --------------------------------------------------------------------------- #

def main() -> None:
    p = argparse.ArgumentParser(description="Génération du référentiel diplômes RNCP")
    p.add_argument("--xml", default="", help="export XML local (ne pas télécharger)")
    p.add_argument("--state", default="state/snapshot.json")
    p.add_argument("--outdir", default="docs")
    p.add_argument("--config", default="config/formacodes.txt")
    p.add_argument("--site", default="site/index.html")
    p.add_argument("--min-fiches", type=int, default=5000,
                   help="échec si l'export contient moins de N fiches (garde-fou)")
    p.add_argument("--debug-fiche", default="")
    args = p.parse_args()

    try:
        from zoneinfo import ZoneInfo
        maintenant = datetime.now(ZoneInfo("Europe/Paris"))
    except Exception:  # noqa: BLE001
        maintenant = datetime.now(timezone.utc)
    jour_iso = maintenant.date().isoformat()

    exacts, prefixes = lire_perimetre(Path(args.config))
    if exacts or prefixes:
        perimetre_txt = (f"{len(exacts)} formacode(s)"
                         + (f" + {len(prefixes)} domaine(s)" if prefixes else ""))
    else:
        perimetre_txt = "répertoire RNCP entier"
        log("ATTENTION : périmètre vide, tout le répertoire sera suivi.")
    log(f"Périmètre : {perimetre_txt}")

    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        if args.xml:
            xml_path, source = Path(args.xml), Path(args.xml).name
            log(f"Export local : {xml_path}")
        else:
            log("Recherche du dernier export RNCP sur data.gouv.fr…")
            with requests.Session() as session:
                xml_path, source = telecharger_et_extraire(session, tmpdir)
        log("Lecture de l'export…")
        repertoire = parser_export(xml_path, args.debug_fiche or None)

    if len(repertoire) < args.min_fiches:
        raise SystemExit(
            f"Export anormalement petit ({len(repertoire)} fiches < {args.min_fiches}) : "
            "schéma modifié ? Lancer --debug-fiche RNCP35803 pour vérifier."
        )

    # Instantané précédent
    chemin_state = Path(args.state)
    ancien: dict = {}
    historique: list[dict] = []
    journal: list[dict] = []
    premiere = True
    if chemin_state.exists():
        contenu = json.loads(chemin_state.read_text(encoding="utf-8"))
        if contenu.get("version") == 2:
            ancien = contenu.get("codes", {})
            historique = contenu.get("historique", [])
            journal = contenu.get("journal", [])
            premiere = False
        else:
            log("Instantané v1 détecté : nouvelle base de référence (v2).")

    # Sélection : fiches ACTIVES du périmètre, plus toutes celles déjà suivies
    # quels que soient leurs formacodes (une fiche suivie qui se désactive ou
    # dont les formacodes changent reste visible ; les fiches inactives depuis
    # avant le suivi n'entrent pas).
    selection = {c: f for c, f in repertoire.items()
                 if c in ancien
                 or (f["actif"] and dans_perimetre(f["formacodes"], exacts, prefixes))}
    n_inactives = sum(1 for f in selection.values() if not f["actif"])
    log(f"  {len(selection)} fiches dans le périmètre "
        f"({len(selection) - n_inactives} actives, {n_inactives} inactives déjà suivies)")
    if not selection:
        raise SystemExit("Aucune fiche dans le périmètre : vérifier config/formacodes.txt")

    diff, events = ({}, []) if premiere else comparer(ancien, selection, jour_iso)
    if premiere:
        historique.append({"date": jour_iso, "code": "—",
                           "type": "Initialisation du suivi", "avant": "",
                           "apres": f"{len(selection)} diplômes suivis"})
    historique.extend(events)
    historique = historique[-1000:]
    n_nouveaux = 1 if premiere else len(events)

    # Lignes du fichier : sélection courante + disparues (tombstones rouges)
    lignes: list[tuple[str, dict, str, str]] = []
    for code, f in selection.items():
        evolution, detail = diff.get(code, ("", ""))
        prec = ancien.get(code, {})
        f = dict(f)
        f["premiere_vue"] = prec.get("premiere_vue", jour_iso)
        f["dernier_changement"] = (jour_iso if evolution
                                   else prec.get("dernier_changement", ""))
        lignes.append((code, f, evolution, detail))
    for code, (evolution, detail) in diff.items():
        if evolution == "SUPPRIMÉ":
            f = dict(ancien.get(code, {}))
            f["dernier_changement"] = jour_iso
            lignes.append((code, f, evolution, detail))
    lignes.sort(key=lambda x: (x[1].get("type", ""), x[1].get("intitule", "")))

    compteurs = {e: 0 for e in EVOLUTIONS}
    compteurs["—"] = 0
    for _, _, evolution, _ in lignes:
        compteurs[evolution if evolution in compteurs else "—"] += 1

    meta = {
        "date_generation": maintenant.strftime("%d/%m/%Y %H:%M"),
        "source_export": source,
        "perimetre": perimetre_txt,
        "nb_lignes": len(lignes),
        "compteurs": compteurs,
    }
    log("Bilan : " + ", ".join(f"{compteurs[e]} {e.lower()}" for e in EVOLUTIONS)
        + f", {compteurs['—']} sans changement")

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    log(f"Écriture du classeur : {outdir / 'diplomes_rncp.xlsx'}")
    ecrire_classeur(outdir / "diplomes_rncp.xlsx", lignes, historique,
                    n_nouveaux, meta)
    ecrire_site(outdir, Path(args.site), lignes, meta)
    md, html, sujet = construire_resumes(meta, events, lignes, premiere)
    (outdir / "summary.md").write_text(md, encoding="utf-8")
    (outdir / "summary.html").write_text(html, encoding="utf-8")

    changements = premiere or bool(events)
    journal.append(entree_journal(
        "initialisation" if premiere else ("changements" if changements else "sans_changement"),
        maintenant, export=source, diplomes=len(lignes),
        ajoutes=compteurs["AJOUTÉ"], modifies=compteurs["MODIFIÉ"],
        supprimes=compteurs["SUPPRIMÉ"]))
    journal = journal[-JOURNAL_MAX:]
    ecrire_runs_json(outdir, journal, maintenant)

    # Nouvel instantané
    etat = {}
    for code, f, evolution, _ in lignes:
        if evolution == "SUPPRIMÉ":
            continue  # les disparues sortent du suivi après signalement
        etat[code] = {k: f.get(k) for k in ("intitule", "type", "niveau", "actif",
                                            "date_fin", "successeurs", "formacodes",
                                            "premiere_vue", "dernier_changement")}
    chemin_state.parent.mkdir(parents=True, exist_ok=True)
    chemin_state.write_text(
        json.dumps({"version": 2, "derniere_execution": maintenant.isoformat(),
                    "source_export": source, "codes": etat,
                    "historique": historique, "journal": journal},
                   ensure_ascii=False, indent=1),
        encoding="utf-8")

    resume = (f"{cptr(compteurs)}" if not premiere
              else f"initialisation, {len(lignes)} diplômes")
    ecrire_github_output(changements, sujet, resume)
    step = os.environ.get("GITHUB_STEP_SUMMARY")
    if step:
        with open(step, "a", encoding="utf-8") as f:
            f.write(md)
    log(f"Changements à signaler : {'oui' if changements else 'non'} — {sujet}")


def cptr(c: dict) -> str:
    return (f"{c.get('AJOUTÉ', 0)} ajout(s), {c.get('MODIFIÉ', 0)} modif(s), "
            f"{c.get('SUPPRIMÉ', 0)} suppression(s)")


if __name__ == "__main__":
    main()
