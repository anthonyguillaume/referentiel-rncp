# CLAUDE.md — Référentiel des diplômes RNCP (GIP FCIP, académie de Reims)

Ce fichier est le point d'entrée pour reprendre le projet, par un humain ou par Claude.
Il décrit ce qui existe, pourquoi, et comment intervenir sans casser la chaîne.

## 1. Le projet en trois phrases

Chaque jour à 04:30 UTC, un workflow GitHub Actions télécharge l'export open data du RNCP
(France compétences, data.gouv.fr, ~75 Mo zip / ~470 Mo XML, ~25 700 fiches dont ~5 000 actives),
le filtre sur un périmètre de formacodes, compare avec l'état de la veille, puis publie un
classeur Excel colorisé et un site statique sur GitHub Pages. Les couleurs (ajouté / modifié /
supprimé) sont **cumulées depuis une référence de départ**, actuellement le fichier Excel
« Liens CPF par formacode et diplômes.xlsx » importé le 09/09/2026.

- Dépôt : https://github.com/anthonyguillaume/referentiel-rncp (public, renommé depuis `rncp-watch`)
- Site : https://anthonyguillaume.github.io/referentiel-rncp/ (Pages, branche `main`, dossier `/docs`)
- Historique des runs : https://anthonyguillaume.github.io/referentiel-rncp/runs.html
- Dossier local : `~/Dev/rncp-watch` (nom d'origine conservé, le remote pointe sur le nouveau dépôt)

## 2. Arborescence

| Chemin | Rôle |
|---|---|
| `genere_rncp.py` | Générateur : téléchargement, parsing XML en flux, périmètre, diff, Excel, site, journal, sorties Actions |
| `outils/initialiser_depuis_excel.py` | Réinitialise `state/snapshot.json` depuis un Excel de liens formacode → RNCP (référence) |
| `outils/journaliser_erreur.py` | Inscrit un échec dans le journal des runs (appelé par le workflow `if: failure()`) |
| `config/formacodes.txt` | Périmètre : un formacode par ligne (5 chiffres exact, 3 chiffres = domaine). 700 entrées |
| `site/index.html` | Gabarit du site (format « bundler » Claude Design, voir §6). Recopié dans `docs/` à chaque run |
| `site/runs.html` | Page Historique des générations (HTML classique) |
| `state/snapshot.json` | État versionné : codes suivis avec évolution cumulée, historique des événements, journal des runs, référence |
| `docs/` | Sortie publiée par le bot : `index.html`, `runs.html`, `data.json`, `runs.json`, `diplomes_rncp.xlsx`, `summary.*` |
| `.github/workflows/controle-rncp.yml` | Workflow quotidien + `workflow_dispatch` |
| `tests/generation-1.xml`, `-2.xml` | Exports simulés (1 412 fiches ; la 2 contient 2 ajouts, 3 modifs, 2 suppressions) |
| `Dockerfile`, `.dockerignore` | Image de test locale alignée sur le runner (Python 3.12) |
| `tmp/` | Ignoré par git ; y déposer les analyses ponctuelles (ex. `analyse_liens_cpf.xlsx`) |

## 3. Fonctionnement détaillé

### Sélection des fiches (règle importante)
Une fiche entre dans le suivi si elle est **active** et porte un formacode du périmètre.
Une fiche **déjà suivie reste suivie quels que soient ses formacodes** et même devenue inactive.
Pourquoi : 546 des 1 412 codes du fichier Excel (dont 471 actifs) portent chez France
compétences des formacodes absents des 700 du périmètre ; sans cette règle ils seraient perdus.
Les fiches inactives depuis avant le suivi n'entrent jamais (sinon 4 289 fiches historiques).

### Évolution cumulée depuis la référence
Pour chaque code, `state/snapshot.json` conserve `evolution` (AJOUTÉ / MODIFIÉ / SUPPRIMÉ / vide)
et `detail` (cumul « Libellé : avant → après », fusionné par `fusionner_details`).
- AJOUTÉ = absente de la référence ; reste AJOUTÉ même si modifiée ensuite ; si elle disparaît, elle sort simplement.
- MODIFIÉ = fiche de la référence changée depuis ; redevient « sans changement » si tous les champs reviennent à la valeur d'origine.
- SUPPRIMÉ = fiche de la référence disparue de l'export ; la ligne rouge persiste ; si elle réapparaît → MODIFIÉ « Fiche réapparue ».
- Le diff **du jour** (vs run précédent) alimente le journal des runs, l'onglet Changements, les compteurs « Dont cette génération » de la Synthèse Excel et les sorties du workflow.
- `dernier_changement` = date du dernier événement sur la fiche ; c'est ce qui permet le sélecteur de période sur le site.

Champs comparés : intitulé, type (ABREGE), niveau, actif, date de fin d'enregistrement, successeurs (NOUVELLE_CERTIFICATION), formacodes.

### Journal des runs et erreurs
Chaque run ajoute une entrée dans `journal` (date, statut initialisation / changements / sans_changement / erreur,
export, diplômes, ajoutés, modifiés, supprimés, lien vers le run Actions) → `docs/runs.json` → page `runs.html`.
Le workflow **committe à chaque run**, même sans changement (sinon le journal ne serait pas persisté).
En cas d'échec, `journaliser_erreur.py` inscrit les 3 dernières lignes du log et le commit a lieu quand même ;
le site affiche alors un point rouge sur « Généré le » et un bandeau d'alerte (orange si aucune génération depuis > 48 h).

### Site (`site/index.html`)
Recherche (code, intitulé, formacode), filtres Ajoutés / Modifiés / Supprimés / Fiches inactives, listes Type et
Niveau (avec effectifs), colonnes Formacodes et Formacode simplifié (domaine à 3 chiffres, calculé côté site ; puces, tri sur le premier code,
masquées sous 620 px), liste Formacode simplifié (77 domaines, avec effectifs ; une fiche à plusieurs domaines est comptée sous chacun ;
pas de liste sur les formacodes complets : 1 576 codes distincts, trop pour un sélecteur — la recherche suffit), sélecteur de période des changements (défaut : **7 jours** ; dernier run, 30 jours, depuis le début),
tri par clic sur les en-têtes, colonne « Dernier changement », badges Niveau / Active, infobulles sur les en-têtes
de colonnes et les puces du bandeau, lien Excel, lien Historique, icône GitHub, balises anti-cache et `data.json`
chargé avec horodatage + `cache: no-store`.

### Export France compétences
Publié chaque jour (week-end compris) vers 02:00 UTC sous le nom `export-fiches-rncp-v4-1-AAAA-MM-JJ.zip`.
Le script liste les ressources via l'API data.gouv (User-Agent obligatoire, sinon réponse non JSON) et prend la plus
récente d'après la date du nom. Garde-fou `--min-fiches 5000`. `--debug-fiche RNCP35803` affiche le XML brut d'une fiche.
Structure vérifiée le 09/09/2026 : `NUMERO_FICHE`, `INTITULE`, `ABREGE/CODE`, `NOMENCLATURE_EUROPE/NIVEAU` (NIV5),
`ACTIF` (Oui/Non, au premier niveau), `DATE_FIN_ENREGISTREMENT` (jj/mm/aaaa), `FORMACODES/FORMACODE/CODE`,
`NOUVELLES_CERTIFICATIONS/NOUVELLE_CERTIFICATION/ID_FICHE_NOUVELLE_CERTIFICATION`.
Le Type est vide pour ~1 350 fiches actives : France compétences ne renseigne pas d'ABREGE pour les titres privés (donnée source, pas un bug).

## 4. Configuration GitHub

- Compte à utiliser : `anthonyguillaume` (pas le compte `anthony-guillaume_getlink` actif par défaut dans `gh`).
  Pour les commandes : `export GH_TOKEN=$(gh auth token --user anthonyguillaume)` et pousser avec
  `git push "https://x-access-token:${GH_TOKEN}@github.com/anthonyguillaume/referentiel-rncp.git" main`.
- Variables : `SITE_URL` (utilisée dans les résumés). `MAIL_SERVER`, `MAIL_PORT`, `MAIL_USERNAME`, `MAIL_FROM`, `MAIL_TO`
  et le secret `MAIL_PASSWORD` existent encore mais **l'envoi d'e-mail a été retiré du workflow** le 09/09/2026 à la demande
  de l'utilisateur. Pour le réactiver : remettre une étape `dawidd6/action-send-mail@v3` avant le commit (voir historique git, commit « retrait du mail »).
- Pages : dépôt public obligatoire (plan Free). Rebuild automatique ~1 min après chaque push sur `main`.
- Bot : `controle-rncp[bot]`, commits « Génération RNCP AAAA-MM-JJ — … ». **Toujours `git pull` après un run** avant de modifier en local.
- Licence : MIT (code) ; données sous licence ouverte Etalab 2.0.

## 5. Procédures

### Tester en local (avant tout push)
```bash
docker build -t rncp-watch .
# hors ligne, rapide
docker run --rm rncp-watch python genere_rncp.py --xml tests/generation-1.xml --min-fiches 0 --state /tmp/s.json --outdir /tmp/d
# réel, dans un dossier temporaire (ne pas écraser state/ et docs/ committés par le bot)
docker run --rm -v "$PWD/tmp:/t" rncp-watch python genere_rncp.py --state /t/s.json --outdir /t/d
```
Pour éviter de retélécharger 75 Mo à chaque essai : extraire une fois `export_rncp.xml` dans `tmp/` et utiliser `--xml`.
Le site se teste en servant un dossier contenant `index.html`, `runs.html`, `data.json`, `runs.json` : `python3 -m http.server 8766`.

### Lancer un run à la demande
`gh workflow run "Contrôle RNCP" -R anthonyguillaume/referentiel-rncp` puis `gh run watch <id>`, puis `git pull`.

### Changer le périmètre
Éditer `config/formacodes.txt`, committer, pousser. Les nouvelles fiches actives apparaîtront en vert au run suivant.
Les fiches déjà suivies ne sortent jamais du suivi par ce biais.

### Réinitialiser la référence depuis un nouvel Excel
```bash
docker run --rm -v "$PWD:/app" rncp-watch python outils/initialiser_depuis_excel.py "Nouveau fichier.xlsx"   # télécharge l'export
```
Le script lit tous les onglets (colonnes « Nomenclature RNCP » / « Code RNCP »), considère les codes comme actifs et sans
successeur, conserve le journal des runs, puis committer `state/snapshot.json` et lancer un run. Le run suivant
affiche les écarts par rapport à ce fichier (le 09/09/2026 : 2 704 ajoutés, 229 modifiés, 0 supprimé).

### Modifier le gabarit du site (voir §6), puis
`cp site/index.html docs/index.html` (idem `runs.html`) et committer les deux : le bot ne recopie le gabarit qu'au run suivant,
la copie manuelle permet une mise en ligne immédiate.

## 6. Pièges connus

- **Format du gabarit `site/index.html`** : page « bundler » Claude Design. Le vrai HTML est une chaîne JSON sur la ligne
  qui commence par `"<!DOCTYPE` (dans `<script type="__bundler/template">`). Pour l'éditer sans le casser :
  `json.loads` de cette ligne → modifier la chaîne → `json.dumps(h, ensure_ascii=False).replace('</', '<\\u002F')`.
  Le ré-encodage a été vérifié sans perte. Les balises meta anti-cache sont aussi dans l'en-tête statique du wrapper
  (lignes 3-7) car c'est lui qui est servi. Les blocs `__bundler/manifest` et `ext_resources` doivent rester `{}` et `[]`.
- **`overflow:hidden` sur `.table-carte` cassait le `position:sticky` des en-têtes** (décalés de deux lignes) → `overflow:clip`.
- **Infobulles** : un seul élément `#infobulle` en `position:fixed`, événements délégués sur `document`
  (`.th-aide`, `.aide`), car les puces du bandeau sont créées après le chargement des données.
- **Hook `rtk`** dans ce Claude Code : les sorties JSON de `curl` sont résumées en schéma ; écrire dans un fichier puis lire avec Python.
  `rtk find` ne supporte pas `-exec`.
- **API data.gouv** : sans User-Agent la réponse n'est pas du JSON.
- **Git** : ne jamais committer `docs/*.xlsx` ou `docs/data.json` générés localement (le bot s'en charge) ;
  ne jamais pousser sans accord explicite de l'utilisateur ; pas de trailer Co-Authored-By.
- Le fichier v1 `data/liens_cpf_formacode_rncp.xlsx` est encore dans l'historique git (dépôt public, non purgé, choix de l'utilisateur).
- Historique des événements plafonné à 10 000 (`HISTORIQUE_MAX`), journal des runs à 1 500 (`JOURNAL_MAX`).
- **Cron non déclenché le 10/09/2026 à 04:30 UTC** (lendemain du renommage du dépôt). Le fichier de workflow a été retouché et poussé
  le 10/09 pour réenregistrer la planification, puis un run manuel lancé à 08:26 UTC. Vérifier qu'un run automatique a bien eu lieu le 11/09 ;
  sinon, voir la page Actions du dépôt (workflow désactivé ?) et retoucher à nouveau le fichier.
- L'export France compétences du jour n'est pas toujours disponible à 04:30 UTC (le 10/09 à 08:30 UTC, le dernier publié était celui du 09/09) :
  un run « sans changement » sur l'export de la veille est normal.

## 7. Journal des décisions et réalisations (09/09/2026)

1. Test réel sous Docker : parseur validé sur le vrai export (25 668 fiches, RNCP35803 → successeur RNCP41832).
   En-tête de tableau décalé corrigé (`overflow:clip`).
2. Périmètre : les 700 formacodes couvraient 7 737 fiches dont 4 289 inactives → règle « actives seulement » (base 3 448).
3. Dépôt existant `rncp-watch` (v1 : `check_rncp.py`, `data/`) migré en v2, passé en **public** pour Pages, variables mail conservées,
   `SITE_URL` posée, premier run OK, cron quotidien réactivé (il avait été désactivé le 25/08/2026 sur la v1).
4. Site : infobulles sur les colonnes, filtres Type / Niveau, badges Niveau / Active, puces du bandeau explicitées, anti-cache.
5. Analyse du fichier « Liens CPF par formacode et diplômes.xlsx » (1 412 codes, 700 formacodes, identiques au périmètre) :
   175 remplacées, 22 inactives, 32 avec successeur publié, 60 échéances avant fin 2026, 2 704 actifs du périmètre absents.
   Livrable ponctuel : `tmp/analyse_liens_cpf.xlsx` (non versionné).
6. Référence réinitialisée depuis cet Excel (`outils/initialiser_depuis_excel.py`) ; règle « une fiche suivie reste suivie » ;
   **e-mail retiré** du workflow ; puces « Export France compétences » / « Périmètre suivi ».
7. Page Historique des générations (`runs.html`, `runs.json`), journal des runs, commit à chaque run, journalisation des échecs.
8. Évolution **cumulée** depuis la référence (choix utilisateur), alerte rouge si dernier run en échec, puce Référence ajoutée puis retirée.
9. Colonne « Dernier changement » + tri par colonnes ; sélecteur de période (défaut 7 jours) ; icône GitHub ; licence MIT.
10. Dépôt renommé `referentiel-rncp` ; `SITE_URL` et remote mis à jour.

### 10/09/2026

11. Site : colonnes « Formacodes » (codes complets) et « Formacode simplifié » (domaine à 3 chiffres, calculé côté site) ;
    liste déroulante sur le formacode simplifié (77 domaines). Une liste sur les formacodes complets a été ajoutée puis retirée
    (1 576 entrées, jugée inutilisable). En-têtes longs sur deux lignes pour que les onze colonnes tiennent dans la carte.
12. Run planifié de 04:30 UTC absent (lendemain du renommage) : fichier de workflow retouché et poussé, run manuel lancé à 08:26 UTC
    (sans changement, export du 09/09 car celui du 10/09 n'était pas encore publié sur data.gouv).

État au 10/09/2026 11:00 : 4 116 codes suivis (3 919 actifs, 197 inactifs), 2 704 ajoutés / 229 modifiés / 0 supprimé
par rapport à la référence, 6 générations au journal, prochain run automatique attendu le 11/09/2026 à 04:30 UTC (à vérifier, voir §6).
