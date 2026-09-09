# Contrôle RNCP — référentiel généré

Chaque jour, ce dépôt **génère** le référentiel des diplômes RNCP de votre périmètre depuis l'export open data officiel de France compétences (data.gouv.fr, mise à jour quotidienne, licence ouverte). Plus de fichier à maintenir : les diplômes nouvellement publiés, modifiés ou disparus sont détectés automatiquement, et un site web permet de chercher et consulter le tout.

## Ce que produit chaque génération

`docs/diplomes_rncp.xlsx` — le classeur complet. Onglet **Diplômes** : une ligne par fiche (code, intitulé, type, niveau, fiche active, échéance, remplacée par, formacodes, lien France compétences cliquable, évolution et son détail avant → après, date du dernier changement, suivi depuis), sous forme de tableau Excel filtrable. Code couleur par rapport à la génération précédente : **vert** = ligne ajoutée, **orange** = ligne modifiée, **rouge** = fiche désactivée ou disparue de l'export. Une fiche disparue reste affichée en rouge le temps d'une génération puis sort du fichier ; sa trace demeure dans l'onglet **Changements** (journal daté complet, lignes de la génération courante surlignées en bleu). L'onglet **Synthèse** donne les compteurs et la liste des lignes ayant évolué, et un onglet par domaine (« Formacode 110 », « Formacode 114 »…) reprend la vue par formacode simplifié.

Le gabarit du site vit dans `site/index.html` (design GIP FCIP produit avec Claude Design) et est recopié dans `docs/` à chaque génération. Si vous régénérez ce gabarit avec Claude Design, veillez à ce qu'il n'embarque pas de données d'exemple : les blocs `__bundler/manifest` et `__bundler/ext_resources` doivent être vides (`{}` et `[]`) pour que la page charge le vrai `data.json`.

`docs/index.html` + `docs/data.json` — le **site web** (servi par GitHub Pages) : recherche instantanée par code, intitulé ou formacode, filtres Ajoutés / Modifiés / Supprimés / Fiches inactives, détail par fiche avec lien vers France compétences, et bouton de téléchargement de l'Excel.

Le site et l'état ne sont committés que si des changements sont détectés. L'état de comparaison vit dans `state/snapshot.json`, versionné.

## Le périmètre

`config/formacodes.txt` définit ce qui est suivi : un formacode par ligne, 5 chiffres pour un formacode exact, 3 chiffres pour couvrir tout un domaine (ex. `114` couvre 11421, 11454…). Le fichier est initialisé avec les 700 formacodes de votre ancien fichier de liens CPF — le périmètre couvert est donc identique au départ, à la différence près que **les nouveaux diplômes publiés dans ces formacodes apparaîtront désormais tout seuls, en vert**. Un fichier vide (hors commentaires) suit le répertoire entier (~5 000 fiches actives) ; attendez-vous alors à des alertes quotidiennes nombreuses. Élargir ou réduire le périmètre = éditer ce fichier dans GitHub, rien d'autre.

Seules les fiches **actives** du périmètre entrent dans le suivi : les fiches déjà inactives (remplacées ou échues avant le début du suivi) sont ignorées, ce qui évite de traîner des milliers de fiches historiques. Une fiche suivie reste suivie quels que soient ses formacodes ; si elle se désactive, elle reste dans le fichier (ligne orange le jour du changement, puis « Fiche active : Non » avec son successeur) tant qu'elle figure dans l'export.

La base de référence peut être (ré)initialisée à partir d'un fichier Excel de liens formacode → RNCP avec `python outils/initialiser_depuis_excel.py fichier.xlsx` : ses codes deviennent la référence, considérés comme actifs et sans successeur ; la génération suivante signale en orange ceux qui ont été remplacés ou désactivés depuis, et en vert les diplômes du périmètre absents du fichier. Attention : environ 40 % des codes d'un tel fichier portent chez France compétences des formacodes différents de ceux du fichier ; c'est la règle « une fiche suivie reste suivie » qui garantit leur suivi.

Limite à connaître : une fiche sans aucun formacode renseigné chez France compétences est invisible d'un périmètre par formacodes.

## Mise en place (une fois, ~10 minutes)

1. Créer un dépôt GitHub et y déposer tout le contenu de ce dossier en conservant l'arborescence — attention aux éléments cachés `.github/` et `.gitignore` (⌘⇧. dans le Finder pour les voir, ou création directe dans l'interface GitHub).

2. Settings → Secrets and variables → Actions → Variables. Créer la variable `SITE_URL` avec l'adresse du site Pages (facultatif, reprise dans le résumé de chaque run).


3. Activer le site : Settings → **Pages** → Build and deployment → Source « Deploy from a branch » → branche `main`, dossier `/docs` → Save. L'URL affichée est celle à mettre dans `SITE_URL`. Point d'attention : sur un dépôt **privé**, GitHub Pages nécessite un plan payant (Pro/Team/Enterprise) ; le contenu du site étant exclusivement de l'open data France compétences, un dépôt public est une alternative acceptable si vous retirez toute donnée interne du dépôt.

4. Onglet **Actions** : activer les workflows si demandé, puis « Contrôle RNCP » → « Run workflow » pour la première génération (~2 minutes : téléchargement de l'export ~70 Mo, génération, commit de `docs/` et `state/`).

5. C'est tout. Génération quotidienne à 04:30 UTC (cron GitHub en UTC, léger glissement possible). En cas d'échec technique, GitHub notifie l'auteur du workflow.

## Au quotidien

Rien à faire : le fichier et le site se régénèrent seuls, l'équipe consulte le site pour chercher un code et télécharge l'Excel depuis le site. Les seules interventions possibles : éditer `config/formacodes.txt` pour ajuster le périmètre, et « Run workflow » pour forcer une génération. Le site et le fichier committés reflètent la dernière génération **avec changements** ; le contrôle tourne bien tous les jours même quand rien ne bouge.

## Exécution locale et options

```
pip install -r requirements.txt
python genere_rncp.py                                  # télécharge le dernier export
python genere_rncp.py --xml export-fiches-rncp.xml     # export déjà téléchargé
python genere_rncp.py --debug-fiche RNCP35803          # affiche le XML brut d'une fiche
```

`--min-fiches` (défaut 5000) fait échouer la génération si l'export paraît anormalement petit — garde-fou contre un changement de schéma silencieux chez France compétences. Le lecteur XML est volontairement tolérant (codes par motif, deux formes acceptées pour les fiches de remplacement, niveau extrait défensivement) ; si un jour le format change vraiment, `--debug-fiche` montre la structure reçue en dix secondes.

## Migration depuis la v1

L'ancien mode (comparaison d'un fichier Excel maintenu à la main) est remplacé : supprimer `check_rncp.py` et le dossier `data/` du dépôt. L'ancien `state/snapshot.json` est détecté et ignoré, la première génération v2 repart d'une base de référence propre.
