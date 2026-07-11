# Contrôle RNCP

Compare chaque jour les codes RNCP du fichier « Liens CPF par formacode et diplômes » avec l'export open data officiel de France compétences (mis à jour quotidiennement sur data.gouv.fr, licence ouverte), et envoie un e-mail uniquement quand quelque chose a changé.

## Ce que produit chaque exécution

Une copie annotée du classeur (`output/Liens_CPF_controle.xlsx`) — le fichier d'origine n'est jamais modifié. L'onglet principal reçoit huit colonnes supplémentaires (Statut, Contrôle France compétences, Code(s) à jour, Intitulé à jour, Échéance fiche, Formacodes de la fiche à jour, Changé ?, Changement depuis le dernier envoi) et chaque ligne est colorée selon son statut. Les onglets par formacode sont également colorés. Deux onglets sont ajoutés en tête : **Synthèse** (compteurs + liste des codes à traiter) et **Changements** (journal daté de toutes les évolutions détectées depuis la mise en place).

Ce qui vient de changer est mis en évidence **en bleu** à trois endroits : la colonne « Changé ? » (OUI/NON) suivie du détail « Changement depuis le dernier envoi » (avant → après) sur l'onglet principal, la puce ● en tête de la Synthèse (les nouveautés remontent en premier dans chaque statut), et les lignes surlignées de l'onglet Changements. Les trois onglets sont équipés de filtres automatiques Excel : pour ne voir que ce qui vient de bouger, filtrer « Changé ? » sur OUI ; une ligne orange avec « Changé ? » à NON est un point déjà signalé lors d'un envoi précédent.

Les statuts : **vert** (RAS — fiche active, intitulé conforme), **orange** (à vérifier — intitulé officiel différent, successeur déjà publié alors que la fiche est encore active, ou échéance à moins de 180 jours), **rouge** (fiche inactive : remplacée, avec le ou les nouveaux codes indiqués en suivant la chaîne de remplacement jusqu'à la fiche active, ou expirée sans successeur), **gris** (code introuvable dans le répertoire, saisie à vérifier).

L'e-mail n'est envoyé que si des changements sont détectés par rapport à la veille (l'état de référence est conservé dans `state/snapshot.json`, versionné dans le dépôt). La toute première exécution envoie un état des lieux complet.

## Mise en place (une fois, ~10 minutes)

1. Créer un dépôt GitHub **privé** (le fichier contient des données métier) et y déposer tout le contenu de ce dossier — soit par `git push`, soit via l'interface web (« Add file → Upload files », en conservant l'arborescence, notamment `.github/workflows/`).

2. Dans le dépôt : Settings → Secrets and variables → Actions → « New repository secret », créer les secrets SMTP ci-dessous.

| Où | Nom | Contenu | Exemple |
|---|---|---|---|
| Secret | `MAIL_SERVER` | serveur SMTP | `smtp-relay.brevo.com` |
| Secret | `MAIL_PORT` | port SMTP | `587` |
| Secret | `MAIL_USERNAME` | identifiant SMTP | `xxx@smtp-brevo.com` |
| Secret | `MAIL_PASSWORD` | mot de passe / clé SMTP | — |
| **Variable** | `MAIL_TO` | destinataires, **séparés par des virgules** | `marielle@…, anthony@…, equipe@…` |
| Secret | `MAIL_FROM` | expéditeur (facultatif, défaut : `MAIL_USERNAME`) | `noreply@…` |

Les destinataires se saisissent dans l'onglet **Variables** (même écran que les secrets) : la liste reste lisible et modifiable par toute l'équipe, sans passer par un secret masqué. Ajouter ou retirer une adresse = éditer la variable, rien d'autre. (`MAIL_TO` en secret fonctionne aussi, la variable est prioritaire.)

Fournisseurs qui fonctionnent bien : **Brevo** (300 mails/jour gratuits, le plus simple), **Gmail** avec un mot de passe d'application, ou le SMTP **Microsoft 365** de l'entreprise si l'authentification SMTP y est autorisée. Pour un serveur en TLS implicite (port 465), décommenter `secure: true` dans le workflow.

3. Onglet **Actions** du dépôt : activer les workflows si GitHub le demande, ouvrir « Contrôle RNCP » → « Run workflow » pour la première exécution. Elle télécharge l'export du jour (~70 Mo), établit l'état de référence, committe `state/` et `output/`, et envoie l'e-mail d'état des lieux.

4. À noter : l'état de référence n'est enregistré qu'après l'envoi réussi du mail — en cas d'échec SMTP, le job passe en erreur et les mêmes changements sont re-signalés au contrôle suivant, aucune alerte n'est perdue.

5. C'est tout : le contrôle tourne ensuite chaque jour à 04:30 UTC (06:30 à Paris l'été). Le cron GitHub est en UTC et peut glisser de quelques minutes. En cas d'échec technique (data.gouv indisponible, schéma modifié…), GitHub notifie automatiquement l'auteur du workflow.

## Au quotidien

Quand le fichier de référence évolue (nouveau diplôme, lien CPF ajouté…), remplacer `data/liens_cpf_formacode_rncp.xlsx` directement dans l'interface GitHub (ouvrir le dossier `data/`, « Add file → Upload files », écraser). Le contrôle suivant prendra la nouvelle version ; tout `.xlsx` présent dans `data/` fait l'affaire, le nom importe peu. Le dernier rapport est toujours lisible dans `output/` sur GitHub, joint à chaque e-mail, et conservé 90 jours en artefact de workflow. Pour recevoir le rapport sans attendre un changement : « Run workflow » en cochant « Envoyer le mail même sans changement ».

## Exécution locale et options

```
pip install -r requirements.txt
python check_rncp.py                                   # télécharge le dernier export
python check_rncp.py --xml export-fiches-rncp.xml      # export déjà téléchargé
python check_rncp.py --xml tests/fiches-exemple.xml --min-couverture 0   # démo hors ligne
python check_rncp.py --debug-fiche RNCP35803           # affiche le XML brut d'une fiche
```

`--seuil-jours` règle l'alerte orange d'échéance proche (défaut 180). `--min-couverture` (défaut 0.5) fait échouer l'exécution si moins de la moitié des codes du fichier sont retrouvés dans l'export — c'est le garde-fou contre un changement de schéma silencieux chez France compétences.

## Points d'attention

Le lecteur XML est volontairement tolérant (recherche des codes par motif, deux formes acceptées pour les fiches de remplacement) car France compétences fait évoluer son schéma d'export de temps en temps ; si un jour la couverture chute, lancer `--debug-fiche` sur un code connu montre immédiatement la structure reçue. Les colonnes E et G du fichier source sont des formules : la copie annotée force leur recalcul à l'ouverture dans Excel (elles peuvent apparaître vides dans un simple aperçu). Enfin, l'intitulé officiel est comparé après normalisation (accents, casse, tirets, espaces) et en tenant compte du préfixe de type de diplôme (« BTS - … ») ; un orange « intitulé différent » signale donc un vrai écart de fond, pas une différence typographique.
