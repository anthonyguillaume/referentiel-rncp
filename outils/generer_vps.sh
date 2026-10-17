#!/usr/bin/env bash
# Génération quotidienne sur le VPS : mêmes étapes que .github/workflows/controle-rncp.yml.
# Crontab (le cron Ubuntu suit le fuseau du système, ici Europe/Paris : minuit à Paris toute l'année) :
#   5 0 * * * flock -n /tmp/rncp.lock ~/referentiel-rncp/outils/generer_vps.sh >> ~/rncp.log 2>&1
set -uo pipefail
cd "$(dirname "$0")/.."

git pull --rebase --quiet || exit 1
docker build -q -t rncp-watch . >/dev/null || exit 1

# Monté en volume : le conteneur travaille sur le clone, avec l'utilisateur courant (fichiers non root).
lancer() { docker run --rm --user "$(id -u):$(id -g)" -v "$PWD:/app" -e SITE_URL -e GITHUB_OUTPUT=/app/.sortie rncp-watch "$@"; }
export SITE_URL="${SITE_URL:-https://anthonyguillaume.github.io/referentiel-rncp/}"

rm -f .sortie
if ! lancer python genere_rncp.py 2>&1 | tee generation.log; then
  lancer python outils/journaliser_erreur.py generation.log
fi

resume=$(sed -n 's/^resume=//p' .sortie 2>/dev/null)
[ -n "$resume" ] || resume="échec de la génération"
git add state docs
git diff --cached --quiet || git commit -q -m "Génération RNCP $(date -u +%F) — $resume"
git pull --rebase --quiet || true
git push --quiet
