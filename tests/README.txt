Démo hors ligne (sans accès à data.gouv.fr) :
  python genere_rncp.py --xml tests/generation-1.xml --min-fiches 0 --state /tmp/s.json --outdir /tmp/demo1
  python genere_rncp.py --xml tests/generation-2.xml --min-fiches 0 --state /tmp/s.json --outdir /tmp/demo2
La génération 2 contient 2 ajouts, 3 modifications et 2 suppressions par rapport à la 1.
