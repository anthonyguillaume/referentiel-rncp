# Image de test locale, alignée sur le runner GitHub Actions (Python 3.12)
FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
# Exécution par défaut : génération complète (téléchargement réel data.gouv)
CMD ["python", "genere_rncp.py"]
