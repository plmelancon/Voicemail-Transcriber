#!/usr/bin/env python3
"""
Soumet un fichier installateur à VirusTotal pour analyse.

Nécessite la variable d'environnement VIRUSTOTAL_API_KEY ou
la clé dans le registre Windows: HKCU\Environment\VIRUSTOTAL_API_KEY

Usage:
    python submit_to_virustotal.py Output/OutlookEmailSaver_Setup_X.X.X.exe

ATTENTION: Les fichiers soumis à VirusTotal deviennent PUBLICS.
Ne JAMAIS soumettre de fichiers contenant des données client.
"""

import os
import sys
import time
import hashlib
from pathlib import Path

try:
    import requests
except ImportError:
    print("ERREUR: Le module 'requests' est requis.")
    print("Installez-le avec: pip install requests")
    sys.exit(1)


VIRUSTOTAL_API_URL = "https://www.virustotal.com/api/v3"
ALLOWED_EXTENSIONS = {".exe"}
FORBIDDEN_PATTERNS = ["config.json", "settings.json", ".msg", ".eml", ".pdf"]


def get_api_key() -> str:
    """Récupère la clé API depuis l'environnement ou le registre Windows."""
    # Essayer la variable d'environnement
    api_key = os.environ.get("VIRUSTOTAL_API_KEY")
    if api_key:
        return api_key.strip()

    # Essayer le registre Windows
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment") as key:
            api_key, _ = winreg.QueryValueEx(key, "VIRUSTOTAL_API_KEY")
            if api_key:
                return api_key.strip()
    except (ImportError, FileNotFoundError, OSError):
        pass

    return ""


def validate_file(file_path: Path) -> bool:
    """
    Vérifie que le fichier est sécuritaire à soumettre.

    Retourne True si le fichier peut être soumis, False sinon.
    """
    # Vérifier l'extension
    if file_path.suffix.lower() not in ALLOWED_EXTENSIONS:
        print(f"ERREUR: Seuls les fichiers {ALLOWED_EXTENSIONS} peuvent être soumis.")
        print(f"Fichier fourni: {file_path.suffix}")
        return False

    # Vérifier le nom du fichier
    filename_lower = file_path.name.lower()
    for pattern in FORBIDDEN_PATTERNS:
        if pattern in filename_lower:
            print(f"ERREUR: Le fichier contient un pattern interdit: {pattern}")
            return False

    # Vérifier que c'est bien un installateur VoicemailTranscriber
    if "voicemailtranscriber_setup" not in filename_lower:
        print("AVERTISSEMENT: Ce fichier ne semble pas être un installateur VoicemailTranscriber.")
        return False

    return True


def calculate_sha256(file_path: Path) -> str:
    """Calcule le hash SHA256 du fichier."""
    sha256_hash = hashlib.sha256()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            sha256_hash.update(chunk)
    return sha256_hash.hexdigest()


def check_existing_report(api_key: str, sha256: str) -> dict | None:
    """Vérifie si un rapport existe déjà pour ce hash."""
    headers = {"x-apikey": api_key}
    url = f"{VIRUSTOTAL_API_URL}/files/{sha256}"

    try:
        response = requests.get(url, headers=headers, timeout=30)
        if response.status_code == 200:
            return response.json()
        return None
    except requests.RequestException:
        return None


def upload_file(api_key: str, file_path: Path) -> str:
    """
    Soumet le fichier à VirusTotal.

    Retourne l'ID d'analyse.
    """
    headers = {"x-apikey": api_key}

    # Pour les fichiers > 32MB, il faut une URL d'upload spéciale
    file_size = file_path.stat().st_size
    if file_size > 32 * 1024 * 1024:
        # Obtenir l'URL d'upload pour gros fichiers
        url = f"{VIRUSTOTAL_API_URL}/files/upload_url"
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        upload_url = response.json()["data"]
    else:
        upload_url = f"{VIRUSTOTAL_API_URL}/files"

    print(f"Upload en cours... ({file_size / 1024 / 1024:.2f} MB)")

    with open(file_path, "rb") as f:
        files = {"file": (file_path.name, f)}
        response = requests.post(upload_url, headers=headers, files=files, timeout=300)

    response.raise_for_status()
    data = response.json()

    return data["data"]["id"]


def wait_for_analysis(api_key: str, analysis_id: str, timeout: int = 600) -> dict:
    """
    Attend la fin de l'analyse et retourne les résultats.

    Args:
        api_key: Clé API VirusTotal
        analysis_id: ID de l'analyse
        timeout: Temps maximum d'attente en secondes (défaut: 10 minutes)
    """
    headers = {"x-apikey": api_key}
    url = f"{VIRUSTOTAL_API_URL}/analyses/{analysis_id}"

    start_time = time.time()

    while time.time() - start_time < timeout:
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        data = response.json()

        status = data["data"]["attributes"]["status"]

        if status == "completed":
            return data
        elif status == "queued":
            print("En file d'attente...")
        elif status == "in-progress":
            print("Analyse en cours...")

        time.sleep(15)  # Attendre 15 secondes entre chaque vérification

    raise TimeoutError(f"L'analyse n'a pas terminé dans les {timeout} secondes.")


def print_report(results: dict, sha256: str) -> None:
    """Affiche un rapport lisible des résultats."""
    stats = results["data"]["attributes"]["stats"]

    print("\n" + "=" * 60)
    print("RAPPORT VIRUSTOTAL")
    print("=" * 60)
    print(f"SHA256: {sha256}")
    print(f"Lien:   https://www.virustotal.com/gui/file/{sha256}")
    print("-" * 60)
    print(f"Malveillant:  {stats.get('malicious', 0)}")
    print(f"Suspect:      {stats.get('suspicious', 0)}")
    print(f"Non détecté:  {stats.get('undetected', 0)}")
    print(f"Inoffensif:   {stats.get('harmless', 0)}")
    print(f"Échec:        {stats.get('failure', 0)}")
    print(f"Non supporté: {stats.get('type-unsupported', 0)}")
    print("=" * 60)

    # Afficher les détections si présentes
    malicious = stats.get('malicious', 0)
    suspicious = stats.get('suspicious', 0)

    if malicious > 0 or suspicious > 0:
        print("\nDÉTECTIONS:")
        results_dict = results["data"]["attributes"].get("results", {})
        for engine, result in results_dict.items():
            category = result.get("category", "")
            if category in ("malicious", "suspicious"):
                print(f"  - {engine}: {result.get('result', 'N/A')} ({category})")
    else:
        print("\n[OK] Aucune detection malveillante!")

    print()


def main():
    """Point d'entrée principal."""
    if len(sys.argv) < 2:
        print("Usage: python submit_to_virustotal.py <chemin_vers_installateur.exe>")
        print("\nExemple:")
        print("  python submit_to_virustotal.py Output/OutlookEmailSaver_Setup_0.0.30.exe")
        sys.exit(1)

    file_path = Path(sys.argv[1])

    # Vérifier que le fichier existe
    if not file_path.exists():
        print(f"ERREUR: Fichier introuvable: {file_path}")
        sys.exit(1)

    # Obtenir la clé API
    api_key = get_api_key()
    if not api_key:
        print("ERREUR: Clé API VirusTotal non trouvée.")
        print("\nConfigurez la clé avec:")
        print('  [Environment]::SetEnvironmentVariable("VIRUSTOTAL_API_KEY", "votre-clé", "User")')
        print("\nOu définissez la variable d'environnement VIRUSTOTAL_API_KEY")
        sys.exit(1)

    # Valider le fichier (sécurité données client)
    if not validate_file(file_path):
        sys.exit(1)

    print(f"Fichier: {file_path}")
    print(f"Taille:  {file_path.stat().st_size / 1024 / 1024:.2f} MB")

    # Calculer le hash
    print("Calcul du hash SHA256...")
    sha256 = calculate_sha256(file_path)
    print(f"SHA256:  {sha256}")

    # Vérifier si un rapport existe déjà
    print("\nVérification des rapports existants...")
    existing = check_existing_report(api_key, sha256)

    if existing:
        print("Un rapport existe déjà pour ce fichier.")
        last_analysis = existing["data"]["attributes"].get("last_analysis_stats", {})
        if last_analysis:
            print_report({"data": {"attributes": {"stats": last_analysis}}}, sha256)
        print("Abandon: rapport existant pour ce SHA256, pas de re-soumission automatique.")
        sys.exit(0)

    # Soumettre le fichier
    print("\nSoumission à VirusTotal...")
    try:
        analysis_id = upload_file(api_key, file_path)
        print(f"Upload réussi! ID d'analyse: {analysis_id}")
    except requests.RequestException as e:
        print(f"ERREUR lors de l'upload: {e}")
        sys.exit(1)

    # Attendre les résultats
    print("\nAttente des résultats d'analyse (peut prendre quelques minutes)...")
    try:
        results = wait_for_analysis(api_key, analysis_id)
        print_report(results, sha256)
    except TimeoutError as e:
        print(f"\nATTENTION: {e}")
        print(f"Vous pouvez consulter les résultats plus tard:")
        print(f"https://www.virustotal.com/gui/file/{sha256}")
    except requests.RequestException as e:
        print(f"ERREUR lors de la récupération des résultats: {e}")
        print(f"Consultez les résultats manuellement:")
        print(f"https://www.virustotal.com/gui/file/{sha256}")


if __name__ == "__main__":
    main()
