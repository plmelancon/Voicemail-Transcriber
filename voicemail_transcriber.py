# -*- coding: utf-8 -*-
"""
Voicemail Transcriber pour Outlook
Transcrit automatiquement les messages vocaux (WAV) recus par courriel
Utilise OpenAI Whisper pour la transcription

Auteur: Claude pour Pierre-Luc Melancon
"""

import os
import sys
import tempfile
import logging
from datetime import datetime, timedelta
from pathlib import Path

# Force UTF-8 output (sys.stdout is None when frozen with console=False)
if sys.stdout is not None and sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# Frozen-aware app directory (works both dev and PyInstaller)
def _get_app_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent

# Configuration du logging — logs go to %LOCALAPPDATA% (writable even when app is in Program Files)
LOG_DIR = Path(os.environ.get('LOCALAPPDATA', str(_get_app_dir()))) / "VoicemailTranscriber" / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / f"voicemail_transcriber_{datetime.now().strftime('%Y-%m')}.log"

_handlers = [logging.FileHandler(LOG_FILE, encoding='utf-8')]
if sys.stdout is not None:
    _handlers.append(logging.StreamHandler(sys.stdout))

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=_handlers,
)
logger = logging.getLogger(__name__)

# ============================================================================
# CONFIGURATION - Modifier selon vos besoins
# ============================================================================

CONFIG = {
    # Modele Whisper: tiny, base, small, medium, large, turbo
    # turbo = large-v3-turbo (rapide et precis, recommande)
    "whisper_model": "turbo",
    
    # Langue (fr pour francais, None pour detection auto)
    "language": "fr",
    
    # Mots-cles dans le sujet pour identifier les messages vocaux
    # Ajoutez d'autres termes selon votre fournisseur
    "voicemail_keywords": [
        "voicemail", "voice mail", "messagerie vocale", 
        "message vocal", "vm:", "vmail", "missed call",
        "appel manque", "boite vocale"
    ],
    
    # Extensions audio supportees
    "audio_extensions": [".wav", ".mp3", ".m4a", ".ogg", ".wma"],
    
    # Nombre de jours a remonter pour chercher des messages
    "days_to_search": 7,
    
    # Marqueur ajoute au sujet apres transcription
    "processed_marker": "[TRANSCRIT]",
    
    # Prefixe pour la transcription dans le corps du message
    "transcription_header": "\n\n" + "="*50 + "\n[TRANSCRIPTION DU MESSAGE VOCAL]\n" + "="*50 + "\n\n",
    "transcription_footer": "\n\n" + "="*50 + "\n[Transcrit automatiquement par Whisper]\n" + "="*50
}

# ============================================================================
# CLASSES ET FONCTIONS
# ============================================================================

class WhisperTranscriber:
    """Gere la transcription audio avec Whisper"""
    
    def __init__(self, model_name: str = "medium", language: str = "fr"):
        self.model_name = model_name
        self.language = language
        self.model = None
        
    def load_model(self):
        """Charge le modele Whisper (lazy loading)"""
        if self.model is None:
            logger.info(f"Chargement du modele Whisper '{self.model_name}'...")
            try:
                import whisper
            except ImportError:
                raise RuntimeError(
                    "openai-whisper n'est pas installe. "
                    "Utilisez voicemail_service.py qui utilise faster-whisper."
                )
            self.model = whisper.load_model(self.model_name)
            logger.info("Modele charge avec succes")
        return self.model
    
    def transcribe(self, audio_path: str) -> str:
        """Transcrit un fichier audio en texte"""
        model = self.load_model()
        logger.info(f"Transcription de: {audio_path}")
        
        try:
            result = model.transcribe(
                audio_path,
                language=self.language,
                task="transcribe"
            )
            text = result["text"].strip()
            logger.info(f"Transcription terminee: {len(text)} caracteres")
            return text
        except Exception as e:
            logger.error(f"Erreur de transcription: {e}")
            raise


class OutlookVoicemailProcessor:
    """Gere l'interaction avec Outlook pour traiter les messages vocaux"""
    
    def __init__(self, config: dict):
        self.config = config
        self.transcriber = WhisperTranscriber(
            model_name=config["whisper_model"],
            language=config["language"]
        )
        self.outlook = None
        self.namespace = None
        
    def connect_outlook(self):
        """Etablit la connexion avec Outlook"""
        if self.outlook is None:
            import win32com.client
            logger.info("Connexion a Outlook...")
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            logger.info("Connecte a Outlook")
        return self.outlook, self.namespace

    def is_voicemail(self, mail_item) -> bool:
        """Verifie si un courriel est un message vocal (par le sujet uniquement)"""
        subject = mail_item.Subject.lower() if mail_item.Subject else ""
        
        # Verifie si deja traite
        if self.config["processed_marker"].lower() in subject:
            return False
        
        # Verifie les mots-cles dans le sujet UNIQUEMENT
        # Ne traite PAS les courriels avec audio qui ne sont pas des messages vocaux
        for keyword in self.config["voicemail_keywords"]:
            if keyword.lower() in subject:
                # Confirme qu'il y a une piece jointe audio
                if mail_item.Attachments.Count > 0:
                    for i in range(1, mail_item.Attachments.Count + 1):
                        att = mail_item.Attachments.Item(i)
                        ext = Path(att.FileName).suffix.lower()
                        if ext in self.config["audio_extensions"]:
                            return True
                return False
        
        return False


    def extract_audio_attachment(self, mail_item) -> tuple:
        """Extrait la piece jointe audio vers un fichier temporaire"""
        for i in range(1, mail_item.Attachments.Count + 1):
            att = mail_item.Attachments.Item(i)
            ext = Path(att.FileName).suffix.lower()
            
            if ext in self.config["audio_extensions"]:
                temp_dir = tempfile.mkdtemp()
                temp_path = os.path.join(temp_dir, att.FileName)
                att.SaveAsFile(temp_path)
                logger.info(f"Piece jointe extraite: {att.FileName}")
                return temp_path, att.FileName
        
        return None, None
    
    def add_transcription_to_email(self, mail_item, transcription: str, audio_filename: str):
        """Ajoute la transcription au corps du courriel"""
        try:
            header = self.config["transcription_header"]
            footer = self.config["transcription_footer"]
            
            transcription_block = f"{header}{transcription}{footer}"
            transcription_block += f"\n[Fichier source: {audio_filename}]"
            
            if mail_item.BodyFormat == 2:  # HTML
                html_transcription = transcription_block.replace("\n", "<br>")
                html_block = f"<div style='background-color:#f0f0f0; padding:10px; margin:10px 0; border-left:4px solid #007bff;'><pre style='white-space:pre-wrap;'>{html_transcription}</pre></div>"
                mail_item.HTMLBody = mail_item.HTMLBody + html_block
            else:
                mail_item.Body = mail_item.Body + transcription_block
            
            if self.config["processed_marker"] not in mail_item.Subject:
                mail_item.Subject = f"{self.config['processed_marker']} {mail_item.Subject}"
            
            mail_item.Save()
            logger.info(f"Transcription ajoutee au courriel: {mail_item.Subject}")
            return True
            
        except Exception as e:
            logger.error(f"Erreur lors de l'ajout de la transcription: {e}")
            return False


    def process_single_email(self, mail_item) -> bool:
        """Traite un seul courriel de messagerie vocale"""
        subject = mail_item.Subject or "(Sans sujet)"
        logger.info(f"Traitement: {subject}")
        
        audio_path, audio_filename = self.extract_audio_attachment(mail_item)
        if not audio_path:
            logger.warning(f"Aucune piece jointe audio trouvee dans: {subject}")
            return False
        
        try:
            transcription = self.transcriber.transcribe(audio_path)
            
            if transcription:
                success = self.add_transcription_to_email(mail_item, transcription, audio_filename)
                return success
            else:
                logger.warning("Transcription vide")
                return False
                
        finally:
            try:
                if audio_path and os.path.exists(audio_path):
                    os.remove(audio_path)
                    temp_dir = os.path.dirname(audio_path)
                    if os.path.exists(temp_dir):
                        os.rmdir(temp_dir)
            except Exception as e:
                logger.warning(f"Impossible de supprimer le fichier temporaire: {e}")
    
    def get_inbox_folder(self):
        """Recupere le dossier Boite de reception"""
        _, namespace = self.connect_outlook()
        return namespace.GetDefaultFolder(6)  # 6 = olFolderInbox


    def find_voicemails(self, days_back: int = None) -> list:
        """Trouve tous les messages vocaux non traites"""
        if days_back is None:
            days_back = self.config["days_to_search"]
        
        inbox = self.get_inbox_folder()
        
        cutoff_date = datetime.now() - timedelta(days=days_back)
        cutoff_str = cutoff_date.strftime("%m/%d/%Y")
        
        restriction = f"[ReceivedTime] >= '{cutoff_str}'"
        
        voicemails = []
        items = inbox.Items.Restrict(restriction)
        items.Sort("[ReceivedTime]", True)
        
        logger.info(f"Recherche de messages vocaux depuis {cutoff_str}...")
        
        for item in items:
            try:
                if self.is_voicemail(item):
                    voicemails.append(item)
                    logger.debug(f"Message vocal trouve: {item.Subject}")
            except Exception as e:
                logger.warning(f"Erreur lors de l'analyse d'un element: {e}")
                continue
        
        logger.info(f"{len(voicemails)} message(s) vocal(aux) trouve(s)")
        return voicemails
    
    def process_all_voicemails(self, days_back: int = None) -> dict:
        """Traite tous les messages vocaux non traites"""
        results = {"processed": 0, "failed": 0, "skipped": 0}
        
        voicemails = self.find_voicemails(days_back)
        
        if not voicemails:
            logger.info("Aucun message vocal a traiter")
            return results
        
        for mail in voicemails:
            try:
                success = self.process_single_email(mail)
                if success:
                    results["processed"] += 1
                else:
                    results["failed"] += 1
            except Exception as e:
                logger.error(f"Erreur lors du traitement: {e}")
                results["failed"] += 1
        
        logger.info(f"Traitement termine: {results}")
        return results


# ============================================================================
# POINT D'ENTREE
# ============================================================================

def main():
    """Fonction principale"""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Transcrit automatiquement les messages vocaux recus par courriel"
    )
    parser.add_argument(
        "--days", "-d",
        type=int,
        default=CONFIG["days_to_search"],
        help=f"Nombre de jours a remonter (defaut: {CONFIG['days_to_search']})"
    )
    parser.add_argument(
        "--model", "-m",
        type=str,
        default=CONFIG["whisper_model"],
        choices=["tiny", "base", "small", "medium", "large", "turbo"],
        help=f"Modele Whisper a utiliser (defaut: {CONFIG['whisper_model']})"
    )
    parser.add_argument(
        "--test", "-t",
        action="store_true",
        help="Mode test: affiche les messages vocaux trouves sans les traiter"
    )
    
    args = parser.parse_args()
    CONFIG["whisper_model"] = args.model
    
    print("\n" + "="*60)
    print(" TRANSCRIPTEUR DE MESSAGES VOCAUX OUTLOOK")
    print("="*60 + "\n")
    
    processor = OutlookVoicemailProcessor(CONFIG)
    
    if args.test:
        print("[TEST] Recherche de messages vocaux...\n")
        voicemails = processor.find_voicemails(args.days)
        
        if voicemails:
            print(f"\n{len(voicemails)} message(s) vocal(aux) trouve(s):\n")
            for i, mail in enumerate(voicemails, 1):
                received = mail.ReceivedTime.strftime("%Y-%m-%d %H:%M")
                print(f"  {i}. [{received}] {mail.Subject}")
                for j in range(1, mail.Attachments.Count + 1):
                    att = mail.Attachments.Item(j)
                    print(f"      -> {att.FileName}")
        else:
            print("Aucun message vocal non traite trouve.")
    else:
        print(f"Traitement des messages vocaux des {args.days} derniers jours...\n")
        print(f"   Modele Whisper: {args.model}")
        print(f"   Langue: {CONFIG['language']}\n")
        
        results = processor.process_all_voicemails(args.days)
        
        print("\n" + "="*60)
        print(" RESUME")
        print("="*60)
        print(f"   Traites avec succes: {results['processed']}")
        print(f"   Echecs: {results['failed']}")
        print("="*60 + "\n")


if __name__ == "__main__":
    main()
