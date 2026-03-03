@echo off
chcp 65001 > nul
title Transcripteur Vocal - Service
cd /d "%~dp0"
python voicemail_service.py %*
pause
