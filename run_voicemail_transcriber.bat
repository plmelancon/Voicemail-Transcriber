@echo off
chcp 65001 > nul
title Transcripteur de Messages Vocaux

echo.
echo ========================================
echo  Transcripteur de Messages Vocaux
echo ========================================
echo.
echo Appuyez sur une touche pour lancer le traitement...
pause > nul

cd /d "%~dp0"
python voicemail_transcriber.py %*

echo.
echo ========================================
echo Traitement termine!
echo ========================================
echo.
pause
