@echo off
title Target SSW - Chrome com CDP (porta 9222)
echo =========================================================
echo   Feche TODAS as janelas do Chrome antes de continuar
echo   Depois este script abre o Chrome pronto p/ automacao
echo =========================================================
echo.
pause
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --start-maximized
echo Chrome aberto com CDP na porta 9222.
echo Agora faca login no SSW e rode rodar_automacao_playwright.bat
pause
