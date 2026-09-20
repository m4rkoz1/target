@echo off
title Enviar para o GitHub
cd /d "%~dp0"
echo =========================================================
echo       SUBINDO PROJETO TARGET SSW PARA O GITHUB
echo =========================================================
echo.
git remote -v | findstr origin >nul
if %errorlevel% neq 0 (
    echo Nenhum repositorio remoto configurado.
    set /p REPO_URL="Cole a URL do seu repositorio no GitHub (ex: https://github.com/usuario/repo.git): "
    git remote add origin %REPO_URL%
)
echo.
echo Enviando arquivos para o GitHub...
git add .
git commit -m "update: automacao de lancamento SSW Tela 475" 2>nul
git branch -M main
git push -u origin main
echo.
echo Processo finalizado!
pause
