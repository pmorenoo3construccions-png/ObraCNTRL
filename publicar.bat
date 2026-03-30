@echo off
REM ════════════════════════════════════════════════════════════
REM  publicar.bat — OBRACTRL · O3 Construccions
REM  Actualitza el dashboard i el puja a GitHub Pages
REM  Doble-clic per executar manualment, o programat cada divendres
REM ════════════════════════════════════════════════════════════

echo.
echo  ╔══════════════════════════════════════════╗
echo  ║  OBRACTRL — Publicant dashboard...       ║
echo  ╚══════════════════════════════════════════╝
echo.

REM ── Canviar al directori on es troba aquest .bat ──────────
REM    (la carpeta del repo GitHub, p.ex. Documents\GitHub\ObraCNTRL)
cd /d "%~dp0"

REM ── 1. Generar dashboard amb dades fresques dels Excel ─────
REM    (el script llegeix els Excel de GOOGLE DRIVE\OBRAS PEDRO
REM     i escriu index.html en aquesta mateixa carpeta)
echo [1/3] Llegint Excel i generant index.html...
py "%~dp0generar_dashboard.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  ❌ ERROR generant el dashboard. Comprova els Excel i Python.
    echo     Possibles causes:
    echo     - Python no instal·lat o no al PATH
    echo     - Falta: pip install openpyxl xlrd
    echo     - Algun Excel obert per un altre programa
    pause
    exit /b 1
)

REM ── 2. Git: afegir index.html i fer commit ─────────────────
echo.
echo [2/3] Fent commit a Git...
git add index.html
git commit -m "Dashboard auto-update %date% %time:~0,5%"
if %ERRORLEVEL% NEQ 0 (
    echo  ℹ️  Sense canvis nous al dashboard (git sense modificacions).
)

REM ── 3. Pujar a GitHub ──────────────────────────────────────
echo.
echo [3/3] Pujant a GitHub Pages...
git push origin main
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  ❌ ERROR pujant a GitHub.
    echo     Comprova la connexió i les credencials de Git.
    echo     Executa: git remote -v  per verificar la URL del repositori.
    pause
    exit /b 1
)

echo.
echo  ✅ Dashboard publicat correctament!
echo     Visible a: https://pmorenoo3construccions-png.github.io/ObraCNTRL/
echo.
timeout /t 4 /nobreak > nul
