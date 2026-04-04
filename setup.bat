@echo off
chcp 65001 >nul
echo ============================================
echo   Installation — Bulletins Scolaires
echo ============================================
echo.

:: Vérifier Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERREUR] Python n'est pas installe.
    echo Telechargez-le sur https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Créer le dossier fonts
if not exist "static\fonts" mkdir "static\fonts"

:: Installer les dépendances
echo [1/3] Installation des dependances Python...
python -m pip install --upgrade pip >nul 2>&1
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERREUR] Echec de l'installation des dependances.
    pause
    exit /b 1
)
echo     OK

:: Télécharger les polices
echo [2/3] Telechargement des polices...
python -c "
import urllib.request, os, sys

fonts_dir = os.path.join('static', 'fonts')

fonts = {
    'Amiri-Regular.ttf':   'https://github.com/aliftype/amiri/raw/main/fonts/Amiri-Regular.ttf',
    'Amiri-Bold.ttf':      'https://github.com/aliftype/amiri/raw/main/fonts/Amiri-Bold.ttf',
    'DejaVuSans.ttf':      'https://cdn.jsdelivr.net/npm/dejavu-fonts-ttf@2.37.3/ttf/DejaVuSans.ttf',
    'DejaVuSans-Bold.ttf': 'https://cdn.jsdelivr.net/npm/dejavu-fonts-ttf@2.37.3/ttf/DejaVuSans-Bold.ttf',
}

headers = {'User-Agent': 'Mozilla/5.0'}
for name, url in fonts.items():
    dest = os.path.join(fonts_dir, name)
    if os.path.exists(dest):
        print(f'  {name} deja present')
        continue
    print(f'  Telechargement {name}...')
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = resp.read()
        with open(dest, 'wb') as f:
            f.write(data)
        print(f'  {name} OK')
    except Exception as e:
        print(f'  [ERREUR] {name}: {e}')
        sys.exit(1)
"
if errorlevel 1 (
    echo [ERREUR] Echec du telechargement des polices.
    echo Verifiez votre connexion internet.
    pause
    exit /b 1
)
echo     OK

echo [3/3] Installation terminee !
echo.
echo ============================================
echo   Lancez l'application avec : run.bat
echo ============================================
echo.
pause
