@echo off
title Crear EtiquetadorCuadros.exe
cd /d "%~dp0"

rem Busca Python. Prefiere 3.12 / 3.11 (el OCR no siempre funciona con las mas nuevas)
set PY=
for %%V in (3.12 3.11 3.13 3.10) do (
  if not defined PY (
    py -%%V -c "import sys" >nul 2>nul && set "PY=py -%%V"
  )
)
rem Si no hay 3.12/3.11 pero esta el instalador nuevo de Python ("py install"),
rem instala Python 3.12 automaticamente (lo necesita el OCR)
if not defined PY (
  where py >nul 2>nul && (
    echo.
    echo No hay Python 3.12. Instalandolo para poder incluir el OCR...
    py install 3.12
    py -3.12 -c "import sys" >nul 2>nul && set "PY=py -3.12"
  )
)
if not defined PY ( where py >nul 2>nul && set "PY=py" )
if not defined PY ( where python >nul 2>nul && set "PY=python" )
if not defined PY (
  echo.
  echo No se encuentra Python en este ordenador.
  echo Instala Python 3.12 de 64 bits desde https://www.python.org/downloads/
  echo ^(marca "Add python.exe to PATH"^) y vuelve a ejecutar este archivo.
  pause
  exit /b 1
)

echo.
echo Usando:
%PY% -c "import sys, platform; print('  Python', sys.version.split()[0], '-', platform.machine())"

echo.
echo [1/4] Instalando librerias (PDF, Excel, DXF/DWG)...
%PY% -m pip install --upgrade pymupdf openpyxl ezdxf pillow numpy pyinstaller
if errorlevel 1 goto error

echo.
echo [2/4] Instalando el reconocimiento de texto (OCR) para planos escaneados...
set OCR=1
%PY% -m pip install --upgrade rapidocr onnxruntime
if errorlevel 1 set OCR=0

if "%OCR%"=="1" (
  echo.
  echo [3/4] Descargando los modelos del OCR ^(solo esta vez, necesita internet^)...
  %PY% -c "from rapidocr import RapidOCR; RapidOCR()"
  if errorlevel 1 set OCR=0
)

echo.
echo [4/4] Creando el .exe (tarda unos minutos)...
set EXTRA=--collect-data ezdxf
if "%OCR%"=="1" set EXTRA=%EXTRA% --collect-all rapidocr --collect-all onnxruntime
%PY% -m PyInstaller --noconfirm --onefile --windowed --name EtiquetadorCuadros %EXTRA% etiquetador.py
if errorlevel 1 goto error

echo.
echo ============================================================
echo  Listo. El programa esta en:  dist\EtiquetadorCuadros.exe
echo  Puedes copiar ese .exe a cualquier PC con Windows.
echo  No necesita Python ni internet para funcionar.
if "%OCR%"=="1" (
  echo  OCR ^(planos escaneados^): INCLUIDO
) else (
  echo.
  echo  OCR ^(planos escaneados^): NO INCLUIDO
  echo  El programa funciona con PDF con texto, DXF y Excel, pero no
  echo  lee planos escaneados ni imagenes. Para incluirlo, instala
  echo  Python 3.12 de 64 bits ^(x64^) desde python.org y vuelve a
  echo  ejecutar este archivo.
)
echo ============================================================
explorer dist
pause
exit /b 0

:error
echo.
echo Algo ha fallado. Copia el texto de arriba y pasamelo.
pause
exit /b 1
