@echo off
REM ============================================================================
REM Script para ejecutar el Resumen Diario de Outlook
REM Modo: BORRADO ACTIVADO (elimina correos antiguos de conversaciones)
REM ============================================================================

REM Cambiar al directorio del script
cd /d "%~dp0"

REM Activar el entorno virtual de Python (si existe)
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
)

REM Ejecutar el script con la opción de borrado
python outlook_daily_digest.py --delete

REM Registrar la ejecución en un log
echo [%date% %time%] Resumen diario ejecutado >> outlook_digest_log.txt

REM Opcional: Descomentar la siguiente línea si quieres ver el resultado
REM pause
