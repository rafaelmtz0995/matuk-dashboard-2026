@echo off
cd /d "%~dp0"
echo ============================================================
echo   MATUK Dashboard 2026 - Actualizacion completa
echo ============================================================
echo.

echo [1/4] Regenerando dashboard completo (con Flujo de Caja)...
python regenerar_dashboard_v2_local.py
if errorlevel 1 (
    echo ERROR al regenerar dashboard.html. Revisa los archivos Excel y vuelve a intentar.
    pause
    exit /b 1
)
echo.

echo [2/4] Subiendo dashboard.html a GitHub via git...
rem Limpiar archivos lock huerfanos que pueden quedar de operaciones anteriores
if exist ".git\index.lock" (
    echo   Limpiando index.lock huerfano...
    del /f ".git\index.lock"
)
if exist ".git\HEAD.lock" (
    echo   Limpiando HEAD.lock huerfano...
    del /f ".git\HEAD.lock"
)
git add dashboard.html
git commit -m "Actualizar dashboard - %date%"
git push
if errorlevel 1 (
    echo ERROR al hacer git push.
    pause
    exit /b 1
)
echo.

echo [3/4] Actualizando dashboard de Costos/Pasivo (index.html)...
python actualizar_dashboard.py
if errorlevel 1 (
    echo ERROR al actualizar index.html.
    pause
    exit /b 1
)
echo.

echo ============================================================
echo   LISTO - Ambos dashboards actualizados
echo   https://rafaelmtz0995.github.io/matuk-dashboard-2026/dashboard.html
echo   https://rafaelmtz0995.github.io/dashboard-matuk/
echo ============================================================
echo.
pause
