@echo off
cls
echo ===============================
echo  SUBIENDO CAMBIOS A GITHUB...
echo ===============================

:: Cambia "Tu mensaje aquí" por algo más útil si lo deseas
set /p MSG="Escribe un mensaje para el commit: "

git add .
git commit -m "%MSG%"
git push origin main

echo.
echo ✅ Cambios subidos correctamente.
pause
