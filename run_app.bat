@echo off
cd /d "%~dp0"
"C:\Users\Serge\AppData\Local\Programs\Python\Python312\python.exe" -m streamlit run app.py
if errorlevel 1 (
    echo.
    echo Une erreur est survenue au lancement de l'application.
    pause >nul
)
