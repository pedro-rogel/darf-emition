@echo off
cd /d "%~dp0"
if not exist "venv\Scripts\pythonw.exe" (
    where pythonw.exe >nul 2>nul
    if errorlevel 1 (
        powershell.exe -NoProfile -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Instale o Python 3 para Windows e abra novamente este arquivo.','Emissao de DARF')"
        exit /b 1
    )
    start "" pythonw.exe "%~dp0iniciar.pyw"
    exit /b 0
)
start "" "%~dp0venv\Scripts\pythonw.exe" "%~dp0gui.pyw"
