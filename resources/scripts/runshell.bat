set "current_dir=%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%current_dir%\GetAllPnPDevices.ps1" -Wait