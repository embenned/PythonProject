$ErrorActionPreference = 'Stop'

$python = 'C:/Users/bkrol/AppData/Local/Programs/Python/Python314/python.exe'

& $python -m pip install -r requirements.txt

& $python -m PyInstaller `
  --noconfirm `
  --clean `
  --onefile `
  --windowed `
  --name TrainTypeDataPrepTool `
  --add-data 'data;data' `
  --add-data 'templates;templates' `
  main.py

Write-Host 'Build complete. Check the dist\ folder for TrainTypeDataPrepTool.exe.'