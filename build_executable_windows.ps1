$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

if (-not (Test-Path '.venv\Scripts\python.exe')) {
    Write-Host 'Ambiente virtual .venv não encontrado. Crie-o primeiro com:'
    Write-Host '  python -m venv .venv'
    exit 1
}

& '.\.venv\Scripts\python.exe' -m pip install --upgrade pip
& '.\.venv\Scripts\python.exe' -m pip install -r requirements.txt pyinstaller

Remove-Item -Recurse -Force dist, build, automacao_afiliados.spec -ErrorAction SilentlyContinue

& '.\.venv\Scripts\python.exe' -m PyInstaller `
    --onefile `
    --name automacao_afiliados `
    --add-data "imagens;imagens" `
    --icon=imagens\icone.png `
    --hidden-import=PIL._tkinter_finder `
    main.py

Write-Host ''
Write-Host 'Executável gerado em: dist\automacao_afiliados.exe'
