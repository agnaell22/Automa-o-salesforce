$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$Desktop = [Environment]::GetFolderPath("Desktop")
$BuildDirectory = Join-Path $ProjectRoot "build"
$ExecutableName = "Automacao Salesforce.exe"
$BuiltExecutable = Join-Path $BuildDirectory $ExecutableName
$DesktopExecutable = Join-Path $Desktop $ExecutableName

if (-not (Test-Path -LiteralPath $Python -PathType Leaf)) {
    throw "Ambiente Python não encontrado em '$Python'. Crie .venv e instale requirements-build.txt."
}
if (-not (Test-Path -LiteralPath $Desktop -PathType Container)) {
    throw "A pasta Área de Trabalho não foi encontrada: '$Desktop'."
}

Push-Location $ProjectRoot
try {
    & $Python -m PyInstaller `
        --clean `
        --noconfirm `
        --onefile `
        --windowed `
        --name "Automacao Salesforce" `
        --collect-all selenium `
        --distpath $BuildDirectory `
        --workpath (Join-Path $BuildDirectory "work") `
        --specpath $BuildDirectory `
        interface.py
    if ($LASTEXITCODE -ne 0) {
        throw "O empacotamento com PyInstaller falhou (código $LASTEXITCODE)."
    }

    try {
        Copy-Item -LiteralPath $BuiltExecutable -Destination $DesktopExecutable -Force
        Write-Output "Executável atualizado: $DesktopExecutable"
    }
    catch {
        $UpdatedExecutable = Join-Path $Desktop "Automacao Salesforce atualizado.exe"
        Copy-Item -LiteralPath $BuiltExecutable -Destination $UpdatedExecutable -Force
        Write-Warning (
            "O executável da Área de Trabalho está aberto e não pôde ser substituído. " +
            "A versão nova foi criada em '$UpdatedExecutable'. Feche a janela antiga " +
            "e execute este script novamente para atualizar o nome principal."
        )
    }
    Write-Output "Os scripts-fonte permanecem no projeto: $ProjectRoot"
}
finally {
    Pop-Location
}
