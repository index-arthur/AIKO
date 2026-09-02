<#
    Gera o instalador do Cadastro de Bordo.

        .\build.ps1                # usa a versao do arquivo VERSION
        .\build.ps1 -Versao 6.2    # forca uma versao

    Duas escolhas do build que nao sao obvias e nao devem ser mexidas sem
    motivo:

    --collect-data sv_ttk   o tema carrega arquivos .tcl em runtime e o
                            PyInstaller nao os acha sozinho. Sem isso o app
                            compilado abre e morre ao aplicar o tema.

    onedir (nao onefile)    o onefile se descompacta em %TEMP% e cria um
                            processo filho a cada abertura - passo bloqueado
                            pelo antivirus corporativo, com a mensagem
                            "CreateProcessW: Acesso negado". O onedir tambem
                            abre bem mais rapido.
#>
param(
    [string]$Versao,
    [switch]$PularInstalador
)

$ErrorActionPreference = "Stop"
$raiz = $PSScriptRoot
Set-Location $raiz

if (-not $Versao) {
    $Versao = (Get-Content (Join-Path $raiz "VERSION") -Raw).Trim()
}
Write-Host "Versao: $Versao" -ForegroundColor Cyan

# A versao no codigo e a da release tem de bater: o updater compara a tag do
# GitHub com a constante VERSION embutida. Divergiu, ou ninguem ve a
# atualizacao ou o app pede update para sempre.
$noCodigo = (Select-String -Path (Join-Path $raiz "main.py") `
             -Pattern '^VERSION = "([^"]+)"').Matches[0].Groups[1].Value
if ($noCodigo -ne $Versao) {
    throw "VERSION do arquivo ($Versao) != VERSION do main.py ($noCodigo). Alinhe os dois."
}

$dist = Join-Path $raiz "_build\dist"
$work = Join-Path $raiz "_build\work"

Write-Host "`n[1/2] Empacotando com PyInstaller..." -ForegroundColor Cyan
python -m PyInstaller --noconfirm --windowed `
    --name "CadastroBordo" `
    --collect-data sv_ttk `
    --hidden-import motor_api `
    --hidden-import motor_vinculo `
    --hidden-import motor_starlink `
    --hidden-import trackit_api_client `
    --distpath $dist --workpath $work --specpath (Join-Path $raiz "_build") `
    main.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller falhou" }

$appDir = Join-Path $dist "CadastroBordo"
if (-not (Test-Path (Join-Path $appDir "CadastroBordo.exe"))) {
    throw "nao encontrei CadastroBordo.exe em $appDir"
}
Write-Host "  pasta do app: $appDir" -ForegroundColor DarkGray

if ($PularInstalador) { Write-Host "`nParando antes do instalador."; exit 0 }

$iscc = @(
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) {
    throw "Inno Setup nao encontrado. Instale com: winget install JRSoftware.InnoSetup --scope user"
}

Write-Host "`n[2/2] Gerando o instalador..." -ForegroundColor Cyan
& $iscc "/DVERSAO=$Versao" "/DORIGEM=$appDir" (Join-Path $raiz "installer.iss")
if ($LASTEXITCODE -ne 0) { throw "Inno Setup falhou" }

$setup = Join-Path $raiz "_setup\CadastroBordo_Setup_$Versao.exe"
if (Test-Path $setup) {
    $mb = (Get-Item $setup).Length / 1MB
    Write-Host ("`nPronto: {0}  ({1:N1} MB)" -f $setup, $mb) -ForegroundColor Green
    Write-Host "Anexe este arquivo na release v$Versao do GitHub." -ForegroundColor DarkGray
} else {
    throw "instalador nao apareceu em $setup"
}
