; Instalador do Cadastro de Bordo (Aiko / TracKit).
;
; Instala POR USUARIO, em %LOCALAPPDATA%\Programs. Sem UAC, sem depender da
; TI - o time nao tem admin na maquina, e exigir elevacao aqui inviabilizaria
; a distribuicao.
;
; O app empacotado e PyInstaller em modo "onedir", nao "onefile": o onefile
; se descompacta em %TEMP% e cria um processo filho a cada abertura, passo
; que o antivirus corporativo bloqueia (CreateProcessW: Acesso negado).
;
; Compilar:
;   ISCC.exe /DVERSAO=6.2 /DORIGEM="caminho\dist\CadastroBordo" installer.iss

#ifndef VERSAO
  #define VERSAO "0.0"
#endif
#ifndef ORIGEM
  #define ORIGEM "dist\CadastroBordo"
#endif

#define APP "Cadastro de Bordo"
#define EXE "CadastroBordo.exe"

[Setup]
AppId={{8F3C1A94-7D62-4E58-9B41-2C7A5E0D6F13}
AppName={#APP}
AppVersion={#VERSAO}
AppVerName={#APP} {#VERSAO}
AppPublisher=Aiko Digital
AppPublisherURL=https://aiko.digital
VersionInfoVersion={#VERSAO}

; Instalacao por usuario: nada de UAC.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
DefaultDirName={autopf}\{#APP}
DefaultGroupName={#APP}
DisableProgramGroupPage=yes
DisableDirPage=auto

; Relativo a ESTE arquivo, nao ao diretorio de onde se chama o compilador.
OutputDir=_setup
OutputBaseFilename=CadastroBordo_Setup_{#VERSAO}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible

; Permite que o updater feche o app sozinho ao atualizar.
CloseApplications=yes
RestartApplications=yes
SetupLogging=yes
UninstallDisplayName={#APP}
UninstallDisplayIcon={app}\{#EXE}

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Files]
Source: "{#ORIGEM}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#APP}"; Filename: "{app}\{#EXE}"
Name: "{userdesktop}\{#APP}"; Filename: "{app}\{#EXE}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na area de trabalho"; \
  GroupDescription: "Atalhos:"

[Run]
; Sem "skipifsilent": e assim que o app volta sozinho depois da atualizacao
; automatica, que roda o instalador em modo /SILENT.
Filename: "{app}\{#EXE}"; Description: "Abrir o {#APP}"; \
  Flags: nowait postinstall
