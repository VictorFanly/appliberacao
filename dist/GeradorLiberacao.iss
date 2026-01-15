[Setup]
AppName=Gerador de Termo de Liberação
AppVersion=1.0.0
AppPublisher=Secretaria Municipal
AppPublisherURL=
AppSupportURL=
AppUpdatesURL=

DefaultDirName={pf}\GeradorLiberacao
DefaultGroupName=Gerador de Termo de Liberação

OutputDir=installer
OutputBaseFilename=Instalador_Gerador_Liberacao
Compression=lzma
SolidCompression=yes

UninstallDisplayIcon={app}\GeradorLiberacao.exe

DisableProgramGroupPage=yes
WizardStyle=modern

ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin

[Languages]
Name: "portuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Files]
Source: "GeradorLiberacao.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{commondesktop}\Gerador de Termo de Liberação"; Filename: "{app}\GeradorLiberacao.exe"
Name: "{group}\Gerador de Termo de Liberação"; Filename: "{app}\GeradorLiberacao.exe"

[Run]
Filename: "{app}\GeradorLiberacao.exe"; Description: "Executar o sistema agora"; Flags: nowait postinstall skipifsilent
