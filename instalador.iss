[Setup]
AppName=Gerador de Termo de Liberação
AppVersion=1.0.0
DefaultDirName={pf}\GeradorLiberacao
DefaultGroupName=Gerador de Termo de Liberação
OutputDir=output
OutputBaseFilename=Instalador_Gerador_Liberacao
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
DisableProgramGroupPage=yes
PrivilegesRequired=admin

[Languages]
Name: "portuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Files]
Source: "dist\GeradorLiberacao.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "datasec"; DestDir: "{app}"; Flags: ignoreversion
Source: "LIBERACAO.docx"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Gerador de Termo de Liberação"; Filename: "{app}\GeradorLiberacao.exe"
Name: "{commondesktop}\Gerador de Termo de Liberação"; Filename: "{app}\GeradorLiberacao.exe"

[Run]
Filename: "{app}\GeradorLiberacao.exe"; Description: "Executar agora"; Flags: nowait postinstall skipifsilent
