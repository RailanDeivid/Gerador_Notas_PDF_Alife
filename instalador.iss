[Setup]
AppName=Gerador de Notas de Débitos
AppVersion=1.0
DefaultDirName={pf}\Gerador de Notas de Débitos
DefaultGroupName=Gerador de Notas de Débitos
OutputDir=.
OutputBaseFilename=Instalador_Gerador_Notas
Compression=lzma
SolidCompression=yes


[Files]
Source: "C:\Users\Raila\Documents\Gerador de Notas de Debitos\Gerador de Notas de Debitos.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Raila\Documents\Gerador de Notas de Debitos\icone.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Gerador de Notas de Débitos"; Filename: "{app}\Gerador de Notas de Debitos.exe"
Name: "{commondesktop}\Gerador de Notas de Débitos"; Filename: "{app}\Gerador de Notas de Debitos.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na área de trabalho"; GroupDescription: "Opções adicionais:"

[Run]
Filename: "{app}\Gerador de Notas de Debitos.exe"; Description: "Executar Gerador de Notas de Débitos"; Flags: nowait postinstall
