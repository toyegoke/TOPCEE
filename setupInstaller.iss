[Setup]
AppName=TOPCEE
AppVersion=2
DefaultDirName={pf}\TOPCEE
DefaultGroupName=TOPCEE
OutputBaseFilename=TOPCEEinstaller
SetupIconFile=app_icon.ico
Compression=lzma
SolidCompression=yes

[Files]
Source: "dist\TOPCEEv2.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{commondesktop}\TOPCEE"; Filename: "{app}\TOPCEEv2.exe"

[Run]
Filename: "{app}\TOPCEEv2.exe"; Description: "Launch TOPCEE"; Flags: nowait postinstall
