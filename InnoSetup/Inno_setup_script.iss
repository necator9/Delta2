; Delta2

[Setup]
AppName=Delta2 
AppVersion=1.01 beta
WizardStyle=modern
DefaultDirName=C:\opt\Delta2
DefaultGroupName=Delta2
UninstallDisplayIcon={app}\Delta2_uninst.exe
Compression=lzma2
SolidCompression=yes
OutputDir={#SourcePath}\dist
OutputBaseFilename=Delta2_setup
DisableWelcomePage=no

[Files]
Source: "{#SourcePath}\dist\run_delta2\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs
;Source: "{#SourcePath}\Readme.txt"; DestDir: "{app}"; Flags: isreadme
Source: "{#SourcePath}\delta.ico"; DestDir: "{app}";

[Icons]
Name: "{autodesktop}\Delta2"; Filename: "{app}\run_delta2.exe"; IconFilename: "{app}\delta.ico";

[Messages]
WelcomeLabel2=This will install [name/ver] on your computer.%n%nConfiguration options will be requested during installation.

[Setup]
PrivilegesRequired=lowest
