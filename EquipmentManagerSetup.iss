
[Setup]
AppName=ניהול ציוד מחשבים
AppVersion=1.0
DefaultDirName={pf}\EquipmentManager
DefaultGroupName=ניהול ציוד
UninstallDisplayIcon={app}pp.ico
OutputDir=.
OutputBaseFilename=EquipmentManagerSetup
Compression=lzma
SolidCompression=yes

[Files]
Source: "C:\ITPROD\equipment_management.db"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\ITPROD\app.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\ITPROD\requirements.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\ITPROD\run_app.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\ITPROD\install.bat"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\הפעל את האפליקציה"; Filename: "{app}\run_app.bat"

[Run]
Filename: "{app}\install.bat"; Description: "התקנת תלותים והפעלת האפליקציה"; Flags: runhidden
