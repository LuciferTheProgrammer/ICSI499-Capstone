[Setup]
PrivilegesRequired=lowest
AppName=FindingsAutomation
AppVersion=1.0
DefaultDirName={localappdata}\FindingsAutomation
DefaultGroupName=FindingsAutomation
OutputDir=C:\Users\David\Documents\ICSI499\installer
OutputBaseFilename=FindingsAutomationSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=C:\Users\David\Downloads\App\app.ico

[Files]
Source: "C:\Users\David\Documents\ICSI499\dist\FindingsAutomation\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Icons]
Name: "{group}\FindingsAutomation"; Filename: "{app}\FindingsAutomation.exe"
Name: "{autodesktop}\FindingsAutomation"; Filename: "{app}\FindingsAutomation.exe"; Tasks: desktopicon