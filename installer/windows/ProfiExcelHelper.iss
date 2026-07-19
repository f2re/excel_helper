#define MyAppName "ПрофиПомощник для Excel"
#define MyAppVersion "1.2.0"
#define MyAppPublisher "f2re"
#ifndef PayloadDir
  #define PayloadDir "..\..\release\payload"
#endif

[Setup]
AppId={{67A1E987-2DF8-4D08-A708-96A39CB7A7A1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://github.com/f2re/excel_helper
AppSupportURL=https://github.com/f2re/excel_helper/issues
DefaultDirName={localappdata}\ProfiExcelHelper
DefaultGroupName=ПрофиПомощник
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=dist
OutputBaseFilename=ProfiExcelHelper-Setup-{#MyAppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
SetupLogging=yes
UninstallDisplayName={#MyAppName}
ChangesAssociations=no
CloseApplications=no
RestartApplications=no

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык шаблона на рабочем столе"; GroupDescription: "Ярлыки:"; Flags: unchecked

[Files]
Source: "{#PayloadDir}\*"; DestDir: "{app}\payload"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "install.ps1"; DestDir: "{app}\installer"; Flags: ignoreversion
Source: "uninstall.ps1"; DestDir: "{app}\installer"; Flags: ignoreversion
Source: "repair.ps1"; DestDir: "{app}\installer"; Flags: ignoreversion
Source: "ProfiInstaller.Common.psm1"; DestDir: "{app}\installer"; Flags: ignoreversion

[Run]
Filename: "{sysnative}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\installer\install.ps1"" -PayloadRoot ""{app}\payload"" -Mode Full -Silent"; StatusMsg: "Установка XLAM и XLTM..."; Flags: runhidden waituntilterminated; Check: IsWin64
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\installer\install.ps1"" -PayloadRoot ""{app}\payload"" -Mode Full -Silent"; StatusMsg: "Установка XLAM и XLTM..."; Flags: runhidden waituntilterminated; Check: not IsWin64

[UninstallRun]
Filename: "{sysnative}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\installer\uninstall.ps1"" -Silent -Force"; Flags: runhidden waituntilterminated; Check: IsWin64
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\installer\uninstall.ps1"" -Silent -Force"; Flags: runhidden waituntilterminated; Check: not IsWin64

[Icons]
Name: "{group}\ПрофиПомощник — новый проект"; Filename: "{userappdata}\Microsoft\Templates\ProfiExcelHelper\ProfiExcelHelper-Template.xltm"
Name: "{group}\Папка установки"; Filename: "{app}"
Name: "{autodesktop}\ПрофиПомощник — новый проект"; Filename: "{userappdata}\Microsoft\Templates\ProfiExcelHelper\ProfiExcelHelper-Template.xltm"; Tasks: desktopicon
