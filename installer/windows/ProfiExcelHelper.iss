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

[Types]
Name: "full"; Description: "Полная установка: XLAM + XLTM + Office.js"
Name: "legacy"; Description: "Excel 2010–2019: XLAM + XLTM"
Name: "addin"; Description: "Только глобальная надстройка XLAM"
Name: "template"; Description: "Только переносимый шаблон XLTM"
Name: "modern"; Description: "Только файлы современной Office.js-надстройки"
Name: "custom"; Description: "Выборочная установка"; Flags: iscustom

[Components]
Name: "addin"; Description: "Глобальная XLAM-надстройка для Excel 2010–2019"; Types: full legacy addin
Name: "template"; Description: "Переносимый XLTM-шаблон проекта"; Types: full legacy template
Name: "modern"; Description: "Манифесты Office.js для Microsoft 365 и Excel 2016"; Types: full modern

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык XLTM-шаблона на рабочем столе"; GroupDescription: "Ярлыки:"; Flags: unchecked; Components: template

[Files]
Source: "{#PayloadDir}\legacy\ProfiExcelHelper-Legacy.xlam"; DestDir: "{app}\payload\legacy"; Flags: ignoreversion; Components: addin
Source: "{#PayloadDir}\template\ProfiExcelHelper-Template.xltm"; DestDir: "{app}\payload\template"; Flags: ignoreversion; Components: template
Source: "{#PayloadDir}\officejs\*"; DestDir: "{app}\payload\officejs"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: modern
Source: "{#PayloadDir}\docs\*"; DestDir: "{app}\payload\docs"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "install.ps1"; DestDir: "{app}\installer"; Flags: ignoreversion
Source: "uninstall.ps1"; DestDir: "{app}\installer"; Flags: ignoreversion
Source: "repair.ps1"; DestDir: "{app}\installer"; Flags: ignoreversion
Source: "ProfiInstaller.Common.psm1"; DestDir: "{app}\installer"; Flags: ignoreversion

[Run]
Filename: "{sysnative}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\installer\install.ps1"" -PayloadRoot ""{app}\payload"" -Mode {code:GetInstallMode} -Silent"; StatusMsg: "Установка выбранных компонентов..."; Flags: runhidden waituntilterminated; Check: IsWin64
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\installer\install.ps1"" -PayloadRoot ""{app}\payload"" -Mode {code:GetInstallMode} -Silent"; StatusMsg: "Установка выбранных компонентов..."; Flags: runhidden waituntilterminated; Check: not IsWin64

[UninstallRun]
Filename: "{sysnative}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\installer\uninstall.ps1"" -Silent"; Flags: runhidden waituntilterminated; Check: IsWin64
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\installer\uninstall.ps1"" -Silent"; Flags: runhidden waituntilterminated; Check: not IsWin64

[Icons]
Name: "{group}\ПрофиПомощник — новый проект"; Filename: "{userappdata}\Microsoft\Templates\ProfiExcelHelper\ProfiExcelHelper-Template.xltm"; Components: template
Name: "{group}\Папка установки"; Filename: "{app}"
Name: "{autodesktop}\ПрофиПомощник — новый проект"; Filename: "{userappdata}\Microsoft\Templates\ProfiExcelHelper\ProfiExcelHelper-Template.xltm"; Tasks: desktopicon; Components: template

[Code]
function ExcelIsRunning: Boolean;
begin
  Result := FindWindowByClassName('XLMAIN') <> 0;
end;

function GetInstallMode(Param: String): String;
var
  AddinSelected, TemplateSelected, ModernSelected: Boolean;
begin
  AddinSelected := WizardIsComponentSelected('addin');
  TemplateSelected := WizardIsComponentSelected('template');
  ModernSelected := WizardIsComponentSelected('modern');

  if AddinSelected and TemplateSelected and ModernSelected then
    Result := 'Full'
  else if AddinSelected and TemplateSelected then
    Result := 'LegacyFull'
  else if AddinSelected and ModernSelected then
    Result := 'AddinModern'
  else if TemplateSelected and ModernSelected then
    Result := 'TemplateModern'
  else if AddinSelected then
    Result := 'AddinOnly'
  else if TemplateSelected then
    Result := 'TemplateOnly'
  else if ModernSelected then
    Result := 'ModernOnly'
  else
    Result := '';
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if (CurPageID = wpSelectComponents) and (GetInstallMode('') = '') then
  begin
    MsgBox('Выберите хотя бы один компонент.', mbError, MB_OK);
    Result := False;
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  if WizardIsComponentSelected('addin') and ExcelIsRunning then
    Result := 'Закройте все окна Microsoft Excel и повторите установку.'
  else
    Result := '';
end;

function InitializeUninstall(): Boolean;
begin
  Result := not ExcelIsRunning;
  if not Result then
    MsgBox('Перед удалением ПрофиПомощника закройте все окна Microsoft Excel.', mbError, MB_OK);
end;
