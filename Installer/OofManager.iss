; Inno Setup script for OOF Manager
; Build with: .\Tools\build-installer.ps1
; Output: Installer\Output\OofManagerSetup.exe

#define MyAppName "OOF Manager"
#define MyAppVersion "1.1.11"
#define MyAppPublisher "OOF Manager"
#define MyAppExeName "OofManager.exe"

[Setup]
AppId={{A4B6F2E2-0E5C-4D6F-9A1A-OOFMGR1234}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\OofManager
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=Output
OutputBaseFilename=OofManagerSetup
Compression=lzma2/ultra
SolidCompression=yes
WizardStyle=modern
PrivilegesRequiredOverridesAllowed=dialog
PrivilegesRequired=lowest
UninstallDisplayIcon={app}\{#MyAppExeName}
MinVersion=10.0
SetupIconFile=..\Resources\oofmanager.ico

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Bundle the entire publish output (app + ExchangeOnlineManagement module).
Source: "..\publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Automatic Power Automate Solution import uses pac. The build script downloads
; this official Microsoft prerequisite before compiling the installer.
Source: "Prerequisites\powerapps-cli-1.0.msi"; DestDir: "{tmp}"; Flags: ignoreversion deleteafterinstall

[InstallDelete]
; v1.0.9 renamed the exe from OofManager.Wpf.exe to OofManager.exe. Remove the
; legacy files on upgrade so an old autostart Run entry pointing at the old
; name doesn't survive (StartupService.EnsureRegistrationIsFresh would
; otherwise rewrite the entry to the still-present old exe path on first
; launch). Safe even on fresh installs because IgnoreErrors is on by default.
Type: files; Name: "{app}\OofManager.Wpf.exe"
Type: files; Name: "{app}\OofManager.Wpf.exe.config"
Type: files; Name: "{app}\OofManager.Wpf.pdb"

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "msiexec.exe"; Parameters: "/i ""{tmp}\powerapps-cli-1.0.msi"" /qn /norestart ALLUSERS=2 MSIINSTALLPERUSER=1"; StatusMsg: "Installing Microsoft PowerApps CLI..."; Check: not PowerAppsCliInstalled; Flags: waituntilterminated
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
function PowerAppsCliInstalled: Boolean;
begin
	Result :=
		FileExists(ExpandConstant('{localappdata}\Microsoft\PowerAppsCLI\pac.cmd')) or
		FileExists(ExpandConstant('{localappdata}\Microsoft\PowerAppsCLI\pac.launcher.exe')) or
		FileExists(ExpandConstant('{localappdata}\Microsoft\PowerAppsCLI\pac.exe')) or
		FileExists(ExpandConstant('{%USERPROFILE}\.dotnet\tools\pac.exe'));
end;
