; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!


; Removed IConform and use of GACInstall to install it to GAC
; Reverted to 64bit install directory on 64bit systems
; Added GACInstall calls to uninstall old IConform assemblies
; Conform Version 3.0.0.10 Released

; Added uninstall old version to installer
; Added check for platform 5.5 to installer
; Enabled setup logging
; Updated SideOfPier document to version 1.1
; Conform Version 3.0.0.11 Released
; Updated to use Inno setup 5.3.7 Unicode 4th February 2010
; Version 3.1.1.17 Released
; 32bit version of version 3.0.0.17
; Updated for Platform 6
; Stores Conform version in the Platform Registry Profile

[Setup]
#define MyAppVer GetVersionNumbersString("..\Conform\bin\Release\Conform.exe") ; define variable
AppID={{CEDF90C4-A415-4903-94D6-9099F7AE41B5}
AppCopyright=Copyright � 2021 Peter Simpson
AppName=ASCOM Conform
AppPublisher=Peter Simpson
AppPublisherURL=https://www.ascom-standards.org
AppSupportURL=https://ascomtalk.groups.io/g/Developer/topics
AppUpdatesURL=https://www.ascom-standards.org
#emit "AppVerName=ASCOM Conform " + MyAppVer
#emit "AppVersion=" + MyAppVer
;ArchitecturesInstallIn64BitMode=x64
Compression=lzma/Max
DefaultDirName={commonpf}\ASCOM\Conform
DefaultGroupName=ASCOM Platform 6\Tools
MinVersion=6.1SP1
#emit "OutputBaseFilename=Conform(" + MyAppVer +")setup"
OutputDir=.\Builds
PrivilegesRequired=admin
SetupIconFile=..\Conform\Resources\ASCOM.ico
SetupLogging=true
SignToolRunMinimized=yes
ShowLanguageDialog=auto
SolidCompression=true
UninstallDisplayIcon={app}\Conform.exe
UsedUserAreasWarning=no
VersionInfoCompany=ASCOM Initiative
VersionInfoCopyright=Peter Simpson
VersionInfoDescription=Tool to validate ASCOM drivers
VersionInfoProductName=ASCOM Conform
#emit "VersionInfoProductVersion=" + MyAppVer
#emit "VersionInfoVersion=" + MyAppVer
WizardImageFile=NewWizardImage.bmp
WizardSmallImageFile=ASCOMLogo.bmp
SignTool = SignConform

[Languages]
Name: english; MessagesFile: compiler:Default.isl

[Tasks]
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked
Name: quicklaunchicon; Description: {cm:CreateQuickLaunchIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked

[Files]
Source: ..\Conform\bin\Release\Conform.exe; DestDir: {app}; Flags: ignoreversion
Source: ..\Conform\bin\Release\Conform.pdb; DestDir: {app}; Flags: ignoreversion
Source: ..\Conform\bin\Release\Conform.pdb; DestDir: {win}\Symbols\dll; Flags: ignoreversion

;32bit version
Source: ..\Conform\bin\x86\Release\Conform.exe; DestDir: {app}; Flags: ignoreversion; DestName: "Conform32.exe"
Source: ..\Conform\bin\x86\Release\Conform.pdb; DestDir: {app}; Flags: ignoreversion; DestName: "Conform32.pdb"
Source: ..\Conform\bin\x86\Release\Conform.pdb; DestDir: {win}\Symbols\dll; Flags: ignoreversion; DestName: "Conform32.pdb"

;Help files
Source: ..\SideOfPier Document\ASCOM SideOfPier Behaviour 1.2.pdf; DestDir: {app}; Flags: ignoreversion
Source: ..\Help\Conform.chm; DestDir: {app}; Flags: ignoreversion

[Registry]
Root: HKLM; Subkey: Software\Microsoft\.NETFramework\v2.0.50727\AssemblyFoldersEx\Conform; ValueType: string; ValueName: ; ValueData: {app}\; Flags: uninsdeletekey
#emit "Root: HKLM; SubKey: SOFTWARE\ASCOM\Platform; ValueType: string; ValueName: ""Conform Version""; ValueData: " + MyAppVer + "; Flags: UninsDeleteValue;" 

[Icons]
Name: {autostartmenu}\ASCOM Conform; Filename: {app}\Conform.exe
Name: {commondesktop}\Conform; Filename: {app}\Conform.exe; Tasks: desktopicon
Name: {userappdata}\Microsoft\Internet Explorer\Quick Launch\Conform; Filename: {app}\Conform.exe; Tasks: quicklaunchicon

[Run]
Filename: {app}\Conform.exe; Description: {cm:LaunchProgram,Conform}; Flags: nowait postinstall skipifsilent

[UninstallRun]

[InstallDelete]
Name: {app}\Conform.exe.config; Type: files
[UninstallDelete]
Name: {app}; Type: dirifempty

[PreCompile]
Name: "..\BuildConform.cmd"; Flags: CmdPrompt

[Code]
const
   MINIMUM_PLATFORM_VERSION = 7.0; // Set this to the minimum required ASCOM Platform version for this application
   MAXIMUM_PLATFORM_VERSION = 7.9; // This must be a number that is larger than the highest release version of Platform 6

//
// Function to return the ASCOM Platform's version number as a double.
//
function PlatformVersion(): Extended;
var
   PlatVerString : String;
begin
   Result := 0.0;  // Initialise the return value in case we can't read the registry
   try
      if RegQueryStringValue(HKEY_LOCAL_MACHINE_32, 'Software\ASCOM','PlatformVersion', PlatVerString) then 
      begin // Successfully read the value from the registry
         Result := StrToFloat(PlatVerString); // Create a double from the X.Y Platform version string
      end;
   except                                                                   
      ShowExceptionMessage;
      Result:= -1.0; // Indicate in the return value that an exception was generated
   end;
end;

//
// Before the installer UI appears, verify that the required ASCOM Platform version is installed.
//
function InitializeSetup(): Boolean;
var
   PlatformVersionNumber : double;
begin
   Result := FALSE;  // Assume failure
   PlatformVersionNumber := PlatformVersion(); // Get the installed Platform version as a double
   // Check whether we have the minimum required Platform or newer and the maximum or older Platform
   If (PlatformVersionNumber >= MINIMUM_PLATFORM_VERSION) and (PlatformVersionNumber <= MAXIMUM_PLATFORM_VERSION) then // We are on a supported OS
      Result := TRUE
   else // We are not on a suypported OS
      // Check whether a Platform is installed
      If PlatformVersionNumber = 0.0 then // No Platform is installed
         MsgBox('No ASCOM Platform is installed. Please install Platform ' + Format('%3.1f', [MINIMUM_PLATFORM_VERSION]) + ' or later from http://www.ascom-standards.org', mbCriticalError, MB_OK)
      else begin // An incompatible Platform is installed
        // Check whether an earlier Platform than the minimum is installed
        If PlatformVersionNumber < MINIMUM_PLATFORM_VERSION then // An earlier Platform is installed
          MsgBox('This version of Conform requires ASCOM Platform ' + Format('%3.1f', [MINIMUM_PLATFORM_VERSION]) + ' or later' + 
            ', but Platform '+ Format('%3.1f', [PlatformVersionNumber]) + ' is installed.' #13#13 'Please install Platform 7 to use this Conform version.', mbCriticalError, MB_OK)
        else // A later Platform is installed that is younger than the maximum supported version
          MsgBox('This version of Conform only runs on ASCOM Platforms from ' + Format('%3.1f', [MINIMUM_PLATFORM_VERSION]) + ' to ' + Format('%3.1f', [MAXIMUM_PLATFORM_VERSION]) + 
            ', but Platform '+ Format('%3.1f', [PlatformVersionNumber]) + ' is installed.' #13#13 'For Platform 7, please use a later version of Conform or preferably Conform Universal, which has greater functionality.', mbCriticalError, MB_OK);
      end;
end;

// Code to enable the installer to uninstall previous versions of itself when a new version is installed
procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
  UninstallExe: String;
  UninstallRegistry: String;
begin
  if (CurStep = ssInstall) then
	begin
      UninstallRegistry := ExpandConstant('Software\Microsoft\Windows\CurrentVersion\Uninstall\{#SetupSetting("AppId")}' + '_is1');
      if RegQueryStringValue(HKLM, UninstallRegistry, 'UninstallString', UninstallExe) then
        begin
          MsgBox('Setup will now remove the previous version.', mbInformation, MB_OK);
          Exec(RemoveQuotes(UninstallExe), ' /SILENT', '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode);
          sleep(1000);    //Give enough time for the install screen to be repainted before continuing
        end
  end;
end;

[InnoIDE_PreCompile]
Name: 'C:\Users\Peter\Documents\Visual Studio Projects\Conform\BuildConform.cmd'; Flags: CmdPrompt;
