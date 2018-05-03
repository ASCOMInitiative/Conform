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
#define MyAppVer GetFileVersion("..\Conform\bin\Release\Conform.exe") ; define variable
AppID={{CEDF90C4-A415-4903-94D6-9099F7AE41B5}
AppCopyright=Copyright � 2017 Peter Simpson
AppName=ASCOM Conform
AppPublisher=Peter Simpson
AppPublisherURL=http://www.ascom-standards.org
AppSupportURL=http://tech.groups.yahoo.com/group/ASCOM-Talk/
AppUpdatesURL=http://www.ascom-standards.org
#emit "AppVerName=ASCOM Conform " + MyAppVer
#emit "AppVersion=" + MyAppVer
;ArchitecturesInstallIn64BitMode=x64
Compression=lzma/Max
DefaultDirName={pf}\ASCOM\Conform
DefaultGroupName=ASCOM Platform 6\Tools
MinVersion=0,5.0.2195sp4
#emit "OutputBaseFilename=Conform(" + MyAppVer +")setup"
OutputDir=.\Builds
PrivilegesRequired=admin
SetupIconFile=..\Conform\Resources\ASCOM.ico
SetupLogging=true
ShowLanguageDialog=auto
SolidCompression=true
UninstallDisplayIcon={app}\Conform.exe
VersionInfoCompany=ASCOM Initiative
VersionInfoCopyright=Peter Simpson
VersionInfoDescription=Tool to validate ASCOM drivers
VersionInfoProductName=ASCOM Conform
#emit "VersionInfoProductVersion=" + MyAppVer
#emit "VersionInfoVersion=" + MyAppVer
WizardImageFile=NewWizardImage.bmp
WizardSmallImageFile=ASCOMLogo.bmp

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
Name: {group}\Conform; Filename: {app}\Conform.exe
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
Name: "C:\Users\Peter\Documents\Visual Studio Projects\Conform\BuildConform.cmd"; Flags: CmdPrompt

[Code]
//
// Before the installer UI appears, verify that the (prerequisite)
// ASCOM Platform 5.5 is installed, including both Helper components.
// Helper is required for all typpes (COM and .NET)!
//
function InitializeSetup(): Boolean;
var
   H : Variant;
   H2 : Variant;
   PlatVerString : String;
   PlatVer : Variant;
   DoubleValue : Variant;
   DoubleValueString : String;
   Separator : String;
begin
   Result := FALSE;  // Assume failure
   try               // Will catch all errors including missing reg data
//     MsgBox('Before DriverHelper.Util', mbCriticalError, MB_OK);
      H := CreateOLEObject('DriverHelper.Util');  // Assure both are available

//      MsgBox('Before DriverHelper.Util2', mbCriticalError, MB_OK);
      H2 := CreateOleObject('DriverHelper2.Util');

//      MsgBox('After DriverHelper.Util2', mbCriticalError, MB_OK);
      DoubleValue := 1.0 / 3.0; // Create a real number of value 0.33333
      DoubleValueString := DoubleValue; // Get the real number as a string including this system's decimal separator
      Separator := Copy(DoubleValueString,2,1); // Parse out the decimal separator

//      MsgBox('After Separator ' + DoubleValueString, mbCriticalError, MB_OK);
      PlatVerString := H2.PlatformVersion; // Get the Platform version string (which has "." as its decimal separator)

//      MsgBox('After H2.PlatformVersion', mbCriticalError, MB_OK);
      StringChangeEx(PlatVerString, '.', Separator, True); // Change the "." to the current system's decimal separator

//      MsgBox('PlatVerString: ' + Separator + ' ' + PlatVerString, mbCriticalError, MB_OK)
      
      PlatVer := PlatVerString; // Create a variant from the modified string containing this system's decimal separator
      if (PlatVer < 6.2) then // Check for installed platform below 6.2
          begin
            MsgBox('ASCOM Platform 6.2 or later is required for Conform 6.3 to work correctly. Please install the latest Platform before continuing; you will find it at http:\\www.ascom-standards.org', mbCriticalError, MB_OK)
          end
          else // Platform 5.5 or greater is installed so allow the installation to proceed
              Result:= TRUE;
   except

      begin
//        ShowExceptionMessage;
        MsgBox('ASCOM Platform 6.2 or later is required for Conform 6.3 to work correctly. Please install the latest Platform before continuing; you will find it at http:\\www.ascom-standards.org', mbCriticalError, MB_OK);
//        Result:= TRUE;
      end;
   end;

end;

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
