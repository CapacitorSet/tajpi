[Setup]
AppName=Tajpi
AppVerName=Tajpi 2.97
DefaultDirName={pf}\Tajpi
DisableProgramGroupPage=yes
ShowLanguageDialog=no
UninstallDisplayIcon={app}\Tajpi.exe
;Admin rights required if VB6 runtimes are being installed.
;PrivilegesRequired=admin

[Files]
Source: "Tajpi.exe"; DestDir: "{app}"
Source: "Tajpi.skr"; DestDir: "{app}"; Flags: onlyifdoesntexist
Source: "Help\eo\Helpo.chm"; DestDir: "{app}"
Source: "Help\en\Helpo (angla).chm"; DestDir: "{app}"
;These are the VB runtime files. Not needed on versions of Windows 2000 and above as they ship with the runtimes
;by default, but future versions of Windows may require them to be installed.
;Source: "runtimes\stdole2.tlb";  DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
;Source: "runtimes\msvbvm60.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
;Source: "runtimes\oleaut32.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
;Source: "runtimes\olepro32.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
;Source: "runtimes\asycfilt.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile
;Source: "runtimes\comcat.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver

[Dirs]
Name: "{app}\"; Permissions: everyone-modify

[Tasks]
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked
Name: quicklaunchicon; Description: "Create a &Quick Launch icon"; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked

[Icons]
Name: "{commondesktop}\Tajpi"; Filename: "{app}\Tajpi.exe"; Tasks: desktopicon
Name: "{commonprograms}\Tajpi\Tajpi"; Filename: "{app}\Tajpi.exe"
Name: "{commonprograms}\Tajpi\Helpo"; Filename: "{app}\Helpo.chm"
Name: "{commonprograms}\Tajpi\Helpo (angla)"; Filename: "{app}\Helpo (angla).chm"
Name: "{commonprograms}\Tajpi\Malinstali Tajpi"; Filename: "{uninstallexe}"
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\Tajpi"; Filename: "{app}\Tajpi.exe"; Tasks: quicklaunchicon

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueName: "Tajpi"; Flags: uninsdeletevalue

[UninstallDelete]
Type: files; Name: "{localappdata}\Tajpi.ini"

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"; LicenseFile: "GPL.txt"

[Run]
Filename: "{app}\Tajpi.exe"; Description: "Launch Tajpi"; Flags: postinstall nowait skipifsilent unchecked

[Code]
function QuitTajpi(): Boolean;
var
  Wnd: HWND;
  abort : boolean;
  resultcode : integer;
begin
  abort := false;
  Wnd := FindWindowByWindowName('Tajpi Cxefa Fenestro');
  if Wnd <> 0 then SendMessage(Wnd, $0010, 0, 0); // WM_CLOSE
  Exec( '>', 'taskkill /f /im Tajpi.exe', '', SW_HIDE, ewNoWait, resultcode );
  result := not(abort);
end;

function InitializeSetup(): Boolean;
begin
  result := QuitTajpi();
end;

function InitializeUninstall(): Boolean;
begin
  result := QuitTajpi();
end;





