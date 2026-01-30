; Inno Setup script for English Word Rotator

#define MyAppName "VocabBuddy"
#define MyAppVersion "0.5"
#define MyAppPublisher "Your Name"
#define MyAppExeName "english_app.exe"
#define MyAppDir "C:\\Users\\Administrator\\Documents\\Englishapp"

[Setup]
AppId={{7D44B3E2-6E43-4B89-9C6A-9F9B8B8A1F7A}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir={#MyAppDir}\installer_output
OutputBaseFilename=VocabBuddySetup
Compression=lzma
SolidCompression=yes
WizardStyle=classic
#ifexist "{#MyAppDir}\app.ico"
SetupIconFile={#MyAppDir}\app.ico
#endif

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "{#MyAppDir}\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

#ifexist "{#MyAppDir}\dist\words.csv"
Source: "{#MyAppDir}\dist\words.csv"; DestDir: "{app}"; Flags: ignoreversion
#endif

#ifexist "{#MyAppDir}\dist\words.txt"
Source: "{#MyAppDir}\dist\words.txt"; DestDir: "{app}"; Flags: ignoreversion
#endif

#ifexist "{#MyAppDir}\Donate-Icon-PNG-Clipart-Background.png"
Source: "{#MyAppDir}\Donate-Icon-PNG-Clipart-Background.png"; DestDir: "{app}"; Flags: ignoreversion
#endif

#ifexist "{#MyAppDir}\My logo.png"
Source: "{#MyAppDir}\My logo.png"; DestDir: "{app}"; Flags: ignoreversion
#endif

#ifexist "{#MyAppDir}\app.ico"
Source: "{#MyAppDir}\app.ico"; DestDir: "{app}"; Flags: ignoreversion
#endif

[Icons]
#ifexist "{#MyAppDir}\app.ico"
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\app.ico"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\app.ico"; Tasks: desktopicon
#else
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
#endif

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
