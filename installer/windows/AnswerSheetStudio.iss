[Setup]
AppName=Answer Sheet Studio
AppVersion=0.1.0
AppPublisher=Answer Sheet Studio
DefaultDirName={autopf}\Answer Sheet Studio
DefaultGroupName=Answer Sheet Studio
DisableDirPage=no
DisableProgramGroupPage=yes
OutputBaseFilename=AnswerSheetStudio-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\..\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion; Excludes: ".git\*;outputs\*;.venv\*;.pycache_tmp\*;installer\*"

[Icons]
Name: "{group}\Answer Sheet Studio"; Filename: "{app}\start_windows_terminal.bat"; WorkingDir: "{app}"
Name: "{commondesktop}\Answer Sheet Studio"; Filename: "{app}\start_windows_terminal.bat"; WorkingDir: "{app}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Run]
Filename: "{app}\start_windows_terminal.bat"; Description: "Launch Answer Sheet Studio"; Flags: nowait postinstall skipifsilent
