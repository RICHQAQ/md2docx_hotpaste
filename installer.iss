#define MyAppName "MD2DOCX HotPaste"
#define MyAppVersion "0.1.3.2"
#define MyAppPublisher "RichQAQ"
#define MyAppExeName "MD2DOCX-HotPaste.exe"

; 如果你用 onefile，SourceDir 就指向 dist
#define BuildDir "dist"
; 如果你用 onedir，且产物目录叫 MD2DOCX-HotPaste，就改成 dist\MD2DOCX-HotPaste
; #define BuildDir "dist\\MD2DOCX-HotPaste"

[Setup]
AppId={{87d83b72-8644-45ed-88d4-aa6c1ce7ce6b}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename=MD2DOCX-HotPaste_pandoc-Setup_v{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"


[CustomMessages]
english.CreateDesktopIcon=Create a desktop shortcut
english.AdditionalOptions=Additional options:
english.AutoStartup=Start automatically with Windows (current user)
english.RunAfterInstall=Launch {#MyAppName} after installation

chinesesimplified.CreateDesktopIcon=创建桌面快捷方式
chinesesimplified.AdditionalOptions=其他选项：
chinesesimplified.AutoStartup=开机自启（当前用户）
chinesesimplified.RunAfterInstall=安装完成后运行 {#MyAppName}

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalOptions}"
Name: "autorun"; Description: "{cm:AutoStartup}"; GroupDescription: "{cm:AdditionalOptions}"

[Files]
Source: "{#BuildDir}\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "third_party\pandoc\*"; DestDir: "{app}\pandoc"; Flags: ignoreversion recursesubdirs createallsubdirs
; 如果是 onedir，再把整个目录下的其他文件全部带上：
; Source: "{#BuildDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:RunAfterInstall}"; Flags: nowait postinstall skipifsilent

[Registry]
; 开机自启（当前用户）
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; \
  ValueType: string; ValueName: "{#MyAppName}"; ValueData: """{app}\{#MyAppExeName}"""; \
  Tasks: autorun; Flags: uninsdeletevalue

[UninstallDelete]
Type: filesandordirs; Name: "{userappdata}\{#MyAppName}"

[UninstallRun]
Filename: "taskkill"; Parameters: "/f /im {#MyAppExeName}"; Flags: runhidden

