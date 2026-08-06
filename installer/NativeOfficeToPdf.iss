; Instalador de NativeOfficeToPdf — Inno Setup 7.
;
; Diseñado para no pedir UAC en ningún momento: instala por usuario en %LOCALAPPDATA%\Programs y
; registra el verbo del menú contextual bajo HKEY_CURRENT_USER. Contrapartida conocida: en un equipo
; compartido hay que instalar con cada cuenta.
;
; Compilación:
;   1) dotnet publish src\NativeOfficeToPdf\NativeOfficeToPdf.csproj -c Release -r win-x64 ^
;        --self-contained false -o installer\payload
;   2) ISCC /DAppVersion=0.1.0 installer\NativeOfficeToPdf.iss
;
; La versión llega por /DAppVersion desde build-package.yml; el valor de abajo es solo el respaldo
; para compilar a mano.

#ifndef AppVersion
  #define AppVersion "0.1.0"
#endif

#define AppName "NativeOfficeToPdf"
#define AppPublisher "n-a-monterocarvajal"
#define AppUrl "https://github.com/n-a-monterocarvajal/NativeOfficeToPDF"
#define ExeName "NativeOfficeToPdf.exe"

; Estos tres valores están duplicados en el código (Shell\ContextMenuRegistrar.cs y
; Converters\SupportedFormats.cs). Hay pruebas que comparan ambas listas: si se cambian acá y no allá,
; el CI se pone en rojo.
#define VerbKeyName "ConvertToPdf"
#define VerbDisplayName "Convertir a PDF"
#define SupportedExtensions ".doc;.docx;.docm;.ppt;.pptx;.pptm"

[Setup]
AppId={{D3908383-BDD7-4AE2-B7A8-AF29323B6E95}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppUrl}
AppSupportURL={#AppUrl}/issues
AppUpdatesURL={#AppUrl}/releases
VersionInfoVersion={#AppVersion}

; Sin UAC: instalación por usuario, siempre.
PrivilegesRequired=lowest
DefaultDirName={localappdata}\Programs\{#AppName}
DisableDirPage=yes
DisableProgramGroupPage=yes
DisableWelcomePage=no
UsePreviousAppDir=yes

; El instalador escribe DOTNET_ROOT en HKCU\Environment cuando trae el runtime por su cuenta; esto
; hace que Inno avise al sistema del cambio en vez de dejarlo para el próximo inicio de sesión.
ChangesEnvironment=yes

; La app es AnyCPU y solo automatiza servidores COM fuera de proceso, así que no hay que casar
; arquitecturas con las de Office.
MinVersion=10.0

OutputDir=output
OutputBaseFilename={#AppName}-Setup
UninstallDisplayIcon={app}\{#ExeName}
UninstallDisplayName={#AppName} {#AppVersion}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern

[Languages]
#if FileExists(AddBackslash(CompilerPath) + "Languages\Spanish.isl")
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
#endif
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; El payload lo produce `dotnet publish` (ver encabezado). No se versiona: installer\payload está
; en .gitignore.
Source: "payload\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Registry]
; Verbo del menú contextual, una entrada por extensión soportada. `uninsdeletekey` en la clave del
; verbo deja el registro limpio al desinstalar.
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.doc\shell\{#VerbKeyName}"; ValueType: string; ValueName: "MUIVerb"; ValueData: "{#VerbDisplayName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.doc\shell\{#VerbKeyName}"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\{#ExeName}"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.doc\shell\{#VerbKeyName}\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#ExeName}"" ""%1"""

Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\{#VerbKeyName}"; ValueType: string; ValueName: "MUIVerb"; ValueData: "{#VerbDisplayName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\{#VerbKeyName}"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\{#ExeName}"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\{#VerbKeyName}\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#ExeName}"" ""%1"""

Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.docm\shell\{#VerbKeyName}"; ValueType: string; ValueName: "MUIVerb"; ValueData: "{#VerbDisplayName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.docm\shell\{#VerbKeyName}"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\{#ExeName}"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.docm\shell\{#VerbKeyName}\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#ExeName}"" ""%1"""

Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.ppt\shell\{#VerbKeyName}"; ValueType: string; ValueName: "MUIVerb"; ValueData: "{#VerbDisplayName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.ppt\shell\{#VerbKeyName}"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\{#ExeName}"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.ppt\shell\{#VerbKeyName}\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#ExeName}"" ""%1"""

Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pptx\shell\{#VerbKeyName}"; ValueType: string; ValueName: "MUIVerb"; ValueData: "{#VerbDisplayName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pptx\shell\{#VerbKeyName}"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\{#ExeName}"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pptx\shell\{#VerbKeyName}\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#ExeName}"" ""%1"""

Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pptm\shell\{#VerbKeyName}"; ValueType: string; ValueName: "MUIVerb"; ValueData: "{#VerbDisplayName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pptm\shell\{#VerbKeyName}"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\{#ExeName}"
Root: HKCU; Subkey: "Software\Classes\SystemFileAssociations\.pptm\shell\{#VerbKeyName}\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#ExeName}"" ""%1"""

; Solo cuando el instalador trae el runtime por su cuenta: el apphost lo busca en DOTNET_ROOT.
Root: HKCU; Subkey: "Environment"; ValueType: expandsz; ValueName: "DOTNET_ROOT"; ValueData: "{localappdata}\Microsoft\dotnet"; Flags: preservestringtype; Check: NeedsRuntimeBootstrap

[Run]
; Instalación del runtime .NET 10 por usuario, sin administrador, con el script oficial de Microsoft.
Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -Command ""& {{ $ErrorActionPreference='Stop'; $d = Join-Path $env:LOCALAPPDATA 'Microsoft\dotnet'; $s = Join-Path $env:TEMP 'dotnet-install.ps1'; Invoke-WebRequest -UseBasicParsing -Uri 'https://dot.net/v1/dotnet-install.ps1' -OutFile $s; & $s -Runtime dotnet -Channel 10.0 -InstallDir $d }"""; StatusMsg: "Instalando el runtime de .NET 10 para este usuario…"; Flags: runhidden waituntilterminated; Check: NeedsRuntimeBootstrap

[Code]
var
  BootstrapRuntime: Boolean;

{ Busca una carpeta shared\Microsoft.NETCore.App\10.* bajo la raíz dada. }
function HasNet10Runtime(const BaseDir: String): Boolean;
var
  Rec: TFindRec;
begin
  Result := False;
  if FindFirst(AddBackslash(BaseDir) + 'shared\Microsoft.NETCore.App\10.*', Rec) then
  try
    repeat
      { 16 = FILE_ATTRIBUTE_DIRECTORY }
      if (Rec.Attributes and 16) <> 0 then
      begin
        Result := True;
        Break;
      end;
    until not FindNext(Rec);
  finally
    FindClose(Rec);
  end;
end;

function DotNetRuntimePresent: Boolean;
begin
  Result := HasNet10Runtime(ExpandConstant('{commonpf64}\dotnet')) or
            HasNet10Runtime(ExpandConstant('{commonpf32}\dotnet')) or
            HasNet10Runtime(ExpandConstant('{localappdata}\Microsoft\dotnet'));
end;

function NeedsRuntimeBootstrap: Boolean;
begin
  Result := BootstrapRuntime;
end;

function InitializeSetup: Boolean;
begin
  Result := True;
  BootstrapRuntime := False;

  if DotNetRuntimePresent then
    Exit;

  if MsgBox('NativeOfficeToPdf necesita el runtime de .NET 10, que no está instalado en este equipo.'
      + #13#10#13#10 'Puede descargarlo e instalarlo ahora solo para su usuario, sin permisos de '
      + 'administrador (unos 70 MB).'
      + #13#10#13#10 '¿Hacerlo ahora?',
      mbConfirmation, MB_YESNO) = IDYES then
    BootstrapRuntime := True
  else
    MsgBox('La instalación va a continuar, pero el conversor no funcionará hasta que instale el '
      + 'runtime de .NET 10, o hasta que use la versión portable autocontenida publicada en los '
      + 'Releases del proyecto.', mbInformation, MB_OK);
end;
