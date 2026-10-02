; ============================================================
; Instalador — ConvertidorBancos v5.1
; Edición completa (incluye Descarga Masiva SAT) | Serial HMAC + hardware binding
; ERSL Consultores 2026
; ============================================================

#define AppName      "ConvertidorBancos"
#define AppVersion   "5.1"
#define AppPublisher "ERSL Consultores"
#define AppExeName   "ConvertidorBancos.exe"
#define SourceDir    "dist\ConvertidorBancos"
; Certificado PUBLICO (.cer) de firma de codigo del editor. Debe estar junto a
; este .iss al compilar. Es SOLO la parte publica (nunca el .pfx/.key privado).
; Registra el editor de confianza para que las macros firmadas del reporte
; (papel de trabajo) corran sin el aviso "Habilitar macros".
#define CertFile     "cert_publisher.cer"
; CN exacto del certificado (para poder desinstalarlo). Debe coincidir con el
; Subject usado al generarlo: New-SelfSignedCertificate -Subject "CN=..."
#define CertCN       "ERSL Consultores"

[Setup]
; AppId PROPIO de la 5.1 — distinto del de la 5.0 (son productos separados).
AppId={{7C4E1B92-3A6D-4F51-9E88-2B7A0C5D14F3}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} v{#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\ConvertidorBancos
DefaultGroupName={#AppName}
AllowNoIcons=yes
OutputDir=instalador_salida
OutputBaseFilename=ConvertidorBancos_Setup_v5.1
SetupIconFile=app_icon.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
; Registro del cert a nivel MAQUINA (todas las cuentas) requiere admin.
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName} v{#AppVersion}
VersionInfoVersion=5.1.0.0
VersionInfoCompany={#AppPublisher}
VersionInfoDescription=ConvertidorBancos v5.1 - Compliance + Conciliacion CFDI + Descarga Masiva SAT + Convertidor PDF a Excel
VersionInfoProductName={#AppName}
VersionInfoProductVersion=5.1.0.0

; ── Firma Authenticode del instalador y del desinstalador ──────────────────
; Requiere un certificado de firma de codigo (idealmente OV/EV de una CA:
; DigiCert, Sectigo, GlobalSign). Un cert AUTOFIRMADO NO detiene el aviso de
; SmartScreen/Defender — solo un cert de CA de confianza construye reputacion.
;
; Configura el sign tool UNA vez en el IDE de Inno Setup:
;   Tools > Configure Sign Tools... > Add
;   Nombre: firmar
;   Comando:
;     "C:\Program Files (x86)\Windows Kits\10\bin\x64\signtool.exe" sign /f "$q{#SignPfx}$q" /p $qTU_PASSWORD$q /fd sha256 /tr http://timestamp.digicert.com /td sha256 $f
; La contrasena queda SOLO en la config local del IDE, nunca en este .iss.
; Luego descomenta las dos lineas de abajo:
;#define SignPfx "C:\ruta\privada\ersl_codesign.pfx"
;SignTool=firmar
;SignedUninstaller=yes

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el Escritorio"; GroupDescription: "Iconos adicionales:"; Flags: unchecked

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Certificado publico -> se extrae a {tmp} solo durante la instalacion y se borra al terminar.
Source: "{#CertFile}"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Icons]
Name: "{group}\{#AppName}";       Filename: "{app}\{#AppExeName}"
Name: "{group}\Desinstalar";      Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
; ── Registrar editor de confianza para macros firmadas (nivel MAQUINA, requiere admin) ──
Filename: "certutil.exe"; Parameters: "-addstore -f TrustedPublisher ""{tmp}\{#CertFile}"""; StatusMsg: "Registrando editor de confianza para macros..."; Flags: runhidden
; Root: SOLO para certificado autofirmado (self-signed). Con cert de CA comercial, borra esta linea.
Filename: "certutil.exe"; Parameters: "-addstore -f Root ""{tmp}\{#CertFile}"""; StatusMsg: "Registrando cadena de confianza del certificado..."; Flags: runhidden
Filename: "{app}\{#AppExeName}"; Description: "Iniciar {#AppName}"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "certutil.exe"; Parameters: "-delstore TrustedPublisher ""{#CertCN}"""; Flags: runhidden; RunOnceId: "DelTrustedPublisher"
Filename: "certutil.exe"; Parameters: "-delstore Root ""{#CertCN}"""; Flags: runhidden; RunOnceId: "DelRoot"

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Messages]
WelcomeLabel1=Bienvenido al asistente de instalación de [name]
WelcomeLabel2=Este programa instalará [name/ver] en su computadora.%n%nDesarrollado por ERSL Consultores.%n%nModulos incluidos:%n  - Convertidor de estados de cuenta PDF a Excel (multi-banco)%n  - Conciliacion Bancaria vs CFDI (XML del SAT)%n  - Descarga Masiva SAT (Web Service + Portal)%n  - Obligaciones Fiscales y Laborales (Compliance)%n  - Generador de Polizas COI (Aspel)%n  - Envio de correo multi-proveedor%n%nBancos soportados:%n  Santander, BBVA, Banregio, Banorte,%n  Inbursa, Banbajio, HSBC, Banamex, Konfio.%n%nRequiere usuario y contraseña al iniciar.%nRequiere serial de activación (XXXXXX-XXXXXX-XXXXXX-XXXXXX), atado al equipo donde se active.%n%nEste instalador registrará el editor de confianza de {#AppPublisher} en este equipo (todas las cuentas de usuario), para que las macros firmadas del reporte de conciliación (papel de trabajo) funcionen sin la advertencia de macros de Excel. Requiere permisos de administrador.%n%nSe recomienda cerrar todas las demás aplicaciones antes de continuar.
