# Arquitectura y decisiones de diseño

Este documento existe para que cualquiera pueda retomar, auditar o extender el proyecto sin volver a
derivar el razonamiento.

## 1. Objetivo y restricciones

- Fidelidad máxima en Word/PowerPoint → PDF: se descarta toda vía que imprima a un driver PDF virtual
  (Ghostscript, "Microsoft Print to PDF", etc.). Solo se usa el mismo motor que "Guardar como PDF" de
  la cinta de Office.
- Entorno de destino: Office 365 corporativo de 64 bits, ya instalado en todos los equipos.
- Permisos de administrador disponibles, pero **sin requerir UAC** ni en el uso diario ni en la
  instalación.
- Código propio y versionado; nada de binarios de terceros sin mantenimiento.

## 2. Arquitectura

Dos piezas independientes:

```
┌──────────────────────────────┐        ┌───────────────────────────────┐
│  Motor + CLI (C#)            │        │  Instalador (Inno Setup 7)    │
│  NativeOfficeToPdf.exe       │◄───────┤  installer/                   │
│  - Converters/               │  copia │  NativeOfficeToPdf.iss        │
│  - Cli/                      │  y     │  Registra HKCU\…\shell\…      │
│  - Shell/ContextMenuRegistrar│  registra                              │
│  - Update/                   │                                        │
└──────────────────────────────┘        └───────────────────────────────┘
```

El motor no sabe nada del menú contextual: es un CLI. El instalador no sabe nada de conversión: copia
un `.exe` y escribe claves de registro. Eso permite versionar y probar cada pieza por separado, y
reemplazar el mecanismo de menú contextual en el futuro —por ejemplo, por una extensión de shell en
C++— sin tocar el motor.

El propio ejecutable sabe registrarse y desregistrarse (`install` / `uninstall`), que es lo que usa la
versión portable y lo que cubren las pruebas. El instalador escribe las mismas claves por su cuenta;
una prueba compara ambas listas de extensiones para que no se desincronicen.

## 3. Motor de conversión

### 3.1 .NET 10, no .NET Framework 4.8

El planteamiento original era .NET Framework 4.8, porque es la ruta histórica de los PIA de Office.
Se descartó: el criterio del repositorio es mantener las dependencias en su versión actual, y este
proyecto es lo bastante pequeño como para que ese esfuerzo sea trivial. Se usa `net10.0-windows`,
framework-dependent, con un zip autocontenido disponible para equipos sin runtime.

Consecuencia a tener presente: `Marshal.GetActiveObject` no existe fuera de .NET Framework.
`Interop/ActiveObject.cs` lo reimplementa con P/Invoke a `CLSIDFromProgID` y `GetActiveObject`.

### 3.2 Por qué enlace tardío y no los PIA

El diseño original pedía enlace temprano con los *Primary Interop Assemblies* de Office publicados en
NuGet. Al implementarlo apareció un problema concreto: `Microsoft.Office.Interop.PowerPoint`
referencia el ensamblado `office` (`Microsoft.Office.Core`), del que salen tipos imprescindibles como
`MsoTriState` —el que se usa para abrir una presentación sin ventana—. Ese ensamblado **no está
publicado por Microsoft en NuGet**: el único paquete que lo trae es `Office 12.0.0`, de 2016, sobre
`net40`, subido por un tercero.

Depender de eso contradice de frente dos criterios del proyecto (dependencias actuales, nada de
binarios de terceros sin mantenimiento) para ganar comprobación en tiempo de compilación de unas
pocas llamadas. Se optó por lo contrario: **cero paquetes NuGet en el proyecto principal** y
automatización por `IDispatch` sobre los servidores COM que Office ya registró en la máquina.

Lo que se pierde —comprobación en compilación— se compensa así:

- `Interop/ComObject.cs` concentra todas las llamadas y permite pasar **argumentos por nombre**, tal
  como se documentan en la referencia VBA de Microsoft. Es exactamente el punto donde un error de
  orden de parámetros sería más probable, y queda cubierto.
- Las constantes (`wdExportFormatPDF = 17`, `ppFixedFormatTypePDF = 2`, …) se declaran con el nombre
  que tienen en la documentación de Office, en el conversor que las usa.
- El guion de humo manual (`docs/02-guion-humo-manual.md`) es el que valida de verdad la conversión;
  ningún CI puede hacerlo, porque el runner no tiene Office.

### 3.3 Reglas de la automatización

- **Reutilizar la instancia activa.** Cada conversor busca primero un Word o PowerPoint en marcha
  (`ActiveObject.TryGet`). Si lo encuentra, lo usa y **no lo cierra**; si no, crea uno y ese sí se
  cierra con `Quit()`. Es lo que evita cerrarle a alguien el Word que tenía abierto con un documento
  sin guardar.
- **Liberar COM explícitamente.** `ComObject.Dispose` llama a `Marshal.ReleaseComObject`, y
  `OfficeApplication.Dispose` remata con `GC.Collect()` y `GC.WaitForPendingFinalizers()`. Sin esto,
  las conversiones repetidas dejan procesos `WINWORD.EXE` y `POWERPNT.EXE` huérfanos: es la causa
  habitual de que estas herramientas "se degraden" con el uso.
- **Restaurar lo que se tocó.** `DisplayAlerts` y `AutomationSecurity` se leen antes y se reponen
  después, porque la instancia puede ser del usuario.
- **Apertura silenciosa.** Word abre con `ReadOnly`, sin agregar a recientes, invisible y con macros
  desactivadas (`msoAutomationSecurityForceDisable`); PowerPoint abre con `WithWindow:=msoFalse`, que
  es su forma soportada de trabajar sin ventana — asignarle `Visible = false` a la aplicación falla.
- **Documentos protegidos.** Word recibe una contraseña deliberadamente inverosímil, para que un
  documento con contraseña falle con excepción en vez de abrir un diálogo modal que nadie va a
  responder. PowerPoint no tiene un parámetro equivalente: ahí un archivo protegido es una limitación
  conocida.
- **`ExportAsFixedFormat`**, no `SaveAs2`. Es el método que Microsoft documenta específicamente para
  exportar a PDF/XPS, y el que da control fino sobre marcadores, propiedades del documento, etiquetas
  de estructura e IRM.
- **AnyCPU a propósito.** La automatización es COM *fuera de proceso*: Word y PowerPoint corren como
  ejecutables aparte, no como DLL cargada en nuestro proceso, así que COM hace el marshalling entre
  arquitecturas distintas y la bitness del binario no tiene que coincidir con la de Office.

## 4. WinExe, y cómo sigue siendo un CLI

El binario se compila como `WinExe`. Una app de consola haría parpadear una ventana negra cada vez
que se convierte desde el Explorador, que es el uso principal.

Eso, por sí solo, dejaría la herramienta muda en la línea de comandos. `Cli/UserOutput.cs` resuelve
el conflicto: al arrancar intenta `AttachConsole(ATTACH_PARENT_PROCESS)`. Si el proceso padre tiene
consola —lo invocó PowerShell o cmd—, escribe ahí como cualquier CLI. Si no la hay, los errores salen
por `MessageBox` y el éxito no muestra nada. La variable `NATIVEOFFICETOPDF_NO_DIALOGS=1` suprime los
diálogos para uso desatendido; el CI la usa para que un fallo no cuelgue el runner en un modal.

El precio de `WinExe`: PowerShell no espera a que el proceso termine. Está documentado en el README
con la forma correcta (`Start-Process -Wait -PassThru`).

## 5. Instalación sin UAC

Windows fusiona dos raíces de registro para resolver el menú contextual:

- `HKEY_LOCAL_MACHINE\Software\Classes` — por máquina, requiere administrador.
- `HKEY_CURRENT_USER\Software\Classes` — por usuario, **no requiere administrador**, y tiene
  prioridad sobre HKLM si hay conflicto.

Se escribe únicamente en
`HKCU\Software\Classes\SystemFileAssociations\<ext>\shell\ConvertToPdf\command`, para las seis
extensiones soportadas. El binario va a `%LOCALAPPDATA%\Programs\NativeOfficeToPdf\`, también una
ruta de usuario. Eso elimina el UAC en la instalación y en el uso, sin necesidad de ninguna extensión
de shell de terceros.

**Contrapartida**: el registro es por usuario, así que en un equipo compartido hay que instalar con
cada cuenta. Para un despliegue centralizado en muchos equipos sería razonable ofrecer además una
variante que escriba en HKLM, una sola vez, vía script de despliegue o GPO; no está implementada
porque el requisito explícito era evitar UAC.

## 6. Chequeo de actualizaciones

`Update/UpdateChecker.cs` consulta `GET /repos/{owner}/{repo}/releases/latest` de la API de GitHub,
de forma anónima. Tres reglas lo gobiernan:

1. **Nunca bloquea**: cinco segundos de tope, y cualquier excepción se traga. Jamás altera el código
   de salida de una conversión.
2. **Nunca molesta dos veces**: una consulta cada veinticuatro horas como mucho, un aviso por versión
   como mucho. El estado vive en `%LOCALAPPDATA%\NativeOfficeToPdf\update-check.json`.
3. **Nunca instala nada**: solo ofrece abrir la página del Release.

La comparación es SemVer (`Update/SemanticVersion.cs`) entre la etiqueta del Release y la versión del
ensamblado. Si el repositorio fuera privado, la API responde 404 sin credenciales y el chequeo
simplemente no encuentra nada; todo lo demás sigue funcionando.

## 7. Estructura del repositorio

```
NativeOfficeToPdf/
├── NativeOfficeToPdf.slnx          Solución en formato slnx
├── src/NativeOfficeToPdf/
│   ├── Program.cs                  Despacho de verbos y códigos de salida
│   ├── AppInfo.cs                  Versión, rutas, identidad del repositorio
│   ├── Cli/                        Parseo de argumentos, rutas, salida al usuario
│   ├── Converters/                 Word, PowerPoint, ciclo de vida de Office
│   ├── Interop/                    P/Invoke, ROT, envoltorio de IDispatch
│   ├── Shell/                      Registro del menú contextual
│   └── Update/                     Chequeo de versiones
├── tests/NativeOfficeToPdf.Tests/  xunit — lógica pura, sin Office
├── installer/NativeOfficeToPdf.iss Inno Setup 7, por usuario
└── docs/                           Este documento y sus vecinos
```

## 8. Referencias

- [`Document.ExportAsFixedFormat` (Word VBA)](https://learn.microsoft.com/es-es/office/vba/api/word.document.exportasfixedformat)
- [`Presentation.ExportAsFixedFormat` (PowerPoint VBA)](https://learn.microsoft.com/es-es/office/vba/api/powerpoint.presentation.exportasfixedformat)
- [`Documents.Open` (Word VBA)](https://learn.microsoft.com/es-es/office/vba/api/word.documents.open)
- [`Presentations.Open` (PowerPoint VBA)](https://learn.microsoft.com/es-es/office/vba/api/powerpoint.presentations.open)
