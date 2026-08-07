# Menú contextual de Windows 11: análisis de la vía sparse package

Estado: **análisis, no implementado**. Este documento existe para que la decisión no haya que volver a
investigarla, y para que quien la retome sepa exactamente qué cuesta.

## 1. El problema

El instalador registra el verbo como una clave clásica en
`HKCU\Software\Classes\SystemFileAssociations\<ext>\shell\ConvertToPdf`. Eso funciona, pero en
Windows 11 las entradas clásicas no aparecen en el menú principal: quedan detrás de **"Mostrar más
opciones"** (o Shift+F10). El menú principal solo muestra comandos que implementan
[`IExplorerCommand`](https://learn.microsoft.com/en-us/windows/apps/desktop/modernize/integrate-packaged-app-with-file-explorer)
y cuya aplicación tiene **identidad de paquete**.

## 2. La vía sin MSIX completo

No hace falta empaquetar la aplicación como MSIX. Basta un **identity package** (o *sparse package*):
un paquete que solo aporta identidad y deja los binarios donde ya los pone el instalador, mediante
`uap10:AllowExternalContent`. La documentación oficial es
[Grant package identity by packaging with external location](https://learn.microsoft.com/en-us/windows/apps/desktop/modernize/grant-identity-to-nonpackaged-apps).

Las piezas serían tres:

| Pieza | Qué es | Dónde vive |
|---|---|---|
| Manifiesto de identidad | `AppxManifest.xml` con `AllowExternalContent`, empaquetado con `MakeAppx pack /nv` y firmado con `SignTool` | nuevo, en el repositorio |
| Handler `IExplorerCommand` | DLL COM in-process que muestra el ítem y lanza `NativeOfficeToPdf.exe` con la ruta | componente nuevo |
| Registro en el instalador | `Add-AppxPackage -Path <pkg>.msix -ExternalLocation <dir>` desde `[Run]`, y `Remove-AppxPackage` al desinstalar | `installer/NativeOfficeToPdf.iss` |

Precedente directo y del mismo perfil que este proyecto —instalación por usuario con Inno Setup—:
[microsoft/vscode-explorer-command](https://github.com/microsoft/vscode-explorer-command), la extensión
de shell de VS Code. Ejemplos de implementación: [microsoft/AppModelSamples — SparsePackages](https://github.com/microsoft/AppModelSamples/tree/master/Samples/SparsePackages)
y [cjee21/IExplorerCommand-Examples](https://github.com/cjee21/IExplorerCommand-Examples).

## 3. La firma no obliga a pedir administrador

Conviene dejarlo claro porque es contraintuitivo y es el punto donde es fácil descartar la vía por
error: el paquete **debe ir firmado**, pero para una instalación **por usuario** el certificado público
se importa en `Cert:\CurrentUser\TrustedPeople`, que **no requiere elevación**. Es la instalación
*por máquina* la que usa `Cert:\LocalMachine\TrustedPeople` y sí exige administrador.

Es decir: con un certificado autofirmado y el diseño por usuario que ya tiene el proyecto, **ningún
paso pide UAC**. Sin firma no hay registro posible: `Add-AppxPackage` falla con `CERT_E_UNTRUSTEDROOT`
(`0x800B0109`).

El coste real de esto no es técnico sino de confianza: meter un certificado propio en el almacén
"Personas de confianza" del usuario significa que ese usuario pasará a confiar en cualquier paquete
firmado con él. Si se implementa, el instalador debe decirlo explícitamente, no colarlo.

## 4. El coste que sí es serio: la DLL vive dentro del Explorador

`IExplorerCommand` **solo admite activación in-process**. La DLL se carga dentro de `explorer.exe` al
construir el menú y permanece ahí toda la sesión. De ahí se derivan tres costes permanentes:

1. **Residente en memoria** del proceso del shell, sin descarga.
2. **Archivo bloqueado**: al desinstalar o actualizar, el `.dll` no se puede borrar mientras Explorer
   siga vivo. Hay que programar el borrado al reiniciar o matar y relanzar `explorer.exe`, con lo que
   la desinstalación limpia pasa a depender de eso.
3. **Sus fallos son fallos del Explorador**: un cuelgue o una excepción degradan el shell del usuario,
   y el menú nuevo impone además un tiempo límite de respuesta.

Compárese con el diseño actual: un verbo de registro **no carga nada** en Explorer. Cero código en el
proceso del shell, cero archivos bloqueados, cero superficie de fallo. Su único defecto es vivir en el
menú secundario.

La elección, por tanto, no es entre dos formas de poner una entrada, sino entre *"sin código en el
shell, menú secundario"* y *"código residente en el shell, menú principal"*.

## 5. C++ o C# con NativeAOT

El handler no comparte una sola línea de lógica con el motor: solo muestra el ítem y lanza el
ejecutable con la ruta. Unas 200 líneas que no se tocan nunca.

| | C++ | C# + NativeAOT |
|---|---|---|
| Cadena de compilación | añade MSVC al CI, hoy solo `dotnet` | sigue siendo `dotnet publish` |
| Precedente | VS Code y prácticamente todo el ecosistema | poco recorrido en extensiones de shell |
| Postura de Microsoft | vía recomendada | lenguajes gestionados desaconsejados in-process |
| Compilación por arquitectura | x64 + arm64 obligatorio | x64 + arm64 obligatorio |

El "desaconsejado" nace de cargar CoreCLR dentro de Explorer, y NativeAOT elimina el runtime, así que
la objeción se debilita; .NET 8+ trae generadores de origen de COM para esto. Pero sigue siendo
terreno poco pisado para extensiones de shell.

**Recomendación: C++**, no por afinidad con el proyecto —no la hay por ninguna de las dos vías— sino
porque el proceso anfitrión es el shell del usuario y ahí el camino trillado vale más que la unidad de
lenguaje. La simplicidad de `AnyCPU` se pierde igual en ambos casos.

## 6. Qué haría falta decidir antes de empezar

- Si el menú principal compensa una DLL residente en el Explorador y una desinstalación que depende de
  reiniciarlo.
- Cómo se obtiene y renueva el certificado, y cómo se le explica al usuario lo que se le pide confiar.
- Dos arquitecturas más en el CI (x64 y arm64) y su firma en el pipeline.
- Si el verbo clásico se mantiene en paralelo para Windows 10, donde el menú nuevo no existe.
