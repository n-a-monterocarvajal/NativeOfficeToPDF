# Lectura obligatoria antes de modificar código

En este orden:

1. `README.md` — qué hace la herramienta y qué promete al usuario.
2. `docs/01-arquitectura.md` — las decisiones y, sobre todo, **por qué** están tomadas así. Casi
   todas las tentaciones de "simplificar" que va a tener ya están discutidas ahí.
3. `docs/03-automatizacion-github-actions.md` — qué valida el CI y qué no.

Según lo que vaya a tocar:

| Si va a tocar… | Lea primero |
|---|---|
| La conversión | `src/NativeOfficeToPdf/Converters/` completo y `Interop/ComObject.cs` |
| El CLI o los códigos de salida | `src/NativeOfficeToPdf/Cli/` y la tabla del `README.md` |
| El menú contextual | `src/NativeOfficeToPdf/Shell/ContextMenuRegistrar.cs` **y** `installer/NativeOfficeToPdf.iss` — escriben las mismas claves y hay una prueba que los compara |
| El chequeo de actualizaciones | `src/NativeOfficeToPdf/Update/` |
| El instalador | `installer/README.md` y el `.iss` |
| Los workflows | `docs/03-automatizacion-github-actions.md` |

## Trampas conocidas

- **`Marshal.GetActiveObject` no existe** en .NET moderno. Está reimplementado en
  `Interop/ActiveObject.cs`; no lo "arregle" volviendo a la API de .NET Framework.
- **Las llamadas a Office van por nombre de parámetro**, no por posición. `ExportAsFixedFormat` tiene
  catorce parámetros; reordenarlos es el error más fácil de cometer y el más difícil de ver.
- **`Visible = false` no funciona en PowerPoint.** La vía soportada es abrir con
  `WithWindow:=msoFalse`.
- **El binario es `WinExe`.** No lo cambie a `Exe` para "ver la salida": rompería el uso principal
  (menú contextual sin parpadeo). La salida por consola ya funciona vía `AttachConsole`, en
  `Cli/UserOutput.cs`.
- **Las extensiones soportadas están en dos sitios** —`Converters/SupportedFormats.cs` y el `.iss`—
  y una prueba los compara. Cambiar uno solo pone el CI en rojo, que es exactamente lo que se busca.
- **El runner no tiene Office.** No escriba pruebas automatizadas que abran Word o PowerPoint.
