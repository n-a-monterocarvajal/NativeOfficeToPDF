# NativeOfficeToPdf

[![CI](https://github.com/n-a-monterocarvajal/NativeOfficeToPDF/actions/workflows/ci.yml/badge.svg)](https://github.com/n-a-monterocarvajal/NativeOfficeToPDF/actions/workflows/ci.yml)

Convierte documentos de Word y presentaciones de PowerPoint a PDF usando **el mismo motor que
"Guardar como PDF" de Office**, a través de una entrada **"Convertir a PDF"** en el menú contextual del
Explorador de Windows.

- **Sin permisos de administrador**, ni al instalar ni al usar.
- **Sin ventanas adicionales** al convertir desde el menú contextual.
- **Funciona también como CLI**, para scripts y conversiones por lote.

## Requisitos

| | |
|---|---|
| Sistema | Windows 10 o posterior |
| Office | Word y/o PowerPoint instalados (Microsoft 365, 2016 o posterior) |
| Runtime | .NET 10 |

## Instalación

Descargue `NativeOfficeToPdf-Setup.exe` de la
[última versión publicada](https://github.com/n-a-monterocarvajal/NativeOfficeToPDF/releases/latest)
y ejecútelo. Instala en `%LOCALAPPDATA%\Programs\NativeOfficeToPdf`, registra el menú contextual para
el usuario actual y deja su entrada en "Aplicaciones instaladas" de Windows.

En un equipo compartido hay que instalarlo con cada cuenta: el registro es por usuario.

**Versión portable**: el zip `…-selfcontained.zip` del mismo Release trae el runtime adentro. Se
descomprime donde sea y se registra el menú contextual con:

```powershell
.\NativeOfficeToPdf.exe install
```

Para quitarlo, `.\NativeOfficeToPdf.exe uninstall`.

## Uso

**Desde el Explorador**: clic derecho sobre un `.docx`, `.doc`, `.docm`, `.pptx`, `.ppt` o `.pptm` →
**Convertir a PDF**. El PDF queda junto al archivo original, con el mismo nombre. Si ya existía, pregunta
antes de reemplazarlo. Si algo falla, se advierte con un cuadro de diálogo; si todo sale bien, no
interrumpe.

Permite selección múltiple de archivos: el Explorador invoca la herramienta una vez por
archivo.

**Desde la línea de comandos**:

```
NativeOfficeToPdf.exe <origen> [destino] [opciones]
NativeOfficeToPdf.exe convert <origen> [destino] [opciones]
NativeOfficeToPdf.exe install | uninstall
NativeOfficeToPdf.exe check-updates
NativeOfficeToPdf.exe --version | --help
```

| Opción | Efecto |
|---|---|
| `--overwrite`, `-o` | Reemplaza el PDF de destino si ya existe |
| `--open` | Abre el PDF al terminar |
| `--quiet`, `-q` | No consulta si hay versiones nuevas de la aplicación |

Si el destino se omite, el PDF se genera junto al original. Si el destino es una carpeta, se usa
adentro el nombre del original.

> **Para scripts**: el binario es una aplicación de Windows sin consola propia, así que PowerShell no
> espera a que termine. Use `Start-Process -Wait -PassThru` para leer el código de salida:
>
> ```powershell
> $p = Start-Process NativeOfficeToPdf.exe -ArgumentList 'informe.docx' -Wait -PassThru
> $p.ExitCode
> ```
>
> La variable de entorno `NATIVEOFFICETOPDF_NO_DIALOGS=1` desactiva todos los cuadros de diálogo.

## Códigos de salida

Compatibles con la convención de OfficeToPDF, por si reutiliza scripts de verificación existentes.

| Código | Significado |
|---|---|
| 0 | Éxito |
| 1 | Fallo genérico (no se generó el PDF sin lanzar excepción) |
| 4 | `check-updates`: hay una versión más nueva de la aplicación |
| 8 | Argumentos inválidos |
| 32 | Extensión no soportada |
| 64 | Archivo de origen no encontrado |
| 2048 | Error durante la automatización de Office |

## Actualizaciones

Después de convertir, la herramienta consulta si hay una versión más nueva publicada. Se notifica una sola vez por versión. Nunca
descarga ni instala nada por su cuenta: solo ofrece abrir la página del Release. `check-updates`
fuerza la consulta.

## Compilar

```powershell
dotnet restore NativeOfficeToPdf.slnx
dotnet build NativeOfficeToPdf.slnx -c Release
dotnet test NativeOfficeToPdf.slnx -c Release
```

El binario queda en `src\NativeOfficeToPdf\bin\Release\net10.0-windows\NativeOfficeToPdf.exe`.
Para armar el instalador, ver [`installer/README.md`](installer/README.md).

## Cómo está hecho, y qué no hace

La arquitectura y las decisiones de diseño están en
[`docs/01-arquitectura.md`](docs/01-arquitectura.md). Lo que conviene saber de entrada:

- Si hay ventanas de Office abiertas, **lo reutiliza** y no lo cierra al terminar. Solo cierra las instancias
  que produce la misma aplicación.
- **Sin soporte para Excel** todavía.
- **Microsoft no soporta oficialmente automatizar Office de forma desatendida**. En uso interactivo
  como este el riesgo es bajo, pero un documento con macros de apertura, protegido por contraseña o
  abierto en Vista Protegida puede fallar o pedir intervención.
- **Sin resumen agregado** al convertir muchos archivos de una vez: cada uno reporta por su cuenta.

## Licencia

[MIT](LICENSE).
