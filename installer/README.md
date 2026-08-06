# Instalador

`NativeOfficeToPdf.iss` es un script de **Inno Setup 7**. Produce `NativeOfficeToPdf-Setup.exe`, un
instalador por usuario que **no pide permisos de administrador**.

## Qué hace

| Paso | Dónde |
|---|---|
| Copia el programa | `%LOCALAPPDATA%\Programs\NativeOfficeToPdf\` |
| Registra "Convertir a PDF" | `HKCU\Software\Classes\SystemFileAssociations\<ext>\shell\ConvertToPdf` |
| Deja entrada de desinstalación | "Aplicaciones instaladas" de Windows |
| Comprueba el runtime .NET 10 | Y ofrece instalarlo solo para el usuario si falta |

Las seis extensiones registradas (`.doc`, `.docx`, `.docm`, `.ppt`, `.pptx`, `.pptm`) están
duplicadas entre este script y el código del programa. Hay una prueba automatizada que compara ambas
listas, así que una desincronización pone el CI en rojo antes de llegar a producción.

## Compilar a mano

Requiere [Inno Setup 7](https://jrsoftware.org/isdl.php) instalado.

```powershell
# 1. Publicar el payload (framework-dependent, la variante que distribuye el instalador)
dotnet publish src\NativeOfficeToPdf\NativeOfficeToPdf.csproj `
  -c Release -r win-x64 --self-contained false -o installer\payload

# 2. Compilar el instalador
& "${env:ProgramFiles(x86)}\Inno Setup 7\ISCC.exe" /DAppVersion=0.1.0 installer\NativeOfficeToPdf.iss
```

El resultado queda en `installer\output\NativeOfficeToPdf-Setup.exe`. Ni `payload\` ni `output\` se
versionan.

Lo normal, de todas formas, es no compilarlo a mano: el workflow `Compilación distribuible` lo hace
en un runner limpio y adjunta el `.sha256`.

## Alternativa sin instalar el runtime

Los Releases publican también un zip **autocontenido**: trae el runtime de .NET adentro, se
descomprime donde sea y se registra con `NativeOfficeToPdf.exe install`. Pesa bastante más, pero no
depende de nada instalado en el equipo.
