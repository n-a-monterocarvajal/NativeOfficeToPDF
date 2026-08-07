<#
.SYNOPSIS
    Genera src/NativeOfficeToPdf/NativeOfficeToPdf.ico a partir del glifo "file-pdf-box" de
    Material Design Icons.

.DESCRIPTION
    El icono se versiona ya generado; este script existe para que sea reproducible y para dejar
    constancia de su origen. No forma parte de la compilación.

    Por qué un rasterizador propio en vez de una biblioteca: el trazo de este glifo usa solo comandos
    absolutos M, H, V, C y Z, así que convertirlo a un GraphicsPath son cuarenta líneas y evita
    añadir una dependencia —o un navegador headless— al proyecto solo para dibujar un icono.

    Por qué el icono lleva color propio y no es el glifo monocromo tal cual: Windows no reajusta los
    iconos según el tema en aplicaciones sin empaquetar (eso son variantes `altform-lightunplated`
    de MSIX). El bitmap es uno solo y se dibuja igual sobre la barra de tareas clara y la oscura, así
    que el color tiene que aguantar ambas.

.NOTICE
    Glifo: Material Design Icons "file-pdf-box" — © Pictogrammers, licencia Apache 2.0.
    https://pictogrammers.com/library/mdi/icon/file-pdf-box/
#>
[CmdletBinding()]
param(
    [string] $Destino = (Join-Path $PSScriptRoot '..\src\NativeOfficeToPdf\NativeOfficeToPdf.ico'),
    [string] $Color = '#D0342C'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

# Trazo original del glifo, viewBox 0 0 24 24.
$Trazo = 'M19 3H5C3.9 3 3 3.9 3 5V19C3 20.1 3.9 21 5 21H19C20.1 21 21 20.1 21 19V5C21 3.9 20.1 3 19 3M9.5 11.5C9.5 12.3 8.8 13 8 13H7V15H5.5V9H8C8.8 9 9.5 9.7 9.5 10.5V11.5M14.5 13.5C14.5 14.3 13.8 15 13 15H10.5V9H13C13.8 9 14.5 9.7 14.5 10.5V13.5M18.5 10.5H17V11.5H18.5V13H17V15H15.5V9H18.5V10.5M12 10.5H13V13.5H12V10.5M7 10.5H8V11.5H7V10.5Z'

$Tamanos = 16, 24, 32, 48, 64, 128, 256

function ConvertFrom-SvgPath {
    <#
      Convierte el atributo `d` en un GraphicsPath. Solo entiende M, L, H, V, C y Z absolutos, que es
      lo que usa este glifo; cualquier otro comando lanza en vez de dibujar algo silenciosamente mal.
      El modo de relleno alterno es el que recorta las letras dentro de la caja.
    #>
    param([string] $D, [double] $Escala)

    $ruta = New-Object System.Drawing.Drawing2D.GraphicsPath
    $ruta.FillMode = [System.Drawing.Drawing2D.FillMode]::Alternate

    $piezas = [regex]::Matches($D, '([MLHVCZmlhvcz])|(-?\d*\.?\d+)')
    $indice = 0
    $comando = ''
    [double]$x = 0; [double]$y = 0
    $figuraAbierta = $false

    # Lee el siguiente número del trazo y lo devuelve ya escalado. El índice va por referencia: una
    # función anidada no comparte el ámbito local de su contenedora.
    $Siguiente = {
        param([ref] $Cursor)
        $valor = [double]::Parse($piezas[$Cursor.Value].Value, [Globalization.CultureInfo]::InvariantCulture)
        $Cursor.Value++
        return $valor * $Escala
    }

    while ($indice -lt $piezas.Count) {
        $pieza = $piezas[$indice].Value
        if ($pieza -match '^[MLHVCZmlhvcz]$') {
            $comando = $pieza
            $indice++
            if ($comando -eq 'Z' -or $comando -eq 'z') {
                if ($figuraAbierta) { $ruta.CloseFigure(); $figuraAbierta = $false }
                continue
            }
        }

        switch ($comando) {
            'M' {
                if ($figuraAbierta) { $ruta.CloseFigure() }
                $ruta.StartFigure()
                $figuraAbierta = $true
                $x = & $Siguiente ([ref]$indice); $y = & $Siguiente ([ref]$indice)
            }
            'L' {
                $nx = & $Siguiente ([ref]$indice); $ny = & $Siguiente ([ref]$indice)
                $ruta.AddLine($x, $y, $nx, $ny); $x = $nx; $y = $ny
            }
            'H' {
                $nx = & $Siguiente ([ref]$indice)
                $ruta.AddLine($x, $y, $nx, $y); $x = $nx
            }
            'V' {
                $ny = & $Siguiente ([ref]$indice)
                $ruta.AddLine($x, $y, $x, $ny); $y = $ny
            }
            'C' {
                $c1x = & $Siguiente ([ref]$indice); $c1y = & $Siguiente ([ref]$indice)
                $c2x = & $Siguiente ([ref]$indice); $c2y = & $Siguiente ([ref]$indice)
                $nx  = & $Siguiente ([ref]$indice); $ny  = & $Siguiente ([ref]$indice)
                $ruta.AddBezier($x, $y, $c1x, $c1y, $c2x, $c2y, $nx, $ny)
                $x = $nx; $y = $ny
            }
            default { throw "Comando SVG no soportado: '$comando'" }
        }
    }

    if ($figuraAbierta) { $ruta.CloseFigure() }
    return $ruta
}

function New-Lienzo {
    param([int] $Lado)

    $bmp = New-Object System.Drawing.Bitmap -ArgumentList ([int]$Lado), ([int]$Lado), ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)

    $ruta = ConvertFrom-SvgPath -D $Trazo -Escala ($Lado / 24.0)
    $brocha = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml($Color))
    $g.FillPath($brocha, $ruta)

    $brocha.Dispose(); $ruta.Dispose(); $g.Dispose()
    return $bmp
}

function ConvertTo-Dib {
    param([System.Drawing.Bitmap] $Bmp)

    $lado = $Bmp.Width
    $rect = New-Object System.Drawing.Rectangle -ArgumentList 0, 0, $lado, $lado
    $datos = $Bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $pixeles = New-Object byte[] ($datos.Stride * $lado)
    [System.Runtime.InteropServices.Marshal]::Copy($datos.Scan0, $pixeles, 0, $pixeles.Length)
    $Bmp.UnlockBits($datos)

    $filaMascara = [Math]::Floor(($lado + 31) / 32) * 4
    $ms = New-Object System.IO.MemoryStream
    $w = New-Object System.IO.BinaryWriter($ms)
    $w.Write([uint32]40)
    $w.Write([int32]$lado)
    $w.Write([int32]($lado * 2))     # alto XOR + máscara AND
    $w.Write([uint16]1)
    $w.Write([uint16]32)
    $w.Write([uint32]0)
    $w.Write([uint32]($lado * $lado * 4 + $filaMascara * $lado))
    $w.Write([int32]0); $w.Write([int32]0); $w.Write([uint32]0); $w.Write([uint32]0)

    for ($fila = $lado - 1; $fila -ge 0; $fila--) { $w.Write($pixeles, $fila * $datos.Stride, $lado * 4) }
    $w.Write((New-Object byte[] ($filaMascara * $lado)))
    $w.Flush()
    return $ms.ToArray()
}

function ConvertTo-Png {
    param([System.Drawing.Bitmap] $Bmp)

    $ms = New-Object System.IO.MemoryStream
    $Bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    return $ms.ToArray()
}

$imagenes = @()
foreach ($lado in $Tamanos) {
    $bmp = New-Lienzo -Lado $lado
    # 256 va como PNG, que es lo que Windows espera en el tamaño extra grande; el resto como DIB.
    $bytes = if ($lado -ge 256) { ConvertTo-Png -Bmp $bmp } else { ConvertTo-Dib -Bmp $bmp }
    $imagenes += [pscustomobject]@{ Lado = $lado; Bytes = $bytes }
    $bmp.Dispose()
}

$rutaDestino = [System.IO.Path]::GetFullPath($Destino)
$fs = New-Object System.IO.FileStream($rutaDestino, [System.IO.FileMode]::Create)
$w = New-Object System.IO.BinaryWriter($fs)
$w.Write([uint16]0); $w.Write([uint16]1); $w.Write([uint16]$imagenes.Count)

$desplazamiento = 6 + 16 * $imagenes.Count
foreach ($img in $imagenes) {
    $dim = if ($img.Lado -ge 256) { 0 } else { $img.Lado }   # 0 significa 256 en el formato ICO
    $w.Write([byte]$dim); $w.Write([byte]$dim); $w.Write([byte]0); $w.Write([byte]0)
    $w.Write([uint16]1); $w.Write([uint16]32)
    $w.Write([uint32]$img.Bytes.Length)
    $w.Write([uint32]$desplazamiento)
    $desplazamiento += $img.Bytes.Length
}
foreach ($img in $imagenes) {
    # Cast explícito: con la propiedad tal cual, PowerShell resuelve Write(byte) y escribe un byte.
    [byte[]]$datosImagen = $img.Bytes
    $w.Write($datosImagen, 0, $datosImagen.Length)
}
$w.Flush(); $w.Close(); $fs.Close()

"{0} ({1:N1} KB, {2} resoluciones)" -f $rutaDestino, ((Get-Item $rutaDestino).Length / 1KB), $imagenes.Count
