# Guion de humo manual

El CI compila, corre las pruebas y comprueba los códigos de salida del binario, pero **no puede
probar una sola conversión**: el runner de GitHub no tiene Office instalado. Todo lo que sigue hay
que hacerlo a mano, en un equipo Windows con Office, antes de publicar una versión.

Marque cada punto. Si alguno falla, no se publica.

## Preparación

- [ ] Equipo con Windows 10 u 11 y Microsoft Office (Word y PowerPoint) instalados.
- [ ] Descargar `NativeOfficeToPdf-Setup.exe` del artefacto del workflow, o compilarlo localmente.
- [ ] Tener a mano un `.docx` con encabezados, tabla e imagen; un `.pptx` de varias diapositivas; un
      `.docx` protegido con contraseña; y un `.txt` cualquiera.

## Instalación

- [ ] Ejecutar el instalador **sin** "Ejecutar como administrador". **No aparece el diálogo de UAC**
      en ningún momento.
- [ ] Termina sin error y el programa queda en `%LOCALAPPDATA%\Programs\NativeOfficeToPdf\`.
- [ ] Aparece "NativeOfficeToPdf" en Configuración → Aplicaciones → Aplicaciones instaladas.
- [ ] Si el equipo no tenía el runtime de .NET 10, el instalador lo ofreció y lo instaló sin pedir
      administrador.

## Conversión desde el menú contextual

- [ ] Clic derecho sobre el `.docx` → aparece **Convertir a PDF** con el icono del programa.
- [ ] Al pulsarlo **no parpadea ninguna ventana de consola**.
- [ ] El PDF aparece junto al original, con el mismo nombre.
- [ ] Abrir el PDF: encabezados, tabla e imagen se ven igual que en Word; los marcadores del panel de
      navegación corresponden a los títulos del documento.
- [ ] Lo mismo con el `.pptx`: una página por diapositiva, sin recortes.
- [ ] Convertir el mismo archivo otra vez → pregunta si reemplazar el PDF existente. Responder que no
      deja el PDF anterior intacto.
- [ ] Seleccionar tres documentos a la vez y convertir: se generan los tres.

## Lo que no debe pasar

- [ ] Abrir Word con un documento **sin guardar**, convertir otro archivo desde el Explorador, y
      comprobar que **el Word del usuario sigue abierto y con sus cambios**.
- [ ] Después de cinco o seis conversiones seguidas, el Administrador de tareas **no** muestra
      procesos `WINWORD.EXE` ni `POWERPNT.EXE` huérfanos (con Office cerrado por el usuario).
- [ ] Convertir el `.docx` protegido con contraseña: falla con un cuadro de diálogo explicativo, **no
      se queda colgado** esperando una contraseña que nadie puede escribir.
- [ ] Clic derecho sobre el `.txt`: **no** aparece la entrada "Convertir a PDF".

## Línea de comandos

Desde PowerShell, en la carpeta del programa:

- [ ] `.\NativeOfficeToPdf.exe --version` imprime nombre y versión en la consola.
- [ ] `.\NativeOfficeToPdf.exe ..\ruta\informe.docx` convierte e imprime la ruta del PDF.
- [ ] `.\NativeOfficeToPdf.exe informe.docx C:\salida\` deja el PDF dentro de esa carpeta.
- [ ] `.\NativeOfficeToPdf.exe check-updates` responde algo (al día, o hay versión nueva) sin
      quedarse colgado.
- [ ] `$p = Start-Process .\NativeOfficeToPdf.exe -ArgumentList 'no-existe.docx' -Wait -PassThru;
      $p.ExitCode` devuelve `64`.

## Desinstalación

- [ ] Desinstalar desde "Aplicaciones instaladas".
- [ ] La carpeta `%LOCALAPPDATA%\Programs\NativeOfficeToPdf\` queda vacía o borrada.
- [ ] La entrada "Convertir a PDF" desaparece del menú contextual de las seis extensiones.
- [ ] En `regedit`, no queda ninguna clave `ConvertToPdf` bajo
      `HKCU\Software\Classes\SystemFileAssociations`.
