# Contrato para agentes de IA

## Regla de esta carpeta

Toda la documentación dirigida a agentes vive en `.agents/` y en ningún otro sitio. Las referencias
van **en un solo sentido**: los archivos de `.agents/` pueden apuntar a `README.md`, `docs/`, `src/`,
`tests/`, `installer/` y los workflows; **ningún archivo fuera de `.agents/` apunta hacia dentro**.

La única excepción es `AGENTS.md` en la raíz, que se reduce a un puntero a este archivo porque las
herramientas lo cargan por convención. `README.md` y `docs/` se escriben para personas y no
mencionan `.agents/`.

Si escribe documentación nueva para agentes, va acá. Si escribe documentación de producto, va a
`docs/` y no puede citar esta carpeta.

## Qué es este proyecto

`NativeOfficeToPdf` convierte Word y PowerPoint a PDF usando el motor nativo de Office
(`ExportAsFixedFormat` por automatización COM), y expone la conversión como entrada "Convertir a PDF"
en el menú contextual del Explorador de Windows. Es una herramienta de escritorio, local, pequeña, sin
servidor ni servicios externos más allá de consultar los Releases del propio repositorio.

Lectura obligatoria antes de tocar código: [`01-lectura-obligatoria.md`](01-lectura-obligatoria.md).

## Principios obligatorios

- **Fidelidad antes que comodidad.** Cualquier vía que no sea el motor nativo de Office —imprimir a un
  PDF virtual, rasterizar, reinterpretar el documento— está descartada de entrada.
- **Sin UAC.** Ni al instalar ni al usar. Todo lo que se escribe va a `HKEY_CURRENT_USER` y a
  `%LOCALAPPDATA%`. Una propuesta que requiera administrador para el uso normal no es aceptable.
- **No romperle el Office al usuario.** Si la instancia de Word o PowerPoint ya estaba en marcha, es
  suya: se reutiliza y no se cierra. Lo que se cambie de su configuración se restaura.
- **No dejar procesos huérfanos.** Todo objeto COM se libera explícitamente. Es la causa clásica de
  degradación en este tipo de herramientas.
- **El chequeo de actualizaciones jamás bloquea ni altera un resultado.** Es accesorio por diseño.
- **Separar motor de presentación.** El motor no sabe del menú contextual; el instalador no sabe de
  conversión.

## Criterio de dependencias: todo en su versión actual

La aplicación es diminuta, así que mantener la punta de cada dependencia cuesta poco y evita deuda.
Se aplica a .NET, Inno Setup, xunit y las acciones de GitHub. **Ante un PR de Dependabot, la opción
por omisión es actualizar, no posponer**; posponer necesita una razón escrita.

Corolario, y es una decisión ya tomada: el proyecto principal **no tiene ni un paquete NuGet**. La
automatización de Office se hace por enlace tardío sobre COM, no con los PIA, porque los PIA de
PowerPoint arrastran un ensamblado que solo existe en NuGet como repaquetado de un tercero de 2016.
El razonamiento completo está en `docs/01-arquitectura.md`, sección "Por qué enlace tardío y no los
PIA". No lo revierta sin leerlo.

Antes de incorporar cualquier dependencia nueva, revisar licencia, mantenimiento, compatibilidad y
reemplazabilidad.

## Pruebas esperadas

`tests/NativeOfficeToPdf.Tests` cubre lo que se puede probar sin Office: parseo de argumentos,
resolución de rutas de destino, formatos soportados, comparación SemVer, lectura de la respuesta de
la API de releases, y la sincronía entre las extensiones declaradas en el código y las del
instalador.

Todo cambio en esa lógica viene con pruebas. Lo que toca COM no se prueba automáticamente: se valida
con el guion de humo manual (`docs/02-guion-humo-manual.md`), y **eso hay que pedirle al usuario que
lo corra** cuando el cambio afecta la conversión.

## Estilo de commits

Commits pequeños y descriptivos, en español:

- `feat(converters): agrega soporte de Excel`
- `fix(interop): libera el objeto Documents antes de cerrar Word`
- `test(cli): cubre destino que es carpeta existente`
- `ci: fija ISCC a la rama 7`
- `docs: documenta el código de salida 4`

## Documentación

Actualizar `docs/` cuando se adopte una decisión técnica importante: motor de conversión, mecanismo de
menú contextual, formato de distribución, chequeo de actualizaciones o estructura de la solución.
