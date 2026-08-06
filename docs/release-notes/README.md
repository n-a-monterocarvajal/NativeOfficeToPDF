# Notas de versión

Un archivo por versión publicada, llamado `vX.Y.Z.md`. El workflow de compilación distribuible los
exige: publicar sin notas falla antes de compilar.

## Formato

La primera línea, si empieza con `# `, se usa como **título del Release**; el resto es el cuerpo. El
workflow le agrega automáticamente la tabla de descargas y los `SHA-256` reales, así que no hay que
escribirlos a mano.

```markdown
# v0.2.0 — Soporte de Excel

## Novedades

- Convierte `.xlsx`, `.xlsm` y `.xls` con `Workbook.ExportAsFixedFormat`.

## Correcciones

- El PDF de destino ya no se bloquea cuando estaba abierto en un visor.

## Notas de actualización

- El instalador reemplaza la versión anterior sin desinstalarla primero.
```

Escrito para quien va a instalar la herramienta, no para quien escribió el código: qué cambia en el
uso, no qué clases se refactorizaron.
