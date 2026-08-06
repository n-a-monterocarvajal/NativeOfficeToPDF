# Automatización en GitHub Actions

GitHub Actions es parte del procedimiento normal de validación y entrega de este repositorio, no un
extra opcional.

| Workflow | Archivo | Finalidad | Disparadores |
|---|---|---|---|
| CI | `.github/workflows/ci.yml` | Restaura, compila en `Release`, corre las pruebas y comprueba los códigos de salida del binario | PR hacia `main`, push a `main`, `workflow_dispatch`. No corre si el cambio es solo documentación |
| Compilación distribuible | `.github/workflows/build-package.yml` | Genera el instalador Inno Setup 7 y/o el zip autocontenido, con su `SHA-256`; opcionalmente publica el Release | `workflow_dispatch` (entradas `version`, `artifact`, `publish`) y push de tag `v*`. Nunca en cada commit |
| Dependabot | `.github/dependabot.yml` | Actualiza NuGet y `github-actions`, agrupado y semanal | Lunes |

## Qué valida el CI, y qué no

Corre en `windows-latest` porque el proyecto es `net10.0-windows` y automatiza COM.

**Sí valida**: que la solución restaura y compila en limpio; que pasan las pruebas de lógica pura
(parseo de argumentos, resolución de rutas, comparación SemVer, lectura de la API de releases,
sincronía entre las extensiones del código y las del instalador); y que el ejecutable arranca y
devuelve los códigos de salida documentados.

**No valida** —y no puede— ninguna conversión real: el runner no tiene Office. Esa parte vive en
[`02-guion-humo-manual.md`](02-guion-humo-manual.md).

El job exporta `NATIVEOFFICETOPDF_NO_DIALOGS=1`. Sin eso, un fallo del paso de humo abriría un cuadro
de diálogo modal y el job se colgaría hasta el timeout en vez de fallar con su mensaje.

## Sin caché de NuGet

A propósito. El proyecto principal no tiene ni un paquete —la automatización de Office es por COM— y
la suite de pruebas trae tres paquetes pequeños. Cachear costaría más de lo que ahorra.

## Inno Setup 7 en el runner

La imagen `windows-latest` trae Inno Setup preinstalado, históricamente en la rama 6. El repositorio
se compila con la 7, así que el paso no busca "un ISCC" cualquiera: filtra por versión mayor ≥ 7 y, si
no encuentra ninguno, descarga e instala la última estable desde `jrsoftware.org` en silencio. El log
del paso imprime qué ISCC terminó usando; conviene mirarlo cuando algo del instalador se comporte
raro.

## Publicar una versión

1. Escribir `docs/release-notes/vX.Y.Z.md` (ver [`release-notes/README.md`](release-notes/README.md)).
2. Actualizar `<Version>` en `src/NativeOfficeToPdf/NativeOfficeToPdf.csproj`.
3. Correr el guion de humo manual sobre el artefacto.
4. Lanzar `Compilación distribuible` desde `main` con `artifact: both` y `publish: yes`.

El workflow valida antes de compilar que la versión tenga formato válido, que existan las notas y que
el tag y el Release no existan ya. La publicación ocurre en un job aparte —el único con permiso de
escritura— que verifica los `SHA-256` de lo descargado antes de crear el Release.

## Ahorro de minutos

El repositorio corre en runners Windows, que se facturan al doble. Por eso:

- `paths-ignore` deja fuera cambios que solo tocan documentación.
- `concurrency` con `cancel-in-progress` descarta runs obsoletos.
- La compilación distribuible nunca corre sola.
- Dependabot agrupa sus PR en uno por ecosistema y por semana.
