# Continuidad entre sesiones

## Estado del proyecto

Repositorio público con v0.1.0 y v0.1.1 publicadas como Releases, así que el chequeo de
actualizaciones funciona. v0.1.1 corrigió errores de conversión encontrados contra Office real. La
conversión básica de `.docx` y `.pptx` (Office cerrado, sin procesos huérfanos) se verificó a mano
el 2026-09-25; el resto del guion de humo manual no consta como ejecutado.

## Lo que está pendiente y por qué

- **Correr el guion de humo manual.** Es la única validación posible de la conversión. Hasta que se
  corra, no dé por buena ninguna afirmación sobre fidelidad ni sobre procesos huérfanos.
- **Excel.** Fuera del alcance de la primera versión, a pedido. Agregarlo es el mismo patrón con
  `Workbook.ExportAsFixedFormat` y una entrada más en `SupportedFormats`, el `.iss` y el `README`.
- **Firma de código.** No hay. Windows muestra la advertencia de SmartScreen al instalar. Inno Setup 7
  trae `ISSigTool` (firmas ECDSA propias, verificadas por el instalador), que resolvería la integridad
  del payload aunque no la reputación ante SmartScreen; no está implementado.
- **Variante HKLM.** Para despliegue centralizado por GPO en muchos equipos. Descartada de momento
  porque el requisito era evitar UAC.

## Antes de decir que algo está listo

1. ¿El CI está en verde en la rama?
2. Si tocó `Converters/` o `Interop/`, ¿le pidió al usuario el guion de humo manual?
3. Si cambió las extensiones soportadas, ¿tocó el código **y** el `.iss` **y** el `README`?
4. Si cambió un código de salida, ¿está en `ExitCodes.cs`, en el `README`, en el texto de ayuda y en
   el paso de humo del CI?
