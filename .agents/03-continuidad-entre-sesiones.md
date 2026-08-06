# Continuidad entre sesiones

## Estado del proyecto

Repositorio inicializado con el motor de conversión, el CLI, el instalador, la automatización y la
documentación. **Nada de esto se ha ejecutado nunca contra Office real.** La primera compilación de
verdad es la del CI; la primera conversión de verdad la tiene que hacer una persona en un equipo con
Office, con `docs/02-guion-humo-manual.md` en la mano.

## Lo que está pendiente y por qué

- **Correr el guion de humo manual.** Es la única validación posible de la conversión. Hasta que se
  corra, no dé por buena ninguna afirmación sobre fidelidad ni sobre procesos huérfanos.
- **Visibilidad del repositorio.** El chequeo de actualizaciones consulta la API de Releases de forma
  anónima. Con el repositorio privado, esa consulta responde 404 y el chequeo no encuentra nada: todo
  lo demás funciona igual, pero la función no sirve hasta que el repositorio sea público y tenga al
  menos un Release.
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
