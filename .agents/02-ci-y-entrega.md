# CI y entrega

El detalle de los workflows está en `docs/03-automatizacion-github-actions.md`. Acá van las reglas que
obligan a un agente.

## Reglas

- **Revisar el CI antes de dar por terminado un cambio.** Un cambio no está terminado mientras su PR
  tenga el CI en rojo, aunque el diff parezca correcto.
- **Una validación local no sustituye al runner limpio.** Solo el runner prueba el build desde cero,
  sin caché de `obj/` ni herramientas instaladas a mano. Si local pasa y el CI falla, la verdad la
  tiene el CI.
- **Este entorno no compila el proyecto.** Los agentes que corren en Linux no tienen SDK de .NET ni
  Windows: la primera compilación real de cualquier cambio es la del CI. Empujar y mirar el run no es
  opcional, es el único modo de verificar.
- **No agregar artefactos, matrices ni tareas programadas al CI** sin motivo técnico: los minutos de
  runner Windows se facturan al doble.
- **No publicar un Release sin pedido explícito del usuario**, igual que no se mergea un PR sin que lo
  pida.
- **No abrir un PR sin que lo pidan.** Trabajar en la rama indicada y empujar.

## Lo que el CI no puede decirle

Ninguna conversión real se prueba en CI: el runner no tiene Office. Cuando un cambio toca
`Converters/` o `Interop/`, el CI en verde significa "compila y la lógica alrededor está bien", no
"convierte bien". En ese caso hay que **pedirle al usuario que corra el guion de humo manual**
(`docs/02-guion-humo-manual.md`) antes de considerar el trabajo entregado, y decírselo explícitamente
en vez de dar por buena la conversión.
