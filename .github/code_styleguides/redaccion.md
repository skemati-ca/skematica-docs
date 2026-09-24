# Guía de redacción

Aplica a todo texto que escriba una persona o un agente en Skemática: ADRs, specs,
planes, changelogs, runbooks, PRs, commits, correos, mensajes de error y copy de
producto.

El objetivo es uno: que quien lee encuentre lo que necesita, lo entienda en la
primera lectura y sepa qué hacer después.

## 1. Antes de escribir

Responde tres preguntas. Si no puedes, todavía no estás listo para escribir.

1. **¿Quién lo va a leer?** Por ejemplo: alguien del equipo sin contexto técnico,
   un agente que ejecutará un plan, un cliente que recibe un correo.
2. **¿Qué necesita hacer o decidir después de leer?**
3. **¿En qué situación lo lee?** Con prisa, durante un incidente, revisando un PR.

Incluye solo lo que esa persona necesita para eso. El contexto adicional va en un
anexo, un enlace o una sección aparte al final.

## 2. Estructura

- **Lo principal primero.** La conclusión, la decisión o la acción va en las
  primeras líneas. El razonamiento viene después.
- **Encabezados que dicen de qué trata la sección.** "Cómo pedir acceso", no
  "Generalidades". Quien solo lea los encabezados debe entender el documento.
- **Un tema por párrafo**, en párrafos cortos.
- **Listas numeradas para pasos en orden**; viñetas para elementos sin orden;
  tablas para comparar opciones o estados.
- **Separa el material de apoyo** (evidencia, historial, detalles técnicos) del
  flujo principal.

## 3. Oraciones y palabras

- **Una idea por oración.** Si una oración pasa de unas 25 palabras, casi siempre
  son dos.
- **Quién hace qué.** Nombra el sujeto y la acción.
  - Evita: "Se procederá a la creación del tag."
  - Escribe: "El workflow crea el tag."
- **Verbos, no sustantivos que esconden verbos.**
  - Evita: "realizar la validación", "tomar una decisión".
  - Escribe: "validar", "decidir".
- **Palabras que el lector usa.** Si un término técnico es necesario, explícalo la
  primera vez que aparece. Desarrolla cada sigla la primera vez.
- **Un concepto, un nombre.** Si empiezas con "versión de partida", no cambies a
  "baseline" o "versión inicial" más adelante.
- **Español por defecto.** Usa el término en inglés solo cuando es el que el lector
  usa en su trabajo (deploy, PR, tag, commit).
- **Cifras exactas**, con unidad y fecha: "59 commits desde el 2026-07-20", no
  "bastantes commits recientes".
- **Sin relleno.** Elimina frases como "Es importante mencionar que", "Cabe
  destacar", "En este sentido".
- **Tono directo y respetuoso.** Sin exageraciones ni frases de venta en documentos
  técnicos. Para correos y copy de producto, sigue la guía de voz de marca de
  `skematica-design` (tratamiento de usted o tú, según el caso).

## 4. Cierre

- **Di qué sigue**: qué debe hacer el lector, quién y cuándo.
- **Instrucciones como pasos numerados**, en imperativo, con el resultado esperado
  de cada paso cuando no sea obvio.
- **Una decisión pendiente se formula como pregunta concreta**, con las opciones y
  una recomendación.

## 5. Lo que no se hace

- No menciones esta guía ni ninguna norma, estándar o método de redacción dentro
  del texto que escribes. Se aplica, no se anuncia.
- No califiques tu propio texto ("explicado de forma sencilla", "en términos
  simples"). El lector lo juzga.
- No uses fórmulas de legibilidad como prueba de que un texto funciona. La prueba
  es que el lector encuentre, entienda y actúe.
- No sacrifiques precisión. En textos legales o contractuales, conserva el término
  exacto y explícalo; no lo reemplaces por uno aproximado.

## 6. Revisión antes de entregar

- [ ] Las primeras líneas dicen lo principal.
- [ ] Los encabezados, leídos solos, cuentan el documento.
- [ ] Cada término técnico y cada sigla se explica la primera vez.
- [ ] El mismo concepto tiene el mismo nombre en todo el texto.
- [ ] No hay oraciones con dos o más ideas que se puedan separar.
- [ ] Las cifras tienen unidad y fecha.
- [ ] El lector sabe qué hacer después.
- [ ] El texto no menciona esta guía ni ninguna norma de redacción.

## 7. Ejemplos

### Entrada de changelog

Evita:

> Se realizó la implementación de mejoras en la propagación de metadatos de
> anotaciones en el listado de herramientas.

Escribe:

> Las 38 herramientas de SECOP muestran su nombre legible en el Directorio de
> Claude.

### Mensaje de error de un check

Evita:

> Error: validación de changelog fallida.

Escribe:

> Falta la nota del cambio. Agrega una línea en `## [Unreleased]` de
> `CHANGELOG.md`, bajo `Fixed`, `Added` o `Changed`, y vuelve a hacer push.

### Apertura de un ADR

Evita:

> En el marco de la evolución de la plataforma y considerando diversos factores
> técnicos y organizacionales, se ha identificado la necesidad de revisar...

Escribe:

> Nada llega a producción sin versión. El deploy a producción se detiene si falta
> el corte de versión.
