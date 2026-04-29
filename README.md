# skematica-docs

**Servidor MCP para editar documentos de oficina de forma nativa — empezando por DOCX, luego XLSX y PPTX — desde Claude, ChatGPT, Gemini o cualquier cliente compatible con MCP.**

## Qué Es

`skematica-docs` es un servidor MCP local que permite a asistentes de IA **leer, editar y anotar** los documentos de oficina que ya tienes en tu máquina — **sin destruir formato, romper historial de revisiones ni corromper hilos de comentarios**.

En lugar de extraer texto plano (que pierde todo el formato) o generar un archivo nuevo (que descarta el original), este servidor realiza **ediciones quirúrgicas** directamente sobre las estructuras OOXML dentro de archivos `.docx`. Tu documento se ve exactamente igual — salvo por los cambios que pediste.

## Usuarios Objetivo

Profesionales que trabajan con documentos a diario y usan LLMs como parte de su flujo de trabajo:

- **Abogados** — revisan contratos, responden a comentarios de revisores, redactan cláusulas sin romper el control de cambios.
- **Equipos comerciales** — actualizan propuestas, ajustan tablas de precios, anotan feedback de clientes.
- **Gerentes** — revisan informes, aprueban/rechazan secciones de documentos, añaden comentarios estructurados.
- **Cualquier profesional** que trabaja con Word y quiere asistencia de IA que *respeta* el documento, no lo reemplaza.

## Capacidades

### DOCX (Fase 1 — Completa, 19 herramientas)

- **Navegación** — metadata del documento, contenido completo, estructura de secciones y contenido por sección.
- **Búsqueda y reemplazo** — búsqueda con contexto, reemplazo quirúrgico que preserva formato.
- **Comentarios** — listar hilos, crear, responder y resolver comentarios con jerarquía completa.
- **Layout** — leer y modificar tamaño de página, orientación y márgenes.
- **Estilos** — listar estilos y aplicarlos a párrafos individuales o rangos.
- **Comparación** — diff a nivel de carácter entre dos versiones del documento.
- **Notas al pie** — listar todas las notas con su texto.
- **Control de cambios** — insertar cambios rastreados (`w:del` / `w:ins`) visibles en el panel de Revisión de Word.

### XLSX (Fase 2 — Planificada)

- Leer hojas de cálculo con formato de celdas preservado.
- Editar valores, fórmulas y rangos de celdas.
- Añadir y responder comentarios a nivel de hoja y celda.

### PPTX (Fase 3 — Planificada)

- Leer contenido de diapositivas incluyendo notas del presentador.
- Editar texto de diapositivas, diseños y notas del presentador.
- Añadir comentarios de presentador a diapositivas.

## Por Qué Existe

La mayoría de herramientas de "IA para documentos" hoy te fuerzan a una de dos malas opciones:

1. **Extraer a texto plano** — el LLM puede leer y sugerir cambios, pero pierdes todo el formato, estilos, control de cambios y comentarios. Luego tienes que aplicar todo manualmente de vuelta en Word.
2. **Generar un documento nuevo** — el LLM crea un `.docx` fresco, pero el original se descarta junto con su historial de revisiones, comentarios de revisores y flujo de aprobación.

Ninguna opción funciona para flujos de trabajo profesionales donde **el historial de ediciones es la pista de auditoría**.

`skematica-docs` resuelve esto operando a nivel OOXML — manipulando las estructuras XML dentro del paquete `.docx` directamente. El resultado es un documento **visualmente idéntico** al original excepto por los cambios específicos que pediste, con todo el formato, comentarios y control de revisiones intacto.

## Inicio Rápido

### Requisitos

- [Node.js](https://nodejs.org/) 20+ instalado.
- Un cliente compatible con MCP (Claude Desktop, ChatGPT con MCP, VS Code, etc.).

### Instalación

```bash
npm install -g @skematica/docs
```

### Configura Tu Cliente MCP

Añade el servidor a la configuración de tu cliente MCP:

```json
{
  "mcpServers": {
    "skematica-docs": {
      "command": "npx",
      "args": ["-y", "@skematica/docs"]
    }
  }
}
```

### Úsalo

Una vez conectado, puedes pedirle a tu asistente de IA cosas como:

> "Lee mi contrato en `~/documents/contrato-v2.docx` y responde al comentario #5 diciendo que aceptamos la revisión."

> "En `~/propuestas/acme-propuesta.docx`, reemplaza 'Q1 2026' con 'Q2 2026' en el párrafo 3."

> "Abre `~/informes/trimestral.docx` y añade un comentario en la sección de ingresos marcándola para revisión."

El servidor maneja toda la complejidad OOXML — anclaje de comentarios, jerarquías de hilos, fragmentación de tramos, markup de revisiones — para que la IA pueda enfocarse en el contenido.

## Arquitectura de un Vistazo

```
┌─────────────────────────────────┐
│   Cliente MCP (Claude, GPT, …)  │
│         JSON-RPC sobre stdio    │
└──────────────┬──────────────────┘
               │
               ▼
┌─────────────────────────────────┐
│     Servidor skematica-docs     │
│   (local, sin auth, sin red)    │
│                                 │
│  19 herramientas DOCX:          │
│   • Navegación (4)              │
│   • Búsqueda y edición (2)      │
│   • Comentarios (4)             │
│   • Layout (4)                  │
│   • Estilos y comparación (3)   │
│   • Avanzado (2)                │
│                                 │
│  Motor: Edición quirúrgica OOXML│
│  (nivel DOM, preserva formato)  │
└──────────────┬──────────────────┘
               │
               ▼
┌─────────────────────────────────┐
│   Archivos .docx locales        │
└─────────────────────────────────┘
```

El servidor se ejecuta **solo localmente** — nunca transmite tus documentos por la red. Todas las operaciones de archivos ocurren en el sistema de archivos de tu máquina.

## Hoja de Ruta

| Fase | Formato | Estado |
|-------|--------|--------|
| 1 | **DOCX** — 19 herramientas: navegación, edición, comentarios, layout, estilos, notas al pie, control de cambios | Completa |
| 2 | **XLSX** — leer, editar, comentarios de celda | Planificada |
| 3 | **PPTX** — leer, editar, notas del presentador | Planificada |

## Licencia

Este proyecto es software libre y de código abierto bajo la **GNU General Public License v3.0 o posterior** (GPL-3.0-or-later). Consulta el archivo [LICENSE](LICENSE) para el texto completo de la licencia.

## Acerca de

`skematica-docs` es el primer proyecto de código abierto de [Skemática](https://skemati.ca) — una capa de inteligencia relacional tipada para descubrimiento emergente sobre datos públicos.
