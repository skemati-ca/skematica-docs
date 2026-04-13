# Contribuir a skematica-docs

## Setup de Desarrollo

```bash
git clone https://github.com/skemati-ca/skematica-docs.git
cd skematica-docs
npm install
```

## Comandos

```bash
npm run dev          # Ejecutar servidor en modo desarrollo
npm run build        # Compilar TypeScript → dist/
npm test             # Ejecutar Vitest
npm run test:coverage # Ejecutar con cobertura
npm run lint         # ESLint
npm run typecheck    # tsc --noEmit
npm run check        # lint + typecheck + tests
```

## Cómo Añadir una Nueva Herramienta

### 1. Crear el handler

```bash
touch src/tools/word-mi-nueva-herramienta.ts
```

```typescript
import { validateDocxPath } from '../validation.js';

export const WORD_MI_NUEVA_HERRAMIENTA_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Ruta absoluta al archivo .docx' },
    // ... tus parámetros
  },
  required: ['filePath'],
} as const;

export async function wordMiNuevaHerramienta(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath } = args as { filePath: string };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  // ... tu lógica

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        // resultado
        _suggestions: {
          // herramientas relacionadas
        },
      }, null, 2),
    }],
  };
}
```

### 2. Registrar en `src/tool-registry.ts`

```typescript
// Import
const mod = await import('./tools/word-mi-nueva-herramienta.js');

// Registrar
tools.set('word_mi_nueva_herramienta', {
  name: 'word_mi_nueva_herramienta',
  description: 'Descripción de lo que hace',
  inputSchema: mod.WORD_MI_NUEVA_HERRAMIENTA_SCHEMA,
  handler: mod.wordMiNuevaHerramienta,
});
```

### 3. Habilitar en `src/config.ts`

Agregar el nombre al array `ALL_TOOLS`.

### 4. Escribir tests

```bash
touch tests/tools/word-mi-nueva-herramienta.test.ts
```

### 5. Documentar

Agregar la herramienta en `docs/mcp-tools.md` con:
- Descripción
- Schema de entrada con ejemplo
- Schema de salida con ejemplo
- Errores comunes

## Style Guide Resumido

- `const`/`let`, nunca `var`. ES modules (`import`/`export`), no `namespace`.
- Named exports, no default exports.
- TypeScript `private`, nunca `#private`. Sin `public`.
- Comillas simples. Template literals solo para interpolación.
- Siempre `===`/`!==`. Semicolón al final.
- Sin `any` — usar `unknown` o tipo específico.
- Sin aserciones de tipo (`as`) ni non-null (`!`).
- `UpperCamelCase` para tipos. `lowerCamelCase` para vars/funciones.

## Pull Requests

1. Rama descriptiva: `feat/word-mi-herramienta` o `fix/validation-bug`
2. Tests nuevos o actualizados
3. `npm run check` pasando en verde
4. Descripción clara del cambio
