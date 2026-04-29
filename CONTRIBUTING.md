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

## Flujo de Git y CI/CD

Este repositorio usa dos ramas permanentes:

- `develop`: base de trabajo diario.
- `main`: rama de producción; publicar desde aquí dispara deploy a npm.

Este flujo aplica a este repo. No define la política para aplicaciones que sí requieran un entorno online de staging.

Reglas:

1. Todo cambio empieza desde `develop`.
2. Crear una rama descriptiva: `feat/<nombre>`, `fix/<nombre>`, `docs/<nombre>` o `chore/<nombre>`.
3. Abrir PR de la rama de trabajo hacia `develop`.
4. `develop` ejecuta CI y valida el paquete con `npm pack --dry-run`.
5. Cuando `develop` esté listo para publicar, abrir PR de `develop` hacia `main`.
6. No se hacen merges directos de ramas de trabajo hacia `main`.
7. El deploy a npm ocurre solo por push/merge a `main`.
8. Si la versión de `package.json` ya existe en npm, el workflow de publicación se salta el publish.

Comandos típicos:

```bash
git switch develop
git pull origin develop
git switch -c feat/word-mi-herramienta

# trabajar, testear y commitear
npm run check
git push -u origin feat/word-mi-herramienta
```

Para publicar una nueva versión:

```bash
git switch develop
npm version patch # o minor/major según corresponda
npm run check
git push origin develop --follow-tags
```

Después se promueve con PR `develop` -> `main`. El workflow `Publish to npm` publica automáticamente el paquete de `main` usando el secret `NPM_TOKEN`.

Configuración requerida en GitHub:

- Secret de repositorio `NPM_TOKEN` con permiso de publicación para `@skematica/docs`.
- Branch protection en `develop` y `main`.
- Required status check: `Lint, Typecheck, Test, Build`.
- Required status check en PRs: `Branch Promotion Policy`.
- `main` debe aceptar merges únicamente desde PRs cuya rama origen sea `develop`.

Para configurar el secret sin exponerlo en logs ni commits:

```bash
gh secret set NPM_TOKEN
```

Pegá el token cuando `gh` lo pida por stdin. Si un token aparece en chat, issues, commits o logs, revocarlo en npm y generar uno nuevo.

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

1. Rama descriptiva creada desde `develop`: `feat/word-mi-herramienta` o `fix/validation-bug`
2. PR hacia `develop`
3. Tests nuevos o actualizados
4. `npm run check` pasando en verde
5. Descripción clara del cambio
