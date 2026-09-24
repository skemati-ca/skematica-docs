#!/usr/bin/env bash
# Revisa los Markdown modificados según conductor/code_styleguides/redaccion.md:
# el texto no debe citar normas de redacción.
# - Falla si un archivo cita la norma por su número.
# - Avisa (sin fallar) si un archivo describe su propio estilo de redacción.
set -euo pipefail

BLOCK_PATTERN='24495'
WARN_PATTERN='lenguaje claro|plain language'

collect_changed_markdown() {
  local base_ref range
  base_ref="${DOCS_BASE_REF:-origin/main}"
  range=""

  if git rev-parse --verify "$base_ref" >/dev/null 2>&1; then
    range="$base_ref...HEAD"
  elif git rev-parse --verify main >/dev/null 2>&1; then
    range="main...HEAD"
  elif git rev-parse --verify HEAD~1 >/dev/null 2>&1; then
    range="HEAD~1...HEAD"
  fi

  {
    if [[ -n "$range" ]]; then
      git diff --name-only "$range" -- '*.md'
    fi
    git diff --name-only --cached -- '*.md'
    git diff --name-only -- '*.md'
    git ls-files --others --exclude-standard -- '*.md'
  } | sort -u | while IFS= read -r file; do
    [[ -z "$file" ]] && continue
    [[ ! -f "$file" ]] && continue
    case "$file" in
      tmp/*) continue ;;
    esac
    printf '%s\n' "$file"
  done
}

mapfile -t files < <(collect_changed_markdown)

if [[ ${#files[@]} -eq 0 ]]; then
  echo "check:writing (changed): no markdown files to verify."
  exit 0
fi

echo "check:writing (changed): ${#files[@]} file(s)."

if grep -n -i -E "$WARN_PATTERN" "${files[@]}"; then
  echo "check:writing: aviso: las líneas anteriores describen el estilo de redacción."
  echo "  Si es texto propio, elimina la mención y aplica la guía sin anunciarla."
  echo "  Si cita a un tercero (por ejemplo, una autoridad), puede quedarse."
fi

if grep -n -E "$BLOCK_PATTERN" "${files[@]}"; then
  echo "check:writing: error: las líneas anteriores citan una norma de redacción."
  echo "  Elimina la mención. Ver conductor/code_styleguides/redaccion.md, sección 5."
  exit 1
fi
