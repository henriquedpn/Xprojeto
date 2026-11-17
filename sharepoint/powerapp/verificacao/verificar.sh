#!/usr/bin/env bash
# Verifica os tipos de todas as fórmulas do app com o motor oficial do Power Fx.
# Requisito: .NET SDK 8+.   Uso: ./verificacao/verificar.sh
set -euo pipefail
AQUI="$(cd "$(dirname "$0")" && pwd)"
TMP="$(mktemp -d)"
python3 "$AQUI/exportar_formulas.py" "$AQUI/.." "$TMP/formulas.json"
dotnet run --project "$AQUI/VerificarFormulas.csproj" -- "$TMP/formulas.json"
