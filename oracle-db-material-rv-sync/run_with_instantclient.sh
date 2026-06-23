#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
CLIENT_DIR="${ORACLE_CLIENT_CONFIG_FILE:-${ORACLE_CLIENT_LIB_DIR:-$SCRIPT_DIR/instantclient/current}}"

if [[ -n "${PYTHON_BIN:-}" ]]; then
  PYTHON_CMD="$PYTHON_BIN"
elif [[ -x "$REPO_DIR/.venv/bin/python" ]]; then
  PYTHON_CMD="$REPO_DIR/.venv/bin/python"
else
  PYTHON_CMD="python3"
fi

export ORACLE_CLIENT_CONFIG_FILE="$CLIENT_DIR"
export ORACLE_CLIENT_LIB_DIR="$CLIENT_DIR"
export LD_LIBRARY_PATH="$CLIENT_DIR${LD_LIBRARY_PATH:+:$LD_LIBRARY_PATH}"

cd "$SCRIPT_DIR"
exec "$PYTHON_CMD" oracle_to_db_material_rv.py "$@"
