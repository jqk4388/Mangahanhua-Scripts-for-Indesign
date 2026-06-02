#!/usr/bin/env bash
# run.sh
# Usage:
#   ./run.sh                    # run scriptRun.jsx in the same folder as run.sh
#   ./run.sh /path/to/script.jsx # run an absolute path
#   ./run.sh myscript.jsx        # run a relative path from current working directory

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DEFAULT_SCRIPT="$SCRIPT_DIR/scriptRun.jsx"

resolve_script_path() {
  local input_path="$1"
  if [[ "$input_path" = /* ]]; then
    printf '%s' "$input_path"
  else
    printf '%s' "$PWD/$input_path"
  fi
}

find_indesign_app() {
  local app_paths=(
    "/Applications/Adobe InDesign 2026/Adobe InDesign 2026.app"
    "/Applications/Adobe InDesign 2025/Adobe InDesign 2025.app"
    "/Applications/Adobe InDesign 2024/Adobe InDesign 2024.app"
    "/Applications/Adobe InDesign 2023/Adobe InDesign 2023.app"
    "/Applications/Adobe InDesign 2022/Adobe InDesign 2022.app"
    "/Applications/Adobe InDesign 2021/Adobe InDesign 2021.app"
    "/Applications/Adobe InDesign/Adobe InDesign.app"
  )

  for app_path in "${app_paths[@]}"; do
    if [ -d "$app_path" ]; then
      printf '%s' "$app_path"
      return 0
    fi
  done
  return 1
}

error_exit() {
  echo "ERROR: $1" >&2
  exit 1
}

if ! command -v osascript >/dev/null 2>&1; then
  error_exit "osascript is required but not found. 请在 macOS 上运行。"
fi

if [ "$#" -gt 0 ]; then
  SCRIPT_PATH="$(resolve_script_path "$1")"
else
  SCRIPT_PATH="$DEFAULT_SCRIPT"
fi

if [ ! -f "$SCRIPT_PATH" ]; then
  error_exit "找不到脚本文件: $SCRIPT_PATH"
fi

INDESIGN_APP_PATH="$(find_indesign_app || true)"
if [ -z "$INDESIGN_APP_PATH" ]; then
  error_exit "未找到 Adobe InDesign，请安装 InDesign 或更新脚本中的应用路径。"
fi

INDESIGN_APP_NAME="$(basename "$INDESIGN_APP_PATH" .app)"

echo "运行脚本: $SCRIPT_PATH"
echo "使用 InDesign: $INDESIGN_APP_NAME"

osascript <<EOF
tell application "$INDESIGN_APP_NAME"
  activate
  do script POSIX file "$SCRIPT_PATH" language javascript
end tell
EOF
