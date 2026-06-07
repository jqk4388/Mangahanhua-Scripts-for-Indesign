#!/bin/bash
# ============================================
# Manga Layout Automation - Shell Launcher for macOS
#
# Features:
# - Launch Adobe InDesign on macOS
# - Execute manga_layout.jsx via AppleScript
# - Write log output to manga_layout_sh.log
#
# Usage:
#   cd src && ./run_manga_layout.sh
# ============================================

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOG_PATH="$SCRIPT_DIR/manga_layout_sh.log"
JSX_PATH="$SCRIPT_DIR/manga_layout.jsx"
CONFIG_PATH="$SCRIPT_DIR/manga_layout_config.json"

IDS_APP_PATHS=(
  "/Applications/Adobe InDesign 2026/Adobe InDesign 2026.app"
  "/Applications/Adobe InDesign 2025/Adobe InDesign 2025.app"
  "/Applications/Adobe InDesign 2024/Adobe InDesign 2024.app"
  "/Applications/Adobe InDesign 2023/Adobe InDesign 2023.app"
  "/Applications/Adobe InDesign 2022/Adobe InDesign 2022.app"
  "/Applications/Adobe InDesign 2021/Adobe InDesign 2021.app"
  "/Applications/Adobe InDesign/Adobe InDesign.app"
)

log() {
  local msg="$1"
  local timestamp
  timestamp="$(date "+%Y-%m-%d %H:%M:%S")"
  printf "%s - %s\n" "$timestamp" "$msg" | tee -a "$LOG_PATH"
}

error_exit() {
  log "ERROR: $1"
  exit 1
}

find_indesign_app() {
  local app_path
  for app_path in "${IDS_APP_PATHS[@]}"; do
    if [ -d "$app_path" ]; then
      printf "%s" "$app_path"
      return 0
    fi
  done
  return 1
}

run_jsx() {
  local app_name="$1"
  osascript <<EOF
tell application "$app_name"
  activate
  do script POSIX file "$JSX_PATH" language javascript
end tell
EOF
}

# Ensure osascript exists
if ! command -v osascript >/dev/null 2>&1; then
  error_exit "osascript is required but not found. Please run on macOS with AppleScript support."
fi

log "===== Manga Layout Shell Launcher Started ====="
log "Script dir: $SCRIPT_DIR"
log "Start time: $(date '+%Y-%m-%d %H:%M:%S')"

if [ ! -f "$JSX_PATH" ]; then
  error_exit "Script file does not exist: $JSX_PATH"
fi

log "Script file: $JSX_PATH"

if [ -f "$CONFIG_PATH" ]; then
  log "Config file: $CONFIG_PATH"
else
  log "Warning: Config file does not exist: $CONFIG_PATH"
  log "Will run with default configuration"
fi

INDESIGN_APP_PATH="$(find_indesign_app || true)"
if [ -z "$INDESIGN_APP_PATH" ]; then
  error_exit "Adobe InDesign application not found in /Applications. Please install InDesign or update the script with the correct path."
fi

INDESIGN_APP_NAME="$(basename "$INDESIGN_APP_PATH" .app)"
log "Found InDesign app: $INDESIGN_APP_PATH"
log "Using AppleScript target: $INDESIGN_APP_NAME"

log "Starting JSX script execution..."
if run_jsx "$INDESIGN_APP_NAME"; then
  log "Script execution completed"
  log "===== Execution successful ====="
  exit 0
else
  error_exit "Script execution failed. Check InDesign and log output."
fi
