#!/usr/bin/env bash
# ==============================================================================
# Break-Even Point Predictor - Server Launcher
# Auto-detects runtime (Ruby, Python 3, or Node.js) and starts the backend.
# ==============================================================================

PORT=${PORT:-8000}
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$DIR"

echo "======================================================="
echo " Starting Break-Even Point Predictor Backend Server..."
echo "======================================================="

# 1. Check Ruby (standard built-in on macOS with sqlite3 & webrick)
if command -v ruby >/dev/null 2>&1; then
  echo "Found Ruby runtime: $(ruby -v)"
  echo "Launching Ruby SQLite3 server on port $PORT..."
  exec ruby server.rb
fi

# 2. Check Python 3
if command -v python3 >/dev/null 2>&1; then
  echo "Found Python runtime: $(python3 --version 2>&1)"
  echo "Launching Python SQLite3 server on port $PORT..."
  exec python3 server.py
fi

# 3. Check Node.js
if command -v node >/dev/null 2>&1; then
  echo "Found Node.js runtime: $(node -v)"
  echo "Launching Node.js server on port $PORT..."
  exec node server.js
fi

echo "Error: No suitable runtime found (Ruby, Python 3, or Node.js required)."
echo "Please install one of them or open index.html directly in your web browser."
exit 1
