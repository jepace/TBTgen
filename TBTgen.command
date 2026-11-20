#!/bin/bash
# Double-click this in Finder to start TBTgen. A Terminal window opens and your browser shows the app.
# Leave the Terminal window open while you use it; close it (or click Quit in the app) to stop.
cd "$(dirname "$0")" || exit 1
for py in /opt/homebrew/bin/python3 /usr/local/bin/python3 python3; do
  if command -v "$py" >/dev/null 2>&1; then exec "$py" tbtgen_app.py "$@"; fi
done
echo "Python 3 not found. Install it with: brew install python"; read -r -p "Press Return to close."
