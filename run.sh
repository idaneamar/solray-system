#!/usr/bin/env bash
# הרץ את מערכת סול-ריי לסלולר
# Usage: bash run.sh
set -e
cd "$(dirname "$0")"

if ! python3 -c "import flask, docx" 2>/dev/null; then
  echo "מתקין חבילות..."
  pip3 install -r requirements.txt
fi

echo "מפעיל שרת בפורט 5051..."
echo "כתובת גישה מהרשת המקומית:"
python3 -c "import socket; s=socket.socket(socket.AF_INET,socket.SOCK_DGRAM); s.connect(('8.8.8.8',80)); print('  http://'+s.getsockname()[0]+':5051'); s.close()" 2>/dev/null || true
python3 main.py
