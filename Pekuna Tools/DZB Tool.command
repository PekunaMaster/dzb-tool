#!/usr/bin/env zsh
osascript -e 'tell application "Terminal" to set visible of front window to false'
cd -- "$(dirname "$0")"
python3 assets/application/script.py >> assets/application/log.txt