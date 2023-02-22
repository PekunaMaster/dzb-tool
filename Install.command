#!/usr/bin/env zsh
cd -- "$(dirname "$0")"
chmod +x "Pekuna Tools/DZB Tool.command"
python3 -m pip install --upgrade pip
python3 -m pip install --upgrade setuptools
python3 -m pip install --upgrade wheel
python3 -m pip install xlwings
python3 -m pip install pyobjc
