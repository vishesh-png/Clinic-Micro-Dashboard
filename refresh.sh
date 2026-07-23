#!/bin/zsh
set -e
cd "$(dirname "$0")"
python3 build_data.py "${1:-source/Micro_Demand.xlsx}"
python3 build_micro.py
echo "Done."
