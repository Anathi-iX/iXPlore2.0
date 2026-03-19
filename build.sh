#!/usr/bin/env bash
set -euxo pipefail

echo "Installing Python dependencies..."
pip install -r requirements.txt
echo "Build complete."