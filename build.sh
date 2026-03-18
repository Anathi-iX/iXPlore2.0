#!/usr/bin/env bash

echo "Installing system dependencies..."

apt-get update
apt-get install -y tesseract-ocr libtesseract-dev libleptonica-dev

echo "Installing Python dependencies..."

pip install -r requirements.txt

echo "Build complete."