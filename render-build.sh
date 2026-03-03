#!/usr/bin/env bash
set -o errexit

pip install --upgrade pip
pip install -r requirements.txt

# Install LibreOffice for DOCX->PDF conversion
apt-get update
apt-get install -y libreoffice
