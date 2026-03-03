#!/usr/bin/env bash
set -e

apt-get update
apt-get install -y libreoffice

pip install -r requirements.txt
