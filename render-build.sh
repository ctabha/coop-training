#!/usr/bin/env bash
set -o errexit

apt-get update
apt-get install -y libreoffice

pip install -r requirements.txt
