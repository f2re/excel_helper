#!/usr/bin/env bash
set -euo pipefail
mkdir -p .certs
if [[ ! -f .certs/key.pem || ! -f .certs/cert.pem ]]; then
  openssl req -x509 -newkey rsa:2048 -nodes -sha256 -days 825 -keyout .certs/key.pem -out .certs/cert.pem -subj "/CN=localhost" -addext "subjectAltName=DNS:localhost,IP:127.0.0.1"
fi
echo "Certificates are in .certs/. Trust .certs/cert.pem before sideloading."
