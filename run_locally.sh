#!/usr/bin/env bash
# =============================================================================
# run_locally.sh — Create the Restaurant Tracker spreadsheet locally
#
# Prerequisites:
#   pip install google-api-python-client google-auth google-auth-httplib2
#
# Set credentials via environment variables before running:
#   export GOOGLE_CLIENT_ID="your_client_id"
#   export GOOGLE_CLIENT_SECRET="your_client_secret"
#   export GOOGLE_REFRESH_TOKEN="your_refresh_token"
#
# Or point to a credentials JSON file:
#   export GOOGLE_CREDENTIALS_FILE="/path/to/credentials.json"
#   (JSON must have keys: client_id, client_secret, refresh_token)
#
# Usage:
#   chmod +x run_locally.sh
#   ./run_locally.sh
# =============================================================================
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "============================================================"
echo " Restaurant Monthly Budget & Operations Tracker — Local Run"
echo "============================================================"
echo ""

# Validate credentials are set
if [ -z "${GOOGLE_CREDENTIALS_FILE:-}" ] && \
   { [ -z "${GOOGLE_CLIENT_ID:-}" ] || [ -z "${GOOGLE_CLIENT_SECRET:-}" ] || [ -z "${GOOGLE_REFRESH_TOKEN:-}" ]; }; then
    echo "ERROR: No credentials set."
    echo ""
    echo "Either set:"
    echo "  export GOOGLE_CREDENTIALS_FILE=/path/to/credentials.json"
    echo "Or all three of:"
    echo "  export GOOGLE_CLIENT_ID=..."
    echo "  export GOOGLE_CLIENT_SECRET=..."
    echo "  export GOOGLE_REFRESH_TOKEN=..."
    exit 1
fi

python3 "$SCRIPT_DIR/create_tracker_python.py"
