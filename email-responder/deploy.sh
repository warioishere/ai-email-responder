#!/bin/bash
# Deploy Email-Assistent to remote server
# Usage: ./deploy.sh user@server

set -e

if [ -z "$1" ]; then
    echo "Usage: ./deploy.sh user@server"
    echo "Example: ./deploy.sh mario@matrixserver"
    exit 1
fi

SERVER="$1"
REMOTE_DIR="/opt/email-assistant"

echo "=== Deploying Email-Assistent to $SERVER ==="

# 1. Push latest code to git
echo "[1/5] Pushing latest code..."
git add -A
git diff --cached --quiet || git commit -m "Deploy: latest changes"
git push origin main

# 2. Setup on remote server
echo "[2/5] Setting up on server..."
ssh "$SERVER" bash -s "$REMOTE_DIR" << 'REMOTE_SETUP'
REMOTE_DIR="$1"
set -e

# Install system dependencies
sudo apt-get update -qq
sudo apt-get install -y -qq python3 python3-venv python3-pip git > /dev/null 2>&1

# Clone or pull repo
if [ -d "$REMOTE_DIR" ]; then
    echo "Updating existing installation..."
    cd "$REMOTE_DIR"
    git pull
else
    echo "Cloning repository..."
    sudo git clone https://github.com/warioishere/ai-email-responder.git "$REMOTE_DIR"
    sudo chown -R $(whoami):$(whoami) "$REMOTE_DIR"
fi

# Create venv and install dependencies
cd "$REMOTE_DIR"
if [ ! -d ".venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv .venv
fi

echo "Installing dependencies..."
.venv/bin/pip install -q anthropic pyyaml httpx caldav icalendar

# Create memory directories
mkdir -p memory/contacts memory/categories

echo "Remote setup done."
REMOTE_SETUP

# 3. Copy config.yaml (contains credentials - not in git)
echo "[3/5] Copying config.yaml..."
scp config.yaml "$SERVER:$REMOTE_DIR/config.yaml"

# 4. Install systemd service
echo "[4/5] Installing systemd service..."
ssh "$SERVER" bash -s "$REMOTE_DIR" << 'REMOTE_SERVICE'
REMOTE_DIR="$1"
USER=$(whoami)

sudo tee /etc/systemd/system/email-assistant.service > /dev/null << EOF
[Unit]
Description=Email Assistant Service
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
User=$USER
WorkingDirectory=$REMOTE_DIR
Environment=PYTHONUNBUFFERED=1
ExecStart=$REMOTE_DIR/.venv/bin/python3 -u email-assistant_anthropic.py
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
EOF

sudo systemctl daemon-reload
sudo systemctl enable email-assistant
sudo systemctl restart email-assistant

echo "Service installed and started."
REMOTE_SERVICE

# 5. Verify
echo "[5/5] Verifying..."
ssh "$SERVER" "sudo systemctl status email-assistant --no-pager | head -10"

echo ""
echo "=== Deploy complete ==="
echo "Logs: ssh $SERVER 'sudo journalctl -fu email-assistant'"
