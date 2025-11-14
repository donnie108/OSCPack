#!/usr/bin/env bash
set -euo pipefail

########################################
# Config
########################################

APP_NAME="OSCPack"     # macOS .app name
ENTRYPOINT="gui.py"    # main GUI script
VENV_DIR="venv"        # virtual environment folder

echo "========================================"
echo "  Building $APP_NAME for macOS"
echo "========================================"

########################################
# Discover version from core.APP_VERSION
########################################

VERSION="dev"

if python3 - << 'EOF' >/dev/null 2>&1
import core
EOF
then
  VERSION=$(python3 - << 'EOF'
import core
print(getattr(core, "APP_VERSION", "dev"))
EOF
  )
else
  echo "Warning: could not import core.py to get APP_VERSION. Using 'dev'."
fi

echo "App version: $VERSION"

########################################
# Ensure venv exists
########################################

if [[ ! -d "$VENV_DIR" ]]; then
  echo "Creating virtual environment in: $VENV_DIR"
  python3 -m venv "$VENV_DIR"
fi

# Activate venv
source "$VENV_DIR/bin/activate"

########################################
# Install dependencies
########################################

echo "Upgrading pip..."
pip install --upgrade pip

if [[ -f "requirements.txt" ]]; then
  echo "Installing dependencies from requirements.txt..."
  pip install -r requirements.txt
else
  echo "requirements.txt not found, installing core deps manually..."
  pip install pypdf Pillow docx2pdf reportlab beautifulsoup4 cryptography
fi

echo "Installing PyInstaller..."
pip install pyinstaller

# Sanity check: cryptography must be importable
python -c "import cryptography; print('cryptography OK:', cryptography.__version__)"

########################################
# Run PyInstaller
########################################

echo "Cleaning old build/dist..."
rm -rf build dist "__pycache__"

echo "Running PyInstaller..."
pyinstaller \
  --noconfirm \
  --clean \
  --name "$APP_NAME" \
  --windowed \
  --hidden-import cryptography \
  --hidden-import cryptography.hazmat \
  "$ENTRYPOINT"

########################################
# Package .app into a zip
########################################

if [[ ! -d "dist/$APP_NAME.app" ]]; then
  echo "ERROR: dist/$APP_NAME.app not found. Build may have failed."
  exit 1
fi

cd dist

ZIP_NAME="${APP_NAME}-macOS-${VERSION}.zip"
echo "Creating zip: $ZIP_NAME"
rm -f "$ZIP_NAME"
zip -r "$ZIP_NAME" "$APP_NAME.app"

cd ..

echo "========================================"
echo "Build complete."
echo "App bundle : dist/$APP_NAME.app"
echo "Zip file   : dist/$ZIP_NAME"
echo "========================================"
echo "You can now distribute dist/$ZIP_NAME to users."
