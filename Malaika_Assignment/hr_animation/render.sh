#!/bin/bash
# render.sh — One-command render script for HRM Dialogue Animation
#
# Usage:
#   bash render.sh          # High quality 1080p60 (production)
#   bash render.sh low      # Low quality 480p15 (fast preview)
#   bash render.sh medium   # Medium quality 720p30

set -e

# Auto-activate venv if present
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
if [ -f "$SCRIPT_DIR/.venv/bin/activate" ]; then
    source "$SCRIPT_DIR/.venv/bin/activate"
    echo "  Using virtualenv: $SCRIPT_DIR/.venv"
fi

SCENE="HRMDialogueScene"
SCRIPT="main.py"

QUALITY=${1:-high}

case "$QUALITY" in
  high)
    QUALITY_FLAG="-qh"
    LABEL="1080p60"
    ;;
  medium)
    QUALITY_FLAG="-qm"
    LABEL="720p30"
    ;;
  low)
    QUALITY_FLAG="-ql"
    LABEL="480p15 (preview)"
    ;;
  *)
    echo "Unknown quality '$QUALITY'. Use: high | medium | low"
    exit 1
    ;;
esac

echo "=============================================="
echo "  HRM Dialogue Animation — Render Script"
echo "=============================================="
echo "  Scene  : $SCENE"
echo "  Quality: $LABEL"
echo ""
echo "  NOTE: First render downloads TTS audio via gTTS."
echo "        Subsequent renders use cached audio."
echo ""

manim -p $QUALITY_FLAG $SCRIPT $SCENE

echo ""
echo "=============================================="
echo "  Done! Output is in media/videos/"
echo "=============================================="
