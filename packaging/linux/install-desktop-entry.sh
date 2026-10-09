#!/bin/sh
# Adds Song Notation Tool to your applications menu (this user only).
# Run it from the unpacked folder:  ./install-desktop-entry.sh
# To remove it again:  rm ~/.local/share/applications/song-notation-tool.desktop
set -e
DIR=$(cd "$(dirname "$0")" && pwd)
APPS="${XDG_DATA_HOME:-$HOME/.local/share}/applications"
mkdir -p "$APPS"
sed -e "s|@DIR@|$DIR|g" "$DIR/song-notation-tool.desktop" > "$APPS/song-notation-tool.desktop"
chmod +x "$APPS/song-notation-tool.desktop"
echo "Added to the applications menu: $APPS/song-notation-tool.desktop"
