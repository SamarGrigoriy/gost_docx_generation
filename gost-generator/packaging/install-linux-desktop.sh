#!/bin/sh
# Установка пункта меню для Linux без прав суперпользователя (ТЗ §14: на Linux
# приложение запускается из исходников, здесь только регистрация в меню).
set -e
HERE=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
APP="$HERE/gost-generator"
BIN_DIR="$HOME/.local/bin"
APPS_DIR="${XDG_DATA_HOME:-$HOME/.local/share}/applications"
ICONS_ROOT="${XDG_DATA_HOME:-$HOME/.local/share}/icons/hicolor"

mkdir -p "$BIN_DIR" "$APPS_DIR" "$ICONS_ROOT/256x256/apps" "$ICONS_ROOT/48x48/apps"
ln -sf "$APP" "$BIN_DIR/gost-generator"
cp "$HERE/gost-generator.desktop" "$APPS_DIR/gost-generator.desktop"
cp "$HERE/../assets/logo.png" "$ICONS_ROOT/256x256/apps/gost-generator.png"
cp "$HERE/../assets/icons/app_64.png" "$ICONS_ROOT/48x48/apps/gost-generator.png" || true
command -v update-desktop-database >/dev/null 2>&1 && update-desktop-database "$APPS_DIR" || true
command -v gtk-update-icon-cache >/dev/null 2>&1 && gtk-update-icon-cache "$ICONS_ROOT" || true

echo "Установлено: $APPS_DIR/gost-generator.desktop"
echo "Запуск: $BIN_DIR/gost-generator (или из меню приложений «ГОСТ-Генератор»)"
echo "Если пункт не появился в меню, завершите сессию или выполните: update-desktop-database $APPS_DIR"
