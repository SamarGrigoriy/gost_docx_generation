PYTHON ?= python3

.PHONY: run test lint build-win-local desktop clean

run:
	$(PYTHON) -m app

test:
	$(PYTHON) -m unittest discover -s tests -v

lint:
	$(PYTHON) -m compileall -q app scripts tests

# exe собирается только под Windows (ограничение PyInstaller); цель — для Wine/WSL-экспериментов
build-win-local:
	pyinstaller build.spec --noconfirm

# Пункт меню в Linux (без sudo): ~/.local/bin + ~/.local/share/applications
desktop:
	sh packaging/install-linux-desktop.sh

clean:
	rm -rf build dist .pycache __pycache__ app/__pycache__ app/core/__pycache__
	find . -name '__pycache__' -type d -prune -exec rm -rf {} +
	find . -name '*.orig.docx' -delete
