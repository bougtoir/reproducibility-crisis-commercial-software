PYTHON ?= .venv/bin/python
export PYTHONPATH := src

.PHONY: setup build check submission-check
setup:
	python3 -m venv .venv
	.venv/bin/python -m pip install --require-hashes -r requirements.lock

build:
	$(PYTHON) -m paper2.build
	$(PYTHON) -m paper2.documents

check:
	$(PYTHON) -m ruff check src tests
	$(PYTHON) -m mypy
	$(PYTHON) -m pytest -q

submission-check:
	$(PYTHON) -m paper2.build --require-submission-ready
