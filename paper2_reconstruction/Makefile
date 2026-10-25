PYTHON ?= .venv/bin/python
export PYTHONPATH := src

.PHONY: setup build check submission-check fetch-vetted-wheels
fetch-vetted-wheels:
	bash scripts/fetch_vetted_wheels.sh
setup:
	python3 -m venv .venv
	.venv/bin/python -m pip install --require-hashes -r requirements.lock

build:
	$(PYTHON) -m paper2.build
	$(PYTHON) -m paper2.documents
	$(PYTHON) -m paper2.verification_qc

check:
	$(PYTHON) -m ruff check src tests
	$(PYTHON) -m mypy
	$(PYTHON) -m pytest -q

submission-check:
	$(PYTHON) -m paper2.build --require-submission-ready

public-check:
	$(MAKE) build
	$(MAKE) check

private-evidence-check:
	@if [ -d /home/ubuntu/paper2_evidence ]; then \
		echo "evidence package mounted: regenerating private-dependent outputs"; \
		$(PYTHON) -m paper2.verification_layer; \
		$(PYTHON) -m paper2.documents; \
		$(PYTHON) -m paper2.verification_qc; \
	else \
		echo "evidence package NOT mounted: committed bytes are authoritative;"; \
		echo "mount /home/ubuntu/paper2_evidence and rerun to validate private refs."; \
	fi
