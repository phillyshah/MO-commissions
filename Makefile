# MO-commissions repo — operations targets.
#
# The Flask generator (app.py) has no make targets; it is started by Dockerfile.commissions inside the
# accounting compose stack, or bare-metal by commission-app.service (being retired).
#
# Everything below is for the Maxx Accounting Portal (accounting/). Run from the repo root on the VPS:
#   make accounting-deploy
#
# JOB for accounting-worker-once is one of the worker's job names, e.g. heartbeat, imap, pipeline,
# notify, reminders, anchors, verify, sessions, digest, controller_daily (see accounting/worker.py).

COMPOSE := docker compose -f docker-compose.accounting.yml

.PHONY: help accounting-up accounting-deploy accounting-build accounting-migrate accounting-migrate-status \
        accounting-logs accounting-ps accounting-down accounting-restart accounting-test \
        accounting-worker-once accounting-shell accounting-config

help:
	@echo "Accounting portal targets (run from the repo root):"
	@echo "  make accounting-up              - start the stack with the images already built"
	@echo "  make accounting-deploy          - git pull origin main, rebuild and restart (the normal deploy)"
	@echo "  make accounting-build           - rebuild images without starting them"
	@echo "  make accounting-migrate         - apply pending SQL migrations to the portal database"
	@echo "  make accounting-migrate-status  - list migrations not yet applied"
	@echo "  make accounting-logs            - follow logs for every service (SERVICE=accounting-worker for one)"
	@echo "  make accounting-ps              - container status"
	@echo "  make accounting-restart         - restart web and worker without rebuilding"
	@echo "  make accounting-down            - stop the stack (volumes are kept)"
	@echo "  make accounting-test            - run the portal test suite locally (ACCOUNTING_ENV=test)"
	@echo "  make accounting-worker-once JOB=heartbeat - run one worker job once, in the foreground"
	@echo "  make accounting-shell           - a Python shell inside accounting-web with settings loaded"
	@echo "  make accounting-config          - print the resolved compose file (validates the YAML)"

accounting-up:
	$(COMPOSE) up -d

accounting-deploy:
	git pull origin main
	$(COMPOSE) up -d --build

accounting-build:
	$(COMPOSE) build

accounting-migrate:
	$(COMPOSE) exec accounting-web python -m accounting.db.migrate

accounting-migrate-status:
	$(COMPOSE) exec accounting-web python -m accounting.db.migrate --status

SERVICE ?=
accounting-logs:
	$(COMPOSE) logs -f --tail=200 $(SERVICE)

accounting-ps:
	$(COMPOSE) ps

accounting-restart:
	$(COMPOSE) restart accounting-web accounting-worker

accounting-down:
	$(COMPOSE) down

accounting-test:
	ACCOUNTING_ENV=test python -m pytest tests/accounting -q

JOB ?= heartbeat
accounting-worker-once:
	$(COMPOSE) run --rm --no-deps accounting-worker python -m accounting.worker --once $(JOB)

accounting-shell:
	$(COMPOSE) exec accounting-web python -i -c "from accounting.config import settings; print('settings loaded; try settings.portal_host')"

accounting-config:
	$(COMPOSE) config
