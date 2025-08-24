# Makefile for flatjson_to_xlsx
# Quick start:
#   make examples                 # creates example inputs under ./examples/
#   make run-sample-array         # generates out.xlsx from array example
#   make run-sample-ndjson        # generates out.xlsx from NDJSON example
#   make run-sample-update        # updates the same PK, replacing the row
#   make run-jira                 # uses examples/jira_array.json if you save your blob there
#   make help

SHELL := bash
CARGO ?= cargo
BIN   ?= flatjson_to_xlsx

# Output settings (override: make run-sample-array OUT=wow.xlsx SHEET=Data PK=key)
OUT   ?= out.xlsx
SHEET ?= Sheet1
PK    ?= key

.PHONY: all help build release run run-sample-array run-sample-ndjson run-sample-update \
        install uninstall clean lint fmt fmt-check test examples sample_array sample_ndjson sample_update run-jira

all: build ## Default target: build in debug mode

help: ## Show this help
	@echo "Targets:"
	@grep -E '^[a-zA-Z0-9_.-]+:.*?## ' $(lastword $(MAKEFILE_LIST)) | \
		awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-20s\033[0m %s\n", $$1, $$2}'
	@echo
	@echo "Variables:"
	@echo "  OUT   Excel output file (default: out.xlsx)"
	@echo "  SHEET Sheet name (default: Sheet1)"
	@echo "  PK    Primary key column(s), comma-separated (default: key)"

build: ## Build debug binary
	$(CARGO) build

release: ## Build optimized release binary
	$(CARGO) build --release

run: release ## Run tool; read stdin; pass through CLI (example: make run SHEET=Data PK=key)
	@cat /dev/stdin | target/release/$(BIN) --out "$(OUT)" --sheet "$(SHEET)" --pk "$(PK)"

install: release ## Install to ~/.cargo/bin
	$(CARGO) install --path . --force

uninstall: ## Uninstall from ~/.cargo/bin
	$(CARGO) uninstall $(BIN)

lint: ## Run clippy with warnings as errors
	$(CARGO) clippy -- -D warnings

fmt: ## Format code
	$(CARGO) fmt

fmt-check: ## Check formatting
	$(CARGO) fmt -- --check

test: ## Run tests (if any)
	$(CARGO) test

clean: ## Remove target dir and example outputs
	$(CARGO) clean
	@rm -f "$(OUT)"

# --------------------- Examples ---------------------

examples: sample_array sample_ndjson sample_update ## Create all example input files

