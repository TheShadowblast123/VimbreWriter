TESTING_ODT   ?= docs/testing.odt
TESTING_OXT   ?= build/testing/vimbrewriter-0.0.0.oxt
TESTING_XBA   ?= build/testing/vimbrewriter.xba
VERSION       ?= 0.0.0

# Extension identifier – must match the one in description.xml
EXT_ID        ?= vimbrewriter

# Output directory for built extensions
DIST_DIR      ?= dist
OXT_FILE      ?= $(DIST_DIR)/vimbrewriter-$(VIMBREWRITER_VERSION).oxt

# Check for unopkg
UNOPKG := $(shell command -v unopkg 2>/dev/null)

# LibreOffice user Basic directory (not used directly, kept for reference)
UNAME_S := $(shell uname -s)
ifeq ($(UNAME_S),Linux)
    BASIC_DIR = $(HOME)/.config/libreoffice/4/user/basic
else ifeq ($(UNAME_S),Darwin)
    BASIC_DIR = $(HOME)/Library/Application Support/LibreOffice/4/user/basic
endif

.PHONY: testing clean install uninstall lint

# ----------------------------------------------------------------------
# Testing – builds, installs, and launches LibreOffice
# ----------------------------------------------------------------------
testing: $(TESTING_OXT)
	@if [ -z "$(UNOPKG)" ]; then \
		echo "Error: unopkg not found in PATH"; \
		exit 1; \
	fi
	@echo "Stopping LibreOffice..."
	-killall soffice.bin 2>/dev/null || true
	@echo "Removing any previous version of the extension..."
	-$(MAKE) uninstall $(EXT_ID) 2>/dev/null || true
	@echo "Installing extension $(TESTING_OXT)..."
	$(UNOPKG) add "$(TESTING_OXT)"
	@echo "Starting LibreOffice with $(TESTING_ODT)..."
	lowriter "$(TESTING_ODT)" --norestore
	# uninstall
	-$(MAKE) uninstall $(EXT_ID)
	@echo "Extension uninstalled successfully."
	

$(TESTING_OXT): src/vimbrewriter.vbs
	# Build extension package with version $(VERSION)
	mkdir -p build/template
	mkdir -p build/testing
	cp -r extension/template/. build/template/
	./compile.sh "src/vimbrewriter.vbs" "build/template/vimbrewriter/vimbrewriter.xba"
	cd build/template ; \
	  sed -i 's/%VIMBREWRITER_VERSION%/$(VERSION)/g' description.xml ; \
	zip -r "../../$(TESTING_OXT)" .
	# cd build/template && zip -r "../../$(TESTING_OXT)" .
	# Also copy the raw XBA for potential fallback use
	mkdir -p build/testing
	cp build/template/vimbrewriter/vimbrewriter.xba "$(TESTING_XBA)"

# ----------------------------------------------------------------------
# Build release extension (requires VIMBREWRITER_VERSION to be set)
# ----------------------------------------------------------------------
extension: clean src/vimbrewriter.vbs
	if [ -z "$$VIMBREWRITER_VERSION" ]; then \
		echo "VIMBREWRITER_VERSION must be set"; \
	else \
		mkdir -p build/template; mkdir -p $(DIST_DIR); \
		cp -r extension/template/. build/template; \
		./compile.sh "src/vimbrewriter.vbs" "build/template/vimbrewriter/vimbrewriter.xba"; \
		cd "build/template"; \
		sed -i "s/%VIMBREWRITER_VERSION%/$$VIMBREWRITER_VERSION/g" description.xml; \
		zip -r "../../$(DIST_DIR)/vimbrewriter-$$VIMBREWRITER_VERSION.oxt" .; \
	fi

# ----------------------------------------------------------------------
# Install the currently built extension (uses $(OXT_FILE))
# ----------------------------------------------------------------------
install: $(OXT_FILE)
	@if [ -z "$(UNOPKG)" ]; then \
		echo "Error: unopkg not found in PATH"; \
		exit 1; \
	fi
	@echo "Make sure LibreOffice is closed before installing."
	@echo "Installing $(OXT_FILE)..."
	$(UNOPKG) add "$(OXT_FILE)"
	@echo "Extension installed successfully."

$(OXT_FILE): extension
	@if [ ! -f "$(OXT_FILE)" ]; then \
		echo "Extension file $(OXT_FILE) not found. Did you set VIMBREWRITER_VERSION?"; \
		exit 1; \
	fi

# ----------------------------------------------------------------------
# Uninstall the extension
# ----------------------------------------------------------------------
uninstall:
	@if [ -z "$(UNOPKG)" ]; then \
		echo "Error: unopkg not found in PATH"; \
		exit 1; \
	fi
	@echo "Make sure LibreOffice is closed before uninstalling."
	$(UNOPKG) remove $(EXT_ID)

# ----------------------------------------------------------------------
# Clean build artifacts
# ----------------------------------------------------------------------
clean:
	rm -rf build

lint:
	./format-code.sh
