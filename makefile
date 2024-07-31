EMPLOYEE_PDFS_DIR = ./employee-pdfs
WORKDAY_DOCX_DIR = ./workday-reports
PAR_TEMPLATE_DIR = ./par-template
FILLED_REPORTS_DIR = ./filled-reports
SCRIPT_DIR = ./script

REQUIREMENTS_FILE = requirements.txt

SCRIPT = $(SCRIPT_DIR)/fill_out_PARs.py

.PHONY: all install_dependencies check_prerequisites run clean

# default target
all: install_dependencies check_prerequisites run

# installs the required Python packages
install_dependencies:
	@echo "Installing dependencies..."
	pip install -r $(REQUIREMENTS_FILE)
	@echo "Verifying installed packages..."
	python -c "import docx; print('python-docx is installed')"
	python -c "import pdf2docx; print('pdf2docx is installed')"

# check the required directories and files
check_prerequisites:
	@if [ ! -d $(EMPLOYEE_PDFS_DIR) ]; then echo "Directory $(EMPLOYEE_PDFS_DIR) does not exist. Please create it and add workday time-tracking PDFs."; exit 1; fi
	@if [ ! "$(shell ls -A $(EMPLOYEE_PDFS_DIR))" ]; then echo "Directory $(EMPLOYEE_PDFS_DIR) is empty. Please add workday time-tracking PDFs."; exit 1; fi
	@if [ ! -d $(PAR_TEMPLATE_DIR) ]; then echo "Directory $(PAR_TEMPLATE_DIR) does not exist. Please create it and add AU's PAR template as PAR-template.docx."; exit 1; fi
	@if [ ! -f $(PAR_TEMPLATE_DIR)/PAR-template.docx ]; then echo "PAR template file not found in $(PAR_TEMPLATE_DIR). Please add PAR-template.docx."; exit 1; fi
	@mkdir -p $(WORKDAY_DOCX_DIR) $(FILLED_REPORTS_DIR)

# run main script
run:
	@echo "Running the main script..."
	python $(SCRIPT)
	@echo "Cleaning up employee pdfs..."
	find $(FILLED_REPORTS_DIR) -type f -name "*.docx" -exec rm -f {} +

# clean up generated files
clean:
	@echo "Cleaning up generated PARs..."
	rm -f $(FILLED_REPORTS_DIR)/*.docx