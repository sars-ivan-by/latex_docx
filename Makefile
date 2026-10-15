PYTHON ?= python3
TEX    ?= example.tex
BASE   ?= out_base.docx
EDITED ?= out_edited.docx
CHANGES ?= out_changes.md
LOG    ?= out_latex_docx.log
BY     ?= user

MD     ?= notes.md

.PHONY: help install test convert convert-md prepare-edit diff pdf

help:
	@echo "latex-docx"
	@echo "  make install             # pip install -e .[omml,dev]"
	@echo "  make test                # pytest"
	@echo "  make convert TEX=a.tex   # → $(BASE)"
	@echo "  make convert-md MD=a.md  # light Markdown → $(BASE)"
	@echo "  make pdf                 # DOCX → PDF через Microsoft Word"
	@echo "  make prepare-edit        # convert + copy → $(EDITED)"
	@echo "  make diff                # $(BASE) vs $(EDITED) → $(CHANGES)"

install:
	$(PYTHON) -m pip install -e ".[omml,dev]"

test:
	$(PYTHON) -m pytest

convert:
	$(PYTHON) -m latex_docx $(TEX) -o $(BASE) --by $(BY) --log $(LOG)

convert-md:
	$(PYTHON) -m latex_docx $(MD) -o $(BASE) --by $(BY) --log $(LOG)

pdf:
	@test -f $(BASE) || (echo "Missing $(BASE). Run: make convert"; exit 1)
	$(PYTHON) -m latex_docx $(BASE) -o $(BASE:.docx=.pdf) --by $(BY) --log $(LOG)

prepare-edit: convert
	cp $(BASE) $(EDITED)
	@echo "Edit $(EDITED) in Word, then: make diff"

diff:
	@test -f $(BASE) || (echo "Missing $(BASE). Run: make convert"; exit 1)
	@test -f $(EDITED) || (echo "Missing $(EDITED). Run: make prepare-edit"; exit 1)
	$(PYTHON) -m latex_docx --diff-docx $(BASE) $(EDITED) -o $(CHANGES) --by $(BY) --log $(LOG)
	@echo "Report: $(CHANGES)"
