# Пример для проекта, где latex-docx уже стоит в окружении.
# Скопируйте Makefile и wrap.py к своим исходникам.
#
#   pip install "latex-docx[omml,pptx]"
#   make convert SRC=sample.md
#   make diff BASE=out_base.docx EDITED=out_edited.docx

PYTHON ?= python3
WRAP   ?= wrap.py
SRC    ?= sample.md
BASE   ?= out_base.docx
EDITED ?= out_edited.docx
CHANGES ?= out_changes.md
TEMPLATE ?= tdj

SLIDES ?= slides.md
PPTX ?= out_slides.pptx
PPTX_TEMPLATE ?=

.PHONY: help convert prepare-edit diff pdf pptx

help:
	@echo "latex-docx как библиотека (через $(WRAP))"
	@echo "  make convert SRC=sample.md BASE=out_base.docx"
	@echo "  make convert SRC=article.tex TEMPLATE=plain"
	@echo "  make prepare-edit"
	@echo "  make diff"
	@echo "  make pdf                 # Word сохраняет PDF рядом с DOCX"
	@echo "  make pptx SLIDES=slides.md PPTX_TEMPLATE=org.pptx"

convert:
	$(PYTHON) $(WRAP) $(SRC) -o $(BASE) --template $(TEMPLATE)

pdf: convert
	$(PYTHON) $(WRAP) $(BASE) -o $(BASE:.docx=.pdf)

pptx:
	$(PYTHON) $(WRAP) $(SLIDES) -o $(PPTX) $(if $(PPTX_TEMPLATE),--pptx-template "$(PPTX_TEMPLATE)",)

prepare-edit: convert
	cp $(BASE) $(EDITED)
	@echo "Правьте $(EDITED) в Word, затем: make diff"

diff:
	@test -f $(BASE) || (echo "Нет $(BASE). Сначала: make convert"; exit 1)
	@test -f $(EDITED) || (echo "Нет $(EDITED). Сначала: make prepare-edit"; exit 1)
	$(PYTHON) $(WRAP) --diff $(BASE) $(EDITED) -o $(CHANGES)
	@echo "Отчёт: $(CHANGES)"
