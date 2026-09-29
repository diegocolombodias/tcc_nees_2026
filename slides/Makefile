PDF=apresentacao_tcc.pdf
MAIN=apresentacao_tcc

.PHONY: all pptx clean

all: $(PDF)

$(PDF): $(MAIN).tex assets/capa_nees.png assets/rodape_nees.png
	pdflatex -interaction=nonstopmode -halt-on-error $(MAIN).tex
	pdflatex -interaction=nonstopmode -halt-on-error $(MAIN).tex

pptx:
	python3 gerar_apresentacao.py

clean:
	rm -f $(MAIN).aux $(MAIN).log $(MAIN).nav $(MAIN).out \
	      $(MAIN).snm $(MAIN).toc
