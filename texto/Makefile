PDF=tcc.pdf
MAIN=main

.PHONY: all clean

all: $(PDF)

$(PDF): main.tex conteudo_artigo.tex referencias.bib figuras/comparacao_alternativas_original.png
	pdflatex -interaction=nonstopmode -halt-on-error $(MAIN).tex
	bibtex $(MAIN)
	pdflatex -interaction=nonstopmode -halt-on-error $(MAIN).tex
	pdflatex -interaction=nonstopmode -halt-on-error $(MAIN).tex
	cp $(MAIN).pdf $(PDF)

clean:
	rm -f $(MAIN).aux $(MAIN).bbl $(MAIN).blg $(MAIN).fdb_latexmk \
	      $(MAIN).fls $(MAIN).log $(MAIN).out $(MAIN).toc $(MAIN).pdf
