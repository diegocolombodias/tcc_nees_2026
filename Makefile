.PHONY: all texto slides clean

all: texto slides

texto:
	$(MAKE) -C texto

slides:
	$(MAKE) -C slides

clean:
	$(MAKE) -C texto clean
	$(MAKE) -C slides clean
