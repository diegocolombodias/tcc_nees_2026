# TCC NEES 2026

Projeto Integrado da Especialização em Gestão de Políticas Públicas Educacionais e Transformação Digital da Universidade Federal de Alagoas (UFAL).

**Título:** Inteligência Artificial de Baixo Consumo para o Monitoramento Preventivo do Bem-Estar Estudantil em Escolas Públicas de Baixa Conectividade

**Autores:** Ivo Calado, Rafael Durelli e Diego Dias

## Estrutura

- `texto/`: fontes LaTeX, bibliografia, figura e PDF final do TCC;
- `slides/`: apresentação em LaTeX/Beamer e PowerPoint, roteiro, assets e PDFs gerados.

## Compilação

Para compilar o texto e os slides:

```bash
make
```

Também é possível compilar cada parte separadamente:

```bash
make texto
make slides
```

O texto final fica em `texto/tcc.pdf`; os slides em `slides/apresentacao_tcc.pdf`.

Dependências principais: TeX Live, `abntex2`, `abntex2cite`, TikZ, Beamer e BibTeX. Para regenerar o PowerPoint, também são necessários Python 3 e `python-pptx`.
