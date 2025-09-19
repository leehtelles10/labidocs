# LabiDocs — Documentador Power BI

[![CI](https://github.com/leehtelles10/labidocs/actions/workflows/ci.yml/badge.svg)](https://github.com/leehtelles10/labidocs/actions)
[![codecov](https://codecov.io/gh/leehtelles10/labidocs/branch/main/graph/badge.svg?token=)](https://codecov.io/gh/leehtelles10/labidocs)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](./LICENSE)

## Descrição
**LabiDocs** é uma ferramenta Streamlit para gerar documentação automatizada de modelos Power BI exportados (PBIX → ZIP contendo `.tmdl`). A aplicação percorre os arquivos `.tmdl`, extrai metadados (tabelas, colunas, medidas, expressões, relacionamentos) e gera um documento Word/PDF profissional.

Funcionalidades principais:
- Extração de tabelas, colunas e colunas calculadas;
- Extração de medidas DAX e expressões;
- Geração de documento `.docx` formatado (com capa, índice e seções);
- Conversão para PDF (via **docx2pdf** em Windows);
- UI Streamlit intuitiva com upload de ZIP e logo.

## Pré-requisitos
- Python 3.10+  
- (Windows) Microsoft Word instalado (necessário para conversão `.docx` → `.pdf`).

## Instalação local
```bash
git clone https://github.com/<seu-usuario>/labidocs.git
cd labidocs
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
