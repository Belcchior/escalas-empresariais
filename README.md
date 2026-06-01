# Escalas Empresariais

[![Python check](https://github.com/Belcchior/escalas-empresariais/actions/workflows/python-check.yml/badge.svg)](https://github.com/Belcchior/escalas-empresariais/actions/workflows/python-check.yml)
![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![License](https://img.shields.io/github/license/Belcchior/escalas-empresariais)
![Last commit](https://img.shields.io/github/last-commit/Belcchior/escalas-empresariais)
![Repo size](https://img.shields.io/github/repo-size/Belcchior/escalas-empresariais)

Aplicacao Python para gerar escalas empresariais, exportar mapas mensais em Excel e produzir relatorios operacionais em Word.

O projeto simula um fluxo de recursos humanos para organizar funcionarios por funcoes, turnos, folgas, regras de descanso e necessidades operacionais.

## Funcionalidades

- Cadastro local de funcoes, colaboradores, disponibilidade e regras.
- Geracao de escala semanal e mensal.
- Validacoes basicas de horas maximas e descanso minimo.
- Exportacao de escala mensal em Excel.
- Geracao de relatorio operacional em Word.
- Persistencia local em JSON para reutilizar a configuracao.

## Requisitos

- Python 3.10+
- Dependencias em `requirements.txt`

## Instalacao

```bash
python -m venv .venv
pip install -r requirements.txt
```

## Execucao

```bash
python escala_app.py
```

## Validacao

```bash
python -m py_compile engine.py escala_app.py estado.py export_excel.py export_relatorio.py
python -m unittest discover -s tests
```

## Ficheiros locais

O projeto gera ficheiros locais como `estado_escala.json`, `Escala_Mensal.xlsx` e relatorios `.docx`. Esses ficheiros ficam fora do Git por conterem dados operacionais ou saidas geradas.

## Roadmap

- Criar dados ficticios de exemplo para demonstracao.
- Melhorar validacoes de conflitos de disponibilidade.
- Separar a interface de terminal da logica de negocio.
- Evoluir para uma interface grafica ou web.
