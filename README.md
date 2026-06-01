# Escalas Empresariais

[![Python check](https://github.com/Belcchior/escalas-empresariais/actions/workflows/python-check.yml/badge.svg)](https://github.com/Belcchior/escalas-empresariais/actions/workflows/python-check.yml)
![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![License](https://img.shields.io/github/license/Belcchior/escalas-empresariais)
![Last commit](https://img.shields.io/github/last-commit/Belcchior/escalas-empresariais)
![Repo size](https://img.shields.io/github/repo-size/Belcchior/escalas-empresariais)

Sistema em Python para geração de escalas empresariais, pensado para simular um fluxo de Recursos Humanos e operações: cadastro de funcionários, funções, turnos, folgas, regras de descanso, necessidades por setor e exportação de resultados.

O objetivo do projeto é automatizar parte do processo de montagem de horários de trabalho, reduzindo conflitos manuais e gerando entregáveis úteis para gestão: planilhas Excel e relatórios operacionais em Word.

## Visão Geral

O sistema permite configurar uma empresa fictícia com:

- Funções ou cargos necessários por turno.
- Funcionários com múltiplas funções.
- Folgas fixas.
- Pedidos de folga.
- Restrições de disponibilidade.
- Limites de horas por dia e por semana.
- Descanso mínimo entre turnos.
- Necessidade de equipe por dia e horário.

Com base nessas informações, o motor tenta distribuir os colaboradores disponíveis e marca vagas em aberto quando não há equipe suficiente.

## Funcionalidades

- Cadastro local de funções, funcionários, disponibilidade e regras.
- Geração de escala semanal.
- Geração de escala mensal.
- Distribuição de colaboradores conforme funções compatíveis.
- Respeito a folgas fixas, pedidos de folga e restrições.
- Controle básico de limite diário e semanal de horas.
- Validação de descanso mínimo entre turnos.
- Identificação de vagas em aberto.
- Exportação da escala mensal para Excel.
- Geração de relatório operacional em Word.
- Persistência local em JSON para reutilizar configurações.

## Estrutura do Projeto

```text
.
├── escala_app.py          # Fluxo principal da aplicação via terminal
├── engine.py              # Motor de geração das escalas
├── estado.py              # Persistência local em JSON
├── export_excel.py        # Exportação da escala para Excel
├── export_relatorio.py    # Análise operacional e relatório Word
├── tests/                 # Testes automatizados
├── examples/              # Dados fictícios de exemplo
├── requirements.txt       # Dependências Python
└── README.md
