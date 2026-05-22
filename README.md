# Dashboard Macro M Wealth

Aplicativo em Streamlit para acompanhamento macro do escritório, com leitura da planilha `Controle de Clientes MWealth 2026.xlsx`.

## Como rodar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Como atualizar a base no GitHub

Substitua o arquivo abaixo e faça novo commit:

```text
data/Controle de Clientes MWealth 2026.xlsx
```

O app usa esse arquivo como base padrão. Também existe opção de upload na sidebar para teste local, mas no deploy via GitHub o fluxo recomendado é substituir a planilha na pasta `data`.

## Estrutura esperada da planilha

A aba principal deve se chamar `Controle Contas` e conter, pelo menos, as colunas:

- Corretora
- Grupo Geral
- Grupo Familiar
- Cliente
- PF/ PJ
- Canal
- Conta
- Consultor
- PL dd/mm/aaaa

As colunas de PL mês a mês são identificadas automaticamente pelo padrão `PL 31/01/2025`, `PL 28/02/2025`, etc.
