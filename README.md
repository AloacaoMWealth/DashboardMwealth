# Dashboard Macro M Wealth

Aplicativo Streamlit para acompanhamento macro do escritório usando a planilha `data/Controle de Clientes MWealth 2026.xlsx`.

## Rodar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Atualizar a base

Substitua o arquivo abaixo no GitHub:

```text
data/Controle de Clientes MWealth 2026.xlsx
```

## Observações

- Contas offshore, incluindo Charles Schwab/Charles Shwab, são convertidas de USD para BRL.
- A cotação USDBRL tem prioridade quando preenchida na planilha, na aba de análise com linha `Dólar Price`.
- Se a cotação não estiver preenchida, o app tenta buscar via `yfinance` usando o fechamento da data/mês do PL.
- A última coluna de PL usada como `PL Atual` é a última coluna com valor preenchido, evitando zerar o dashboard quando o mês D0 ainda está vazio.
