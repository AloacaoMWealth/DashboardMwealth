# Dashboard Macro M Wealth

Dashboard institucional em Streamlit para acompanhamento de PL, meta patrimonial, corretoras, canais, segmentação e análise detalhada por período.

## Como rodar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Atualização da base

Substitua o arquivo abaixo no GitHub:

```text
data/Controle de Clientes MWealth 2026.xlsx
```

O app usa a aba `Controle Contas` e identifica as colunas de PL no padrão `PL dd/mm/aaaa`.

## Conversão offshore

Contas classificadas como Offshore são convertidas de USD para BRL usando o fechamento USDBRL via yfinance para cada data de referência de PL. O app tenta `USDBRL=X` e `BRL=X`.

## Tema

O projeto inclui `.streamlit/config.toml` para forçar tema claro e evitar distorções visuais nos filtros e gráficos.
