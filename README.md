# Norte 3.2 — Importação simplificada

Execute `pip install -r requirements.txt` e `streamlit run app.py`.

Na seção **Adicionar dados do negócio**, escolha: exemplo fictício, planilhas CSV/Excel ou preenchimento na tela.

- CSV: aceita separadores `;`, `,`, tab e `|`, com codificações comuns no Brasil.
- Excel: aceita `.xlsx`.
- Títulos de colunas: reconhece alguns nomes alternativos, como Vencimento, Qtd e Custo.
- Valores: aceita formato brasileiro `1.234,56`.
- Quando há erro de importação, o sistema exibe orientação sem interromper a demonstração.
- O banco SQLite de ações é local; não use dados reais confidenciais em hospedagem pública sem implementar isolamento por usuário e persistência segura.

Esta versão mantém o fluxo de investigação e resultado do Norte 3.1. Não há integração automática com bancos/ERPs nem execução externa de pagamentos ou compras.
