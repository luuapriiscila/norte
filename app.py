"""Norte 3.0: demonstração educacional de investigação financeira explicável."""
from datetime import date, timedelta
from pathlib import Path
import sqlite3
import io
import re
import unicodedata
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Norte 3.0 | Capital Invisível", page_icon="🧭", layout="wide")
DB = Path(__file__).with_name("norte.db")
TODAY = date.today()


def money(value):
    return "R$ " + f"{float(value):,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")


def init_db():
    with sqlite3.connect(DB) as conn:
        conn.execute("""CREATE TABLE IF NOT EXISTS actions
                     (id INTEGER PRIMARY KEY, created TEXT, title TEXT, detail TEXT, status TEXT)""")


def save_action(title, detail):
    with sqlite3.connect(DB) as conn:
        conn.execute("INSERT INTO actions(created,title,detail,status) VALUES(datetime('now'),?,?,?)",
                     (title, detail, "Autorizada — tarefa registrada"))


def load_actions():
    with sqlite3.connect(DB) as conn:
        return pd.read_sql_query("SELECT created AS Data, title AS Tarefa, detail AS Detalhes, status AS Status FROM actions ORDER BY id DESC", conn)


def demo_data():
    bills = pd.DataFrame([
        {"data": TODAY + timedelta(days=3), "tipo": "Saída", "descricao": "Fornecedor", "valor": 1250.0},
        {"data": TODAY + timedelta(days=5), "tipo": "Saída", "descricao": "Aluguel", "valor": 950.0},
        {"data": TODAY + timedelta(days=8), "tipo": "Entrada", "descricao": "Recebimento previsto", "valor": 1300.0},
        {"data": TODAY + timedelta(days=11), "tipo": "Saída", "descricao": "Energia", "valor": 390.0},
        {"data": TODAY + timedelta(days=14), "tipo": "Entrada", "descricao": "Recebimento previsto 2", "valor": 650.0},
    ])
    stock = pd.DataFrame([
        {"produto": "Café 500g", "quantidade": 35, "vendas_30d": 8, "custo_unitario": 18.0},
        {"produto": "Leite 1L", "quantidade": 20, "vendas_30d": 50, "custo_unitario": 5.0},
        {"produto": "Chocolate", "quantidade": 28, "vendas_30d": 7, "custo_unitario": 12.0},
        {"produto": "Biscoito", "quantidade": 16, "vendas_30d": 0, "custo_unitario": 9.0},
    ])
    return bills, stock


def clean_header(value):
    value = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode().lower().strip()
    return re.sub(r"[^a-z0-9]+", "_", value).strip("_")


def standardize_columns(df, kind):
    aliases = {
        "data": ["data", "dia", "vencimento", "data_vencimento", "date"],
        "tipo": ["tipo", "movimentacao", "natureza", "entrada_saida", "operacao"],
        "descricao": ["descricao", "descricao_da_conta", "historico", "detalhe", "conta", "nome"],
        "valor": ["valor", "valor_r", "valor_rs", "valor_em_reais", "montante", "total"],
        "produto": ["produto", "item", "mercadoria", "nome_do_produto", "nome"],
        "quantidade": ["quantidade", "qtd", "qtde", "estoque", "quantidade_em_estoque"],
        "vendas_30d": ["vendas_30d", "vendas_30_dias", "vendas_nos_ultimos_30_dias", "vendidos_30_dias", "vendas", "saida_30d"],
        "custo_unitario": ["custo_unitario", "custo", "preco_de_custo", "valor_unitario", "custo_por_unidade"],
    }
    wanted = ["data", "tipo", "descricao", "valor"] if kind == "contas" else ["produto", "quantidade", "vendas_30d", "custo_unitario"]
    headers = {clean_header(c): c for c in df.columns}
    rename = {}
    for target in wanted:
        for alias in aliases[target]:
            if alias in headers:
                rename[headers[alias]] = target
                break
    result = df.rename(columns=rename)
    missing = [c for c in wanted if c not in result.columns]
    if missing:
        raise ValueError("Não encontramos estas informações: " + ", ".join(missing) + ". Confira os títulos da planilha ou use o preenchimento manual.")
    return result[wanted].copy()


def parse_number(value):
    if pd.isna(value) or str(value).strip() == "":
        return float("nan")
    if isinstance(value, (int, float)):
        return float(value)
    v = str(value).strip().replace("R$", "").replace(" ", "")
    if "," in v and "." in v:
        v = v.replace(".", "").replace(",", ".") if v.rfind(",") > v.rfind(".") else v.replace(",", "")
    elif "," in v:
        v = v.replace(",", ".")
    return pd.to_numeric(v, errors="coerce")


def normalize_bills(df):
    result = standardize_columns(df, "contas")
    result["data"] = pd.to_datetime(result["data"], errors="coerce", dayfirst=True).dt.date
    result["valor"] = result["valor"].map(parse_number)
    def tipo_normalizado(v):
        v = clean_header(v)
        if v in ("entrada", "receita", "recebimento", "credito", "a_receber"):
            return "Entrada"
        if v in ("saida", "despesa", "pagamento", "debito", "a_pagar"):
            return "Saída"
        return ""
    result["tipo"] = result["tipo"].map(tipo_normalizado)
    result["descricao"] = result["descricao"].fillna("").astype(str).str.strip()
    if (result[["data", "valor"]].isna().any().any() or (~result["tipo"].isin(["Entrada", "Saída"])).any()
            or (result["valor"] < 0).any() or result["descricao"].eq("").any()):
        raise ValueError("Confira as datas, os valores e o tipo de movimentação (Entrada ou Saída).")
    return result


def normalize_stock(df):
    result = standardize_columns(df, "estoque")
    for col in ["quantidade", "vendas_30d", "custo_unitario"]:
        result[col] = result[col].map(parse_number)
    result["produto"] = result["produto"].fillna("").astype(str).str.strip()
    if (result[["quantidade", "vendas_30d", "custo_unitario"]].isna().any().any()
            or (result[["quantidade", "vendas_30d", "custo_unitario"]] < 0).any().any()
            or result["produto"].eq("").any()):
        raise ValueError("Confira nome do produto, quantidade, vendas em 30 dias e custo por unidade.")
    return result


def read_spreadsheet(upload):
    data = upload.getvalue()
    if upload.name.lower().endswith(".xlsx"):
        return pd.read_excel(io.BytesIO(data), engine="openpyxl")
    for encoding in ("utf-8-sig", "cp1252", "latin1"):
        try:
            raw = data.decode(encoding)
            for sep in (";", ",", "\t", "|"):
                df = pd.read_csv(io.StringIO(raw), sep=sep, dtype=str)
                if len(df.columns) >= 4:
                    return df
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue
    raise ValueError("Não conseguimos ler a planilha. Envie Excel (.xlsx) ou CSV (.csv), ou preencha os dados manualmente.")


def project(bills, opening, extra=None, receipt_factor=1.0):
    """Projeção diária: agrega movimentos do mesmo dia antes de testar mínimos."""
    flows = bills[bills["data"] >= TODAY].copy()
    if extra is not None:
        flows = pd.concat([flows, pd.DataFrame([extra])], ignore_index=True)
    if flows.empty:
        return pd.DataFrame(columns=["data", "movimento", "saldo"]), float(opening), float(opening)
    flows["movimento"] = flows.apply(
        lambda row: float(row["valor"]) * receipt_factor if row["tipo"] == "Entrada" else -float(row["valor"]), axis=1
    )
    daily = flows.groupby("data", as_index=False)["movimento"].sum().sort_values("data")
    daily["saldo"] = float(opening) + daily["movimento"].cumsum()
    return daily, min(float(opening), float(daily["saldo"].min())), float(daily["saldo"].iloc[-1])


def investigate(stock, months_limit):
    result = stock.copy()
    result["capital_em_estoque"] = result["quantidade"] * result["custo_unitario"]
    result["cobertura_meses"] = result.apply(
        lambda row: float(row["quantidade"]) / float(row["vendas_30d"]) if row["vendas_30d"] > 0 else None,
        axis=1,
    )
    result["sinal"] = result.apply(
        lambda row: "Sem vendas em 30 dias" if row["quantidade"] > 0 and row["vendas_30d"] == 0
        else "Cobertura acima do limite" if row["vendas_30d"] > 0 and row["cobertura_meses"] > months_limit
        else "Sem sinal pela regra", axis=1
    )
    flagged = result[result["sinal"] != "Sem sinal pela regra"].copy()
    return result, flagged


init_db()
st.title("🧭 Norte 3.1")
st.caption("Operação Capital Invisível · observar → investigar → comparar → preparar → autorizar")
st.info("Protótipo demonstrativo com regras explicáveis. Não usa IA generativa nem executa compras, pagamentos ou integrações externas. Os dados iniciais são fictícios.")

with st.sidebar:
    st.header("Configuração da investigação")
    nome = st.text_input("Nome do negócio", "Mercadinho Exemplo")
    saldo = st.number_input("Caixa disponível hoje (R$)", min_value=0.0, value=2300.0, step=100.0)
    reserva = st.number_input("Reserva mínima desejada (R$)", min_value=0.0, value=500.0, step=100.0)
    compra = st.number_input("Compra planejada (R$)", min_value=0.0, value=900.0, step=50.0)
    data_compra = st.date_input("Data de pagamento da compra", value=TODAY + timedelta(days=2), min_value=TODAY)
    limite = st.slider("Sinalizar cobertura acima de (meses)", min_value=1.0, max_value=6.0, value=2.0, step=0.5)
    st.caption("Regra de cobertura ajustável; não é previsão de demanda. Entradas futuras não são garantidas.")

bills, stock = demo_data()
with st.expander("📥 Adicionar dados do negócio", expanded=False):
    st.write("Comece com o exemplo, envie uma planilha Excel/CSV ou preencha os dados na tela. Não é necessário saber programar.")
    modo = st.radio("Como deseja começar?", ["Experimentar com dados de exemplo", "Enviar minhas planilhas", "Preencher na tela"], horizontal=False)
    if modo == "Enviar minhas planilhas":
        st.caption("Aceitamos Excel (.xlsx) e CSV com vírgula ou ponto e vírgula. Também reconhecemos títulos comuns como 'Vencimento', 'Descrição', 'Qtd' e 'Custo'.")
        left, right = st.columns(2)
        with left:
            st.download_button("⬇️ Modelo de contas", bills.assign(data=bills.data.astype(str)).to_csv(index=False, sep=";").encode("utf-8-sig"), "modelo_contas.csv", "text/csv")
            f1 = st.file_uploader("Contas a pagar e receber", type=["csv", "xlsx"], key="bills")
        with right:
            st.download_button("⬇️ Modelo de estoque", stock.to_csv(index=False, sep=";").encode("utf-8-sig"), "modelo_estoque.csv", "text/csv")
            f2 = st.file_uploader("Produtos em estoque", type=["csv", "xlsx"], key="stock")
        if f1 is not None:
            try:
                bills = normalize_bills(read_spreadsheet(f1))
                st.success(f"Contas importadas: {len(bills)} registros.")
                st.dataframe(bills.head(5), hide_index=True, use_container_width=True)
            except Exception as exc:
                st.warning(f"Não foi possível importar as contas: {exc} Os dados de exemplo serão mantidos até corrigir o arquivo.")
        if f2 is not None:
            try:
                stock = normalize_stock(read_spreadsheet(f2))
                st.success(f"Estoque importado: {len(stock)} produtos.")
                st.dataframe(stock.head(5), hide_index=True, use_container_width=True)
            except Exception as exc:
                st.warning(f"Não foi possível importar o estoque: {exc} Os dados de exemplo serão mantidos até corrigir o arquivo.")
        if f1 is None or f2 is None:
            st.info("O que não for enviado continuará usando os dados fictícios de demonstração.")
    elif modo == "Preencher na tela":
        st.write("Edite as tabelas diretamente. Para adicionar linhas, clique no último espaço vazio.")
        st.markdown("**Contas a pagar e receber**")
        edited_bills = st.data_editor(bills, num_rows="dynamic", hide_index=True, use_container_width=True, key="edit_bills",
            column_config={"data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"), "tipo": st.column_config.SelectboxColumn("Tipo", options=["Entrada", "Saída"]), "descricao": "Descrição", "valor": st.column_config.NumberColumn("Valor (R$)", min_value=0.0)})
        st.markdown("**Estoque**")
        edited_stock = st.data_editor(stock, num_rows="dynamic", hide_index=True, use_container_width=True, key="edit_stock",
            column_config={"produto": "Produto", "quantidade": st.column_config.NumberColumn("Quantidade", min_value=0), "vendas_30d": st.column_config.NumberColumn("Vendas em 30 dias", min_value=0), "custo_unitario": st.column_config.NumberColumn("Custo unitário (R$)", min_value=0.0)})
        try:
            bills = normalize_bills(edited_bills.dropna(how="all"))
            stock = normalize_stock(edited_stock.dropna(how="all"))
            st.success("Dados atualizados na análise abaixo.")
        except Exception as exc:
            st.warning(f"Revise os campos preenchidos: {exc} Enquanto isso, a demonstração continua com os dados de exemplo.")
    else:
        st.success("Demonstração pronta. Explore a missão sem enviar nenhum arquivo.")

bills = bills.sort_values("data").reset_index(drop=True)
stock_all, flagged = investigate(stock, limite)
capital_total = float(stock_all["capital_em_estoque"].sum())
capital_flagged = float(flagged["capital_em_estoque"].sum())
base_flow, base_min, base_end = project(bills, saldo)
extra = {"data": data_compra, "tipo": "Saída", "descricao": "Compra planejada", "valor": compra}
flow_purchase, min_purchase, end_purchase = project(bills, saldo, extra)

m1, m2, m3, m4 = st.columns(4)
m1.metric("Caixa hoje", money(saldo))
m2.metric("Capital sob investigação", money(capital_flagged), help="Custo de aquisição dos itens sinalizados; não representa economia recuperada.")
m3.metric("Menor caixa com compra", money(min_purchase))
m4.metric("Produtos sinalizados", str(len(flagged)))
st.caption(f"Capital total de aquisição no estoque: {money(capital_total)}. Projeções incluem apenas lançamentos informados; caixa não inclui estoque.")

t1, t2, t3, t4, t5 = st.tabs(["🔎 Missão Capital Invisível", "🧮 Decisões de compra", "📊 Evidências e dados", "📋 Resultado da investigação", "✅ Ações autorizadas"])

with t1:
    st.subheader("Missão 001 — Onde o capital está comprometido?")
    st.write("O Norte procura possíveis estoques elevados e cruza a descoberta com a decisão de compra e o caixa projetado.")
    if flagged.empty:
        st.success("Nenhum item foi sinalizado pela regra configurada. Isso não significa que o estoque esteja ideal.")
    else:
        st.warning(f"{len(flagged)} produto(s) merecem revisão; {money(capital_flagged)} em custo de aquisição estão associados a esses itens.")
        for _, item in flagged.sort_values("capital_em_estoque", ascending=False).iterrows():
            with st.container(border=True):
                st.markdown(f"**{item['produto']} — {item['sinal']}**")
                coverage = "sem vendas registradas" if pd.isna(item["cobertura_meses"]) else f"{item['cobertura_meses']:.1f} meses de cobertura aproximada"
                st.write(f"Estoque: {item['quantidade']:g} unidades · Vendas nos últimos 30 dias: {item['vendas_30d']:g} · {coverage}.")
                st.write(f"Capital de aquisição associado: **{money(item['capital_em_estoque'])}**.")
                st.caption("O sinal depende de dados corretos, sazonalidade, prazo de reposição e demanda. Não comprova desperdício ou prejuízo.")
                if st.button("Preparar conferência", key=f"stock_{item.name}"):
                    st.session_state["draft"] = (
                        f"Conferir estoque: {item['produto']}",
                        f"Negócio: {nome}. Quantidade: {item['quantidade']:g}; vendas 30 dias: {item['vendas_30d']:g}; custo registrado: {money(item['capital_em_estoque'])}. Verificar validade, demanda, sazonalidade e reposição."
                    )
                    st.success("Tarefa preparada. Revise em 'Ações autorizadas'.")
    st.markdown("#### Relação com o caixa")
    if min_purchase < 0:
        st.error(f"Com a compra planejada, o menor saldo projetado é {money(min_purchase)}: risco de caixa negativo.")
    elif min_purchase < reserva:
        st.warning(f"Com a compra planejada, o menor saldo projetado é {money(min_purchase)}, abaixo da reserva de {money(reserva)}.")
    else:
        st.success(f"Com a compra planejada, o menor saldo projetado é {money(min_purchase)}, acima da reserva definida.")
    st.write(f"Sem a compra, o menor saldo projetado seria {money(base_min)}. A comparação não considera efeitos da compra sobre vendas futuras.")
    if not flagged.empty and compra > 0:
        st.info("Investigação cruzada: há itens sinalizados no estoque e uma compra planejada. Confira se o pedido inclui itens já disponíveis e se a compra é necessária antes de autorizar.")
    st.caption("O Norte ainda não conhece os produtos que compõem a compra planejada; portanto, não afirma haver compra duplicada.")

with t2:
    st.subheader("Comparar alternativas de compra")
    st.write("A projeção inclui pagamentos e recebimentos registrados. Adiar não elimina o custo nem garante abastecimento.")
    later = data_compra + timedelta(days=7)
    options = [
        ("Compra integral", compra, data_compra),
        ("Compra pela metade", compra * 0.5, data_compra),
        ("Adiar 7 dias", compra, later),
    ]
    rows = []
    for title, amount, when in options:
        _, low, end = project(bills, saldo, {"data": when, "tipo": "Saída", "descricao": title, "valor": amount})
        rows.append({"Cenário": title, "Pagamento": str(when), "Desembolso": money(amount), "Menor saldo": money(low), "Saldo final": money(end), "Abaixo da reserva?": "Sim" if low < reserva else "Não"})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    choice = st.selectbox("Qual alternativa deseja preparar?", [r["Cenário"] for r in rows])
    selected = next(item for item in options if item[0] == choice)
    if st.button("Preparar plano para autorização", type="primary"):
        st.session_state["draft"] = (
            f"Revisar compra — {choice}",
            f"Negócio: {nome}. Alternativa: {choice}. Desembolso: {money(selected[1])}; pagamento: {selected[2]}. Valor original: {money(compra)}. Verificar necessidade por produto, risco de ruptura, fornecedores e vencimentos antes de executar."
        )
        st.success("Plano preparado. Vá à aba 'Ações autorizadas' para revisar e registrar.")
    st.markdown("#### Teste de sensibilidade dos recebimentos")
    receipt_pct = st.slider("Percentual dos recebimentos previstos que se concretizam", 0, 100, 100, 10)
    _, stress_min, stress_end = project(bills, saldo, extra, receipt_factor=receipt_pct / 100)
    a, b = st.columns(2)
    a.metric("Menor caixa no cenário", money(stress_min))
    b.metric("Saldo final no cenário", money(stress_end))
    st.caption("Cenário de sensibilidade, não previsão probabilística. Reduz apenas as entradas futuras informadas.")

with t3:
    st.subheader("Evidências da investigação")
    st.write("Projeção diária considerando a compra planejada")
    graph = pd.concat([pd.DataFrame([{"data": TODAY, "saldo": saldo}]), flow_purchase[["data", "saldo"]]], ignore_index=True)
    st.line_chart(graph.set_index("data")["saldo"])
    st.markdown("**Contas e recebimentos**")
    st.dataframe(bills.assign(data=bills["data"].astype(str)), use_container_width=True, hide_index=True)
    st.markdown("**Estoque e regras aplicadas**")
    shown = stock_all.copy()
    shown["cobertura_meses"] = shown["cobertura_meses"].round(2)
    st.dataframe(shown, use_container_width=True, hide_index=True)
    st.caption(f"Critério ilustrativo: cobertura acima de {limite:g} meses ou zero vendas nos últimos 30 dias. Cobertura = quantidade em estoque ÷ vendas em 30 dias. Não é um diagnóstico definitivo.")
    st.download_button("Baixar relatório de investigação CSV", shown.to_csv(index=False).encode("utf-8-sig"), "norte_investigacao.csv", "text/csv")

with t4:
    st.subheader("📋 Resultado da investigação")
    st.caption("Missão 001 — Capital Invisível · relatório gerado com os dados atuais. Valores de exemplo são fictícios até que você importe seus próprios dados.")

    col_a, col_b, col_c = st.columns(3)
    col_a.metric("Capital sob investigação", money(capital_flagged), help="Custo de aquisição dos produtos sinalizados; não é economia recuperada.")
    col_b.metric("Compra planejada", money(compra))
    col_c.metric("Menor caixa com a compra", money(min_purchase))

    st.markdown("#### 1. O que o Norte identificou?")
    if flagged.empty:
        st.success("Nenhum produto sinalizado pelo critério de cobertura selecionado.")
    else:
        st.warning(f"{len(flagged)} produto(s) com possível baixa movimentação representam {money(capital_flagged)} em custo de aquisição registrado.")
        summary = flagged[["produto", "quantidade", "vendas_30d", "capital_em_estoque", "sinal"]].copy()
        summary.columns = ["Produto", "Quantidade", "Vendas em 30 dias", "Capital de aquisição (R$)", "Motivo"]
        st.dataframe(summary, use_container_width=True, hide_index=True)

    if min_purchase < 0:
        st.error(f"Risco de caixa negativo: menor saldo com a compra = {money(min_purchase)}.")
    elif min_purchase < reserva:
        st.warning(f"A compra deixa o caixa abaixo da reserva definida de {money(reserva)}.")
    else:
        st.info(f"O caixa projetado permanece acima da reserva de {money(reserva)}, conforme os dados registrados.")
    st.caption("A análise de estoque e a análise de caixa são sinais complementares. Como o pedido não está detalhado por produto, não é possível afirmar que a compra repete itens existentes.")

    st.markdown("#### 2. Como cada decisão altera o caixa?")
    report_options = [
        ("Compra integral", float(compra), data_compra),
        ("Compra pela metade", float(compra) / 2, data_compra),
        ("Adiar 7 dias", float(compra), data_compra + timedelta(days=7)),
    ]
    comparisons = []
    for label, amount, when in report_options:
        _, low, end = project(bills, saldo, {"data": when, "tipo": "Saída", "descricao": label, "valor": amount})
        comparisons.append({"Alternativa": label, "Pagamento": str(when), "Desembolso (R$)": amount,
                            "Menor caixa (R$)": low, "Saldo final (R$)": end,
                            "Margem acima da reserva (R$)": low - reserva})
    comp_df = pd.DataFrame(comparisons)
    st.dataframe(comp_df, use_container_width=True, hide_index=True,
                 column_config={"Desembolso (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
                                "Menor caixa (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
                                "Saldo final (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
                                "Margem acima da reserva (R$)": st.column_config.NumberColumn(format="R$ %.2f")})
    selected_label = st.selectbox("Alternativa que deseja analisar e preparar", comp_df["Alternativa"].tolist(), key="resultado_escolha")
    chosen = next(row for row in comparisons if row["Alternativa"] == selected_label)
    original = comparisons[0]
    preserved_now = float(compra) - (chosen["Desembolso (R$)"] if chosen["Pagamento"] == str(data_compra) else 0.0)
    st.metric("Desembolso evitado na data original", money(preserved_now),
              help="Valor não pago na data inicialmente planejada. Adiar não elimina a obrigação e não é economia realizada.")
    st.write(f"**Menor caixa nesta alternativa:** {money(chosen['Menor caixa (R$)'])} "
             f"(diferença de {money(chosen['Menor caixa (R$)'] - original['Menor caixa (R$)'])} em relação à compra integral).")
    if chosen["Menor caixa (R$)"] < reserva:
        st.warning("Mesmo nesta alternativa, o caixa projetado fica abaixo da reserva. É necessário revisar o plano.")
    else:
        st.success("Nesta simulação, o menor caixa projetado não fica abaixo da reserva escolhida.")

    st.markdown("#### 3. Próxima decisão")
    st.write("Conferir a necessidade de reposição por produto, prazo do fornecedor, validade, sazonalidade e risco de ruptura antes de aprovar qualquer compra.")
    if st.button("Preparar esta decisão para autorização", type="primary", key="resultado_preparar"):
        st.session_state["draft"] = (
            f"Missão 001 — revisar compra: {selected_label}",
            f"Negócio: {nome}. Capital sob investigação: {money(capital_flagged)}; "
            f"alternativa: {selected_label}; desembolso: {money(chosen['Desembolso (R$)'])}; "
            f"data: {chosen['Pagamento']}; menor caixa projetado: {money(chosen['Menor caixa (R$)'])}; "
            "conferir produtos, demanda, fornecedores e vencimentos. Ação preparada, não executada."
        )
        st.success("Decisão preparada! Abra a aba 'Ações autorizadas' para revisar e registrar.")

    report_lines = [
        "NORTE 3.1 — RESULTADO DA INVESTIGAÇÃO", "Missão 001 — Capital Invisível",
        f"Negócio: {nome}", f"Data da análise: {TODAY}",
        "Dados: fornecidos pelo usuário ou cenário fictício demonstrativo.",
        f"Produtos sinalizados: {len(flagged)}", f"Capital sob investigação: {money(capital_flagged)}",
        f"Compra planejada: {money(compra)}", f"Reserva definida: {money(reserva)}",
        f"Menor caixa com compra integral: {money(min_purchase)}", "",
        "ALTERNATIVAS SIMULADAS",
    ]
    for row in comparisons:
        report_lines.append(f"{row['Alternativa']} | {row['Pagamento']} | desembolso {money(row['Desembolso (R$)'])} | menor caixa {money(row['Menor caixa (R$)'])} | saldo final {money(row['Saldo final (R$)'])}")
    report_lines += ["", "PRODUTOS SINALIZADOS"]
    for _, item in flagged.iterrows():
        report_lines.append(f"{item['produto']}: {item['quantidade']:g} un.; vendas 30d {item['vendas_30d']:g}; capital {money(item['capital_em_estoque'])}; {item['sinal']}")
    report_lines += ["", "LIMITAÇÕES: estimativas baseadas em regras e dados informados; não comprovam desperdício ou economia. A compra não está detalhada por produto. Recebimentos não são garantidos. Nenhuma ação financeira é executada."]
    st.download_button("📥 Baixar resultado da investigação (.txt)", "\n".join(report_lines).encode("utf-8"),
                       file_name="norte_resultado_investigacao.txt", mime="text/plain", key="resultado_baixar")

with t5:
    st.subheader("Revisar e autorizar uma tarefa")
    draft = st.session_state.get("draft")
    if draft:
        st.markdown(f"**{draft[0]}**")
        st.write(draft[1])
        # Checkbox não tem key persistente: evita modificar widget já criado no mesmo ciclo.
        approved = st.checkbox("Li e autorizo o registro desta tarefa no histórico local.")
        if st.button("Autorizar e registrar", type="primary", disabled=not approved):
            save_action(*draft)
            st.session_state.pop("draft", None)
            st.rerun()
    else:
        st.info("Nenhuma tarefa preparada. Escolha uma descoberta ou alternativa nas outras abas.")
    history = load_actions()
    if not history.empty:
        st.markdown("**Histórico de tarefas autorizadas**")
        st.dataframe(history, use_container_width=True, hide_index=True)
    else:
        st.caption("Nenhuma tarefa registrada até o momento.")
    st.caption("Autorizar aqui apenas salva um registro no SQLite local. Não cria pedidos, transferências ou mensagens externas.")

st.divider()
st.caption("Norte 3.1 — MVP educacional. Dados fictícios por padrão. Análises são indicativas e não substituem avaliação financeira profissional.")
