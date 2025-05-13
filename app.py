import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl

# Configuração da página
st.set_page_config(page_title="Análise de Estoque", layout="wide")
st.title("📈 Análise de Cobertura de Estoque")

# Upload do arquivo
uploaded_file = st.file_uploader("Carregue seu arquivo Excel (análise.xlsx)", type=["xlsx"])

if uploaded_file:
    # Carregar dados
    df = pd.read_excel(uploaded_file)

    # Validação das colunas obrigatórias
    required_cols = ["Filial", "Cobertura Atual", "Vlr Estoque Tmk", "Mercadoria", "Saldo Pedido"]
    if not all(col in df.columns for col in required_cols):
        st.error("⚠️ Arquivo inválido! Verifique se contém as colunas: 'Filial', 'Cobertura Atual', 'Vlr Estoque Tmk', 'Mercadoria', 'Saldo Pedido'.")
        st.stop()

    # Renomear colunas
    df = df.rename(columns={
        "Vlr Estoque Tmk": "valor_estoque",
        "Cobertura Atual": "cobertura_dias",
        "Filial": "filial",
        "Saldo Pedido": "saldo_pedido"
    })

    # Filtrar dados válidos (saldo e cobertura positivos)
    df = df[(df['cobertura_dias'] > 0) & (df['saldo_pedido'] > 0)].copy()

    # 📌 Cobertura Média Ponderada por Filial (continua usando valor_estoque)
    st.subheader("📌 Cobertura Média Ponderada por Filial")

    cobertura = (
        df.groupby("filial")
        .apply(lambda grupo: pd.Series({
            "Dias de Cobertura": np.average(grupo["cobertura_dias"], weights=grupo["valor_estoque"]),
            "Saldo Pedido Total": grupo["saldo_pedido"].sum()
        }))
        .round(2)
        .reset_index()
        .rename(columns={"filial": "Filial"})
    )

    styled_cobertura = cobertura.style \
        .format({"Dias de Cobertura": "{:.2f}", "Saldo Pedido Total": "R$ {:,.2f}"}) \
        .set_properties(**{'text-align': 'center'}) \
        .set_table_styles([{'selector': 'th', 'props': [('text-align', 'center')]}])

    st.dataframe(styled_cobertura, use_container_width=True)

    # 📊 Distribuição por Faixa de Cobertura (usando Saldo Pedido)
    st.subheader("📊 Distribuição por Faixa de Cobertura (Saldo de Pedido)")

    # Criar faixas de cobertura conforme solicitado
    df['faixa'] = pd.cut(
        df['cobertura_dias'],
        bins=[0, 15, 30, 45, 60, np.inf],
        labels=["0-15 dias", "16-30 dias", "31-45 dias", "46-60 dias", "Maior que 60 dias"]
    )

    # Agrupar por filial e faixa, somando o saldo de pedido
    resumo = df.groupby(['filial', 'faixa'])['saldo_pedido'].sum().unstack().fillna(0)

    # Adicionar coluna TOTAL por filial
    resumo['TOTAL'] = resumo.sum(axis=1)

    # Ordenar colunas
    ordem_colunas = [
        "0-15 dias", "16-30 dias", "31-45 dias",
        "46-60 dias", "Maior que 60 dias", "TOTAL"
    ]
    resumo = resumo[[col for col in ordem_colunas if col in resumo.columns]]

    # Exibir tabela formatada
    styled_resumo = resumo.style \
        .format("R$ {:,.2f}") \
        .set_properties(**{'text-align': 'center'}) \
        .set_table_styles([{'selector': 'th', 'props': [('text-align', 'center')]}])

    st.dataframe(styled_resumo, use_container_width=True)

    # 📥 Exportar para Excel
    output_final = io.BytesIO()
    with pd.ExcelWriter(output_final, engine='xlsxwriter') as writer:
        cobertura.to_excel(writer, sheet_name='Cobertura Média', index=False)
        resumo.to_excel(writer, sheet_name='Distribuição por Faixa')

    st.download_button(
        label="📥 Baixar Relatório Completo (Excel)",
        data=output_final.getvalue(),
        file_name="relatorio_estoque.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.warning("Por favor, carregue um arquivo Excel para análise.")