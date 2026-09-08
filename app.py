# app.py
import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(
    page_title="Auto-BI & Manipulador de Planilhas Inteligente",
    page_icon="📑",
    layout="wide"
)

# Estilização visual customizada
st.markdown("""
<style>
    .main-header {
        font-size: 26px;
        font-weight: 700;
        color: #0284C7;
        margin-bottom: 5px;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 15px;
    }
    .stTabs [data-baseweb="tab"] {
        font-weight: 600;
        font-size: 15px;
    }
    .metric-box {
        background-color: #1E293B;
        border: 1px solid #334155;
        padding: 15px;
        border-radius: 10px;
        color: #F8FAFC;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">📑 Auto-BI: Cruzador de Dados & Tabelas Dinâmicas</div>', unsafe_allow_html=True)
st.caption("Suba qualquer planilha (.xlsx, .xls, .csv), faça PROCV visual sem fórmulas, crie tabelas dinâmicas e gere gráficos instantâneos.")
st.markdown("---")

# Função para leitura universal de planilhas
def carregar_planilha(arquivo):
    if arquivo.name.endswith('.csv'):
        try:
            return pd.read_csv(arquivo, sep=None, engine='python')
        except Exception:
            return pd.read_csv(arquivo, encoding='latin1', sep=';')
    else:
        return pd.read_excel(arquivo)

# Função para exportação em Excel
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=True, sheet_name='Resultado_Processado')
    return output.getvalue()

aba_procv, aba_pivot = st.tabs([
    "🔗 1. Super PROCV / PROCX (Cruzamento de 2 Planilhas)",
    "📊 2. Tabela Dinâmica & Gráficos Automáticos"
])

# =========================================================
# ABA 1: SUPER PROCV / CRUZAMENTO DE BASES
# =========================================================
with aba_procv:
    st.subheader("Cruzamento Inteligente de Bases (PROCV Visual)")
    st.write("Junte duas tabelas utilizando uma coluna comum (ex.: CPF, Matrícula, Código, CA, ID).")
    
    col_up1, col_up2 = st.columns(2)
    
    with col_up1:
        st.markdown("**Planilha Principal (Base A)**")
        file_a = st.file_uploader("Suba a tabela principal:", type=["xlsx", "xls", "csv"], key="file_a")
        
    with col_up2:
        st.markdown("**Planilha de Consulta / Procura (Base B)**")
        file_b = st.file_uploader("Suba a tabela que contém os dados a buscar:", type=["xlsx", "xls", "csv"], key="file_b")

    if file_a and file_b:
        df_a = carregar_planilha(file_a)
        df_b = carregar_planilha(file_b)
        
        st.success(f"Bases carregadas com sucesso! Base A: {df_a.shape[0]} linhas | Base B: {df_b.shape[0]} linhas")
        
        st.markdown("---")
        st.markdown("### Configuração do Cruzamento")
        
        c1, c2, c3 = st.columns(3)
        with c1:
            col_chave_a = st.selectbox("Coluna-Chave na Planilha Principal (Base A):", df_a.columns)
        with c2:
            col_chave_b = st.selectbox("Coluna-Chave na Planilha de Procura (Base B):", df_b.columns)
        with c3:
            tipo_juncao = st.selectbox(
                "Tipo de PROCV / Junção:",
                options=["left", "inner", "outer"],
                format_func=lambda x: {
                    "left": "PROCV Padrão (Mantém todos da Base A e traz correspondentes da B)",
                    "inner": "Apenas Correspondências (Mantém apenas registros em ambas)",
                    "outer": "Completo (Mantém tudo de ambas as bases)"
                }[x]
            )

        colunas_disponiveis_b = [c for c in df_b.columns if c != col_chave_b]
        colunas_para_trazer = st.multiselect(
            "Selecione quais colunas da Planilha B deseja adicionar à Principal:",
            options=colunas_disponiveis_b,
            default=colunas_disponiveis_b[:3] if len(colunas_disponiveis_b) >= 3 else colunas_disponiveis_b
        )

        if st.button("⚡ Executar Cruzamento de Dados (PROCV)", type="primary"):
            cols_finais_b = [col_chave_b] + colunas_para_trazer
            df_b_filtrado = df_b[cols_finais_b].drop_duplicates(subset=[col_chave_b])
            
            df_resultado_merge = pd.merge(
                df_a,
                df_b_filtrado,
                left_on=col_chave_a,
                right_on=col_chave_b,
                how=tipo_juncao
            )
            
            st.session_state["df_merge_result"] = df_resultado_merge
            st.session_state["df_pivot_ativo"] = df_resultado_merge
            st.success(f"Cruzamento concluído com sucesso! {df_resultado_merge.shape[0]} linhas geradas.")

        if "df_merge_result" in st.session_state:
            df_res = st.session_state["df_merge_result"]
            st.markdown("### 📋 Pré-visualização da Planilha Cruzada")
            st.dataframe(df_res.head(20), use_container_width=True)
            
            excel_data = to_excel(df_res)
            st.download_button(
                label="📥 Baixar Planilha Consolidada (Excel .xlsx)",
                data=excel_data,
                file_name="base_cruzada_procv.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# =========================================================
# ABA 2: TABELA DINÂMICA & GRÁFICOS
# =========================================================
with aba_pivot:
    st.subheader("Construtor de Tabela Dinâmica & Visualizações Gráficas")
    
    origem_dados = st.radio(
        "Origem dos dados para análise:",
        options=["Subir uma nova planilha", "Usar resultado do cruzamento (Aba 1)"],
        horizontal=True
    )
    
    df_pivot = None
    
    if origem_dados == "Subir uma nova planilha":
        file_pivot = st.file_uploader("Selecione sua planilha para análise:", type=["xlsx", "xls", "csv"], key="file_pivot")
        if file_pivot:
            df_pivot = carregar_planilha(file_pivot)
    else:
        if "df_pivot_ativo" in st.session_state:
            df_pivot = st.session_state["df_pivot_ativo"]
        else:
            st.info("Nenhuma base encontrada da Aba 1. Selecione 'Subir uma nova planilha' acima.")

    if df_pivot is not None:
        st.write(f"**Base Carregada:** {df_pivot.shape[0]} linhas × {df_pivot.shape[1]} colunas")
        st.dataframe(df_pivot.head(5), use_container_width=True)
        
        st.markdown("---")
        st.markdown("### Configurar Dimensões da Tabela Dinâmica")
        
        col_c1, col_c2, col_c3, col_c4 = st.columns(4)
        
        with col_c1:
            eixo_linha = st.selectbox("Linhas do Agrupamento (Index):", df_pivot.columns)
        with col_c2:
            eixo_coluna = st.selectbox("Colunas (Opcional):", ["Nenhum"] + list(df_pivot.columns))
        with col_c3:
            eixo_valor = st.selectbox("Coluna de Valores / Métricas:", df_pivot.columns)
        with col_c4:
            tipo_agregacao = st.selectbox(
                "Operação de Cálculo:",
                options=["count", "sum", "mean", "min", "max"],
                format_func=lambda x: {
                    "count": "Contagem de Ocorrências",
                    "sum": "Soma Numérica",
                    "mean": "Média",
                    "min": "Valor Mínimo",
                    "max": "Valor Máximo"
                }[x]
            )

        col_col = None if eixo_coluna == "Nenhum" else eixo_coluna
        
        try:
            pivot_table = df_pivot.pivot_table(
                index=eixo_linha,
                columns=col_col,
                values=eixo_valor,
                aggfunc=tipo_agregacao,
                fill_value=0
            )
            
            st.markdown("### 📋 Tabela Dinâmica Processada")
            st.dataframe(pivot_table, use_container_width=True)
            
            excel_pivot = to_excel(pivot_table)
            st.download_button(
                label="📥 Baixar Tabela Dinâmica (Excel .xlsx)",
                data=excel_pivot,
                file_name="tabela_dinamica_processada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.markdown("---")
            st.markdown("### 📈 Gráficos Automáticos")
            
            df_chart = pivot_table.reset_index()
            tipo_grafico = st.selectbox("Selecione o Modelo de Gráfico:", ["Barras", "Linhas", "Rosca / Pizza"])
            
            if tipo_grafico == "Barras":
                if col_col:
                    fig = px.bar(df_chart, x=eixo_linha, y=pivot_table.columns, barmode="group")
                else:
                    fig = px.bar(df_chart, x=eixo_linha, y=eixo_valor, color=eixo_linha)
            elif tipo_grafico == "Linhas":
                if col_col:
                    fig = px.line(df_chart, x=eixo_linha, y=pivot_table.columns, markers=True)
                else:
                    fig = px.line(df_chart, x=eixo_linha, y=eixo_valor, markers=True)
            elif tipo_grafico == "Rosca / Pizza":
                if not col_col:
                    fig = px.pie(df_chart, names=eixo_linha, values=eixo_valor, hole=0.45)
                else:
                    st.warning("Para gráficos de Pizza/Rosca, selecione 'Nenhum' no campo de Colunas acima.")
                    fig = None
                    
            if fig:
                fig.update_layout(
                    plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)",
                    margin=dict(t=20, l=10, r=10, b=20)
                )
                st.plotly_chart(fig, use_container_width=True)
                
        except Exception as e:
            st.error(f"Erro na agregação. Para operações de Soma ou Média, certifique-se de que a coluna de valores contém dados numéricos válidos. Detalhes: {e}")
