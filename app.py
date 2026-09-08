# app.py
import streamlit as st
import pandas as pd
import plotly.express as px
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">📑 Auto-BI: Cruzador de Dados & Tabelas Dinâmicas</div>', unsafe_allow_html=True)
st.caption("Suba qualquer planilha (.xlsx, .xls, .csv), faça PROCV visual sem fórmulas, crie tabelas dinâmicas e gere relatórios com design executivo.")
st.markdown("---")

# Função para leitura universal de planilhas
def carregar_planilha(arquivo):
    if arquivo.name.endswith('.csv'):
        try:
            df = pd.read_csv(arquivo, sep=None, engine='python')
        except Exception:
            df = pd.read_csv(arquivo, encoding='latin1', sep=';')
    else:
        df = pd.read_excel(arquivo)
    
    df.columns = [str(col).strip() for col in df.columns]
    return df

# Função de Exportação Profissional em Excel (Estilo Executivo com Cores e Fórmulas)
def exportar_excel_profissional(df, titulo="RELATÓRIO CONSOLIDADO"):
    df_export = df.reset_index() if isinstance(df, pd.DataFrame) else df
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dados_Processados"
    ws.views.sheetView[0].showGridLines = True
    
    num_cols = len(df_export.columns)
    
    # 1. Faixa de Título
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
    cell_top = ws.cell(row=1, column=1, value=titulo.upper())
    cell_top.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    cell_top.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    cell_top.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30
    
    # 2. Cabeçalho das Colunas
    for col_idx, col_name in enumerate(df_export.columns, start=1):
        c = ws.cell(row=2, column=col_idx, value=str(col_name))
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="0284C7", end_color="0284C7", fill_type="solid")
        c.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")
    ws.row_dimensions[2].height = 24
    
    thin_border = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )
    
    # 3. Dados com Cores Alternadas (Zebra)
    for r_idx, row in enumerate(df_export.itertuples(index=False), start=3):
        bg_color = "F8FAFC" if r_idx % 2 == 0 else "FFFFFF"
        fill_zebra = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = Font(name="Calibri", size=11, color="1E293B")
            cell.fill = fill_zebra
            cell.border = thin_border
            
            # Formatação numérica
            if isinstance(val, (int, float)):
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0.00"
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                
        ws.row_dimensions[r_idx].height = 20
        
    # 4. Linha de Totalização
    total_row = len(df_export) + 3
    ws.cell(row=total_row, column=1, value="TOTAL GERAL").font = Font(name="Calibri", size=11, bold=True, color="0F172A")
    ws.cell(row=total_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(row=total_row, column=1).fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    
    for c_idx in range(2, num_cols + 1):
        col_letter = get_column_letter(c_idx)
        # Verifica se é coluna numérica para somar
        val_teste = df_export.iloc[:, c_idx - 1]
        if pd.api.types.is_numeric_dtype(val_teste):
            c_tot = ws.cell(row=total_row, column=c_idx, value=f"=SUM({col_letter}3:{col_letter}{total_row-1})")
            c_tot.number_format = "#,##0.00"
        else:
            c_tot = ws.cell(row=total_row, column=c_idx, value="-")
            c_tot.alignment = Alignment(horizontal="center", vertical="center")
            
        c_tot.font = Font(name="Calibri", size=11, bold=True, color="0F172A")
        c_tot.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
        c_tot.border = Border(top=Side(style='thin', color='94A3B8'), bottom=Side(style='double', color='0F172A'))
        
    ws.row_dimensions[total_row].height = 24
    
    # 5. Ajuste automático da largura das colunas
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)
        
    output = io.BytesIO()
    wb.save(output)
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
            
            excel_data = exportar_excel_profissional(df_res, titulo="Base Consolidada - PROCV Automático")
            st.download_button(
                label="📥 Baixar Planilha Consolidada (Excel .xlsx Formatado)",
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
        colunas_lista = list(df_pivot.columns)
        
        with col_c1:
            eixo_linha = st.selectbox("Linhas do Agrupamento (Index):", colunas_lista, index=0)
        with col_c2:
            eixo_coluna = st.selectbox("Colunas (Opcional):", ["Nenhum"] + colunas_lista, index=0)
        with col_c3:
            idx_val = 1 if len(colunas_lista) > 1 else 0
            eixo_valor = st.selectbox("Coluna de Valores / Métricas:", colunas_lista, index=idx_val)
        with col_c4:
            tipo_agregacao = st.selectbox(
                "Operação de Cálculo:",
                options=["sum", "count", "mean", "min", "max"],
                format_func=lambda x: {
                    "sum": "Soma Numérica",
                    "count": "Contagem de Ocorrências",
                    "mean": "Média",
                    "min": "Valor Mínimo",
                    "max": "Valor Máximo"
                }[x]
            )

        try:
            df_calc = df_pivot.copy()
            if tipo_agregacao in ["sum", "mean", "min", "max"]:
                df_calc[eixo_valor] = pd.to_numeric(df_calc[eixo_valor], errors='coerce').fillna(0)

            if eixo_coluna == "Nenhum":
                if tipo_agregacao == "count":
                    pivot_table = df_calc.groupby(eixo_linha)[[eixo_valor]].count().rename(columns={eixo_valor: f"Qtd de {eixo_valor}"})
                elif tipo_agregacao == "sum":
                    pivot_table = df_calc.groupby(eixo_linha)[[eixo_valor]].sum().rename(columns={eixo_valor: f"Soma de {eixo_valor}"})
                elif tipo_agregacao == "mean":
                    pivot_table = df_calc.groupby(eixo_linha)[[eixo_valor]].mean().rename(columns={eixo_valor: f"Média de {eixo_valor}"})
                elif tipo_agregacao == "min":
                    pivot_table = df_calc.groupby(eixo_linha)[[eixo_valor]].min().rename(columns={eixo_valor: f"Mínimo de {eixo_valor}"})
                elif tipo_agregacao == "max":
                    pivot_table = df_calc.groupby(eixo_linha)[[eixo_valor]].max().rename(columns={eixo_valor: f"Máximo de {eixo_valor}"})
            else:
                pivot_table = df_calc.pivot_table(
                    index=eixo_linha,
                    columns=eixo_coluna,
                    values=eixo_valor,
                    aggfunc=tipo_agregacao,
                    fill_value=0
                )
            
            st.markdown("### 📋 Tabela Dinâmica Processada")
            st.dataframe(pivot_table, use_container_width=True)
            
            # Exportação com Layout Executivo
            excel_pivot = exportar_excel_profissional(pivot_table, titulo=f"Tabela Dinâmica — {tipo_agregacao.upper()} de {eixo_valor}")
            st.download_button(
                label="📥 Baixar Tabela Dinâmica (Excel .xlsx Formatado)",
                data=excel_pivot,
                file_name="tabela_dinamica_formatada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.markdown("---")
            st.markdown("### 📈 Gráficos Automáticos")
            
            df_chart = pivot_table.reset_index()
            col_metrica = pivot_table.columns[0] if eixo_coluna == "Nenhum" else None
            
            tipo_grafico = st.selectbox("Selecione o Modelo de Gráfico:", ["Barras", "Linhas", "Rosca / Pizza"])
            
            if tipo_grafico == "Barras":
                if eixo_coluna != "Nenhum":
                    fig = px.bar(df_chart, x=eixo_linha, y=pivot_table.columns, barmode="group")
                else:
                    fig = px.bar(df_chart, x=eixo_linha, y=col_metrica, color=eixo_linha)
            elif tipo_grafico == "Linhas":
                if eixo_coluna != "Nenhum":
                    fig = px.line(df_chart, x=eixo_linha, y=pivot_table.columns, markers=True)
                else:
                    fig = px.line(df_chart, x=eixo_linha, y=col_metrica, markers=True)
            elif tipo_grafico == "Rosca / Pizza":
                if eixo_coluna == "Nenhum":
                    fig = px.pie(df_chart, names=eixo_linha, values=col_metrica, hole=0.45)
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
            st.error(f"Erro no processamento: {e}")
