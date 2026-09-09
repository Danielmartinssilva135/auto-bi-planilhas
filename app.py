# Mapeamento cronológico para ordenação correta dos meses
ORDEM_MESES = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

def carregar_planilha(arquivo):
    if arquivo.name.endswith('.csv'):
        try:
            df = pd.read_csv(arquivo, sep=None, engine='python')
        except Exception:
            df = pd.read_csv(arquivo, encoding='latin1', sep=';')
    else:
        df = pd.read_excel(arquivo)
    
    df.columns = [str(col).strip() for col in df.columns]
    
    # Remove automaticamente linhas residuais de "TOTAL" ou "TOTAL GERAL" da base bruta
    col_zero = df.columns[0]
    df = df[~df[col_zero].astype(str).str.strip().str.upper().isin(["TOTAL", "TOTAL GERAL", "SUBTOTAL"])]
    
    # Se a primeira coluna contiver meses, preserva a ordem cronológica
    if df[col_zero].astype(str).isin(ORDEM_MESES).any():
        df[col_zero] = pd.Categorical(df[col_zero], categories=ORDEM_MESES, ordered=True)
        df = df.sort_values(by=col_zero)
        
    return df
