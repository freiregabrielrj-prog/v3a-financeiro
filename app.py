import streamlit as st
import pandas as pd
import unicodedata
from datetime import datetime
import plotly.graph_objects as go
import openpyxl
import base64

# 1. CONFIGURAÇÃO DE PÁGINA (Deve ser o primeiro comando Streamlit)
st.set_page_config(page_title="V3A Financeiro", layout="wide", initial_sidebar_state="collapsed")




# =================================================================
# 2. SISTEMA DE ACESSO (LOGIN) - VERSÃO FINAL CORRIGIDA
# =================================================================

# Inicializa o estado de autenticação
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

def validar_senha():
    # Busca a senha digitada no widget
    senha = st.session_state.get("senha_input")
    
    if senha == "cashflow":
        st.session_state.autenticado = True
        # Limpa o campo para a próxima execução, mas mantém a chave ativa
        st.session_state["senha_input"] = "" 
    elif senha != "" and senha is not None:
        # Só exibe erro se o campo NÃO estiver vazio e a senha for errada
        st.error("Senha incorreta. Tente novamente.")

# Se não estiver autenticado, desenha a tela de login e para a execução do resto do app
if not st.session_state.autenticado:
    st.markdown("""
        <style>
        /* RESET E CENTRALIZAÇÃO */
        .main .block-container {
            max-width: 100% !important;
            padding: 0 !important;
            display: flex !important;
            justify-content: center !important;
            align-items: center !important;
            height: 100vh !important;
            background-color: white;
        }

        /* CONTAINER DO CARD DE LOGIN */
        .login-wrapper {
            width: 100% !important;
            max-width: 320px !important; 
            display: flex !important;
            flex-direction: column !important;
            align-items: center !important; 
            text-align: center !important;
            padding: 20px;
        }

        /* ESTILO DO BOTÃO */
        .stButton button {
            width: 100% !important;
            background-color: #B8860B !important;
            color: white !important;
            font-weight: bold !important;
            height: 45px !important;
            border-radius: 8px !important;
            border: none !important;
            margin-top: 10px !important;
        }

        /* ESTILO DO CAMPO DE TEXTO */
        .stTextInput input {
            text-align: center !important;
            height: 45px !important;
            border-radius: 8px !important;
        }

        /* ESCONDE CABEÇALHOS E MENUS DURANTE LOGIN */
        [data-testid="stHeader"], [data-testid="stSidebar"], [data-testid="stToolbar"], label {
            display: none !important;
        }
        </style>
    """, unsafe_allow_html=True)

    # Início do Layout de Login
    st.markdown('<div class="login-wrapper">', unsafe_allow_html=True)
    
    # Logo
    try:
        st.image("image_3.png", width=160)
    except:
        st.markdown("<h3 style='text-align:center;'>V3A</h3>", unsafe_allow_html=True)

    st.markdown("<h2 style='margin-top: 20px; color: #1A1A1A; font-family: sans-serif;'>Acesso Restrito</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #666; font-size: 14px; margin-bottom: 25px;'>Digite a senha para acessar o painel</p>", unsafe_allow_html=True)

    # Widget de entrada de senha
    # Usamos on_change para validar quando o usuário aperta ENTER
    st.text_input("Senha", type="password", key="senha_input", placeholder="Sua senha aqui", on_change=validar_senha)
    
    # Botão de acesso
    if st.button("Acessar Relatório"):
        validar_senha()
        if st.session_state.autenticado:
            st.rerun() # Força o app a recarregar e esconder o login imediatamente
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Interrompe a execução aqui enquanto não logar
    st.stop()




# =================================================================
# 3. CARREGAMENTO DE DADOS (Só inicia após o login)
# =================================================================
# Seu código de load_all_v3a_data() continua aqui...
        
    
    # Estrutura Visual
    # Usamos o markdown para envolver os componentes do Streamlit no nosso Card CSS
    st.markdown('<div class="login-card">', unsafe_allow_html=True)
    
    # Renderização da Logo
    try:
        with open("image_3.png", "rb") as f:
            logo_b64 = base64.b64encode(f.read()).decode()
            st.markdown(f'<img src="data:image/png;base64,{logo_b64}" style="width: 130px; margin-bottom: 20px; filter: drop-shadow(0 0 10px rgba(0,0,0,0.2));">', unsafe_allow_html=True)
    except:
        pass

    st.markdown("<h2 style='margin: 0; font-weight: 700; letter-spacing: -0.5px;'>Painel Executivo</h2>", unsafe_allow_html=True)
    st.markdown("<p style='opacity: 0.7; font-size: 14px; margin-bottom: 2rem;'>Financeiro V3A</p>", unsafe_allow_html=True)
    
    # Campo de Senha
    st.text_input("SENHA DE ACESSO", type="password", key="senha_digitada", on_change=validar_senha, placeholder="••••••••")
    
    # Botão de Login
    st.button("Entrar no Sistema", on_click=validar_senha)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.stop()
    

    # Criamos a estrutura do card usando uma única div central
    # Isso evita que o Streamlit tente dividir em colunas 0.5, 1, 0.5
    st.markdown('<div class="main-login-container">', unsafe_allow_html=True)
    st.markdown('<div class="login-card">', unsafe_allow_html=True)
    
    # Logo
    try:
        with open("image_3.png", "rb") as f:
            logo_b64 = base64.b64encode(f.read()).decode()
            st.markdown(f'<img src="data:image/png;base64,{logo_b64}" style="width: 140px; margin-bottom: 20px;">', unsafe_allow_html=True)
    except:
        pass

    st.markdown("<h2 style='margin: 0; font-family: Segoe UI;'>Acesso Restrito</h2>", unsafe_allow_html=True)
    st.markdown("<p style='opacity: 0.8; margin-bottom: 30px;'>Painel Executivo Financeiro</p>", unsafe_allow_html=True)
    
    # O Streamlit renderiza os widgets dentro da div acima
    st.text_input("SENHA DE ACESSO", type="password", key="senha_digitada", on_change=validar_senha, placeholder="Digite a senha...")
    st.button("DESBLOQUEAR", on_click=validar_senha)
    
    st.markdown('</div></div>', unsafe_allow_html=True)
    
    st.stop()
    
    
# =================================================================
# 3. RESTANTE DO CÓDIGO (Só executa se a senha estiver correta)
# =================================================================

def normalize_id(text):
    if not isinstance(text, str): return ""
    text = "".join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
    return "".join(filter(str.isalnum, text)).lower()

# 2. MOTOR DE FORMATAÇÃO
def fmt(val, is_pct=False, is_variance_col=False):
    if pd.isna(val) or val == "-" or val == "n.a.": return "n.a."
    try:
        v_float = float(val)
        if is_pct:
            p_val = v_float * 100 if abs(v_float) <= 1.01 else v_float
            formatted_val = f"{p_val:.0f}%"
            if is_variance_col:
                color = "#dc3545" if v_float > 0.001 else "#28a745"
                return f'<span style="color: {color}; font-weight: bold; display: block; text-align: center;">{formatted_val}</span>'
            return formatted_val
        v_abs = int(round(abs(v_float)))
        formatted_num = f"{v_abs:,}".replace(",", ".")
        return f"({formatted_num})" if v_float < -0.001 else formatted_num
    except: return str(val)

def safe_float(v):
    try:
        if pd.isna(v) or str(v).strip() in ["", "-", "nan"]: return 0.0
        if isinstance(v, str): v = v.replace('.', '').replace(',', '.')
        return float(v)
    except: return 0.0

def safe_div(n, d):
    try: return n / abs(d) if abs(d) > 0.0001 else 0.0
    except: return 0.0

# =================================================================
# 3. CARREGAMENTO DE DADOS (VERSÃO OTIMIZADA PARA VELOCIDADE)
# =================================================================
import openpyxl

@st.cache_data(ttl=86400)
def load_all_v3a_data():
    file_path = "dados_dre.xlsx"
    # Abre o motor do pandas uma vez
    xls = pd.ExcelFile(file_path)
    
    def read_sheet_with_comments(sheet_name, max_cols=18):
        try:
            # ADICIONADO: read_only=True e data_only=True para não travar
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            ws = wb[sheet_name]
            data_rows = []
            for row in ws.iter_rows(max_col=max_cols):
                row_data = []
                for cell in row:
                    val = cell.value
                    # Verifica se existe comentário
                    if hasattr(cell, 'comment') and cell.comment:
                        comment_text = cell.comment.text.strip().replace('\n', ' ').replace('"', "'")
                        val = f"{val if val is not None else ''}||{comment_text}"
                    row_data.append(val)
                data_rows.append(row_data)
            
            # ADICIONADO: Fecha o arquivo explicitamente
            wb.close()
            
            df = pd.DataFrame(data_rows)
            if not df.empty:
                df.columns = df.iloc[0]
                return df[1:].reset_index(drop=True)
            return pd.DataFrame()
        except Exception as e:
            print(f"Erro ao ler comentários: {e}")
            return pd.DataFrame()

    def clean_sheet(name, skip=0, columns=None):
        try:
            df = pd.read_excel(xls, name, skiprows=skip, usecols=columns)
            cols = [c for c in df.columns if "Unnamed" not in str(c)]
            return df[cols].dropna(how='all', axis=0)
        except: 
            return pd.DataFrame()

    # --- LEITURA DE METADADOS ---
    try:
        df_att = pd.read_excel(xls, "DAT ATT", header=None)
        data_att_final = "<br>".join(df_att.iloc[:, 0].dropna().astype(str).tolist())
    except:
        data_att_final = "Data não disponível"

    try:
        df_obs = pd.read_excel(xls, "Principais Observações", usecols="A:B")
        lista_obs_orc = df_obs.iloc[:, 0].dropna().astype(str).tolist() 
        lista_obs_dre = df_obs.iloc[:, 1].dropna().astype(str).tolist() 
    except:
        lista_obs_orc, lista_obs_dre = [], []

    # --- PROCESSAMENTO BUDGET ---
    try:
        df_orc_comentarios = read_sheet_with_comments("Budget", max_cols=18)
        mask_corte = df_orc_comentarios.iloc[:, 0].astype(str).str.contains("Banco do Brasil", case=False, na=False)
        df_orc_limpo = df_orc_comentarios.iloc[:mask_corte.idxmax()].copy() if mask_corte.any() else df_orc_comentarios.copy()
    except:
        df_orc_limpo = pd.DataFrame()

    dict_res = {
        "DRE": clean_sheet("DRE Gerencial", 3),
        "DRE Gerencial": clean_sheet("DRE Gerencial", 3),
        "E26": pd.read_excel(xls, "EBITDA 2026"),
        "E25": pd.read_excel(xls, "EBITDA 2025"),
        "ORC_RAW": df_orc_limpo,
        "ORC_ANUAL": df_orc_limpo,
        "ORC_YTD_BRUTO": pd.read_excel(xls, "Budget YTD", header=None),
        "OBS_DRE": lista_obs_dre,
        "OBS_ORC": lista_obs_orc,
        "ATUALIZACAO": data_att_final 
    }

    abas_extras = {"DRE N1": "DRE N1", "DRE N2": "DRE N2", "DRE N3": "DRE N3", "DRE N4": "DRE N4", 
                   "DRE N5": "DRE N5", "DRE N6": "DRE N6", "DRE N7": "DRE N7", "DRE N8": "DRE N8", 
                   "DRE VENTURES - OUTROS": "DRE VENTURES - OUTROS"}
    
    for k, v in abas_extras.items(): 
        dict_res[k] = clean_sheet(v, 3)

    # Fecha o motor do pandas
    xls.close()
    return dict_res

data = load_all_v3a_data()

# =================================================================
# 4. DEFINIÇÃO DO ÍNDICE AUTOMÁTICO (FECHAMENTO)
# =================================================================
mapa_meses_idx = {
    "jan": 0, "fev": 1, "mar": 2, "abr": 3, "mai": 4, "jun": 5,
    "jul": 6, "ago": 7, "set": 8, "out": 9, "nov": 10, "dez": 11
}

try:
    texto_ref_fechamento = data["ATUALIZACAO"].lower()
    index_fechamento = 1 # Padrão: Fevereiro
    
    for abr_mes, valor_idx in mapa_meses_idx.items():
        if abr_mes in texto_ref_fechamento:
            index_fechamento = valor_idx
            break
except:
    index_fechamento = 1

meses_lista_full = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

# 4. ESTILO CSS GLOBAL CONSOLIDADO (VERSÃO FINAL COM AJUSTES VISUAIS)
st.markdown("""
    <style>
    :root {
        --space-lg: 32px;
        --v3a-gold: #B8860B;
        --text-dark: #1A1A1A;
        --cor-nuvem: #f0f0f0; /* Cor Nuvem solicitada */
    }
    [data-testid="stAppViewContainer"] { background-color: #F8F9FA !important; }
    
    [data-testid="stVerticalBlock"] > div { gap: 0rem !important; margin-bottom: 0px !important; }
    
    [data-testid="stHorizontalBlock"] {
        padding-top: 0px !important; /* Remove o excesso que causa a diferença */
        margin-top: 0px !important;
    }

    .section-container { margin-bottom: var(--space-lg) !important; width: 100%; display: block; }
    
    /* Cabeçalhos dos Quadros - Força alinhamento horizontal */
    .header-container { 
        display: flex !important;
        align-items: center !important; 
        padding-top: 0px !important;
        margin-top: 0px !important;
        margin-bottom: 15px !important;
    }

    /* Número Amarelo Simples ao lado do texto */
    .quadro-num { 
        background-color: transparent !important;
        color: #ffcb05 !important; 
        font-weight: bold !important; 
        font-size: 24px !important; 
        margin-right: 17px !important; /* Espaço entre número e texto */
        line-height: 1 !important;
        display: inline-block !important;
        width: auto !important;
        height: auto !important;
    }

    .quadro-titulo { 
        display: inline-block !important;
        font-weight: 700; 
        color: var(--text-dark); 
        font-size: 15px; 
        text-transform: uppercase; 
        font-family: 'Segoe UI';
        line-height: 1 !important;
    }

    /* 2) DRE GERENCIAL – LINHAS FILHAS (Branco Puro) */
    .row-child-dre { background-color: #FFFFFF !important; }

    /* 4 e 5) EBITDA, MARGEM POR ÁREA E TOP 10 (Fundo Nuvem nos Totais/Resultados) */
    .y-yellow, .row-total-ma, .row-total-top { 
        background-color: var(--cor-nuvem) !important;
    }

    /* Ajuste de Container de KPIs para Grid Responsivo */
    .kpi-container { 
        display: grid; 
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); 
        gap: 12px; 
        margin-bottom: 25px; 
        width: 100%;
    }
    .kpi-card {
        background-color: white; padding: 15px; border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border-left: 5px solid #ccc; font-family: 'Segoe UI', sans-serif;
        box-sizing: border-box;
    }

    /* ESSENCIAL: Container para Scroll Horizontal em Tabelas */
    .table-responsive-container {
        width: 100%;
        overflow-x: auto !important;
        -webkit-overflow-scrolling: touch;
        display: block;
    }

    /* Media Query para telas de Notebook Pequeno e Mobile */
    @media (max-width: 1024px) {
        .header-container { flex-direction: column !important; align-items: flex-start !important; gap: 8px; }
        .quadro-titulo { font-size: 16px !important; }
        .kpi-value { font-size: 18px !important; }
    }
    
    iframe { width: 100% !important; border: none !important; }
            
    .kpi-label { font-size: 11px; font-weight: 700; color: #888; text-transform: uppercase; margin-bottom: 4px; }
    .kpi-value { font-size: 22px; font-weight: 700; color: var(--text-dark); }
    .periodo-label { font-size: 14px; font-weight: 600; color: #666; font-family: 'Segoe UI'; margin-right: 10px; white-space: nowrap; display: flex; align-items: center; height: 38px; }
    div[data-testid="stSelectbox"] { margin-top: 0px !important; }
    div[data-testid="stSelectbox"] label { display: none !important; }
    div[data-testid="stSelectbox"] input { pointer-events: none !important; caret-color: transparent !important; }
    .stPlotlyChart { background-color: transparent !important; border-radius: 12px !important; padding: 15px !important; box-shadow: none !important; border: none !important; margin-bottom: var(--space-lg) !important; }
    iframe { border: none !important; margin: 0 !important; }
    
    /* ... todos os seus estilos anteriores (root, kpi-card, etc) ... */

    /* ADICIONE AQUI O NOVO CSS PARA APROXIMAR OS BOTÕES DO QUADRO */
    [data-testid="stVerticalBlock"] > div:has(div.header-container) + div {
        margin-top: -10px !important;
    }

    div[data-testid="stSegmentedControl"] {
        margin-bottom: 0px !important;
        padding-bottom: 0px !important;
    }

    /* Aproveite e adicione a cor amarela/dourada para os botões se ainda não tiver */
    div[data-testid="stSegmentedControl"] button[aria-checked="true"] {
        background-color: #B8860B !important;
        color: white !important;
    }
    
    /* 1. Zera o padding do container principal do Streamlit */
    .block-container {
        padding-top: 1rem !important; /* Estava em 6rem por padrão */
        padding-bottom: 0rem;
        margin-top: -20px; /* Puxa o conteúdo mais para cima ainda */
    }

    /* 2. Remove o header nativo do Streamlit que ocupa espaço vazio */
    header {
        visibility: hidden;
        height: 0px;
    }
    </style>
""", unsafe_allow_html=True)

# --- INICIALIZAÇÃO DO ESTADO DE NAVEGAÇÃO ---
if 'pagina' not in st.session_state:
    st.session_state.pagina = "DRE"

def mudar_pagina(nome_pagina):
    st.session_state.pagina = nome_pagina

# --- ESTILO ADICIONAL PARA OS BOTÕES DE NAVEGAÇÃO ---
st.markdown("""
    <style>
    div[data-testid="column"] button {
        background-color: #262626 !important;
        color: #888 !important;
        border: none !important;
        font-weight: 600 !important;
        text-transform: uppercase !important;
        border-radius: 0px !important;
        height: 45px !important;
        width: 100% !important;
        margin: 0 !important;
        border-bottom: 4px solid transparent !important;
    }
    div[data-testid="column"] button:hover {
        color: #ffcb05 !important;
    }
    </style>
""", unsafe_allow_html=True)

# ESTILO COMPARTILHADO (Injetado nos IFrames para o quadro branco não sumir)
shared_card_style = """
<style>
    body { margin: 0; padding: 0; background: transparent; font-family: 'Segoe UI', sans-serif; }
    .v3a-card {
        background-color: #FFFFFF !important; border-radius: 12px !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05) !important; border: 1px solid #E9ECEF !important;
        padding: 20px !important; margin: 5px !important; /* Respiro para sombra */
    }
    .obs-item { display: flex; margin-bottom: 10px; align-items: flex-start; line-height: 1.4; }
    .obs-number { font-weight: 700; color: #B8860B; min-width: 25px; font-size: 14px; }
    .obs-text { color: #444; font-size: 14px; text-align: left; }
</style>
"""

# --- BLOCO DO TÍTULO REVISADO (MAIS COMPACTO) ---
try:
    data = load_all_v3a_data()
    data_att = data.get("ATUALIZACAO", "Data não disponível")

    import base64
    try:
        with open("image_3.png", "rb") as f:
            encoded_logo = base64.b64encode(f.read()).decode()
        # REDUZIDO: height de 270px para 80px (ou o tamanho que preferir)
        logo_html = f'<img src="data:image/png;base64,{encoded_logo}" style="height: 150px; width: auto; object-fit: contain; margin: 0;">'
    except:
        logo_html = "" 

    # HTML Unificado do Cabeçalho - Ajustado para ocupar menos espaço vertical
    html_header = (
        f'<div style="margin: 0; padding: 0; width: 100%;">'
        f'  <div style="display: flex; align-items: center; gap: 20px; padding: 0px 0 5px 0; margin-top: -10px;">' # margin-top negativa para subir
        f'    {logo_html}'
        f'    <div style="display: flex; flex-direction: column; justify-content: center; margin: 0; padding: 0;">'
        f'      <h1 style="color: #636466; font-family: \'Segoe UI\', sans-serif; font-size: 26px; font-weight: 650; text-transform: uppercase; margin: 0; padding: 0; line-height: 1.1;">'
        f'Painel Executivo Financeiro</h1>'
        f'      <p style="color: #666666; font-family: \'Segoe UI\', sans-serif; font-size: 12px; margin: 4px 0 0 0; padding: 0; font-weight: 300;">'
        f'{data_att}</p>'
        f'    </div>'
        f'  </div>'
        f'  <hr style="margin: 5px 0 15px 0; border: 0; border-top: 1px solid #E9ECEF; padding: 0;">' # HR mais fina
        f'</div>'
    )
    
    st.markdown(html_header, unsafe_allow_html=True)

    # --- PASSO 2: INTERFACE DE NAVEGAÇÃO (BOTÕES) ---
    col_nav1, col_nav2, col_nav3, col_nav4, col_nav_esp = st.columns([1, 1, 1, 1, 4])

    with col_nav1:
        if st.button("DRE", use_container_width=True, key="btn_dre"):
            mudar_pagina("DRE")
    with col_nav2:
        if st.button("Orçamento", use_container_width=True, key="btn_orc"):
            mudar_pagina("Orçamento")
    with col_nav3:
        if st.button("Receitas", use_container_width=True, key="btn_vend"):
            mudar_pagina("Receitas")

    # Linha de destaque para simular a navegação do print
    st.markdown('<div style="margin-bottom: 30px;"></div>', unsafe_allow_html=True)

except Exception as e:
    st.error(f"Erro ao carregar cabeçalho ou dados: {e}")
    st.stop()

# --- DEFINIÇÕES GLOBAIS (FORA DE QUALQUER IF PARA NÃO DAR NAMEERROR) ---
meses_pt = {1:'Janeiro', 2:'Fevereiro', 3:'Março', 4:'Abril', 5:'Maio', 6:'Junho', 7:'Julho', 8:'Agosto', 9:'Setembro', 10:'Outubro', 11:'Novembro', 12:'Dezembro'}
meses_abr = {1:'Jan', 2:'Fev', 3:'Mar', 4:'Abr', 5:'Mai', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Set', 10:'Out', 11:'Nov', 12:'Dez'}
meses_lista_full = [meses_pt[i] for i in range(1, 13)]

def render_obs_card(df, section_name):
    try:
        if section_name not in df.columns: return ""
        mensagens = df[section_name].dropna().astype(str).tolist()
        mensagens = [m.replace('\xa0', ' ').strip() for m in mensagens if m.strip() != "" and m.lower() != "nan"]
        if not mensagens: return ""
        
        itens_html = ""
        for i, msg in enumerate(mensagens, 1):
            import re
            clean_msg = re.sub(r'<[^>]*>', '', msg)
            itens_html += f'<div class="obs-item"><div class="obs-number">{i}.</div><div class="obs-text">{clean_msg}</div></div>'
        
        return f"""{shared_card_style}<div class="v3a-card">
            <h3 style="color: #1A1A1A; margin-top: 0; font-size: 15px; font-weight: 700; text-transform: uppercase; 
                        border-bottom: 1px solid #EEE; padding-bottom: 10px; margin-bottom: 15px;">
                Principais Observações - {section_name}
            </h3>
            {itens_html}
        </div>"""
    except:
        return ""
    
#=================PRIMEIRA PAGINA - DRE==================#
#========================================================#

if st.session_state.pagina == "DRE":
    # --- KPI EXTRACTION (TOP LEVEL) ---
    df_kpi_raw = data["DRE"].copy()
    def get_accum(label_part):
        try:
            mask = df_kpi_raw.iloc[:,0].astype(str).str.contains(label_part, case=False, na=False, regex=True)
            return safe_float(df_kpi_raw[mask].iloc[0]["Acumulado"])
        except:
            return 0.0

    rb_val = get_accum("Receita Bruta")
    lb_val = get_accum("Lucro Bruto")
    mb_raw = get_accum("Margem Bruta")
    lo_val = get_accum("Lucro Operacional")
    ll_val = get_accum("Lucro Liquido")
    mb_rounded = int(round(mb_raw * 100)) if abs(mb_raw) <= 1.0 else int(round(mb_raw))

    st.markdown(f"""
        <div class="kpi-container">
            <div class="kpi-card" style="border-left-color: #B8860B;"><div class="kpi-label">Receita Bruta (Acum.)</div><div class="kpi-value">R$ {fmt(rb_val)}</div></div>
            <div class="kpi-card" style="border-left-color: #28a745;"><div class="kpi-label">Lucro Bruto (Acum.)</div><div class="kpi-value">R$ {fmt(lb_val)}</div></div>
            <div class="kpi-card" style="border-left-color: #007bff;"><div class="kpi-label">Margem Bruta</div><div class="kpi-value">{mb_rounded}%</div></div>
            <div class="kpi-card" style="border-left-color: #8B4513;"><div class="kpi-label">Lucro Operacional (Acum.)</div><div class="kpi-value">R$ {fmt(lo_val)}</div></div>
            <div class="kpi-card" style="border-left-color: #28a745;"><div class="kpi-label">Lucro Liquido (Acum.)</div><div class="kpi-value">R$ {fmt(ll_val)}</div></div>
        </div>
        """, unsafe_allow_html=True)


#====================# QUADRO 1: DRE GERENCIAL (VERSÃO OTIMIZADA MOBILE) #====================#
    
    col_dre_t, col_dre_s = st.columns([3.2, 1], gap="small", vertical_alignment="bottom")

    with col_dre_t:
        st.markdown('<div class="header-container" style="margin-top: 0px; padding-top: 0px; margin-bottom: 5px;"><div class="quadro-num">01.</div><div class="quadro-titulo">DRE Gerencial 2026</div></div>', unsafe_allow_html=True)

    with col_dre_s:
        # Mapeamento das abas
        opcoes_dre = {
            "Total Geral": "DRE Gerencial",
            "Núcleo 1": "DRE N1",
            "Núcleo 2": "DRE N2",
            "Núcleo 3": "DRE N3",
            "Núcleo 4": "DRE N4",
            "Núcleo 5": "DRE N5",
            "Núcleo 6": "DRE N6",
            "Núcleo 7": "DRE N7",
            "Núcleo 8": "DRE N8",
            "Ventures - Outros": "DRE VENTURES - OUTROS"
        }
        
        c1dre, c2dre, c3dre = st.columns([0.6, 1.2, 1.0], vertical_alignment="bottom")
        
        with c1dre:
            st.markdown('<div style="padding-bottom: 8px; text-align: right;"><span class="periodo-label" style="font-size: 13px;">Núcleo:</span></div>', unsafe_allow_html=True)
        
        with c2dre:
            selecao_label = st.selectbox(
                "Núcleo", 
                list(opcoes_dre.keys()), 
                index=0,
                key="selector_dre_global_final",
                label_visibility="collapsed"
            )
        
        aba_selecionada = opcoes_dre[selecao_label]

    # --- LÓGICA DE DADOS DA DRE ---
    df_dre = data[aba_selecionada].copy()
    
    map_abr_cols = {v: meses_abr[k] for k, v in meses_pt.items()}
    df_dre.columns = [map_abr_cols.get(str(col).strip(), str(col)) for col in df_dre.columns]

    parents_dre = ["+ Receita Bruta", "- Custo", "- Imposto", "- Despesas", "+ Outras Receitas", "- Outras Despesas", "- Investimentos"]
    subs_hierarquia = ["ondemand", "ventures"]

    # RENDERIZAÇÃO DO HTML
    html_dre = f"""
        {shared_card_style}
        <div class="v3a-card" style="margin-top: 0px; padding: 15px !important;">
            <style>
                /* CLINICA QUADRO 04: Fonte 11px e Scroll Touch */
                .v3a-table {{ width: 100%; border-collapse: collapse; color: #000; font-size: 11px; font-family: 'Segoe UI', sans-serif; }} 
                .v3a-table th {{ position: sticky; top: 0; z-index: 10; background: #1A1A1A; color: #FFF; padding: 10px 5px; text-align: center; white-space: nowrap; }} 
                .v3a-table thead th:first-child {{ text-align: center !important; }}
                
                .v3a-table td {{ padding: 8px 5px; border-bottom: 1px solid #F0F0F0; white-space: nowrap; text-align: center; }} 
                .v3a-table td:first-child {{ text-align: left; min-width: 170px; white-space: normal; font-weight: bold; }} 
                
                .row-parent {{ background-color: #FFFFFF !important; font-weight: bold; cursor: pointer; }} 
                .row-sub {{ background-color: #FFFFFF !important; font-weight: bold; cursor: pointer; display: none; }} 
                .row-child-dre {{ background-color: #FFFFFF !important; display: none; }} 
                .arrow {{ display: inline-block; width: 15px; color: #666; font-size: 10px; }}
                
                /* Ajuste de identação para mobile: Menos espaço vazio na esquerda */
                .indent-sub {{ padding-left: 20px !important; }}
                .indent-child {{ padding-left: 35px !important; font-weight: normal !important; color: #666; }}

                .responsive-scroll-dre {{ 
                    width: 100%; 
                    overflow-x: auto !important; 
                    -webkit-overflow-scrolling: touch; 
                }}
            </style>
            
            <script>
                function toggleDRE(id, type) {{ 
                    if (type === 'p1') {{ 
                        const children = document.getElementsByClassName('dre-p1-child-' + id); 
                        const arrow = document.getElementById('ad-' + id + 'p1'); 
                        if(!children.length) return;
                        const isOpening = children[0].style.display !== 'table-row'; 
                        for (let el of children) {{ el.style.display = isOpening ? 'table-row' : 'none'; }} 
                        if (!isOpening) {{ 
                            const netos = document.querySelectorAll('.dre-neto-of-p1-' + id); 
                            netos.forEach(n => n.style.display = 'none'); 
                        }} 
                        arrow.innerHTML = isOpening ? '▼' : '▶'; 
                    }} else if (type === 'p2') {{ 
                        const netos = document.getElementsByClassName('dre-p2-child-' + id); 
                        const arrow = document.getElementById('ad-' + id + 'p2'); 
                        if(!netos.length) return;
                        const isOpening = netos[0].style.display !== 'table-row'; 
                        for (let el of netos) {{ el.style.display = isOpening ? 'table-row' : 'none'; }} 
                        arrow.innerHTML = isOpening ? '▼' : '▶'; 
                    }} 
                }}
            </script>
            
            <div class="responsive-scroll-dre">
                <table class="v3a-table">
                    <thead><tr>"""

    for col in df_dre.columns: 
        html_dre += f"<th>{col}</th>"
    html_dre += "</tr></thead><tbody>"

    cp1, cp2, p_ctx_h = 0, 0, False
    for i, row in df_dre.iterrows():
        desc = str(row.iloc[0]).strip()
        dn = normalize_id(desc)
        is_pct = "%" in desc or "Margem" in desc
        is_result = desc.startswith('=')
        
        if any(p in desc for p in parents_dre):
            cp1 += 1
            p_ctx_h = any(x in dn for x in ["receitabruta", "custo", "imposto"])
            html_dre += f'<tr class="row-parent" onclick="toggleDRE({cp1}, \'p1\')"><td><span id="ad-{cp1}p1" class="arrow">▶</span> {desc}</td>'
        elif dn in subs_hierarquia and p_ctx_h:
            cp2 += 1
            html_dre += f'<tr class="row-sub dre-p1-child-{cp1} dre-neto-of-p1-{cp1}" onclick="toggleDRE({cp2}, \'p2\')"><td class="indent-sub"><span id="ad-{cp2}p2" class="arrow">▶</span> {desc}</td>'
        elif is_result:
            html_dre += f'<tr style="background-color: #f0f0f0 !important; font-weight: bold;"><td>{desc}</td>'
        elif is_pct:
            html_dre += f'<tr style="background-color: #FFFFFF !important; font-style: italic; color: #666;"><td>{desc}</td>'
        else:
            cl_cls = f"dre-p2-child-{cp2} dre-neto-of-p1-{cp1}" if p_ctx_h else f"dre-p1-child-{cp1}"
            html_dre += f'<tr class="row-child-dre {cl_cls}"><td class="indent-child">{desc}</td>'
        
        for v in row[1:]:
            html_dre += f'<td>{fmt(v, is_pct=is_pct)}</td>'
        html_dre += '</tr>'

    html_dre += """
                    </tbody>
                </table>
            </div> 
        </div> """

    st.components.v1.html(html_dre, height=550, scrolling=False)


#====================# QUADRO 2: MARGEM BRUTA POR ÁREA 2026 (VERSÃO MOBILE OK) #====================#
    
    st.markdown('<div style="padding-top: 10px;"></div>', unsafe_allow_html=True) 
    st.markdown("""
        <div class="header-container" style="margin-top: 0px; margin-bottom: 15px; padding: 0;">
            <div class="quadro-num">02.</div>
            <div class="quadro-titulo">Margem Bruta por Área 2026</div>
        </div>
    """, unsafe_allow_html=True)

    try:
        # 1. Carregamento e Normalização (Mantendo sua lógica original)
        df_ma_origem = pd.read_excel("dados_dre.xlsx", sheet_name="Margem por Faixa", skiprows=18)
        df_ma_origem = df_ma_origem.loc[:, ~df_ma_origem.columns.str.contains('^Unnamed')]

        col_map_ma = {normalize_id(c): c for c in df_ma_origem.columns}
        c_nuc, c_cli = col_map_ma.get("nucleo"), col_map_ma.get("cliente")
        c_rec, c_imp, c_cus = col_map_ma.get("receitabruta"), col_map_ma.get("imposto"), col_map_ma.get("custo")

        df_ma_origem[c_nuc] = df_ma_origem[c_nuc].astype(str).str.strip()
        df_ma_origem[c_cli] = df_ma_origem[c_cli].astype(str).str.strip()
        df_ma_origem['nuc_norm'] = df_ma_origem[c_nuc].apply(normalize_id)
        
        for col_f in [c_rec, c_imp, c_cus]:
            df_ma_origem[col_f] = df_ma_origem[col_f].apply(safe_float)

        df_ma_origem['lb_row'] = df_ma_origem[c_rec] - df_ma_origem[c_cus].abs() - df_ma_origem[c_imp].abs()
        
        nucleos_lista = ["Núcleo 1", "Núcleo 2", "Núcleo 3", "Núcleo 4", "Núcleo 5", "Núcleo 6", "Núcleo 7", "Núcleo 8", "Ventures - outros"]
        dados_hierarquia = {}

        for nuc in nucleos_lista:
            n_norm = normalize_id(nuc)
            df_nuc = df_ma_origem[df_ma_origem['nuc_norm'] == n_norm].copy()
            df_cli_group = df_nuc.groupby(c_cli).agg({
                c_rec: 'sum', c_imp: 'sum', c_cus: 'sum', 'lb_row': 'sum'
            }).reset_index()
            df_cli_group = df_cli_group.sort_values(by=[c_rec, 'lb_row'], ascending=False)

            dados_hierarquia[nuc] = {
                "receita": df_nuc[c_rec].sum(),
                "imposto": df_nuc[c_imp].sum(),
                "custo": df_nuc[c_cus].sum(),
                "lb": df_nuc['lb_row'].sum(),
                "clientes": df_cli_group.to_dict('records')
            }

        # 3. HTML / CSS / JS (APLICANDO A CLÍNICA DO QUADRO 04)
        html_ma = f"""
        <div style="background-color: white; padding: 15px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #E9ECEF; margin-bottom: 40px;">
            <style>
                /* CLINICA QUADRO 04: Fonte 11px e Scroll Touch */
                .ma-table {{ width: 100%; border-collapse: collapse; color: #000; font-size: 11px; font-family: 'Segoe UI', sans-serif; }}
                .ma-table thead th {{ position: sticky; top: 0; z-index: 10; background: #1A1A1A; color: #FFF; padding: 10px 5px; text-align: center; white-space: nowrap; }}
                
                /* Controle de largura da primeira coluna */
                .ma-table td:first-child {{ text-align: left; width: 160px; font-weight: bold; white-space: normal; }}
                .ma-table td {{ padding: 8px 5px; border-bottom: 1px solid #F0F0F0; white-space: nowrap; text-align: center; }}
                
                .row-cli-ma {{ background-color: #FFFFFF !important; display: none; font-size: 10px; color: #666; }}
                .row-total-final-ma {{ background-color: #f0f0f0 !important; font-weight: bold; color: #000 !important; }}
                .arr-ma {{ display: inline-block; width: 15px; color: #666; font-size: 10px; transition: 0.2s; }}
                .indent-cli {{ padding-left: 20px !important; font-weight: normal; font-style: italic; }}

                /* DIV DE SCROLL IGUAL AO QUADRO 04 */
                .responsive-scroll-ma {{ 
                    width: 100%; 
                    overflow-x: auto !important; 
                    -webkit-overflow-scrolling: touch; 
                    display: block;
                }}
            </style>
            
            <script>
                function toggleMa(id) {{
                    let rows = document.getElementsByClassName('cli-of-' + id);
                    let arrow = document.getElementById('icon-' + id);
                    if (rows.length > 0) {{
                        let isOpening = rows[0].style.display === 'none' || rows[0].style.display === '';
                        for (let i = 0; i < rows.length; i++) {{
                            rows[i].style.display = isOpening ? 'table-row' : 'none';
                        }}
                        arrow.innerHTML = isOpening ? '▼' : '▶';
                    }}
                }}
            </script>

            <div class="responsive-scroll-ma">
                <table class="ma-table">
                    <thead>
                        <tr>
                            <th>Unidade / Cliente</th>
                            <th>Rec. Bruta</th>
                            <th>Custo</th>
                            <th>Imposto</th>
                            <th>Lucro Bruto</th>
                            <th>Mg (%)</th>
                        </tr>
                    </thead>
                    <tbody>"""

        tot_r, tot_c, tot_i, tot_lb = 0.0, 0.0, 0.0, 0.0

        for i, nuc_name in enumerate(nucleos_lista):
            info = dados_hierarquia[nuc_name]
            mg_n = safe_div(info["lb"], info["receita"])
            tot_r += info["receita"]; tot_c += info["custo"]; tot_i += info["imposto"]; tot_lb += info["lb"]

            html_ma += f"""
            <tr style="background-color: #FFFFFF; cursor: pointer;" onclick="toggleMa({i})">
                <td><span id="icon-{i}" class="arr-ma">▶</span> {nuc_name}</td>
                <td>{fmt(info["receita"])}</td>
                <td>{fmt(info["custo"])}</td>
                <td>{fmt(info["imposto"])}</td>
                <td>{fmt(info["lb"])}</td>
                <td>{fmt(mg_n, True)}</td>
            </tr>"""

            for cli in info["clientes"]:
                if abs(cli[c_rec]) > 0.01 or abs(cli[c_cus]) > 0.01 or abs(cli[c_imp]) > 0.01:
                    mg_c = safe_div(cli['lb_row'], cli[c_rec])
                    html_ma += f"""
                    <tr class="row-cli-ma cli-of-{i}">
                        <td class="indent-cli">{cli[c_cli]}</td>
                        <td>{fmt(cli[c_rec])}</td>
                        <td>{fmt(cli[c_cus])}</td>
                        <td>{fmt(cli[c_imp])}</td>
                        <td>{fmt(cli['lb_row'])}</td>
                        <td>{fmt(mg_c, True)}</td>
                    </tr>"""

        tot_mg = safe_div(tot_lb, tot_r)
        html_ma += f"""
                    <tr class="row-total-final-ma">
                        <td>TOTAL GERAL</td>
                        <td>{fmt(tot_r)}</td>
                        <td>{fmt(tot_c)}</td>
                        <td>{fmt(tot_i)}</td>
                        <td>{fmt(tot_lb)}</td>
                        <td>{fmt(tot_mg, True)}</td>
                    </tr>
                </tbody>
            </table>
            </div>
        </div>"""

        st.components.v1.html(html_ma, height=500, scrolling=False)

    except Exception as e:
        st.error(f"Erro ao processar Margem por Área: {e}")

#====================# QUADRO 03: VISÃO EBITDA YOY 2026 x 2025 (AJUSTE DE LARGURA) #====================#
    
    col_eb_t, col_eb_s = st.columns([3.2, 1], gap="small", vertical_alignment="bottom")
        
    with col_eb_t: 
        st.markdown("""
            <div class="header-container" style="margin-top: 0px; margin-bottom: 10px; padding: 0;">
                <div class="quadro-num">03.</div>
                <div class="quadro-titulo">VISÃO EBITDA YOY 2026 x 2025</div>
            </div>
        """, unsafe_allow_html=True)
        
    with col_eb_s:
        c1eb, c2eb, c3eb = st.columns([0.6, 1.2, 1.0], vertical_alignment="bottom")
        with c1eb:
            st.markdown('<div style="padding-bottom: 8px; text-align: right;"><span class="periodo-label">Período:</span></div>', unsafe_allow_html=True)
        with c2eb:
            mes_eb = st.selectbox(
                "Período", 
                meses_lista_full, 
                index=index_fechamento, 
                key=f"sel_ebitda_dinamico_{index_fechamento}", 
                label_visibility="collapsed"
            )

    idx_eb = meses_lista_full.index(mes_eb)
    c_idx, e26, e25 = 7 + idx_eb, data["E26"], data["E25"]
    t_parents = ["+ Receita Bruta", "- Imposto IBS CBS (ISS PIS COFINS)", "= Receita Liquida", "- Custo", "= Lucro Bruto", "% Margem Bruta (sem IR e CSLL)", "% Sobre Receita Liquida (sem IR / CSLL)", "- Despesas", "= Ebitda", "% Sobre a Receita Liquida", "- Imposto IR CSLL", "+ - Outras Receitas E Despesas", "- Investimentos", "= Lucro Liquido", "% S/ Rec Liq"]

    html_yoy = f"""
    <div style="background-color: #FFFFFF; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #E9ECEF; padding: 15px; margin-bottom: 0px;">
        <style>
            .y-table {{ width: 100%; border-collapse: collapse; color: #000; font-size: 11px; font-family: 'Segoe UI', sans-serif; table-layout: fixed; }} 
            .y-table thead tr th {{ background: #1A1A1A; color: #FFF; padding: 8px 2px; text-align: center; white-space: nowrap; }} 
            
            /* AJUSTE: Aumento da largura e centralização do título 'Categorias' */
            .y-table thead th:first-child {{ 
                text-align: center !important; 
                width: 220px; 
                white-space: normal; 
            }} 
            
            .y-table td {{ padding: 8px 4px; border-bottom: 1px solid #F0F0F0; white-space: nowrap; text-align: center; }} 
            
            /* AJUSTE: Largura mínima maior para evitar quebra de linha excessiva */
            .y-table td:first-child {{ 
                text-align: left; 
                font-weight: bold; 
                white-space: normal; 
                min-width: 210px; 
                padding-left: 10px;
            }} 
            
            .y-yellow {{ background-color: #f0f0f0 !important; font-weight: bold; }} 
            .y-pct-clean {{ font-weight: normal !important; font-style: italic !important; }} 
            .h-grp {{ text-align: center !important; font-weight: bold; border-bottom: 2px solid #FFF !important; font-size: 10px; }}
            
            .responsive-scroll-yoy {{ 
                width: 100%; 
                overflow-x: auto !important; 
                -webkit-overflow-scrolling: touch; 
                display: block;
            }}
        </style>
        
        <div class="responsive-scroll-yoy">
            <table class="y-table">
                <thead>
                    <tr>
                        <th rowspan="2">Categorias</th>
                        <th colspan="4" class="h-grp" style="background:#333;">MÊS {mes_eb.upper()}</th>
                        <th colspan="4" class="h-grp" style="background:#444;">ACUM. {mes_eb.upper()}</th>
                        <th colspan="4" class="h-grp" style="background:#555;">TOTAL ANO</th>
                    </tr>
                    <tr>
                        <th style="text-align: center;">26</th><th>25</th><th>$</th><th>%</th>
                        <th>26</th><th>25</th><th>$</th><th>%</th>
                        <th>26</th><th>25</th><th>$</th><th>%</th>
                    </tr>
                </thead>
                <tbody>"""

    for target_eb in t_parents:
        try:
            r26 = e26[e26.iloc[:,1].astype(str).str.strip() == target_eb].iloc[0]
            r25 = e25[e25.iloc[:,1].astype(str).str.strip() == target_eb].iloc[0]
            nm_eb = str(r26.iloc[1]).strip()
            is_p_eb = "%" in nm_eb or "Margem" in nm_eb
            f26, f25 = safe_float(r26.iloc[c_idx]), safe_float(r25.iloc[c_idx])
            a26 = sum([safe_float(v_val) for v_val in r26.iloc[7:c_idx+1]])
            a25 = sum([safe_float(v_val) for v_val in r25.iloc[7:c_idx+1]])
            t26, t25 = safe_float(r26.iloc[5]), safe_float(r25.iloc[5])
            
            if is_p_eb:
                vfn, vfp, van, vap, vtn, vtp = "n.a.", "n.a.", "n.a.", "n.a.", "n.a.", "n.a."
            else:
                vfn, vfp = fmt(f26-f25), fmt(safe_div(f26-f25, f25), True, False)
                van, vap = fmt(a26-a25), fmt(safe_div(a26-a25, a25), True, False)
                vtn, vtp = fmt(t26-t25), fmt(safe_div(t26-t25, t25), True, False)
            
            cls_eb = "y-yellow" if nm_eb.startswith('=') else ""
            if is_p_eb: cls_eb += " y-pct-clean"
            
            html_yoy += f"""<tr class="{cls_eb}"><td>{nm_eb}</td><td>{fmt(f26, is_p_eb)}</td><td>{fmt(f25, is_p_eb)}</td><td>{vfn}</td><td>{vfp}</td><td>{fmt(a26, is_p_eb)}</td><td>{fmt(a25, is_p_eb)}</td><td>{van}</td><td>{vap}</td><td>{fmt(t26, is_p_eb)}</td><td>{fmt(t25, is_p_eb)}</td><td>{vtn}</td><td>{vtp}</td></tr>"""
        except:
            continue

    html_yoy += """
                </tbody>
            </table>
        </div> 
    </div>
    """

    st.components.v1.html(html_yoy, height=630, scrolling=False)
    
    

    # --- POSIÇÃO 1: OBSERVAÇÕES DRE GERENCIAL (LÊ COLUNA B) ---
    st.markdown('<div style="margin-top: 40px;">' + render_obs_card(pd.DataFrame({"DRE Gerencial": data["OBS_DRE"]}), "DRE Gerencial") + '</div>', unsafe_allow_html=True)

    st.write("") # Apenas para fechar o bloco se estiver vazio


#=================SEGUNDA PAGINA - ORÇAMENTO==================#
#=============================================================#


elif st.session_state.pagina == "Orçamento":
    
#====================# KPI CARDS (LEITURA DIRETA BUDGET YTD) #====================#
    
    # Usamos o DataFrame bruto (sem header) para facilitar a localização das células
    df_ytd = data["ORC_YTD_BRUTO"]

    try:
        # No Pandas com header=None:
        # Linha 6 do Excel = Índice 5
        # Linha 2 do Excel = Índice 1
        # Colunas: K=10, O=14, P=15, Q=16, R=17
        
        k_orcado_ano  = abs(safe_float(df_ytd.iloc[5, 14])) # Célula O6
        k_gasto_prev  = abs(safe_float(df_ytd.iloc[5, 15])) # Célula P6
        k_var_rs_orc  = safe_float(df_ytd.iloc[5, 16])      # Célula Q6
        k_var_pct_orc = safe_float(df_ytd.iloc[5, 17])      # Célula R6
        
        # Índice de Despesas Operacionais na Célula K2
        k_indice_desp = safe_float(df_ytd.iloc[1, 10])      # Célula K2
        
    except Exception as e:
        # Se a aba mudar ou a célula estiver vazia, zera para não quebrar o site
        k_orcado_ano = k_gasto_prev = k_var_rs_orc = k_var_pct_orc = k_indice_desp = 0.0

    # Cor da borda para as Variações: Verde se positivo (economia), Vermelho se negativo (estouro)
    cor_var = "#28a745" if k_var_rs_orc >= 0 else "#dc3545"

    st.markdown(f"""
        <div class="kpi-container" style="margin-bottom: 20px;">
            <div class="kpi-card" style="border-left-color: #B8860B;">
                <div class="kpi-label">TOTAL ORÇADO (ANO)</div>
                <div class="kpi-value">R$ {fmt(k_orcado_ano)}</div>
            </div>
            <div class="kpi-card" style="border-left-color: #B8860B;">
                <div class="kpi-label">TOTAL GASTO + PREV (ANO)</div>
                <div class="kpi-value">R$ {fmt(k_gasto_prev)}</div>
            </div>
            <div class="kpi-card" style="border-left-color: {cor_var};">
                <div class="kpi-label">VARIAÇÃO R$ (ANO)</div>
                <div class="kpi-value">R$ {fmt(k_var_rs_orc)}</div>
            </div>
            <div class="kpi-card" style="border-left-color: {cor_var};">
                <div class="kpi-label">VARIAÇÃO % (ANO)</div>
                <div class="kpi-value">{fmt(k_var_pct_orc, is_pct=True)}</div>
            </div>
            <div class="kpi-card" style="border-left-color: #1A1A1A;">
                <div class="kpi-label">ÍNDICE DESP. OPERAC. (ANO)</div>
                <div class="kpi-value">{fmt(k_indice_desp, is_pct=True)}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)


#====================# QUADRO 1: ORÇAMENTO YTD - ORÇADO X REALIZADO (LÓGICA CORRIGIDA) #====================#
    
    st.markdown('<div style="padding-top: 10px;"></div>', unsafe_allow_html=True) 
    col_y_t, col_y_s = st.columns([3.2, 1], vertical_alignment="bottom") 

    with col_y_t:
        st.markdown('<div class="header-container" style="margin-top: 0px; padding-top: 0px; margin-bottom: 5px;"><div class="quadro-num">01.</div><div class="quadro-titulo">Orçamento YTD - Orçado x Realizado</div></div>', unsafe_allow_html=True)

    with col_y_s:
        c_label_orc, c_select_orc, c_spacer = st.columns([0.6, 1.2, 1.0], vertical_alignment="bottom")
        with c_label_orc:
            st.markdown('<div style="padding-bottom: 8px; text-align: right;"><span class="periodo-label">Período:</span></div>', unsafe_allow_html=True)
        with c_select_orc:
            mes_ytd = st.selectbox(
                "", 
                meses_lista_full, 
                index=index_fechamento, 
                key=f"sel_ytd_dinamico_{index_fechamento}", 
                label_visibility="collapsed"
            )
            
    idx_y = meses_lista_full.index(mes_ytd)
    df_raw_orc_data, r_idx_ytd_orc = data["ORC_RAW"], 6 + idx_y
    t_om, t_rm, t_oa, t_ra, t_ot, t_rt = 0.0, 0.0, 0.0, 0.0, 0.0, 0.0
    p_orc_l_names = ["genteegestao", "financeiro", "diretoria", "comunicacao", "comercial", "operacoes", "ventures", "ondemand"]

    html_bytd = f"""
    <div style="background-color: white; padding: 15px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #E9ECEF; margin-bottom: 40px;">
        <style>
            .b-table {{ width: 100%; border-collapse: collapse; color: #000; font-size: 11px; font-family: 'Segoe UI', sans-serif; table-layout: fixed; }} 
            .b-table thead th {{ position: sticky; top: 0; z-index: 10; background: #1A1A1A; color: #FFF; padding: 8px 2px; text-align: center; white-space: nowrap; }} 
            .b-table thead th:first-child {{ text-align: left !important; width: 160px; white-space: normal; }} 
            .b-table td {{ padding: 8px 4px; border-bottom: 1px solid #F0F0F0; white-space: nowrap; text-align: center; }} 
            .b-table td:first-child {{ text-align: left; font-weight: bold; white-space: normal; min-width: 150px; }} 
            .row-pai-estilo {{ background-color: #FFFFFF !important; }} 
            .row-total-final {{ background-color: #f0f0f0 !important; font-weight: bold; color: #000 !important; }} 
            .h-grp {{ text-align: center !important; font-weight: bold; border-bottom: 2px solid #FFF !important; font-size: 10px; }}
            .responsive-scroll-bytd {{ width: 100%; overflow-x: auto !important; -webkit-overflow-scrolling: touch; display: block; }}
        </style>
        
        <div class="responsive-scroll-bytd">
            <table class="b-table">
                <thead>
                    <tr>
                        <th rowspan="2">Centro de Custo</th>
                        <th colspan="4" class="h-grp" style="background:#333;">MÊS {mes_ytd[:3].upper()}</th>
                        <th colspan="4" class="h-grp" style="background:#444;">ACUM. {mes_ytd[:3].upper()}</th>
                        <th colspan="4" class="h-grp" style="background:#555;">TOTAL ANO</th>
                    </tr>
                    <tr>
                        <th>Orç.</th><th>Real.</th><th>$</th><th>%</th>
                        <th>Orç.</th><th>Real.</th><th>$</th><th>%</th>
                        <th>Orç.</th><th>Real.</th><th>$</th><th>%</th>
                    </tr>
                </thead>
                <tbody>"""

    for _, row_orc_ytd in df_raw_orc_data.iterrows():
        lab_orc = str(row_orc_ytd.iloc[0]).strip()
        if not lab_orc or lab_orc in ["nan", "Orçado", "Gasto"]: continue
        if normalize_id(lab_orc) in p_orc_l_names:
            # 1. VALORES BASE
            o_m = safe_float(row_orc_ytd.iloc[2])
            r_m = safe_float(row_orc_ytd.iloc[r_idx_ytd_orc])
            o_a = o_m * (idx_y + 1)
            r_a = sum([safe_float(v) for v in row_orc_ytd.iloc[6 : r_idx_ytd_orc+1]])
            o_t = safe_float(row_orc_ytd.iloc[1])
            r_t = safe_float(row_orc_ytd.iloc[3])
            
            t_om += o_m; t_rm += r_m; t_oa += o_a; t_ra += r_a; t_ot += o_t; t_rt += r_t
            
            # 2. LÓGICA DE VARIAÇÃO % (IGUAL DRE/EBITDA): (R - O) / O
            # Como é DESPESA, multiplicamos por -1 no fmt para que ECONOMIA (Real < Orçado) fique VERDE
            v_m_p = safe_div(r_m - o_m, o_m)
            v_a_p = safe_div(r_a - o_a, o_a)
            v_t_p = safe_div(r_t - o_t, o_t)
            
            html_bytd += f"""<tr class="row-pai-estilo">
                <td>{lab_orc}</td>
                <td>{fmt(o_m)}</td><td>{fmt(r_m)}</td><td>{fmt(o_m-r_m)}</td><td>{fmt(v_m_p * -1, True, True)}</td>
                <td>{fmt(o_a)}</td><td>{fmt(r_a)}</td><td>{fmt(o_a-r_a)}</td><td>{fmt(v_a_p * -1, True, True)}</td>
                <td>{fmt(o_t)}</td><td>{fmt(r_t)}</td><td>{fmt(o_t-r_t)}</td><td>{fmt(v_t_p * -1, True, True)}</td>
            </tr>"""

    # 3. TOTAIS GERAIS
    gt_vmp = safe_div(t_rm - t_om, t_om)
    gt_vap = safe_div(t_ra - t_oa, t_oa)
    gt_vtp = safe_div(t_rt - t_ot, t_ot)
    
    html_bytd += f"""<tr class="row-total-final">
        <td>TOTAL GERAL</td>
        <td>{fmt(t_om)}</td><td>{fmt(t_rm)}</td><td>{fmt(t_om-t_rm)}</td><td>{fmt(gt_vmp * -1, True, True)}</td>
        <td>{fmt(t_oa)}</td><td>{fmt(t_ra)}</td><td>{fmt(t_oa-t_ra)}</td><td>{fmt(gt_vap * -1, True, True)}</td>
        <td>{fmt(t_ot)}</td><td>{fmt(t_rt)}</td><td>{fmt(t_ot-t_rt)}</td><td>{fmt(gt_vtp * -1, True, True)}</td>
    </tr></tbody></table></div></div>"""
    
    st.components.v1.html(html_bytd, height=500, scrolling=False)
    
    
    

#====================# QUADRO 2: ORÇAMENTO ANUAL (VERSÃO MOBILE OK + TOOLTIPS) #====================#
    
    st.markdown('<div class="header-container" style="margin-top: 20px;"><div class="quadro-num">02.</div><div class="quadro-titulo">Orçamento Anual</div></div>', unsafe_allow_html=True)
    
    df_o_a_orc = data["ORC_ANUAL"].copy()
    
    # Ajuste de nomes de colunas para serem curtos (Mobile First)
    df_o_a_orc.columns = [meses_abr.get(pd.to_datetime(c_val).month, str(c_val)).upper() if isinstance(c_val, datetime) else str(c_val) for c_val in df_o_a_orc.columns]
        
    html_orc_final = f"""
        <div style="background-color: white; padding: 15px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #E9ECEF; margin-bottom: 40px;">
            <style>
                /* CLINICA QUADRO 04: Fonte 11px e Scroll Touch */
                .v3a-table-o {{ width: 100%; border-collapse: collapse; color: #000; font-size: 11px; font-family: 'Segoe UI', sans-serif; }} 
                .v3a-table-o thead th {{ position: sticky; top: 0; z-index: 10; background: #1A1A1A; color: #FFF; padding: 10px 4px; text-align: center; white-space: nowrap; }} 
                
                /* Controle de largura da primeira coluna para não espremer no mobile */
                .v3a-table-o td:first-child {{ text-align: left; width: 180px; font-weight: bold; white-space: normal; }} 
                .v3a-table-o td {{ padding: 8px 4px; border-bottom: 1px solid #F0F0F0; white-space: nowrap; text-align: center; }} 
                
                .row-p1-orc {{ background-color: #FFFFFF !important; font-weight: bold; cursor: pointer; }} 
                .row-p2-orc {{ background-color: #FFFFFF !important; font-weight: bold; cursor: pointer; display: none; }} 
                .row-child-o-orc {{ background-color: #FFFFFF !important; display: none; }} 
                .arrow-o-orc {{ display: inline-block; width: 15px; color: #B8860B; font-size: 10px; }}
                
                /* Indentação Mobile First (Reduzida) */
                .indent-orc-p2 {{ padding-left: 20px !important; font-weight: normal !important; }}
                .indent-orc-neta {{ padding-left: 35px !important; font-weight: normal !important; font-style: italic; color: #666; font-size: 10px; }}
                
                .has-tooltip {{ border-bottom: 1px dotted #B8860B; cursor: help; display: inline-block; }}

                /* DIV DE SCROLL TOUCH SUAVE */
                .responsive-scroll-anual {{ 
                    width: 100%; 
                    overflow-x: auto !important; 
                    -webkit-overflow-scrolling: touch; 
                    display: block;
                }}
            </style>
            
            <script>
                function toggleOrc(id_o, type_o) {{ 
                    let t_o = (type_o === 'p1') ? 'p1-child-of-' + id_o : 'p2-child-of-' + id_o; 
                    let els_o = document.getElementsByClassName(t_o); 
                    if (els_o.length===0) return; 
                    let isOp_o = els_o[0].style.display !== 'table-row'; 
                    for (let el_o of els_o) {{ el_o.style.display = isOp_o ? 'table-row' : 'none'; }} 
                    if (!isOp_o && type_o === 'p1') {{ 
                        let netos_o = document.querySelectorAll('.neto-of-p1-' + id_o); 
                        netos_o.forEach(n_o => n_o.style.display = 'none'); 
                        let arrows_o = document.querySelectorAll('[id^="ao-' + id_o + '"]');
                        arrows_o.forEach(a => a.innerHTML = '▶');
                    }} 
                    document.getElementById('ao-' + id_o + type_o).innerHTML = isOp_o ? '▼' : '▶'; 
                }}
            </script>
            
            <div class="responsive-scroll-anual">
                <table class="v3a-table-o">
                    <thead><tr>"""
        
    for col_o in df_o_a_orc.columns: 
        html_orc_final += f"<th>{col_o}</th>"
        
    html_orc_final += "</tr></thead><tbody>"
        
    cp1o_orc, cp2o_orc = 0, 0
    for _, row_o_vals in df_o_a_orc.iterrows():
        dr_orc = str(row_o_vals.iloc[0]).strip()
        dn_orc = normalize_id(dr_orc)
        
        if dn_orc in p_orc_l_names:
            cp1o_orc += 1
            html_orc_final += f'<tr class="row-p1-orc" onclick="toggleOrc({cp1o_orc}, \'p1\')"><td><span id="ao-{cp1o_orc}p1" class="arrow-o-orc">▶</span> {dr_orc}</td>'
        elif any(x_match in dn_orc for x_match in ["despesas", "honorario", "prospeccoes", "concorrencia"]):
            cp2o_orc += 1
            html_orc_final += f'<tr class="row-p2-orc p1-child-of-{cp1o_orc} neto-of-p1-{cp1o_orc}" onclick="toggleOrc({cp2o_orc}, \'p2\')"><td class="indent-orc-p2"><span id="ao-{cp2o_orc}p2" class="arrow-o-orc">▶</span> {dr_orc}</td>'
        else:
            html_orc_final += f'<tr class="row-child-o-orc p2-child-of-{cp2o_orc} neto-of-p1-{cp1o_orc}"><td class="indent-orc-neta">{dr_orc}</td>'
        
        for j_o, v_o_cell in enumerate(row_o_vals[1:]):
            is_pct_col_o = (j_o == 4) 
            tooltip_txt = ""
            valor_para_formatar = v_o_cell
            
            if isinstance(v_o_cell, str) and "||" in v_o_cell:
                partes = v_o_cell.split("||")
                valor_para_formatar = partes[0] if partes[0] != "" else 0
                tooltip_txt = partes[1]

            val_formatado = fmt(safe_float(valor_para_formatar) * -1, is_pct=True, is_variance_col=True) if is_pct_col_o else fmt(valor_para_formatar)
            
            if tooltip_txt:
                html_orc_final += f'<td title="{tooltip_txt}"><span class="has-tooltip">{val_formatado} 💬</span></td>'
            else:
                html_orc_final += f'<td>{val_formatado}</td>'
        
        html_orc_final += '</tr>'
            
    html_orc_final += "</tbody></table></div></div>"
    
    st.components.v1.html(html_orc_final, height=500, scrolling=False)
    
    
    #====================# QUADRO 3: DASHBOARD ORÇADO X REALIZADO #====================#
    
    st.markdown('<div class="header-container"><div class="quadro-num">03.</div><div class="quadro-titulo">Dashboard Orçado x Realizado - Anual 2026</div></div>', unsafe_allow_html=True)
    cc_labels_orc, cc_orcado_orc, cc_realizado_orc = [], [], []
    for _, row_dash in df_raw_orc_data.iterrows():
        lab_dash = str(row_dash.iloc[0]).strip()
        if normalize_id(lab_dash) in p_orc_l_names:
            cc_labels_orc.append(lab_dash); cc_orcado_orc.append(abs(safe_float(row_dash.iloc[1]))); cc_realizado_orc.append(abs(safe_float(row_dash.iloc[3])))
    fig_dash = go.Figure()
    fig_dash.add_trace(go.Bar(x=cc_labels_orc, y=cc_orcado_orc, name='Orçado Anual', marker_color='#FDE085'))
    fig_dash.add_trace(go.Bar(x=cc_labels_orc, y=cc_realizado_orc, name='Realizado Anual', marker_color='#FFCB05'))
    fig_dash.update_layout(
        template="simple_white", barmode='group', height=450, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        title={'text': "Orçado Anual vs Realizado Anual por Centro de Custo - 2026", 'x': 0.5, 'xanchor': 'center'},
        yaxis=dict(showgrid=True, gridcolor='#EEE', tickprefix="R$ "), xaxis=dict(showgrid=True, gridcolor='#EEE'), font=dict(family="Segoe UI"),
        margin=dict(l=20, r=20, t=80, b=20)
        )
    st.plotly_chart(fig_dash, use_container_width=True, config={'displayModeBar': False})

    # --- POSIÇÃO 2: OBSERVAÇÕES ORÇAMENTO ---
    st.markdown('<div style="margin-top: 40px;">' + render_obs_card(pd.DataFrame({"Orçamento": data["OBS_ORC"]}), "Orçamento") + '</div>', unsafe_allow_html=True)

    st.write("") 


#================= TERCEIRA PAGINA - RECEITAS =====================#
elif st.session_state.pagina == "Receitas":
    
    #====================# 01. KPIs DE RECEITAS (LAYOUT ORÇAMENTO REVISADO) #====================#
    try:
        # Leitura direta da aba correta no Excel
        df_receitas = pd.read_excel("dados_dre.xlsx", sheet_name="Margem por Faixa", skiprows=18)
        df_receitas = df_receitas.loc[:, ~df_receitas.columns.str.contains('^Unnamed')]
        
        # Mapeamento de colunas por normalização (para evitar erro de maiúsculas/acentos)
        col_map_v = {normalize_id(c): c for c in df_receitas.columns}
        c_nuc = col_map_v.get("nucleo")
        c_rec = col_map_v.get("receitabruta")
        c_cus = col_map_v.get("custo")
        c_imp = col_map_v.get("imposto")
        c_job = col_map_v.get("numerojob") # Usando o número do Job para contagem única
        c_proj = col_map_v.get("nomeprojeto") or col_map_v.get("projeto")

        # Limpeza e conversão para float
        for c in [c_rec, c_cus, c_imp]:
            if c: df_receitas[c] = df_receitas[c].apply(safe_float)

        # Definição dos IDs dos Núcleos para Segmentação
        n_ondemand = [normalize_id(x) for x in ["Núcleo 1", "Núcleo 2", "Núcleo 3", "Núcleo 4", "Núcleo 5", "Núcleo 6"]]
        n_ventures = [normalize_id(x) for x in ["Núcleo 7", "Núcleo 8", "Ventures - outros", "Ventures"]]

        # --- CÁLCULOS DOS VALORES ---
        # 1. Receitas por Segmento
        df_receitas['n_id'] = df_receitas[c_nuc].apply(normalize_id)
        rec_ondemand = df_receitas[df_receitas['n_id'].isin(n_ondemand)][c_rec].sum()
        rec_ventures = df_receitas[df_receitas['n_id'].isin(n_ventures)][c_rec].sum()
        
        # 2. Projetos Aprovados (Contagem Única de Jobs)
        total_projetos_unicos = df_receitas[c_job].nunique() if c_job else 0
        
        # 3. Margem Bruta Total %
        rec_total = df_receitas[c_rec].sum()
        custo_abs = df_receitas[c_cus].abs().sum()
        imp_abs = df_receitas[c_imp].abs().sum()
        lucro_bruto = rec_total - custo_abs - imp_abs
        mg_total_pct = safe_div(lucro_bruto, rec_total)
        
        # 4. Ticket Médio REVISADO (Receita Total / Projetos Únicos)
        ticket_medio = safe_div(rec_total, total_projetos_unicos)

        # Cor dinâmica para a margem
        cor_margem_v = "#28a745" if mg_total_pct >= 0.25 else ("#FFCB05" if mg_total_pct > 0 else "#dc3545")

        # --- RENDERIZAÇÃO NO FORMATO KPI-CARD DO ORÇAMENTO ---
        st.markdown(f"""
            <div class="kpi-container" style="margin-bottom: 20px; margin-top: 20px;">
                <div class="kpi-card" style="border-left-color: #B8860B;">
                    <div class="kpi-label">RECEITA ON DEMAND</div>
                    <div class="kpi-value">R$ {fmt(rec_ondemand)}</div>
                </div>
                <div class="kpi-card" style="border-left-color: #B8860B;">
                    <div class="kpi-label">RECEITA VENTURES</div>
                    <div class="kpi-value">R$ {fmt(rec_ventures)}</div>
                </div>
                <div class="kpi-card" style="border-left-color: #1A1A1A;">
                    <div class="kpi-label">PROJETOS APROVADOS</div>
                    <div class="kpi-value">{int(total_projetos_unicos)}</div>
                </div>
                <div class="kpi-card" style="border-left-color: {cor_margem_v};">
                    <div class="kpi-label">MARGEM BRUTA TOTAL</div>
                    <div class="kpi-value">{fmt(mg_total_pct, is_pct=True)}</div>
                </div>
                <div class="kpi-card" style="border-left-color: #1A1A1A;">
                    <div class="kpi-label">TICKET MÉDIO</div>
                    <div class="kpi-value">R$ {fmt(ticket_medio)}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Erro ao processar KPIs de Receitas: {e}")



    #====================# 02. TOP 10 CLIENTES #====================#
    col_t10_tit, col_t10_box = st.columns([5.5, 2.5], gap="small", vertical_alignment="bottom")

    with col_t10_tit:
        st.markdown('<div class="header-container" style="margin-top: 40px; margin-bottom: 10px;"><div class="quadro-num">04.</div><div class="quadro-titulo">Top 10 Clientes</div></div>', unsafe_allow_html=True)

    with col_t10_box:
        st.markdown('<div style="display: flex; justify-content: flex-end; margin-bottom: 10px;">', unsafe_allow_html=True)
        filtro_top10 = st.segmented_control(
            "Segmento:", ["Todos", "On Demand", "Ventures"], 
            selection_mode="single", default="Todos", 
            label_visibility="collapsed", key="radio_top10_vendas"
        )
        st.markdown('</div>', unsafe_allow_html=True)

    try:
        df_top10_raw = pd.read_excel("dados_dre.xlsx", sheet_name="Margem por Faixa", skiprows=18)
        df_top10_raw = df_top10_raw.loc[:, ~df_top10_raw.columns.str.contains('^Unnamed')]

        col_map_t10 = {normalize_id(c): c for c in df_top10_raw.columns}
        c_nucleo_t10   = col_map_t10.get("nucleo")
        c_cliente_t10  = col_map_t10.get("cliente")
        c_receita_t10  = col_map_t10.get("receitabruta")
        c_imposto_t10  = col_map_t10.get("imposto")
        c_custo_t10    = col_map_t10.get("custo")
        c_projeto_t10  = col_map_t10.get("nomeprojeto")

        if not c_nucleo_t10 or not c_cliente_t10:
            st.error("Colunas críticas não encontradas no Top 10 Clientes.")
        else:
            on_demand_ids = [normalize_id(x) for x in ["Núcleo 1", "Núcleo 2", "Núcleo 3", "Núcleo 4", "Núcleo 5", "Núcleo 6"]]
            ventures_ids = [normalize_id(x) for x in ["Núcleo 7", "Núcleo 8", "Ventures - outros", "Ventures"]]

            if filtro_top10 == "On Demand":
                df_filtrado_t10 = df_top10_raw[df_top10_raw[c_nucleo_t10].apply(normalize_id).isin(on_demand_ids)].copy()
            elif filtro_top10 == "Ventures":
                df_filtrado_t10 = df_top10_raw[df_top10_raw[c_nucleo_t10].apply(normalize_id).isin(ventures_ids)].copy()
            else:
                df_filtrado_t10 = df_top10_raw.copy()

            for c_f_t in [c_receita_t10, c_imposto_t10, c_custo_t10]:
                if c_f_t: df_filtrado_t10[c_f_t] = df_filtrado_t10[c_f_t].apply(safe_float)

            agg_dict_t10 = {}
            if c_receita_t10: agg_dict_t10[c_receita_t10] = 'sum'
            if c_imposto_t10: agg_dict_t10[c_imposto_t10] = 'sum'
            if c_custo_t10:   agg_dict_t10[c_custo_t10] = 'sum'
            if c_projeto_t10: agg_dict_t10[c_projeto_t10] = 'count'

            df_cli_t10 = df_filtrado_t10.groupby(c_cliente_t10).agg(agg_dict_t10).reset_index()
            df_cli_t10['lucro_recalc'] = df_cli_t10[c_receita_t10] - df_cli_t10[c_custo_t10].abs() - df_cli_t10[c_imposto_t10].abs()
            df_cli_t10 = df_cli_t10.sort_values(by=c_receita_t10, ascending=False).reset_index(drop=True)

            total_receita_gt10 = df_cli_t10[c_receita_t10].sum()
            total_imposto_gt10 = df_cli_t10[c_imposto_t10].sum()
            total_custo_gt10   = df_cli_t10[c_custo_t10].sum()
            total_lucro_gt10   = df_cli_t10['lucro_recalc'].sum()
            total_projetos_gt10 = df_cli_t10[c_projeto_t10].sum() if c_projeto_t10 else 0

            df_top10_f = df_cli_t10.head(10).copy()
            df_outros_t10 = df_cli_t10.iloc[10:].copy()

            if not df_outros_t10.empty:
                new_row_t10 = {
                    c_cliente_t10: 'Outros',
                    c_receita_t10: df_outros_t10[c_receita_t10].sum(),
                    c_imposto_t10: df_outros_t10[c_imposto_t10].sum(),
                    c_custo_t10: df_outros_t10[c_custo_t10].sum(),
                    'lucro_recalc': df_outros_t10['lucro_recalc'].sum(),
                    c_projeto_t10: df_outros_t10[c_projeto_t10].sum() if c_projeto_t10 else 0
                }
                df_final_t10 = pd.concat([df_top10_f, pd.DataFrame([new_row_t10])], ignore_index=True)
            else:
                df_final_t10 = df_top10_f

            df_final_t10['pct_receita'] = df_final_t10[c_receita_t10].apply(lambda x: safe_div(x, total_receita_gt10))
            df_final_t10['margem_calc'] = df_final_t10.apply(lambda r_t: safe_div(r_t['lucro_recalc'], r_t[c_receita_t10]), axis=1)

            html_top10 = f"""
            <div style="background-color: #FFFFFF; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #E9ECEF; padding: 20px; margin-bottom: 0px;">
                <style>
                    .top-table {{ width: 100%; border-collapse: collapse; color: #000; font-size: 11px; font-family: 'Segoe UI', sans-serif; }}
                    .top-table thead th {{ background: #1A1A1A; color: #FFF; padding: 10px; text-align: center; white-space: nowrap; }}
                    .top-table th:nth-child(2) {{ text-align: left; }}
                    .top-table td {{ padding: 8px; border-bottom: 1px solid #F0F0F0; text-align: center; white-space: nowrap; }}
                    .top-table td:nth-child(2) {{ text-align: left; font-weight: bold; white-space: normal; min-width: 150px; }}
                    .row-total-top {{ background-color: #f0f0f0 !important; font-weight: bold; }}
                    .row-outros {{ font-style: italic; color: #666; }}
                    .responsive-scroll {{ width: 100%; overflow-x: auto; -webkit-overflow-scrolling: touch; }}
                </style>
                
                <div class="responsive-scroll">
                    <table class="top-table">
                        <thead><tr>
                            <th>#</th><th>Cliente</th><th>Projetos</th><th>Receita Bruta</th><th>% Part s/ Rec Total</th><th>Imposto</th><th>Custo</th><th>Lucro Bruto</th><th>Margem</th>
                        </tr></thead>
                        <tbody>"""

            for i_t, row_t10 in df_final_t10.iterrows():
                cls_name_t10 = "row-outros" if row_t10[c_cliente_t10] == "Outros" else ""
                idx_label_t10 = i_t + 1 if row_t10[c_cliente_t10] != "Outros" else "-"
                m_val_t10 = row_t10.get('margem_calc', 0)
                m_color_t10 = "#28a745" if m_val_t10 >= 0 else "#dc3545"
                
                html_top10 += f"""<tr class="{cls_name_t10}">
                    <td>{idx_label_t10}</td>
                    <td>{row_t10[c_cliente_t10]}</td>
                    <td>{int(row_t10[c_projeto_t10]) if c_projeto_t10 else 0}</td>
                    <td>{fmt(row_t10.get(c_receita_t10, 0))}</td>
                    <td>{fmt(row_t10.get('pct_receita', 0), True)}</td>
                    <td>{fmt(row_t10.get(c_imposto_t10, 0))}</td>
                    <td>{fmt(row_t10.get(c_custo_t10, 0))}</td>
                    <td>{fmt(row_t10.get('lucro_recalc', 0))}</td>
                    <td style="color: {m_color_t10}; font-weight: bold;">{fmt(m_val_t10, True)}</td>
                </tr>"""

            t_mg_gt10 = safe_div(total_lucro_gt10, total_receita_gt10)
            html_top10 += f"""
                        <tr class="row-total-top">
                            <td>-</td><td>TOTAL GERAL</td><td>{int(total_projetos_gt10)}</td><td>{fmt(total_receita_gt10)}</td><td>100%</td>
                            <td>{fmt(total_imposto_gt10)}</td><td>{fmt(total_custo_gt10)}</td><td>{fmt(total_lucro_gt10)}</td>
                            <td style="color: {'#28a745' if t_mg_gt10 >= 0 else '#dc3545'}; font-weight: bold;">{fmt(t_mg_gt10, True)}</td>
                        </tr>
                        </tbody>
                    </table>
                </div>
            </div>"""

            st.components.v1.html(html_top10, height=510, scrolling=False)

    except Exception as e:
        st.error(f"Erro ao carregar Top 10 Clientes: {e}")
        
    # --- GRÁFICO DE CONCENTRAÇÃO DE RECEITA (DONUT) ---
    try:
        df_pie_t10 = df_final_t10[df_final_t10[c_receita_t10] > 0].copy()

        if not df_pie_t10.empty:
            import plotly.express as px

            fig_pie_t10 = px.pie(
                df_pie_t10,
                values=c_receita_t10,
                names=c_cliente_t10,
                hole=0,
                color_discrete_sequence=[
                    "#1A1A1A", "#FFCB05", "#424242", "#C59B00", "#757575", 
                    "#FDE085", "#9E9E9E", "#333333", "#E1AD01", "#BDBDBD", "#D2D2D2"
                ]
            )

            fig_pie_t10.update_traces(
                textposition='outside',
                textinfo='none',
                automargin=True,
                texttemplate='%{label} %{percent:.0%}', 
                hovertemplate="<b>%{label}</b><br>Receita: R$ %{value:,.0f}<br>Percentual: %{percent:.0%}<extra></extra>",
                marker=dict(line=dict(color='#FFFFFF', width=2))
            )

            fig_pie_t10.update_layout(
                showlegend=False,
                margin=dict(l=50, r=50, t=30, b=30),
                height=450,
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(family="Segoe UI", size=12),
                uniformtext_minsize=10,
                uniformtext_mode='hide',
                autosize=True
            )

            st.markdown("""
                <style>
                    div[data-testid="stPlotlyChart"] {
                        background-color: transparent !important;
                        border: none !important;
                        box-shadow: none !important;
                        padding: 0 !important;
                    }
                    iframe {
                        overflow: hidden !important;
                    }
                    .element-container, .stPlotlyChart {
                        overflow: hidden !important;
                    }
                </style>
            """, unsafe_allow_html=True)
            
            st.plotly_chart(
                fig_pie_t10, 
                use_container_width=True, 
                config={
                    'responsive': True, 
                    'displayModeBar': False
                }
            )
        
        else:
            st.warning("Dados insuficientes para gerar o gráfico de concentração de receita.")

    except Exception as e:
        st.error(f"Erro ao gerar gráfico de pizza: {e}")


# --- INÍCIO DO BLOCO DE RODAPÉ PERSONALIZADO (SOLUÇÃO DEFINITIVA) ---
import base64

def get_base64_img(file_path_img):
    try:
        with open(file_path_img, "rb") as f_img:
            data_img = f_img.read()
        return base64.b64encode(data_img).decode()
    except:
        return None

encoded_graphic_foot = get_base64_img("image_4.png")

footer_text_html = """
<div style="text-align: center; margin-top: 50px; padding-top: 20px; border-top: 1px solid #E9ECEF;">
    <p style="color: #888888; font-family: 'Segoe UI', sans-serif; font-size: 14px; font-weight: 600; margin-bottom: 10px;">
        V3A.AG<br>
        <span style="font-size: 10px; opacity: 0.8; letter-spacing: 2px;">
            RIO DE JANEIRO &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; SÃO PAULO &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; BRASÍLIA &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; SALVADOR
        </span>
    </p>
</div>
"""

st.markdown(footer_text_html, unsafe_allow_html=True)

if encoded_graphic_foot:
    img_html_foot = f"""
    <div style="text-align: center; padding-bottom: 30px;">
        <img src="data:image/png;base64,{encoded_graphic_foot}" style="height: 50px; width: auto; object-fit: contain; opacity: 0.7;">
    </div>
    """
    st.markdown(img_html_foot, unsafe_allow_html=True)