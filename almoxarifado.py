import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import base64
import unicodedata

#  >    executar  >   python -m streamlit run almoxarifado.py

# =====================================================
# CONFIGURAÇÕES DE CORES E ESTILOS - ALMOXARIFADO
# =====================================================

CORES_STATUS = {
    "ESTOQUE":      "#4caf50",
    "RESERVADO":    "#ff9800",
    "INATIVO":      "#94a3b8",
    "BLOQUEADO":    "#ff3c00",
    "EM ANÁLISE":   "#2196f3",
    "DEVOLVIDO":    "#ffffff",
}

CORES_KPI = {
    "TOTAL":      {"border": "#29b6f6", "text": "#29b6f6", "shadow": "rgba(41, 182, 246, 0.75)"},
    "OPERACAO":   {"border": "#ffb300", "text": "#ffb300", "shadow": "rgba(255, 179, 0, 0.75)"},
    "DISPONIVEIS":{"border": "#00e676", "text": "#00e676", "shadow": "rgba(0, 230, 118, 0.75)"},
    "MANUTENCAO": {"border": "#ff3d00", "text": "#ff3d00", "shadow": "rgba(255, 61, 0, 0.75)"}
}

CORES_POSICAO = ['#00d4ff', '#ff6b00', '#ffcc00', '#00ff88', '#bf5fff', '#ff2d55',
                 '#00ffea', '#ff9500', '#b8ff3c', '#ff3caa']

CORES_HEADER = {
    "background": "#252525", "border": "#ffffff",
    "title": "#ffffff", "subtitle": "#888", "dot": "#4caf50"
}

CORES_SIRENE = {
    "base_top": "#666", "base_bottom": "#333",
    "light_top": "#ff5722", "light_bottom": "#d32f2f",
    "light_glow": "rgba(255, 87, 34, 0.8)",
    "light_top_alt": "#ff9800", "light_bottom_alt": "#ff5722",
    "light_glow_alt": "rgba(255, 152, 0, 1)",
    "beam": "rgba(255, 87, 34, 0.3)"
}

CORES_DISPONIBILIDADE = {
    "background_start": "#000000", "background_end": "#000000",
    "border": "#2e5a2e", "label": "#000000fa",
    "valor": "#000000", "subtitle": "#000000"
}

CORES_INTERFACE = {
    "fundo_gradiente_start": "#141414", "fundo_gradiente_end": "#0a0a0a",
    "painel_background": "#1a1a1a", "painel_border": "#484848",
    "sidebar_background": "#111111", "sidebar_border": "#484848",
    "texto_principal": "#f5f5f5", "texto_secundario": "#bbbbbb",
    "grid": "#3a3a3a", "botao_primary": "#2196f3", "botao_hover": "#1976d2",
    "tabela_header_bg": "#111111", "tabela_row_bg": "#1a1a1a",
    "tabela_row_hover": "#252525"
}

# =====================================================
# CONFIGURAÇÃO DA PÁGINA
# =====================================================
st.set_page_config(
    page_title="CONTROLE | ALMOXARIFADO",
    layout="wide",
    page_icon="📦",
    initial_sidebar_state="expanded"
)


def get_base64_image(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except Exception:
        return None


def show_loading_screen(placeholder):
    img_base64 = get_base64_image("luft.png")
    if img_base64:
        loading_html = f"""
        <style>
        .loading-overlay {{
            position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
            background-color: #1a1a1a; display: flex; flex-direction: column;
            justify-content: center; align-items: center; z-index: 9999;
        }}
        .loading-image {{
            max-width: 80vw; max-height: 80vh; object-fit: contain;
            animation: pulse 2s ease-in-out infinite;
        }}
        .loading-text {{
            color: #ffffff; font-size: 24px; font-weight: 700;
            margin-top: 30px; text-align: center;
            animation: blink 1.5s ease-in-out infinite;
        }}
        @keyframes pulse {{ 0%, 100% {{ opacity: 0.8; transform: scale(1); }} 50% {{ opacity: 1; transform: scale(1.05); }} }}
        @keyframes blink {{ 0%, 100% {{ opacity: 0.5; }} 50% {{ opacity: 1; }} }}
        </style>
        <div class="loading-overlay">
            <img src="data:image/png;base64,{img_base64}" class="loading-image" alt="Loading">
            <div class="loading-text">CARREGANDO DADOS...</div>
        </div>
        """
    else:
        loading_html = """
        <style>
        .loading-overlay {
            position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
            background-color: #1a1a1a; display: flex; flex-direction: column;
            justify-content: center; align-items: center; z-index: 9999;
        }
        .loading-text {
            color: #ffffff; font-size: 32px; font-weight: 700;
            text-align: center; animation: blink 1.5s ease-in-out infinite;
        }
        @keyframes blink { 0%, 100% { opacity: 0.5; } 50% { opacity: 1; } }
        </style>
        <div class="loading-overlay"><div class="loading-text">📦 CARREGANDO DADOS...</div></div>
        """
    placeholder.markdown(loading_html, unsafe_allow_html=True)


# =====================================================
# CSS CUSTOMIZADO
# =====================================================
def load_custom_css():
    st.markdown(f"""
    <style>
    * {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif !important; }}
    header[data-testid="stHeader"] {{ background-color: rgba(0,0,0,0) !important; backdrop-filter: none !important; }}
    header[data-testid="stHeader"] > div:first-child {{ background-color: transparent !important; }}
    button[kind="header"] {{ color: white !important; }}
    .main .block-container {{ padding-top: 2rem !important; }}
    .stApp {{ background: linear-gradient(135deg, {CORES_INTERFACE["fundo_gradiente_start"]} 0%, {CORES_INTERFACE["fundo_gradiente_end"]} 100%) !important; }}
    .block-container {{ padding-top: 1.5rem !important; padding-bottom: 1rem !important; padding-left: 2rem !important; padding-right: 2rem !important; }}
    div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlockBorderWrapper"] {{
        background-color: {CORES_INTERFACE["painel_background"]} !important;
        border: 1px solid {CORES_INTERFACE["painel_border"]} !important;
        border-radius: 10px !important; padding: 25px !important;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.6) !important;
    }}
    .card-title {{
        font-weight: 800; font-size: 1.25rem; color: {CORES_INTERFACE["texto_principal"]} !important;
        text-transform: uppercase; margin-bottom: 20px; letter-spacing: 1.5px;
        display: flex; align-items: center; gap: 10px;
        text-shadow: 0 0 10px rgba(255,255,255,0.12);
    }}
    .kpi-card {{
        background-color: {CORES_INTERFACE["painel_background"]}; border-radius: 12px;
        padding: 25px 15px 15px 15px; text-align: center; width: 100%; box-sizing: border-box;
    }}
    .kpi-card .kpi-label {{
        font-size: 1.15rem; font-weight: 800; text-transform: uppercase;
        letter-spacing: 1.8px; margin-bottom: 10px;
    }}
    .kpi-card .kpi-value {{
        font-size: 3.2rem; font-weight: 900; line-height: 1.1;
        margin-bottom: 12px; letter-spacing: -1px;
    }}
    .kpi-azul {{
        border: 2px solid {CORES_KPI["TOTAL"]["border"]};
        box-shadow: 0 0 22px {CORES_KPI["TOTAL"]["shadow"]}, 0 0 6px {CORES_KPI["TOTAL"]["border"]}, 0 4px 20px rgba(0,0,0,0.5);
    }}
    .kpi-azul .kpi-label {{ color: {CORES_KPI["TOTAL"]["text"]}; text-shadow: 0 0 10px {CORES_KPI["TOTAL"]["shadow"]}; }}
    .kpi-azul .kpi-value {{ color: {CORES_KPI["TOTAL"]["text"]}; text-shadow: 0 0 18px {CORES_KPI["TOTAL"]["shadow"]}; }}
    .kpi-laranja {{
        border: 2px solid {CORES_KPI["OPERACAO"]["border"]};
        box-shadow: 0 0 22px {CORES_KPI["OPERACAO"]["shadow"]}, 0 0 6px {CORES_KPI["OPERACAO"]["border"]}, 0 4px 20px rgba(0,0,0,0.5);
    }}
    .kpi-laranja .kpi-label {{ color: {CORES_KPI["OPERACAO"]["text"]}; text-shadow: 0 0 10px {CORES_KPI["OPERACAO"]["shadow"]}; }}
    .kpi-laranja .kpi-value {{ color: {CORES_KPI["OPERACAO"]["text"]}; text-shadow: 0 0 18px {CORES_KPI["OPERACAO"]["shadow"]}; }}
    .kpi-verde {{
        border: 2px solid {CORES_KPI["DISPONIVEIS"]["border"]};
        box-shadow: 0 0 22px {CORES_KPI["DISPONIVEIS"]["shadow"]}, 0 0 6px {CORES_KPI["DISPONIVEIS"]["border"]}, 0 4px 20px rgba(0,0,0,0.5);
    }}
    .kpi-verde .kpi-label {{ color: {CORES_KPI["DISPONIVEIS"]["text"]}; text-shadow: 0 0 10px {CORES_KPI["DISPONIVEIS"]["shadow"]}; }}
    .kpi-verde .kpi-value {{ color: {CORES_KPI["DISPONIVEIS"]["text"]}; text-shadow: 0 0 18px {CORES_KPI["DISPONIVEIS"]["shadow"]}; }}
    .kpi-vermelho {{
        border: 2px solid {CORES_KPI["MANUTENCAO"]["border"]};
        box-shadow: 0 0 22px {CORES_KPI["MANUTENCAO"]["shadow"]}, 0 0 6px {CORES_KPI["MANUTENCAO"]["border"]}, 0 4px 20px rgba(0,0,0,0.5);
    }}
    .kpi-vermelho .kpi-label {{ color: {CORES_KPI["MANUTENCAO"]["text"]}; text-shadow: 0 0 10px {CORES_KPI["MANUTENCAO"]["shadow"]}; }}
    .kpi-vermelho .kpi-value {{ color: {CORES_KPI["MANUTENCAO"]["text"]}; text-shadow: 0 0 18px {CORES_KPI["MANUTENCAO"]["shadow"]}; }}

    /* ====== DIALOG — estilos base ====== */
    div[role="dialog"] > div {{
        max-width: 92vw !important; width: 92vw !important; max-height: 90vh !important;
        padding: clamp(16px, 3vw, 36px) clamp(14px, 3vw, 40px) !important;
        border-radius: 14px !important; overflow-y: auto !important; box-sizing: border-box !important;
    }}
    [data-testid="stDialog"], [data-testid="stDialog"] > div, div[role="dialog"],
    div[role="dialog"] > div, div[role="dialog"] > div > div,
    div[role="dialog"] section, div[role="dialog"] [data-testid="stVerticalBlock"] {{
        background-color: #1e1e1e !important; color: #ffffff !important;
    }}
    div[role="dialog"] p, div[role="dialog"] span, div[role="dialog"] label,
    div[role="dialog"] div, div[role="dialog"] h1, div[role="dialog"] h2,
    div[role="dialog"] h3, div[role="dialog"] small {{ color: #ffffff !important; }}
    div[role="dialog"] input, div[role="dialog"] textarea, div[role="dialog"] select {{
        background-color: #2a2a2a !important; color: #ffffff !important; border-color: #444 !important;
    }}
    div[role="dialog"] [data-baseweb="select"] > div {{
        background-color: #2a2a2a !important; color: #ffffff !important; border-color: #444 !important;
    }}
    div[role="dialog"] [data-testid="stCaptionContainer"] * {{ color: #aaaaaa !important; font-size: 0.85rem !important; }}
    div[role="dialog"] ::-webkit-scrollbar {{ width: 6px; }}
    div[role="dialog"] ::-webkit-scrollbar-track {{ background: #1a1a1a; }}
    div[role="dialog"] ::-webkit-scrollbar-thumb {{ background: #555; border-radius: 3px; }}
    div[role="dialog"] hr {{ border-color: #444 !important; }}
    div[role="dialog"] button[aria-label="Close"],
    div[role="dialog"] button[aria-label="Fechar"],
    div[role="dialog"] button[aria-label="close"],
    div[role="dialog"] button[aria-label="fechar"],
    div[role="dialog"] [data-testid="stBaseButton-headerNoPadding"],
    div[role="dialog"] button[data-testid="stBaseButton-headerNoPadding"] {{
        background: #c62828 !important; background-color: #c62828 !important;
        border-radius: 6px !important; border: 2px solid #ef5350 !important;
        box-shadow: 0 0 12px rgba(198,40,40,0.9) !important; opacity: 1 !important;
        width: 32px !important; height: 32px !important; min-width: 32px !important;
        padding: 0 !important; color: #ffffff !important;
    }}
    div[role="dialog"] button[aria-label="Close"]:hover,
    div[role="dialog"] button[aria-label="Fechar"]:hover,
    div[role="dialog"] [data-testid="stBaseButton-headerNoPadding"]:hover,
    div[role="dialog"] button[data-testid="stBaseButton-headerNoPadding"]:hover {{
        background: #c62828 !important; background-color: #c62828 !important;
        box-shadow: 0 0 30px rgba(239,83,80,1), 0 0 10px rgba(255,100,100,1) !important; opacity: 1 !important;
    }}
    div[role="dialog"] button[aria-label="Close"] svg,
    div[role="dialog"] [data-testid="stBaseButton-headerNoPadding"] svg,
    div[role="dialog"] button[data-testid="stBaseButton-headerNoPadding"] svg {{
        color: #ffffff !important; fill: #ffffff !important; stroke: #ffffff !important;
    }}
    div[role="dialog"] button[kind="secondary"] {{
        background: #c62828 !important; color: #ffffff !important;
        border: 2px solid #ef5350 !important; border-radius: 8px !important;
        font-size: 1rem !important; font-weight: 900 !important; padding: 8px 24px !important;
        box-shadow: 0 0 18px rgba(198,40,40,0.75) !important;
    }}

    /* ====== MINI CARDS ====== */
    .mini-card-item {{
        border-radius: 10px; padding: clamp(8px, 1.5vw, 14px) clamp(8px, 1.5vw, 12px);
        margin-bottom: 0; box-sizing: border-box; width: 100%;
    }}
    .mini-card-codigo {{
        font-size: clamp(0.75rem, 1.2vw, 0.9rem); font-weight: 900; letter-spacing: 1px;
        font-family: monospace; display: block; margin-bottom: 5px;
    }}
    .mini-card-descricao {{
        font-size: clamp(0.55rem, 0.85vw, 0.68rem); font-weight: 700; display: block;
        margin-bottom: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
    }}
    .mini-card-status-badge {{
        font-size: clamp(0.5rem, 0.9vw, 0.62rem); font-weight: 800; text-transform: uppercase;
        letter-spacing: 0.8px; padding: 3px 8px; border-radius: 20px; display: inline-block;
        margin-bottom: 5px; line-height: 1.4;
    }}
    .mini-card-saldo {{
        font-size: clamp(0.62rem, 1vw, 0.78rem); font-weight: 600; color: #cccccc !important;
        white-space: nowrap; overflow: hidden; text-overflow: ellipsis; display: block; margin-top: 2px;
    }}

    /* ====== CARD COMPLETO ====== */
    .card-info-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(min(100%, 155px), 1fr)); gap: 8px; }}
    .card-info-item {{ display: flex; flex-direction: column; gap: 3px; padding: 8px 10px; border-radius: 7px; min-width: 0; word-break: break-word; }}
    .card-info-label {{ font-size: 0.6rem; text-transform: uppercase; letter-spacing: 0.9px; font-weight: 700; opacity: 0.85; }}
    .card-info-value {{ color: #ffffff !important; font-size: 0.9rem; font-weight: 600; line-height: 1.3; }}

    /* ====== MEDIA QUERIES ====== */
    @media (min-width: 1400px) {{ div[role="dialog"] > div {{ padding: 36px 48px !important; }} }}
    @media (max-width: 900px) {{
        [data-testid="stDialog"] > div, div[role="dialog"] > div {{
            max-width: 99vw !important; width: 99vw !important; padding: 16px 12px !important;
        }}
    }}
    @media (max-width: 600px) {{
        [data-testid="stDialog"] > div, div[role="dialog"] > div {{
            max-width: 100vw !important; width: 100vw !important;
            max-height: 100vh !important; border-radius: 0 !important; padding: 12px 8px !important;
        }}
    }}

    .mini-card-expand-btn button {{
        background-color: transparent !important; color: #ffffff !important;
        border: 1px solid rgba(255,255,255,0.25) !important; border-radius: 5px !important;
        font-size: 0.65rem !important; font-weight: 700 !important; padding: 2px 8px !important;
        letter-spacing: 0.5px !important; width: 100% !important; transition: all 0.2s !important; margin-top: 0 !important;
    }}
    .mini-card-expand-btn button:hover {{ background-color: #cc0000 !important; border: 1px solid #ffffff !important; box-shadow: 0 0 12px #ffffff, 0 0 24px rgba(255,0,0,0.8) !important; color: white !important; }}
    .btn-voltar button {{
        background-color: #2a2a2a !important; color: #ffffff !important;
        border: 1px solid #555 !important; border-radius: 8px !important;
        font-size: 0.8rem !important; font-weight: 700 !important;
        letter-spacing: 0.5px !important; transition: all 0.2s !important;
    }}
    .btn-voltar button:hover {{ background-color: #cc0000 !important; border: 1px solid #ffffff !important; box-shadow: 0 0 12px #ffffff, 0 0 24px rgba(255,0,0,0.8) !important; color: white !important; }}

    /* ====== BOTÃO FULLSCREEN ====== */
    .btn-fullscreen button {{
        background-color: transparent !important; color: rgba(255,255,255,0.45) !important;
        border: 1px solid rgba(255,255,255,0.18) !important; border-radius: 6px !important;
        font-size: 1rem !important; font-weight: 400 !important; padding: 2px 8px !important;
        line-height: 1.2 !important; min-height: 0 !important; height: 28px !important;
        width: auto !important; transition: all 0.2s !important; box-shadow: none !important;
    }}
    .btn-fullscreen button:hover {{
        background-color: #cc0000 !important; border: 1px solid #ffffff !important;
        box-shadow: 0 0 12px #ffffff, 0 0 24px rgba(255,0,0,0.8) !important; color: white !important;
    }}

    .main-header {{
        background-color: {CORES_HEADER["background"]}; padding: 25px 30px; border-radius: 10px;
        border: 3px solid {CORES_HEADER["border"]}; margin-bottom: 25px;
        box-shadow: 0 0 12px rgba(255, 255, 255, 0.4), 0 4px 15px rgba(0, 0, 0, 0.3);
        display: flex; align-items: center; gap: 20px;
    }}
    .main-header h1 {{ color: {CORES_HEADER["title"]}; font-size: 2rem; font-weight: 800; margin: 0; margin-bottom: 8px; letter-spacing: 0.5px; text-align: center; }}
    .main-header p {{ color: {CORES_HEADER["subtitle"]}; font-size: 0.9rem; margin: 0; }}
    .header-center {{ flex: 1; display: flex; align-items: center; justify-content: center; }}
    .header-logo {{ max-height: 120px; max-width: 750px; width: 95%; height: auto; object-fit: contain; margin: 0; display: block; filter: drop-shadow(0 2px 8px rgba(255, 255, 255, 0.2)); transition: all 0.3s ease; }}
    .header-logo:hover {{ filter: drop-shadow(0 4px 12px rgba(255, 255, 255, 0.4)); transform: scale(1.02); }}
    .header-logo-placeholder {{ font-size: 1.8rem; font-weight: 800; color: {CORES_HEADER["title"]}; margin-bottom: 8px; letter-spacing: 1px; }}
    .mini-disponibilidade {{
        background: linear-gradient(135deg, {CORES_DISPONIBILIDADE["background_start"]} 0%, {CORES_DISPONIBILIDADE["background_end"]} 100%);
        border: 1px solid {CORES_DISPONIBILIDADE["border"]}; border-radius: 10px; padding: 18px 28px;
        text-align: center; flex-shrink: 0; min-width: 170px;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.4), inset 0 1px 0 rgba(255,255,255,0.05);
    }}
    .mini-disponibilidade .mini-label {{ font-size: 0.65rem; font-weight: 700; color: {CORES_DISPONIBILIDADE["label"]}; text-transform: uppercase; letter-spacing: 1.2px; margin-bottom: 6px; }}
    .mini-disponibilidade .mini-valor {{ font-size: 3rem; font-weight: 700; color: {CORES_DISPONIBILIDADE["valor"]}; line-height: 1.1; margin: 0; }}
    .sirene-container {{ position: relative; width: 60px; height: 60px; flex-shrink: 0; }}
    .sirene-base {{ position: absolute; bottom: 0; left: 50%; transform: translateX(-50%); width: 50px; height: 15px; background: linear-gradient(180deg, {CORES_SIRENE["base_top"]} 0%, {CORES_SIRENE["base_bottom"]} 100%); border-radius: 0 0 8px 8px; }}
    .sirene-light {{ position: absolute; top: 5px; left: 50%; transform: translateX(-50%); width: 40px; height: 35px; background: linear-gradient(180deg, {CORES_SIRENE["light_top"]} 0%, {CORES_SIRENE["light_bottom"]} 100%); border-radius: 50% 50% 20% 20%; box-shadow: 0 0 20px {CORES_SIRENE["light_glow"]}; animation: giroflex 1s infinite; }}
    .sirene-light::before {{ content: ''; position: absolute; top: 5px; left: 5px; width: 30px; height: 25px; background: linear-gradient(135deg, rgba(255, 255, 255, 0.3) 0%, transparent 100%); border-radius: 50% 50% 20% 20%; }}
    .sirene-beam {{ position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); width: 0; height: 0; border-left: 60px solid transparent; border-right: 60px solid transparent; border-top: 40px solid {CORES_SIRENE["beam"]}; animation: beam-rotate 1s infinite; transform-origin: 50% 0%; }}
    @keyframes giroflex {{
        0%, 100% {{ background: linear-gradient(180deg, {CORES_SIRENE["light_top"]} 0%, {CORES_SIRENE["light_bottom"]} 100%); box-shadow: 0 0 20px {CORES_SIRENE["light_glow"]}; }}
        50% {{ background: linear-gradient(180deg, {CORES_SIRENE["light_top_alt"]} 0%, {CORES_SIRENE["light_bottom_alt"]} 100%); box-shadow: 0 0 40px {CORES_SIRENE["light_glow_alt"]}, 0 0 60px rgba(255, 87, 34, 0.6); }}
    }}
    @keyframes beam-rotate {{
        0% {{ transform: translate(-50%, -50%) rotate(0deg); opacity: 0.3; }}
        50% {{ opacity: 0.6; }}
        100% {{ transform: translate(-50%, -50%) rotate(360deg); opacity: 0.3; }}
    }}
    @keyframes pulse {{ 0%, 100% {{ opacity: 1; }} 50% {{ opacity: 0.5; }} }}

    .stButton button {{ background-color: {CORES_INTERFACE["botao_primary"]} !important; color: white !important; border: none !important; border-radius: 6px !important; font-weight: 600 !important; padding: 10px 20px !important; transition: all 0.3s !important; }}
    .stButton button:hover {{ background-color: #cc0000 !important; border: 1px solid #ffffff !important; box-shadow: 0 0 12px #ffffff, 0 0 24px rgba(255,0,0,0.8) !important; color: white !important; }}
    .stDownloadButton button {{ background-color: {CORES_INTERFACE["botao_primary"]} !important; color: white !important; border: none !important; border-radius: 6px !important; font-weight: 600 !important; padding: 10px 20px !important; transition: all 0.3s !important; }}
    .stDownloadButton button:hover {{ background-color: #cc0000 !important; border: 1px solid #ffffff !important; box-shadow: 0 0 12px #ffffff, 0 0 24px rgba(255,0,0,0.8) !important; color: white !important; }}

    section[data-testid="stSidebar"] {{ background-color: {CORES_INTERFACE["sidebar_background"]} !important; border-right: 1px solid {CORES_INTERFACE["sidebar_border"]} !important; }}
    section[data-testid="stSidebar"] * {{ color: {CORES_INTERFACE["texto_principal"]} !important; }}
    section[data-testid="stSidebar"] .stMarkdown, section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] span, section[data-testid="stSidebar"] div {{ color: {CORES_INTERFACE["texto_principal"]} !important; }}
    section[data-testid="stSidebar"] label[data-testid="stWidgetLabel"] {{ color: {CORES_INTERFACE["texto_principal"]} !important; font-weight: 600 !important; }}
    section[data-testid="stSidebar"] .stMarkdown small {{ color: #cccccc !important; }}
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] p,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] span,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] small {{ color: #000000 !important; }}
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] [data-testid="stFileUploaderFile"],
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] [class*="uploadedFile"],
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] [class*="UploadedFile"] {{ display: none !important; }}
    section[data-testid="stSidebar"] [data-testid="stFileUploadDropzone"],
    section[data-testid="stSidebar"] [class*="uploadDropzone"],
    section[data-testid="stSidebar"] [class*="fileUploader"],
    section[data-testid="stSidebar"] [class*="FileUploader"] {{ background-color: #ffffff !important; border: 3px solid #ffb300 !important; border-radius: 12px !important; box-shadow: none !important; outline: none !important; width: 100% !important; max-width: 100% !important; box-sizing: border-box !important; overflow: hidden !important; }}
    section[data-testid="stSidebar"] [data-testid="stFileUploadDropzone"] *,
    section[data-testid="stSidebar"] [class*="uploadDropzone"] *,
    section[data-testid="stSidebar"] [class*="fileUploader"] *,
    section[data-testid="stSidebar"] [class*="FileUploader"] * {{ color: #000000 !important; }}
    section[data-testid="stSidebar"] [data-testid="stFileUploadDropzone"] button,
    section[data-testid="stSidebar"] [class*="uploadDropzone"] button,
    section[data-testid="stSidebar"] [class*="fileUploader"] button,
    section[data-testid="stSidebar"] [class*="FileUploader"] button {{ background-color: #ffb300 !important; color: #000000 !important; border: none !important; border-radius: 6px !important; font-weight: 700 !important; transition: all 0.2s !important; width: 100% !important; box-sizing: border-box !important; }}
    section[data-testid="stSidebar"] [data-testid="stFileUploadDropzone"] button:hover,
    section[data-testid="stSidebar"] [class*="uploadDropzone"] button:hover,
    section[data-testid="stSidebar"] [class*="fileUploader"] button:hover,
    section[data-testid="stSidebar"] [class*="FileUploader"] button:hover {{ background-color: #2e7d32 !important; color: #ffffff !important; box-shadow: 0 0 14px rgba(46, 125, 50, 0.8), 0 0 28px rgba(76, 175, 80, 0.5) !important; }}

    .dataframe {{ font-size: 0.85rem !important; color: {CORES_INTERFACE["texto_principal"]} !important; }}
    .dataframe thead tr th {{ background-color: {CORES_INTERFACE["tabela_header_bg"]} !important; color: {CORES_INTERFACE["texto_secundario"]} !important; font-weight: 600 !important; text-transform: uppercase !important; font-size: 0.7rem !important; letter-spacing: 0.5px !important; border-bottom: 1px solid {CORES_INTERFACE["painel_border"]} !important; }}
    .dataframe tbody tr td {{ background-color: {CORES_INTERFACE["tabela_row_bg"]} !important; color: {CORES_INTERFACE["texto_principal"]} !important; border-bottom: 1px solid {CORES_INTERFACE["painel_border"]} !important; }}
    .dataframe tbody tr:hover td {{ background-color: {CORES_INTERFACE["tabela_row_hover"]} !important; }}
    .js-plotly-plot {{ background-color: transparent !important; }}
    div[data-testid="stTooltipContent"],
    div[data-testid="stTooltipContent"] p,
    [data-testid="stTooltipContent"] * {{ background-color: #1a1a1a !important; color: #ffffff !important; }}
    div[data-testid="stMetricDelta"] {{ display: none !important; }}
    .stMarkdown, .stMarkdown p, .stMarkdown span, .stMarkdown div {{ color: {CORES_INTERFACE["texto_principal"]} !important; }}
    h1, h2, h3, h4, h5, h6 {{ color: {CORES_INTERFACE["texto_principal"]} !important; }}
    p {{ color: {CORES_INTERFACE["texto_principal"]} !important; }}
    hr {{ border-color: {CORES_INTERFACE["painel_border"]} !important; }}
    </style>
    """, unsafe_allow_html=True)


# =====================================================
# CONSTANTES
# =====================================================
STATUS_OFICIAIS = ["ESTOQUE"]
STATUS_ADICIONAIS = ["RESERVADO", "INATIVO", "BLOQUEADO", "EM ANÁLISE", "DEVOLVIDO"]
ORDEM_STATUS = STATUS_OFICIAIS + STATUS_ADICIONAIS


# =====================================================
# PROCESSAMENTO DE DADOS
# =====================================================

def _sem_acento(texto):
    """Remove acentos de uma string para comparação flexível."""
    return unicodedata.normalize('NFKD', str(texto)).encode('ASCII', 'ignore').decode('ASCII')


def _normalizar_nomes_colunas(df):
    MAPA = {
        "CODIGO":            "CÓDIGO",
        "DESCRICAO":         "DESCRIÇÃO",
        "POSICAO":           "POSIÇÃO",
        "VALORES UNITARIOS": "VALORES UNITÁRIOS",
        "SALDO TOTAL":       "SALDO TOTAL",
        "VALOR TOTAL":       "VALOR TOTAL",
        "ENTRADA":           "ENTRADA",
        "SAIDA":             "SAIDA",
        "STATUS":            "STATUS",
    }
    renomear = {}
    for col in df.columns:
        chave = _sem_acento(col).strip().upper()
        chave = ' '.join(chave.split())
        oficial = MAPA.get(chave)
        if oficial and col != oficial:
            renomear[col] = oficial
    if renomear:
        df = df.rename(columns=renomear)
    return df


def _renomear_colunas_duplicadas(df):
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        dup_indices = [i for i, x in enumerate(cols) if x == dup]
        for idx, pos in enumerate(dup_indices):
            if idx > 0:
                cols[pos] = f"{dup}_{idx}"
    df.columns = cols
    return df


@st.cache_data
def load_data_from_file(file_source):
    try:
        df = pd.read_excel(file_source, sheet_name=0)

        if df.empty:
            st.error("❌ A planilha está vazia!")
            return pd.DataFrame()

        df.columns = [
            str(c).strip().upper() if pd.notna(c) else f"COL_{i}"
            for i, c in enumerate(df.columns)
        ]

        df = _normalizar_nomes_colunas(df)
        df = _renomear_colunas_duplicadas(df)

        if "CÓDIGO" not in df.columns:
            st.error(
                f"❌ Coluna 'CÓDIGO' não encontrada!  "
                f"Colunas detectadas: {list(df.columns)}"
            )
            return pd.DataFrame()

        df = df[
            df["CÓDIGO"].notna() &
            (df["CÓDIGO"].astype(str).str.strip() != "")
        ].copy()

        if df.empty:
            st.error("❌ Nenhum dado válido encontrado!")
            return pd.DataFrame()

        colunas_texto = ["STATUS", "CÓDIGO", "DESCRIÇÃO", "POSIÇÃO"]
        for col in colunas_texto:
            if col in df.columns:
                df[col] = (
                    df[col].astype(str)
                    .str.strip()
                    .str.upper()
                    .str.replace(r'\s+', ' ', regex=True)
                )
                df[col] = df[col].where(
                    ~df[col].isin(["NAN", "NONE", "NAT", ""]),
                    other=pd.NA
                )

        for col in ["ENTRADA", "SAIDA", "SALDO TOTAL"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        if "SALDO TOTAL" not in df.columns:
            for alias in ["SALDO", "QUANTIDADE"]:
                if alias in df.columns:
                    df["SALDO TOTAL"] = pd.to_numeric(df[alias], errors='coerce').fillna(0)
                    break
            else:
                df["SALDO TOTAL"] = 0

        if "VALORES UNITÁRIOS" in df.columns:
            df["VALORES UNITÁRIOS"] = (
                df["VALORES UNITÁRIOS"].astype(str).str.replace(',', '.', regex=False)
            )
            df["VALORES UNITÁRIOS"] = pd.to_numeric(df["VALORES UNITÁRIOS"], errors='coerce')

        if "VALOR TOTAL" in df.columns:
            df["VALOR TOTAL"] = pd.to_numeric(df["VALOR TOTAL"], errors='coerce').fillna(0)

        if "STATUS" in df.columns:
            df["STATUS"] = df["STATUS"].fillna("ESTOQUE").astype(str).str.upper()
        else:
            df["STATUS"] = "ESTOQUE"

        return df

    except Exception as e:
        st.error(f"❌ Erro ao carregar dados: {str(e)}")
        return pd.DataFrame()


def _cor_saldo(saldo):
    try:
        s = float(saldo)
        if s == 0:
            return CORES_KPI["MANUTENCAO"]["border"]
        elif s <= 3:
            return CORES_KPI["OPERACAO"]["border"]
        else:
            return CORES_KPI["DISPONIVEIS"]["border"]
    except Exception:
        return "#888888"


def aplicar_cor_status(row):
    saldo = row.get("SALDO TOTAL", 0)
    try:
        s = float(saldo)
    except Exception:
        s = 0
    if s == 0:
        bg, fg = "#2a0a0a", "#ff3d00"
    elif s <= 3:
        bg, fg = "#2a1a00", "#ffb300"
    else:
        bg, fg = "#0a1f0a", "#00e676"
    return [f'background-color: {bg}; color: {fg}'] * len(row)


# =====================================================
# GRÁFICOS
# =====================================================

def criar_grafico_status(df_filtrado):
    df = df_filtrado.copy()
    zerados = df[df["SALDO TOTAL"] == 0][["CÓDIGO", "DESCRIÇÃO", "SALDO TOTAL"]].copy()
    alertas = df[(df["SALDO TOTAL"] > 0) & (df["SALDO TOTAL"] <= 3)][["CÓDIGO", "DESCRIÇÃO", "SALDO TOTAL"]].copy()
    zerados["CATEGORIA"] = "SEM ESTOQUE"
    alertas["CATEGORIA"] = "ALERTA"
    combined = pd.concat([alertas, zerados], ignore_index=True)

    if combined.empty:
        fig = go.Figure()
        fig.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=320,
            annotations=[dict(
                text="✅ Nenhum alerta ou item zerado",
                x=0.5, y=0.5, xref="paper", yref="paper",
                showarrow=False, font=dict(color="#00e676", size=16, family="Arial Black")
            )]
        )
        return fig

    labels = combined.apply(
        lambda r: (str(r["DESCRIÇÃO"])[:28] + "…") if len(str(r["DESCRIÇÃO"])) > 30 else str(r["DESCRIÇÃO"]),
        axis=1
    ).tolist()
    saldos = combined["SALDO TOTAL"].fillna(0).tolist()
    cores = [
        CORES_KPI["MANUTENCAO"]["border"] if cat == "SEM ESTOQUE" else CORES_KPI["OPERACAO"]["border"]
        for cat in combined["CATEGORIA"]
    ]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=labels, x=saldos, orientation='h',
        marker=dict(color=cores, line=dict(color='rgba(255,255,255,0.15)', width=1)),
        text=[str(int(s)) for s in saldos], textposition='outside',
        textfont=dict(color='#ffffff', size=14, family='Arial Black'),
        hovertemplate='<b>%{y}</b><br>Saldo: %{x}<extra></extra>',
        showlegend=False
    ))
    altura = max(320, len(combined) * 28 + 60)
    fig.update_layout(
        hoverlabel=dict(bgcolor='#1e1e1e', bordercolor='#555',
                        font=dict(size=14, color='#ffffff', family='Arial Black'), namelength=-1),
        height=altura, margin=dict(l=0, r=60, t=10, b=0),
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#f5f5f5', size=13),
        xaxis=dict(showgrid=True, gridcolor=CORES_INTERFACE["grid"],
                   tickfont=dict(size=13, color=CORES_INTERFACE["texto_secundario"])),
        yaxis=dict(showgrid=False, tickfont=dict(size=11, color='#f5f5f5', family='Arial Black'),
                   automargin=True)
    )
    return fig


def criar_grafico_itens(df_filtrado, fullscreen=False):
    df = df_filtrado.copy()
    df["SALDO TOTAL"] = pd.to_numeric(df["SALDO TOTAL"], errors="coerce").fillna(0)
    df_zerado = df[df["SALDO TOTAL"] == 0].copy()

    if df_zerado.empty:
        fig = go.Figure()
        fig.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=420,
            annotations=[dict(
                text="✅ Nenhum item sem estoque",
                x=0.5, y=0.5, xref="paper", yref="paper",
                showarrow=False, font=dict(color="#00e676", size=16, family="Arial Black")
            )]
        )
        return fig

    tem_unit = "VALORES UNITÁRIOS" in df_zerado.columns
    tem_vt   = "VALOR TOTAL" in df_zerado.columns
    if tem_unit:
        df_zerado["_VU"] = pd.to_numeric(df_zerado["VALORES UNITÁRIOS"], errors="coerce").fillna(0)
    elif tem_vt:
        df_zerado["_VU"] = pd.to_numeric(df_zerado["VALOR TOTAL"], errors="coerce").fillna(0)
    else:
        df_zerado["_VU"] = 0.0

    df_zerado = df_zerado.sort_values("_VU", ascending=False).reset_index(drop=True)

    def _fmt_valor(v):
        try:
            return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return "R$ 0,00"

    cor_vermelho = CORES_KPI["MANUTENCAO"]["border"]
    labels  = df_zerado["DESCRIÇÃO"].apply(lambda x: (str(x)[:28] + "…") if len(str(x)) > 30 else str(x)).tolist()
    valores = df_zerado["_VU"].tolist()
    textos  = [_fmt_valor(v) for v in valores]

    ordem   = sorted(range(len(valores)), key=lambda i: valores[i], reverse=True)
    labels  = [labels[i]  for i in ordem]
    valores = [valores[i] for i in ordem]
    textos  = [textos[i]  for i in ordem]

    fig = go.Figure(data=[go.Bar(
        x=labels, y=valores,
        marker=dict(color=cor_vermelho, line=dict(color="rgba(255,255,255,0.20)", width=1)),
        text=textos, textposition="outside",
        textfont=dict(color="#ffffff", size=16, family="Arial Black"),
        hovertemplate="<b>%{x}</b><br>Valor unitário: %{text}<extra></extra>",
        showlegend=False, cliponaxis=False,
    )])

    n       = len(df_zerado)
    altura  = 680 if fullscreen else max(480, n * 30 + 220)
    val_max = max(valores) if valores else 1
    y_max   = val_max * 1.30 if val_max > 0 else 10

    fig.update_layout(
        hoverlabel=dict(bgcolor="#1e1e1e", bordercolor=cor_vermelho,
                        font=dict(size=13, color="#ffffff", family="Arial Black"), namelength=-1),
        height=altura, margin=dict(l=10, r=10, t=80, b=200),
        paper_bgcolor="rgba(0,0,0,0)" if not fullscreen else "#141414",
        plot_bgcolor="rgba(0,0,0,0)" if not fullscreen else "#141414",
        font=dict(color="#f5f5f5", size=13), bargap=0.25,
        xaxis=dict(tickangle=-50, tickfont=dict(size=10, family="Arial Black", color="#f5f5f5"),
                   automargin=True, showgrid=False),
        yaxis=dict(showgrid=True, gridcolor=CORES_INTERFACE["grid"],
                   tickfont=dict(size=12, color=CORES_INTERFACE["texto_secundario"]),
                   tickprefix="R$ ", range=[0, y_max]),
    )
    return fig


def criar_grafico_posicao(posicao_df, fullscreen=False):
    if posicao_df.empty:
        return go.Figure()

    df_plot    = posicao_df if fullscreen else posicao_df.head(10)
    cor_zerado = CORES_KPI["MANUTENCAO"]["border"]
    tem_saldo  = "SALDO_TOTAL" in df_plot.columns
    cores = [
        cor_zerado if (tem_saldo and df_plot.iloc[i].get("SALDO_TOTAL", 1) == 0)
        else CORES_POSICAO[i % len(CORES_POSICAO)]
        for i in range(len(df_plot))
    ]

    def _fmt(v):
        try:
            return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return "R$ 0,00"

    tem_extra      = "SALDO_TOTAL" in df_plot.columns and "VALOR_TOTAL" in df_plot.columns
    tem_itens_info = "ITENS_INFO" in df_plot.columns
    hover_texts    = []
    for _, r in df_plot.iterrows():
        saldo       = int(r.get("SALDO_TOTAL", 0)) if tem_extra else int(r["QUANTIDADE"])
        valor_total = _fmt(r.get("VALOR_TOTAL", 0)) if tem_extra else "—"
        txt = f"<b>Posição: {r['POSIÇÃO']}</b><br>─────────────────────"
        if tem_itens_info and str(r.get("ITENS_INFO", "")).strip():
            txt += f"<br>{r['ITENS_INFO']}"
        else:
            txt += f"<br>Qtd. em estoque: {saldo}<br>Valor total: {valor_total}"
        hover_texts.append(txt)

    tem_saldo_col = "SALDO_TOTAL" in df_plot.columns
    if tem_saldo_col:
        textos_fatia = df_plot["SALDO_TOTAL"].fillna(0).astype(int).apply(
            lambda v: str(v) if v > 0 else "0"
        ).tolist()
    else:
        textos_fatia = df_plot["QUANTIDADE"].astype(str).tolist()

    fig = go.Figure(data=[go.Pie(
        labels=df_plot["POSIÇÃO"], values=df_plot["QUANTIDADE"],
        hole=0.62, marker=dict(colors=cores, line=dict(color='#0a0a0a', width=2)),
        textinfo='none', text=hover_texts, customdata=textos_fatia,
        texttemplate='%{customdata}',
        textfont=dict(color='#ffffff', size=18, family='Arial Black'),
        hovertemplate='%{text}<extra></extra>'
    )])

    if fullscreen:
        fig.update_layout(
            hoverlabel=dict(bgcolor='#1e1e1e', bordercolor='#555',
                            font=dict(size=18, color='#ffffff', family='Arial Black'), namelength=-1),
            height=720, margin=dict(l=10, r=220, t=20, b=20),
            paper_bgcolor='#141414', plot_bgcolor='#141414', showlegend=True,
            legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.01,
                        font=dict(color='#f5f5f5', size=16, family='Arial Black'), bgcolor='rgba(0,0,0,0)')
        )
    else:
        fig.update_layout(
            hoverlabel=dict(bgcolor='#1e1e1e', bordercolor='#555',
                            font=dict(size=18, color='#ffffff', family='Arial Black'), namelength=-1),
            height=320, margin=dict(l=0, r=0, t=0, b=0),
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', showlegend=True,
            legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02,
                        font=dict(color='#f5f5f5', size=14, family='Arial Black'), bgcolor='rgba(0,0,0,0)')
        )
    return fig


# =====================================================
# HELPERS
# =====================================================

def _hex_to_rgba(hex_color, alpha=0.12):
    h = hex_color.lstrip('#')
    r, g, b = tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
    return f"rgba({r},{g},{b},{alpha})"


def _html_mini_card(item, cor_override=None):
    status    = str(item.get("STATUS", "")).strip()
    codigo    = str(item.get("CÓDIGO", "—")).strip()
    descricao = str(item.get("DESCRIÇÃO", "—")).strip()[:50]
    saldo     = item.get("SALDO TOTAL", 0)
    cor_item  = cor_override if cor_override else _cor_saldo(saldo)
    cor_fundo = _hex_to_rgba(cor_item, 0.22)
    cor_borda = _hex_to_rgba(cor_item, 0.70)
    try:
        saldo_str = str(int(float(saldo)))
    except Exception:
        saldo_str = "—"
    return f"""
    <div class="mini-card-item" style="background:{cor_fundo};border:2px solid {cor_borda};border-left:5px solid {cor_item};box-shadow: 0 0 10px {_hex_to_rgba(cor_item, 0.18)};">
        <span class="mini-card-codigo" style="color:{cor_item};text-shadow:0 0 10px {_hex_to_rgba(cor_item, 0.5)};">📦 {codigo}</span>
        <span class="mini-card-descricao" style="color:#e0e0e0;">{descricao}</span>
        <span class="mini-card-status-badge" style="background:{_hex_to_rgba(cor_item, 0.30)};border:1px solid {cor_item};color:{cor_item};">{status}</span>
        <span class="mini-card-saldo">📊 Saldo: <b style="color:{cor_item};">{saldo_str}</b></span>
    </div>
    """


def _html_card_completo(item, cor_override=None):
    status    = str(item.get("STATUS", "")).strip()
    codigo    = str(item.get("CÓDIGO", "—")).strip()
    descricao = str(item.get("DESCRIÇÃO", "—")).strip()

    cor_status = cor_override if cor_override else CORES_STATUS.get(status, _cor_saldo(item.get("SALDO TOTAL", 0)))
    cor_fundo  = _hex_to_rgba(cor_status, 0.28)
    cor_borda  = _hex_to_rgba(cor_status, 0.80)
    cor_ib     = _hex_to_rgba(cor_status, 0.12)
    cor_ib2    = _hex_to_rgba(cor_status, 0.35)

    campos_cabecalho  = {"CÓDIGO", "DESCRIÇÃO", "STATUS"}
    campos_prioridade = ["POSIÇÃO", "ENTRADA", "SAIDA", "SALDO TOTAL", "VALORES UNITÁRIOS", "VALOR TOTAL"]
    todos_campos      = list(item.index)
    campos_extras     = [c for c in todos_campos if c not in campos_cabecalho and c not in campos_prioridade]
    ordem_final       = campos_prioridade + campos_extras

    def _valor_valido(val):
        if val is None:
            return False
        try:
            if pd.isna(val):
                return False
        except (TypeError, ValueError):
            pass
        s = str(val).strip().upper()
        return s not in ("", "NAN", "NONE", "NAT", "UNNAMED")

    info_items = ""
    for campo in ordem_final:
        val = item.get(campo)
        if not _valor_valido(val):
            continue
        val_str = str(val).strip()
        try:
            f = float(val_str)
            if f == int(f):
                val_str = str(int(f))
        except (ValueError, TypeError):
            pass
        label = str(campo).replace("_", " ").title()
        info_items += f"""
        <div style="display:flex; flex-direction:column; gap:6px; padding:16px 18px;
                    border-radius:10px; background:{cor_ib}; border:1px solid {cor_ib2};
                    min-width:0; word-break:break-word;">
            <span style="font-size:0.72rem; text-transform:uppercase; letter-spacing:1.1px;
                         font-weight:700; color:{cor_status}; opacity:0.9;">{label}</span>
            <span style="color:#ffffff; font-size:1.15rem; font-weight:600; line-height:1.3;">{val_str}</span>
        </div>"""

    divider  = f'<hr style="border:none; border-top:1px solid {cor_status}; margin:16px 0 14px 0; opacity:0.4;">' if info_items else ""
    info_blk = f'<div style="display:grid; grid-template-columns:repeat(auto-fill, minmax(clamp(140px, 18vw, 220px), 1fr)); gap:clamp(8px, 1.2vw, 14px);">{info_items}</div>' if info_items else ""

    return f"""
    <div style="border-radius:14px; padding:30px 28px; margin-bottom:14px; box-sizing:border-box; width:100%;
        background:{cor_fundo}; border:2px solid {cor_borda}; border-left:8px solid {cor_status};
        box-shadow:0 0 24px {_hex_to_rgba(cor_status, 0.30)}, 0 4px 20px rgba(0,0,0,0.5);">
        <div style="display:flex; flex-wrap:wrap; align-items:center; justify-content:space-between; gap:14px;">
            <div style="display:flex; flex-wrap:wrap; align-items:center; gap:14px; min-width:0;">
                <span style="font-size:clamp(1.6rem, 4vw, 2.2rem); font-weight:900; letter-spacing:4px;
                             font-family:monospace; white-space:nowrap; color:{cor_status};
                             text-shadow:0 0 20px {_hex_to_rgba(cor_status, 0.7)};">📦 {codigo}</span>
            </div>
            <span style="font-size:0.85rem; font-weight:800; text-transform:uppercase; letter-spacing:1.4px;
                         padding:9px 20px; border-radius:20px; white-space:nowrap; flex-shrink:0;
                         background:{_hex_to_rgba(cor_status, 0.35)}; border:2px solid {cor_status};
                         color:{cor_status}; box-shadow:0 0 12px {_hex_to_rgba(cor_status, 0.5)};">{status}</span>
        </div>
        <div style="color:#cccccc; font-size:1rem; margin-top:10px;">{descricao}</div>
        {divider}
        {info_blk}
    </div>"""


# =====================================================
# INTERFACE
# =====================================================

def criar_header(pct_disponiveis=0.0):
    logo_base64 = get_base64_image("logo_luft.png")
    if logo_base64:
        logo_html = f'<img src="data:image/png;base64,{logo_base64}" class="header-logo" alt="Logo">'
    else:
        logo_html = '<div class="header-logo-placeholder">📦 ALMOXARIFADO</div>'
    st.markdown(f"""
    <div class="main-header">
        <div class="sirene-container"><div class="sirene-beam"></div><div class="sirene-light"></div><div class="sirene-base"></div></div>
        <div class="header-center">{logo_html}</div>
        <div class="mini-disponibilidade"><div class="mini-label">DISPONÍVEIS</div><div class="mini-valor">{pct_disponiveis:.1f}%</div></div>
    </div>
    """, unsafe_allow_html=True)


def criar_kpis(df_filtrado):
    df_filtrado = df_filtrado.copy()
    df_filtrado["SALDO TOTAL"] = pd.to_numeric(df_filtrado["SALDO TOTAL"], errors="coerce").fillna(0)

    total       = len(df_filtrado)
    disponiveis = len(df_filtrado[df_filtrado["SALDO TOTAL"] > 3])
    alerta      = len(df_filtrado[(df_filtrado["SALDO TOTAL"] > 0) & (df_filtrado["SALDO TOTAL"] <= 3)])
    zerados     = len(df_filtrado[df_filtrado["SALDO TOTAL"] == 0])

    st.session_state["_kpi_df_TOTAL"]       = df_filtrado
    st.session_state["_kpi_df_DISPONIVEIS"] = df_filtrado[df_filtrado["SALDO TOTAL"] > 3].copy()
    st.session_state["_kpi_df_ALERTA"]      = df_filtrado[(df_filtrado["SALDO TOTAL"] > 0) & (df_filtrado["SALDO TOTAL"] <= 3)].copy()
    st.session_state["_kpi_df_ZERADOS"]     = df_filtrado[df_filtrado["SALDO TOTAL"] == 0].copy()

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown(f'<div class="kpi-card kpi-azul"><div class="kpi-label">TOTAL DE ITENS</div><div class="kpi-value">{total}</div></div>', unsafe_allow_html=True)
        if st.button("🔍 VER TODOS", key="btn_total", use_container_width=True):
            for k in [k for k in st.session_state if k.startswith("_kpi_aberto_") or k.startswith("_kpi_sel_")]:
                st.session_state[k] = False if k.startswith("_kpi_aberto_") else None
            st.session_state["_kpi_aberto_TOTAL"] = True
            st.rerun()

    with col2:
        st.markdown(f'<div class="kpi-card kpi-verde"><div class="kpi-label">DISPONÍVEIS</div><div class="kpi-value">{disponiveis}</div></div>', unsafe_allow_html=True)
        if st.button("🔍 VER DISPONÍVEIS", key="btn_disp", use_container_width=True):
            for k in [k for k in st.session_state if k.startswith("_kpi_aberto_") or k.startswith("_kpi_sel_")]:
                st.session_state[k] = False if k.startswith("_kpi_aberto_") else None
            st.session_state["_kpi_aberto_DISPONIVEIS"] = True
            st.rerun()

    with col3:
        st.markdown(f'<div class="kpi-card kpi-laranja"><div class="kpi-label">ALERTA</div><div class="kpi-value">{alerta}</div></div>', unsafe_allow_html=True)
        if st.button("🔍 VER ALERTA", key="btn_alerta", use_container_width=True):
            for k in [k for k in st.session_state if k.startswith("_kpi_aberto_") or k.startswith("_kpi_sel_")]:
                st.session_state[k] = False if k.startswith("_kpi_aberto_") else None
            st.session_state["_kpi_aberto_ALERTA"] = True
            st.rerun()

    with col4:
        st.markdown(f'<div class="kpi-card kpi-vermelho"><div class="kpi-label">SEM ESTOQUE</div><div class="kpi-value">{zerados}</div></div>', unsafe_allow_html=True)
        if st.button("🔍 VER ZERADOS", key="btn_zero", use_container_width=True):
            for k in [k for k in st.session_state if k.startswith("_kpi_aberto_") or k.startswith("_kpi_sel_")]:
                st.session_state[k] = False if k.startswith("_kpi_aberto_") else None
            st.session_state["_kpi_aberto_ZERADOS"] = True
            st.rerun()

    return disponiveis, alerta, zerados


def mostrar_detalhes_kpi(titulo, cor_hex, df_kpi):
    st.markdown("""
    <style>
    section[data-testid="stSidebar"] { display: none !important; }
    .main .block-container { padding: 1rem 1.5rem !important; max-width: 100% !important; }
    </style>
    """, unsafe_allow_html=True)

    key_sel    = f"_kpi_sel_{titulo}"
    key_busca  = f"_kpi_busca_{titulo}"
    key_aberto = f"_kpi_aberto_{titulo}"

    if key_sel not in st.session_state:
        st.session_state[key_sel] = None
    if key_aberto not in st.session_state:
        st.session_state[key_aberto] = True

    if st.session_state[key_sel] is not None:
        codigo_sel = st.session_state[key_sel]
        resultado  = df_kpi[df_kpi["CÓDIGO"].astype(str).str.strip() == str(codigo_sel).strip()]

        col1, col2, col3 = st.columns([1, 4, 1])
        with col1:
            if st.button("🏠 INÍCIO", key=f"btn_inicio_v_{titulo}", use_container_width=True):
                for k in [k for k in st.session_state if k.startswith("_kpi_aberto_") or k.startswith("_kpi_sel_")]:
                    st.session_state[k] = False if k.startswith("_kpi_aberto_") else None
                st.session_state.pop("_grafico_fs", None)
                st.rerun()
        with col3:
            if st.button("← VOLTAR", key=f"btn_voltar_{titulo}", use_container_width=True, type="secondary"):
                st.session_state[key_sel]    = None
                st.session_state[key_aberto] = True
                st.rerun()
        with col2:
            st.markdown(f"""
            <div style="border-left:4px solid {cor_hex}; padding:8px 16px;
                background:{_hex_to_rgba(cor_hex, 0.18)}; border-radius:0 8px 8px 0;">
                <span style="color:#aaa; font-size:0.68rem; letter-spacing:1px; text-transform:uppercase;">ITEM SELECIONADO</span><br>
                <span style="color:{cor_hex}; font-size:1.25rem; font-weight:900; font-family:monospace; letter-spacing:2px;">{codigo_sel}</span>
            </div>""", unsafe_allow_html=True)

        st.divider()
        if resultado.empty:
            st.warning("Item não encontrado.")
        else:
            st.markdown(_html_card_completo(resultado.iloc[0], cor_override=cor_hex), unsafe_allow_html=True)
        return

    col_inicio, col_espacador, col_fechar = st.columns([1, 5, 1])
    with col_inicio:
        if st.button("🏠 INÍCIO", key=f"btn_inicio_{titulo}", use_container_width=True):
            for k in [k for k in st.session_state if k.startswith("_kpi_aberto_") or k.startswith("_kpi_sel_")]:
                st.session_state[k] = False if k.startswith("_kpi_aberto_") else None
            st.session_state.pop("_grafico_fs", None)
            st.rerun()
    with col_fechar:
        if st.button("✕  FECHAR", key=f"btn_fechar_{titulo}", use_container_width=True, type="secondary"):
            for k in [k for k in st.session_state if k.startswith("_kpi_aberto_") or k.startswith("_kpi_sel_")]:
                st.session_state[k] = False if k.startswith("_kpi_aberto_") else None
            st.rerun()

    st.markdown(f"""
    <div style="display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px;
        border-left:6px solid {cor_hex}; padding:clamp(10px,2vw,18px) clamp(12px,2.5vw,24px);
        background:{_hex_to_rgba(cor_hex, 0.22)}; border-radius:0 10px 10px 0; margin-bottom:20px;">
        <div>
            <div style="color:{cor_hex}; font-size:clamp(0.65rem,1vw,0.82rem); font-weight:700; letter-spacing:1.4px; text-transform:uppercase;">CATEGORIA SELECIONADA</div>
            <div style="color:#fff; font-size:clamp(1.1rem,2.5vw,1.7rem); font-weight:800; margin-top:6px;">{titulo}</div>
        </div>
        <div style="background:{cor_hex}; color:#fff; font-size:clamp(1.4rem,3vw,2.4rem); font-weight:900;
            padding:clamp(6px,1vw,12px) clamp(14px,2vw,28px); border-radius:12px; min-width:60px; text-align:center;
            box-shadow: 0 0 18px {_hex_to_rgba(cor_hex, 0.6)};">{len(df_kpi)}</div>
    </div>
    """, unsafe_allow_html=True)

    if df_kpi.empty:
        st.info("Nenhum item encontrado nesta categoria.")
        return

    if "STATUS" in df_kpi.columns and df_kpi["STATUS"].nunique() > 1:
        resumo = df_kpi["STATUS"].value_counts().reset_index()
        resumo.columns = ["STATUS", "QTD"]
        cols_r = st.columns(min(len(resumo), 5))
        for i, (_, row) in enumerate(resumo.iterrows()):
            cor = CORES_STATUS.get(row["STATUS"], "#888")
            with cols_r[i % len(cols_r)]:
                st.markdown(f"""
                <div style="border:1px solid {cor}; border-radius:8px; padding:10px 12px; text-align:center;
                    background:{_hex_to_rgba(cor, 0.22)}; margin-bottom:10px;">
                    <div style="color:{cor}; font-size:1.5rem; font-weight:800;">{row['QTD']}</div>
                    <div style="color:#bbb; font-size:0.62rem; text-transform:uppercase; letter-spacing:0.5px; line-height:1.4; margin-top:2px;">{row['STATUS']}</div>
                </div>""", unsafe_allow_html=True)
        st.divider()

    busca = st.text_input(
        "🔍 Busca rápida", placeholder="Código, descrição, posição...",
        key=key_busca, label_visibility="collapsed"
    )

    df_exibir = df_kpi.copy()
    if busca.strip():
        mask = df_exibir.apply(
            lambda col: col.astype(str).str.upper().str.contains(busca.strip().upper(), na=False)
        ).any(axis=1)
        df_exibir = df_exibir[mask]

    df_exibir      = df_exibir.reset_index(drop=True)
    total_exibindo = len(df_exibir)

    icone = '🔎' if busca.strip() else '📋'
    st.markdown(f"""<span style="font-size:1rem; color:#aaaaaa;">
    {icone} Exibindo <b>{total_exibindo}</b> item(ns) — Clique em <b>▶ ABRIR</b> para ver os detalhes completos
    </span>""", unsafe_allow_html=True)

    if df_exibir.empty:
        st.warning("Nenhum item corresponde ao filtro.")
        return

    NUM_COLS = 4
    for row_start in range(0, total_exibindo, NUM_COLS):
        cols = st.columns(NUM_COLS)
        for col_idx in range(NUM_COLS):
            item_idx = row_start + col_idx
            if item_idx >= total_exibindo:
                break
            item   = df_exibir.iloc[item_idx]
            codigo = str(item.get("CÓDIGO", "—")).strip()
            with cols[col_idx]:
                st.markdown(_html_mini_card(item, cor_override=cor_hex), unsafe_allow_html=True)
                if st.button("▶ ABRIR", key=f"open_{titulo}_{item_idx}_{codigo}", use_container_width=True):
                    st.session_state[key_sel] = codigo
                    st.rerun()


# =====================================================
# PAINÉIS
# =====================================================

def criar_painel_status(df_filtrado):
    df = df_filtrado.copy()
    df["SALDO TOTAL"] = pd.to_numeric(df["SALDO TOTAL"], errors="coerce").fillna(0)

    tem_unit = "VALORES UNITÁRIOS" in df.columns
    tem_vt   = "VALOR TOTAL" in df.columns

    if tem_unit:
        df["_VU"]      = pd.to_numeric(df["VALORES UNITÁRIOS"], errors="coerce").fillna(0)
        df["_VT_CALC"] = df["_VU"] * df["SALDO TOTAL"]
    elif tem_vt:
        df["_VU"]      = 0.0
        vt_col         = pd.to_numeric(df["VALOR TOTAL"], errors="coerce").fillna(0)
        df["_VT_CALC"] = vt_col.where(df["SALDO TOTAL"] > 0, other=0.0)
    else:
        df["_VU"]      = 0.0
        df["_VT_CALC"] = 0.0

    def _soma(mask, zerado=False):
        return float(df.loc[mask, "_VU" if zerado else "_VT_CALC"].sum())

    mask_disp   = df["SALDO TOTAL"] > 3
    mask_alerta = (df["SALDO TOTAL"] > 0) & (df["SALDO TOTAL"] <= 3)
    mask_zerado = df["SALDO TOTAL"] == 0
    mask_total  = df["SALDO TOTAL"] > 0

    vt_total  = _soma(mask_total)
    vt_disp   = _soma(mask_disp)
    vt_alerta = _soma(mask_alerta)
    vt_zerado = _soma(mask_zerado, zerado=True)

    def _fmt(v):
        try:
            return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return "R$ 0,00"

    def _pct(v, total):
        if total <= 0:
            return 2
        return max(2, round((v / total) * 100, 1))

    kpis = [
        ("TOTAL DE ITENS", _fmt(vt_total),  CORES_KPI["TOTAL"]["border"],      CORES_KPI["TOTAL"]["shadow"],      100),
        ("DISPONÍVEIS",    _fmt(vt_disp),   CORES_KPI["DISPONIVEIS"]["border"], CORES_KPI["DISPONIVEIS"]["shadow"], _pct(vt_disp,   vt_total)),
        ("ALERTA",         _fmt(vt_alerta), CORES_KPI["OPERACAO"]["border"],    CORES_KPI["OPERACAO"]["shadow"],    _pct(vt_alerta, vt_total)),
        ("SEM ESTOQUE",    _fmt(vt_zerado), CORES_KPI["MANUTENCAO"]["border"],  CORES_KPI["MANUTENCAO"]["shadow"],  _pct(vt_zerado, vt_total)),
    ]

    rows_html = ""
    for label, valor, cor, sombra, pct in kpis:
        rows_html += f'''
        <div style="display:flex;align-items:center;gap:18px;background:rgba(255,255,255,0.04);
            border:1.5px solid {cor};border-left:6px solid {cor};border-radius:10px;
            padding:14px 20px;margin-bottom:10px;
            box-shadow:0 0 14px {sombra},0 2px 8px rgba(0,0,0,0.4);box-sizing:border-box;">
            <div style="min-width:220px;max-width:220px;">
                <div style="color:{cor};font-size:0.72rem;font-weight:800;text-transform:uppercase;
                    letter-spacing:1.4px;margin-bottom:4px;text-shadow:0 0 8px {sombra};">{label}</div>
                <div style="color:{cor};font-size:1.5rem;font-weight:900;letter-spacing:-0.5px;
                    line-height:1.1;text-shadow:0 0 14px {sombra};">{valor}</div>
            </div>
            <div style="flex:1;">
                <div style="background:rgba(255,255,255,0.07);border-radius:8px;height:22px;width:100%;overflow:hidden;">
                    <div style="height:22px;width:{pct}%;background:linear-gradient(90deg,{cor}cc,{cor});
                        border-radius:8px;box-shadow:0 0 12px {sombra};display:flex;align-items:center;
                        justify-content:flex-end;padding-right:8px;box-sizing:border-box;">
                        <span style="color:#fff;font-size:0.7rem;font-weight:800;white-space:nowrap;">{pct:.1f}%</span>
                    </div>
                </div>
            </div>
        </div>'''

    with st.container(border=True):
        st.markdown('<div class="card-title">💰 VALOR TOTAL POR SITUAÇÃO</div>', unsafe_allow_html=True)
        st.markdown(rows_html, unsafe_allow_html=True)


def criar_painel_posicao(posicao_df):
    with st.container(border=True):
        col_t, col_b = st.columns([11, 1])
        with col_t:
            st.markdown('<div class="card-title">📍 DISTRIBUIÇÃO POR POSIÇÃO</div>', unsafe_allow_html=True)
        with col_b:
            st.markdown('<div class="btn-fullscreen">', unsafe_allow_html=True)
            if st.button("⛶", key="fs_btn_posicao"):
                st.session_state["_grafico_fs"] = "posicao"
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
        if not posicao_df.empty:
            st.plotly_chart(criar_grafico_posicao(posicao_df), use_container_width=True, config={'displayModeBar': False})
        else:
            st.info("Nenhuma posição cadastrada")


def criar_painel_itens(df_filtrado):
    with st.container(border=True):
        col_t, col_b = st.columns([11, 1])
        with col_t:
            st.markdown('<div class="card-title">🚨 SEM ESTOQUE — VALOR POR CATEGORIA</div>', unsafe_allow_html=True)
        with col_b:
            st.markdown('<div class="btn-fullscreen">', unsafe_allow_html=True)
            if st.button("⛶", key="fs_btn_itens"):
                st.session_state["_grafico_fs"] = "itens"
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
        st.plotly_chart(criar_grafico_itens(df_filtrado), use_container_width=True, config={'displayModeBar': False})


def mostrar_grafico_fullscreen(grafico_id, df_filtrado, posicao_df):
    titulos = {
        "posicao": "📍 DISTRIBUIÇÃO POR POSIÇÃO",
        "itens":   "🚨 SEM ESTOQUE — VALOR POR CATEGORIA",
    }
    titulo = titulos.get(grafico_id, "")

    st.markdown("""
    <style>
    section[data-testid="stSidebar"] { display: none !important; }
    .stApp { background: #141414 !important; }
    .main .block-container { padding: 1rem 1.5rem !important; max-width: 100% !important; background: #141414 !important; }
    </style>
    """, unsafe_allow_html=True)

    col_titulo, col_fechar = st.columns([10, 2])
    with col_titulo:
        st.markdown(f'<div class="card-title" style="font-size:1.3rem;">{titulo}</div>', unsafe_allow_html=True)
    with col_fechar:
        if st.button("✕  Fechar", key="fs_fechar", use_container_width=True):
            st.session_state.pop("_grafico_fs", None)
            st.rerun()

    st.markdown("<hr style='border-color:#484848;margin:0 0 12px 0;'>", unsafe_allow_html=True)

    if grafico_id == "posicao":
        fig = criar_grafico_posicao(posicao_df, fullscreen=True)
    elif grafico_id == "itens":
        fig = criar_grafico_itens(df_filtrado, fullscreen=True)
    else:
        st.stop()

    st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': True})


# =====================================================
# SIDEBAR
# =====================================================
def criar_sidebar(loading_placeholder):
    with st.sidebar:
        st.header("🎛️ FILTROS DO ALMOXARIFADO")
        st.divider()

        uploaded_file = st.file_uploader(
            "📁 CARREGAR ARQUIVO", type=['xlsx', 'xls'],
            label_visibility="visible"
        )

        df_base = pd.DataFrame()

        if uploaded_file is not None:
            show_loading_screen(loading_placeholder)
            df_base = load_data_from_file(uploaded_file)
            loading_placeholder.empty()
            if not df_base.empty:
                st.success(f"✅ {len(df_base)} itens carregados!")
             

                # ── Seção de relatório ────────────────────────────────────
                st.divider()
                st.markdown(
                    "<div style='color:#ffb300;font-weight:800;font-size:0.8rem;"
                    "text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;'>"
                    "📊 RELATÓRIO DE ALERTAS</div>",
                    unsafe_allow_html=True
                )

                df_tmp = df_base.copy()
                df_tmp["SALDO TOTAL"] = pd.to_numeric(
                    df_tmp.get("SALDO TOTAL", 0), errors="coerce"
                ).fillna(0)
                df_alerta = df_tmp[
                    (df_tmp["SALDO TOTAL"] > 0) & (df_tmp["SALDO TOTAL"] <= 3)
                ].copy()
                df_zerado = df_tmp[df_tmp["SALDO TOTAL"] == 0].copy()

                n_alerta = len(df_alerta)
                n_zerado = len(df_zerado)

                st.markdown(
                    f"<div style='font-size:0.72rem;color:#aaa;margin-bottom:8px;'>"
                    f"🟠 Alerta: <b style='color:#ffb300;'>{n_alerta}</b> itens &nbsp;|&nbsp; "
                    f"🔴 Zerados: <b style='color:#ff3d00;'>{n_zerado}</b> itens</div>",
                    unsafe_allow_html=True
                )

                # Botão visualizar (abre preview na tela principal)
                if st.button(
                    "👁️ VISUALIZAR RELATÓRIO",
                    key="btn_preview_relatorio",
                    use_container_width=True
                ):
                    st.session_state["_relatorio_preview"] = True
                    st.rerun()

                # Botão baixar
                if n_alerta > 0 or n_zerado > 0:
                    try:
                        from gerar_relatorio import gerar_bytes_relatorio
                        rel_bytes = gerar_bytes_relatorio(df_alerta, df_zerado)
                        st.download_button(
                            label="⬇️ BAIXAR RELATÓRIO (.pdf)",
                            data=rel_bytes,
                            file_name="relatorio_alertas_almoxarifado.pdf",
                            mime="application/pdf",
                            key="btn_download_relatorio",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"Erro ao gerar relatório: {e}")
                else:
                    st.info("✅ Nenhum alerta ou item zerado para exportar.")
        else:
            st.info("⬆️ Faça upload da planilha")

        return df_base


# =====================================================
# PREVIEW DO RELATÓRIO (tela principal)
# =====================================================
def mostrar_preview_relatorio(df_base):
    """Exibe preview formatado dos itens de alerta e zerados na tela principal."""
    st.markdown("""
    <style>
    section[data-testid="stSidebar"] { display: none !important; }
    .main .block-container { padding: 1rem 1.5rem !important; max-width: 100% !important; }
    </style>
    """, unsafe_allow_html=True)

    df_tmp = df_base.copy()
    df_tmp["SALDO TOTAL"] = pd.to_numeric(df_tmp.get("SALDO TOTAL", 0), errors="coerce").fillna(0)
    df_alerta = df_tmp[(df_tmp["SALDO TOTAL"] > 0) & (df_tmp["SALDO TOTAL"] <= 3)].copy()
    df_zerado = df_tmp[df_tmp["SALDO TOTAL"] == 0].copy()

    # Ordena por saldo total decrescente
    if not df_alerta.empty:
        df_alerta = df_alerta.sort_values("SALDO TOTAL", ascending=False)
    if not df_zerado.empty:
        df_zerado = df_zerado.sort_values("SALDO TOTAL", ascending=False)

    col_titulo, col_fechar = st.columns([9, 1])
    with col_titulo:
        st.markdown(
            "<div class='card-title' style='font-size:1.4rem;'>📊 RELATÓRIO DE ALERTAS — PREVIEW</div>",
            unsafe_allow_html=True
        )
    with col_fechar:
        if st.button("✕ FECHAR", key="btn_fechar_preview", use_container_width=True):
            st.session_state.pop("_relatorio_preview", None)
            st.rerun()

    st.markdown("<hr style='border-color:#484848;margin:4px 0 16px 0;'>", unsafe_allow_html=True)

    # Botão de download no topo do preview também
    n_alerta = len(df_alerta)
    n_zerado = len(df_zerado)
    if n_alerta > 0 or n_zerado > 0:
        try:
            from gerar_relatorio import gerar_bytes_relatorio
            rel_bytes = gerar_bytes_relatorio(df_alerta, df_zerado)
            st.download_button(
                label="⬇️ BAIXAR RELATÓRIO (.pdf)",
                data=rel_bytes,
                file_name="relatorio_alertas_almoxarifado.pdf",
                mime="application/pdf",
                key="btn_download_preview_top",
                use_container_width=False
            )
        except Exception as e:
            st.error(f"Erro ao gerar relatório: {e}")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Seção laranja ────────────────────────────────────────────────────
    cor_lrj = CORES_KPI["OPERACAO"]["border"]
    st.markdown(
        f"<div style='border-left:6px solid {cor_lrj};padding:10px 18px;"
        f"background:rgba(255,179,0,0.12);border-radius:0 8px 8px 0;margin-bottom:12px;'>"
        f"<span style='color:{cor_lrj};font-weight:800;font-size:1rem;text-transform:uppercase;letter-spacing:1px;'>"
        f"⚠️ ALERTA — ESTOQUE BAIXO (saldo 1–3)</span>"
        f"<span style='color:#aaa;font-size:0.8rem;margin-left:12px;'>{n_alerta} itens</span></div>",
        unsafe_allow_html=True
    )
    if df_alerta.empty:
        st.info("Nenhum item em alerta.")
    else:
        colunas_show = [c for c in df_alerta.columns if not c.startswith("_")]
        st.dataframe(
            df_alerta[colunas_show].style.apply(aplicar_cor_status, axis=1),
            hide_index=True, use_container_width=True, height=min(400, 40 + n_alerta * 36)
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Seção vermelha ───────────────────────────────────────────────────
    cor_verm = CORES_KPI["MANUTENCAO"]["border"]
    st.markdown(
        f"<div style='border-left:6px solid {cor_verm};padding:10px 18px;"
        f"background:rgba(255,61,0,0.12);border-radius:0 8px 8px 0;margin-bottom:12px;'>"
        f"<span style='color:{cor_verm};font-weight:800;font-size:1rem;text-transform:uppercase;letter-spacing:1px;'>"
        f"🚨 SEM ESTOQUE — ITENS ZERADOS (saldo 0)</span>"
        f"<span style='color:#aaa;font-size:0.8rem;margin-left:12px;'>{n_zerado} itens</span></div>",
        unsafe_allow_html=True
    )
    if df_zerado.empty:
        st.info("Nenhum item zerado. ✅")
    else:
        colunas_show = [c for c in df_zerado.columns if not c.startswith("_")]
        st.dataframe(
            df_zerado[colunas_show].style.apply(aplicar_cor_status, axis=1),
            hide_index=True, use_container_width=True, height=min(500, 40 + n_zerado * 36)
        )


# =====================================================
# MAIN
# =====================================================
def main():
    load_custom_css()

    st.markdown("""<script>
    (function(){
        var style = document.createElement('style');
        style.textContent = '[class*="TooltipContent"],[class*="tooltip"]{background:#1a1a1a!important;color:#fff!important;border:1px solid #555!important;}[class*="TooltipContent"] *{color:#fff!important;}';
        document.head.appendChild(style);
    })();
    </script>""", unsafe_allow_html=True)

    loading_placeholder = st.empty()
    df_base = criar_sidebar(loading_placeholder)

    if df_base.empty:
        st.markdown("""
        <style>
        .centered-warning { display: flex; justify-content: center; align-items: center; min-height: 60vh; text-align: center; }
        .warning-box { background-color: #1e1e1e; border: 2px solid #ff9800; border-radius: 15px; padding: 40px 60px; box-shadow: 0 0 20px rgba(255, 152, 0, 0.3); }
        .warning-icon { font-size: 4rem; margin-bottom: 20px; }
        .warning-text { font-size: 1.3rem; color: #ffffff; font-weight: 600; line-height: 1.6; }
        </style>
        <div class="centered-warning">
            <div class="warning-box">
                <div class="warning-icon">📦</div>
                <div class="warning-text">Carregue um arquivo Excel na barra lateral<br>para visualizar os dados.</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.stop()

    # ── Preview do relatório (ocupa tela inteira) ──────────────────────
    if st.session_state.get("_relatorio_preview"):
        mostrar_preview_relatorio(df_base)
        st.stop()

    df_filtrado = df_base.copy()
    df_filtrado["SALDO TOTAL"] = pd.to_numeric(df_filtrado["SALDO TOTAL"], errors="coerce").fillna(0)

    total    = len(df_filtrado)
    zerados  = len(df_filtrado[df_filtrado["SALDO TOTAL"] == 0])
    pct_disp = ((total - zerados) / total * 100) if total > 0 else 0

    criar_header(pct_disp)
    criar_kpis(df_filtrado)
    st.markdown("<br>", unsafe_allow_html=True)

    # ── Modo fullscreen KPI ──────────────────────────────────────────────
    kpi_map = [
        ("TOTAL",       CORES_KPI["TOTAL"]["border"],      "_kpi_df_TOTAL"),
        ("DISPONIVEIS", CORES_KPI["DISPONIVEIS"]["border"], "_kpi_df_DISPONIVEIS"),
        ("ALERTA",      CORES_KPI["OPERACAO"]["border"],    "_kpi_df_ALERTA"),
        ("ZERADOS",     CORES_KPI["MANUTENCAO"]["border"],  "_kpi_df_ZERADOS"),
    ]
    for nome, cor, df_key in kpi_map:
        key_sel    = f"_kpi_sel_{nome}"
        key_aberto = f"_kpi_aberto_{nome}"
        if (st.session_state.get(key_sel) or st.session_state.get(key_aberto)) and df_key in st.session_state:
            mostrar_detalhes_kpi(nome, cor, st.session_state[df_key])
            st.stop()

    # ── Prepara dados de posição ─────────────────────────────────────────
    posicao_df = pd.DataFrame()
    if "POSIÇÃO" in df_filtrado.columns:
        df_pos = df_filtrado[
            df_filtrado["POSIÇÃO"].notna() &
            (df_filtrado["POSIÇÃO"].astype(str).str.strip() != "")
        ].copy()
        df_pos["SALDO TOTAL"] = pd.to_numeric(df_pos["SALDO TOTAL"], errors="coerce").fillna(0)

        if "VALORES UNITÁRIOS" in df_pos.columns:
            df_pos["_VT"] = pd.to_numeric(df_pos["VALORES UNITÁRIOS"], errors="coerce").fillna(0) * df_pos["SALDO TOTAL"]
        elif "VALOR TOTAL" in df_pos.columns:
            vt = pd.to_numeric(df_pos["VALOR TOTAL"], errors="coerce").fillna(0)
            df_pos["_VT"] = vt.where(df_pos["SALDO TOTAL"] > 0, other=0.0)
        else:
            df_pos["_VT"] = 0.0

        def _fmt_vu(v):
            try:
                return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except Exception:
                return "—"

        tem_desc   = "DESCRIÇÃO" in df_pos.columns
        tem_vu     = "VALORES UNITÁRIOS" in df_pos.columns
        tem_vt_col = "VALOR TOTAL" in df_pos.columns

        def _itens_info(grp):
            partes = []
            for _, r in grp.iterrows():
                saldo = float(r.get("SALDO TOTAL", 0))
                desc  = str(r.get("DESCRIÇÃO", "—"))[:40] if tem_desc else "—"
                vu    = _fmt_vu(r.get("VALORES UNITÁRIOS", 0)) if tem_vu else "—"
                if tem_vu:
                    vt = float(r.get("VALORES UNITÁRIOS", 0)) * saldo
                elif tem_vt_col:
                    vt = float(r.get("VALOR TOTAL", 0)) if saldo > 0 else 0.0
                else:
                    vt = 0.0
                qtd         = int(saldo)
                status_icon = "🔴" if qtd == 0 else "📦"
                partes.append(
                    f"{status_icon} Descrição: {desc}<br>"
                    f"     Quantidade: {qtd}<br>"
                    f"     Valor unitário: {vu}<br>"
                    f"     Valor total: {_fmt_vu(vt)}"
                )
            return "<br>".join(partes)

        try:
            itens_info_map = df_pos.groupby("POSIÇÃO").apply(
                _itens_info, include_groups=False
            ).to_dict()
        except TypeError:
            itens_info_map = df_pos.groupby("POSIÇÃO").apply(_itens_info).to_dict()

        pos_counts = df_pos.groupby("POSIÇÃO").agg(
            QUANTIDADE=("POSIÇÃO", "count"),
            SALDO=("SALDO TOTAL", "sum"),
            VALOR=("_VT", "sum")
        ).reset_index()
        pos_counts["ITENS_INFO"] = pos_counts["POSIÇÃO"].map(itens_info_map).fillna("")
        pos_counts               = pos_counts.sort_values("QUANTIDADE", ascending=False)

        if not pos_counts.empty:
            posicao_df = pos_counts.rename(columns={"SALDO": "SALDO_TOTAL", "VALOR": "VALOR_TOTAL"})
            posicao_df["QUANTIDADE"] = posicao_df["QUANTIDADE"].clip(lower=1)

    # ── Modo fullscreen gráfico ──────────────────────────────────────────
    if st.session_state.get("_grafico_fs"):
        mostrar_grafico_fullscreen(st.session_state["_grafico_fs"], df_filtrado, posicao_df)
        st.stop()

    col1, col2 = st.columns(2)
    with col1:
        criar_painel_status(df_filtrado)
    with col2:
        criar_painel_posicao(posicao_df)

    st.markdown("<br>", unsafe_allow_html=True)
    criar_painel_itens(df_filtrado)
    st.markdown("<br>", unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown('<div class="card-title">📋 DETALHAMENTO DO ESTOQUE</div>', unsafe_allow_html=True)
        colunas = ["CÓDIGO", "DESCRIÇÃO", "STATUS"]
        if "POSIÇÃO" in df_filtrado.columns:
            colunas.append("POSIÇÃO")
        colunas += ["ENTRADA", "SAIDA", "SALDO TOTAL"]
        colunas  = [c for c in colunas if c in df_filtrado.columns]

        df_display = df_filtrado[colunas].copy()
        for col in ["ENTRADA", "SAIDA", "SALDO TOTAL"]:
            if col in df_display.columns:
                df_display[col] = df_display[col].fillna(0).astype(int)

        st.dataframe(
            df_display.style.apply(aplicar_cor_status, axis=1),
            hide_index=True, use_container_width=True, height=400
        )


if __name__ == "__main__":
    main()
