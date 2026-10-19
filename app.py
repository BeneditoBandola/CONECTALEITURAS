import base64
from datetime import date, datetime, timedelta
import difflib
import io
from io import BytesIO
import json
import math
import os
import re
import subprocess
import tempfile
import time
import unicodedata
import urllib.parse
from fpdf import FPDF
import folium
from PIL import Image
import pandas as pd
from streamlit_folium import st_folium
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from pathlib import Path
from playwright.sync_api import sync_playwright
import sys

# --- 1. CONFIGURAÇÕES E CONSTANTES (TELA CHEIA / WIDE) ---
st.set_page_config(
    page_title="Minassal TV Performance - Executivo",
    layout="wide",
    page_icon="🐾",
    initial_sidebar_state="collapsed",
)

CONFIG_FILE = "config_metas_minassal.json"
CONFIG_PRECOS_FILE = "config_precos_minassal.json"
CONFIG_FOTOS_FILE = "config_fotos_promotores.json"

PASTA_PRODUTOS = "fotos"
ARQUIVO_CLIENTES_COORD = "clientes com coordenadas.xlsx"

# --- DEFINIÇÃO DOS PERÍODOS PARA A AUTOMAÇÃO (P10 a P13) ---
PERIODOS = {
    "P10": {"inicio": date(2026, 9, 6), "fim": date(2026, 10, 3)},
    "P11": {"inicio": date(2026, 10, 4), "fim": date(2026, 10, 31)},
    "P12": {"inicio": date(2026, 11, 1), "fim": date(2026, 11, 28)},
    "P13": {"inicio": date(2026, 11, 29), "fim": date(2027, 1, 2)},
}

def obter_datas_periodo(nome_periodo):
    hoje = date.today()
    if nome_periodo in PERIODOS:
        p_inicio = PERIODOS[nome_periodo]["inicio"]
        p_fim_oficial = PERIODOS[nome_periodo]["fim"]
        if hoje < p_inicio:
            p_fim = p_inicio
        elif p_inicio <= hoje <= p_fim_oficial:
            p_fim = hoje
        else:
            p_fim = p_fim_oficial
        return p_inicio.strftime("%Y-%m-%d"), p_fim.strftime("%Y-%m-%d")
    return "2026-09-06", hoje.strftime("%Y-%m-%d")

def baixar_planilha_pdvpet(periodo_escolhido="P11"):
    data_de_str, data_ate_str = obter_datas_periodo(periodo_escolhido)
    
    try:
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], capture_output=True, text=True)
    except Exception:
        pass

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, slow_mo=500)
        context = browser.new_context()
        page = context.new_page()
        try:
            page.goto("https://www.pdvpet.com.br/")
            page.get_by_placeholder("Digite seu CPF/CNPJ").fill("137.318.006-40")
            page.get_by_placeholder("Digite sua senha").fill("Fizzvini1234@")
            page.get_by_role("button", name="Entrar").click()
            page.get_by_role("link", name="Questionários").click()
            page.locator("#DataDe").fill(data_de_str)
            page.locator("#DataAte").fill(data_ate_str)
            page.get_by_role("button", name="Buscar").click()
            time.sleep(2)
            with page.expect_download() as download_info:
                page.get_by_role("button", name="Exportar").click()
                page.get_by_role("link", name="Abrir").click()
            download = download_info.value
            pasta_script = Path(__file__).parent if "__file__" in locals() else Path.cwd()
            caminho_salvar = pasta_script / "relatorio_questionarios.csv"
            download.save_as(caminho_salvar)
        except Exception as e:
            raise Exception(f"Erro na automação Playwright: {e}")
        finally:
            time.sleep(3)
            browser.close()

# --- DICIONÁRIO DE TEMAS DE CORES PARA LEITURA EXECUTIVA ---
TEMAS_CORES = {
    "1. Padrão Atual (Clean Executivo)": {
        "bg_app": "#F1F5F9",
        "bg_header": "#F1F5F9",
        "texto": "#0F172A",
        "primario": "#2563EB",
        "secundario": "#1E3A8A",
        "card_bg": "#FFFFFF",
        "borda": "#CBD5E1",
    },
    "2. Midnight Dark (Modo Noturno)": {
        "bg_app": "#0F172A",
        "bg_header": "#0F172A",
        "texto": "#F8FAFC",
        "primario": "#3B82F6",
        "secundario": "#60A5FA",
        "card_bg": "#1E293B",
        "borda": "#334155",
    },
    "3. Slate Moderno (Cinza Metálico)": {
        "bg_app": "#E2E8F0",
        "bg_header": "#E2E8F0",
        "texto": "#1E293B",
        "primario": "#475569",
        "secundario": "#0F172A",
        "card_bg": "#FFFFFF",
        "borda": "#94A3B8",
    },
    "4. Emerald Forest (Verde Esmeralda)": {
        "bg_app": "#ECFDF5",
        "bg_header": "#ECFDF5",
        "texto": "#064E3B",
        "primario": "#059669",
        "secundario": "#065F46",
        "card_bg": "#FFFFFF",
        "borda": "#A7F3D0",
    },
    "5. Ocean Blue (Azul Oceano Profundo)": {
        "bg_app": "#EFF6FF",
        "bg_header": "#EFF6FF",
        "texto": "#1E3A8A",
        "primario": "#1D4ED8",
        "secundario": "#1E40AF",
        "card_bg": "#FFFFFF",
        "borda": "#BFDBFE",
    },
    "6. Sunset Warm (Quente / Âmbar)": {
        "bg_app": "#FFFBEB",
        "bg_header": "#FFFBEB",
        "texto": "#78350F",
        "primario": "#D97706",
        "secundario": "#92400E",
        "card_bg": "#FFFFFF",
        "borda": "#FDE68A",
    },
    "7. Royal Purple (Roxo Realeza)": {
        "bg_app": "#F5F3FF",
        "bg_header": "#F5F3FF",
        "texto": "#4C1D95",
        "primario": "#7C3AED",
        "secundario": "#5B21B6",
        "card_bg": "#FFFFFF",
        "borda": "#DDD6FE",
    },
    "8. Crimson Executive (Vinho / Executivo)": {
        "bg_app": "#FFF1F2",
        "bg_header": "#FFF1F2",
        "texto": "#881337",
        "primario": "#E11D48",
        "secundario": "#9F1239",
        "card_bg": "#FFFFFF",
        "borda": "#FECDD3",
    },
    "9. Cyberpunk Neon (Contraste Máximo)": {
        "bg_app": "#030712",
        "bg_header": "#030712",
        "texto": "#F9FAFB",
        "primario": "#06B6D4",
        "secundario": "#38BDF8",
        "card_bg": "#111827",
        "borda": "#1F2937",
    },
    "10. Coffee Minimal (Minimalista Café)": {
        "bg_app": "#FDFBF7",
        "bg_header": "#FDFBF7",
        "texto": "#3D2914",
        "primario": "#8C6239",
        "secundario": "#593E25",
        "card_bg": "#FFFFFF",
        "borda": "#E6DCCD",
    },
}

# --- DICIONÁRIO DE CÓDIGOS MINASSAL PARA OS PRODUTOS SMALL BAGS ---
MAPA_CODIGOS_MINASSAL = {
    "98985": "CHAMP ADULTO CARNE & CEREAL 900G",
    "98980": "CHAMP FILHOTES 900G",
    "98679": "KITEKAT DRY AD MIX DE CARNES 900G",
    "98989": "PED NUTRICAO ESSENCIAL DRY AD AO LEITE 900G",
    "98982": "PED NUTRICAO ESSENCIAL DRY AD CARNE 900G",
    "98933": "PEDIGREE DRY AD CARNE&FRANG 2,7KG",
    "98934": "PEDIGREE DRY AD CARNE&FRANG 900G",
    "98914": "PEDIGREE DRY AD CARNE&VEG 2,7KG",
    "98915": "PEDIGREE DRY AD CARNE&VEG 900G",
    "98911": "PEDIGREE DRY AD RP 2,7KG",
    "98912": "PEDIGREE DRY AD RP 900G",
    "98930": "PEDIGREE DRY FIL 2,7KG",
    "98931": "PEDIGREE DRY FIL 900G",
    "98899": "WHI DRY AD CARNE 500G",
    "98898": "WHI DRY AD CARNE 900G",
    "98941": "WHI DRY AD FRANGO 900G",
    "98937": "WHI DRY AD PEIXE 500G",
    "98936": "WHI DRY AD PEIXE 900G",
    "98944": "WHI DRY FILHOTES CARNE 500G",
    "98942": "WHI DRY FILHOTES CARNE 900G",
    "98903": "WHI DRY GATO CAST CARNE 500G",
    "98902": "WHI DRY GATO CAST CARNE 900G",
    "98946": "WHI DRY GATOS CAST PEIXE 900G",
}

PERIODOS_CONFIG = {
    "P10": {
        "inicio": date(2026, 9, 6),
        "fim": date(2026, 10, 3),
        "label": "P10 (06/09 - 03/10)",
    },
    "P11": {
        "inicio": date(2026, 10, 4),
        "fim": date(2026, 10, 31),
        "label": "P11 (04/10 - 31/10)",
    },
    "P12": {
        "inicio": date(2026, 11, 1),
        "fim": date(2026, 11, 28),
        "label": "P12 (01/11 - 28/11)",
    },
    "P13": {
        "inicio": date(2026, 11, 29),
        "fim": date(2027, 1, 2),
        "label": "P13 (29/11 - 02/01)",
    },
}

FILIAIS = [
    "Minassal Ltda - Pocos de Caldas",
    "Minassal Ltda - Sao Joao da Boa Vista",
    "Minassal Ltda - Sao Jose do Rio Preto",
    "Minassal Ltda - Juiz de Fora",
]

CORES_FUNDO_FILIAIS = {
    "Minassal Ltda - Pocos de Caldas": "#F0FDF4",
    "Minassal Ltda - Sao Joao da Boa Vista": "#FDF2F8",
    "Minassal Ltda - Sao Jose do Rio Preto": "#EFF6FF",
    "Minassal Ltda - Juiz de Fora": "#FEFCE8",
}

ITENS_P = ["Small Bags", "Ponto Extra de Sachês/Petiscos"]
ITENS_S = [
    "Combos Virtuais Sachês",
    "Combos Virtuais Petiscos",
    "Sheba Creamy - Leve 2 Pague 1",
]
TODOS_ITENS = ITENS_P + ITENS_S
ITENS_RODAPE = ["Combos Virtuais Sachês (KiteKat, Champ, Optimum)"]

METAS_PADRAO = {
    "Minassal Ltda - Pocos de Caldas": {
        "Small Bags": 110,
        "Ponto Extra de Sachês/Petiscos": 29,
        "Combos Virtuais Sachês": 10,
        "Combos Virtuais Petiscos": 10,
        "Sheba Creamy - Leve 2 Pague 1": 10,
    },
    "Minassal Ltda - Sao Joao da Boa Vista": {
        "Small Bags": 345,
        "Ponto Extra de Sachês/Petiscos": 55,
        "Combos Virtuais Sachês": 10,
        "Combos Virtuais Petiscos": 10,
        "Sheba Creamy - Leve 2 Pague 1": 10,
    },
    "Minassal Ltda - Sao Jose do Rio Preto": {
        "Small Bags": 233,
        "Ponto Extra de Sachês/Petiscos": 47,
        "Combos Virtuais Sachês": 10,
        "Combos Virtuais Petiscos": 10,
        "Sheba Creamy - Leve 2 Pague 1": 10,
    },
    "Minassal Ltda - Juiz de Fora": {
        "Small Bags": 94,
        "Ponto Extra de Sachês/Petiscos": 20,
        "Combos Virtuais Sachês": 10,
        "Combos Virtuais Petiscos": 10,
        "Sheba Creamy - Leve 2 Pague 1": 10,
    },
}

REGRAS_PRECOS_PADRAO = {
    "KiteKat Adulto Mix de Carnes - Small Bags 0,9Kg": {
        "min": 11.90,
        "max": 11.90,
        "nao_participante": False,
    },
    "Whiskas Sabor CARNE - Small Bags 0,5Kg": {
        "min": 13.90,
        "max": 14.90,
        "nao_participante": False,
    },
    "Whiskas Sabor CARNE CASTRADOS - Small Bags 0,5Kg": {
        "min": 13.90,
        "max": 14.90,
        "nao_participante": False,
    },
    "Whiskas Sabor PEIXE - Small Bags 0,5Kg": {
        "min": 13.90,
        "max": 14.90,
        "nao_participante": False,
    },
    "Whiskas Sabor FILHOTE CARNE - Small Bags 0,5Kg": {
        "min": 13.90,
        "max": 14.90,
        "nao_participante": False,
    },
    "Banana - Biscrok 500g": {
        "min": 0.00,
        "max": 0.00,
        "nao_participante": True,
    },
    "Maçã - Biscrok 500g": {"min": 0.00, "max": 0.00, "nao_participante": True},
    "Champ Adulto Carne e Cereal - Small Bags 0,9Kg": {
        "min": 11.90,
        "max": 11.90,
        "nao_participante": False,
    },
    "Champ Filhotes Carne e Cereal - Small Bags 0,9Kg": {
        "min": 11.90,
        "max": 11.90,
        "nao_participante": False,
    },
    "Pedigree CARNE, FRANGO E CEREAIS - Small Bags 0,9Kg": {
        "min": 18.90,
        "max": 19.90,
        "nao_participante": False,
    },
    "Pedigree FILHOTES - Small Bags 0,9Kg": {
        "min": 18.90,
        "max": 19.90,
        "nao_participante": False,
    },
    "Pedigree RG CARNE E VEGETAIS - Small Bags 0,9Kg": {
        "min": 18.90,
        "max": 19.90,
        "nao_participante": False,
    },
    "Pedigree RP CARNE E VEGETAIS - Small Bags 0,9Kg": {
        "min": 18.90,
        "max": 19.90,
        "nao_participante": False,
    },
    "Pedigree NE AO LEITE - Small Bags 0,9Kg": {
        "min": 14.90,
        "max": 15.90,
        "nao_participante": False,
    },
    "Pedigree NE Carne - Small Bags 0,9Kg": {
        "min": 14.90,
        "max": 15.90,
        "nao_participante": False,
    },
    "Whiskas Sabor CARNE - Small Bags 0,9Kg": {
        "min": 19.90,
        "max": 20.90,
        "nao_participante": False,
    },
    "Whiskas Sabor CARNE CASTRADOS - Small Bags 0,9Kg": {
        "min": 19.90,
        "max": 20.90,
        "nao_participante": False,
    },
    "Whiskas Sabor FRANGO - Small Bags 0,9Kg": {
        "min": 19.90,
        "max": 20.90,
        "nao_participante": False,
    },
    "Whiskas Sabor PEIXE - Small Bags 0,9Kg": {
        "min": 19.90,
        "max": 20.90,
        "nao_participante": False,
    },
    "Whiskas Sabor PEIXE CASTRADOS - Small Bags 0,9Kg": {
        "min": 19.90,
        "max": 20.90,
        "nao_participante": False,
    },
    "Whiskas Sabor FILHOTE CARNE - Small Bags 0,9Kg": {
        "min": 19.90,
        "max": 20.90,
        "nao_participante": False,
    },
    "Pedigree CARNE, FRANGO E CEREAIS - Small Bags 2,7Kg": {
        "min": 49.90,
        "max": 54.90,
        "nao_participante": False,
    },
    "Pedigree FILHOTES - Small Bags 2,7Kg": {
        "min": 49.90,
        "max": 54.90,
        "nao_participante": False,
    },
    "Pedigree RG CARNE E VEGETAIS - Small Bags 2,7Kg": {
        "min": 49.90,
        "max": 54.90,
        "nao_participante": False,
    },
    "Pedigree RP CARNE E VEGETAIS - Small Bags 2,7Kg": {
        "min": 49.90,
        "max": 54.90,
        "nao_participante": False,
    },
}

def carregar_mapeamento_fotos():
    if os.path.exists(CONFIG_FOTOS_FILE):
        try:
            with open(CONFIG_FOTOS_FILE, "r", encoding="utf-8") as f:
                dados_b64 = json.load(f)
                mapeamento_carregado = {}
                for promotor, b64_str in dados_b64.items():
                    img_bytes = base64.b64decode(b64_str)
                    mapeamento_carregado[promotor] = Image.open(BytesIO(img_bytes))
                return mapeamento_carregado
        except:
            return {}
    return {}

def salvar_mapeamento_fotos(map_fotos):
    try:
        dados_b64 = {}
        for promotor, img_obj in map_fotos.items():
            if isinstance(img_obj, Image.Image):
                buffered = BytesIO()
                img_obj.save(buffered, format="PNG")
                dados_b64[promotor] = base64.b64encode(buffered.getvalue()).decode("utf-8")
        with open(CONFIG_FOTOS_FILE, "w", encoding="utf-8") as f:
            json.dump(dados_b64, f, indent=4)
    except Exception as e:
        st.error(f"Erro ao salvar mapeamento de fotos: {e}")

if "pasta_fotos" not in st.session_state:
    st.session_state["pasta_fotos"] = "fotos"

if "mapeamento_manual_fotos" not in st.session_state:
    st.session_state["mapeamento_manual_fotos"] = carregar_mapeamento_fotos()

if "periodo_ativo" not in st.session_state:
    st.session_state["periodo_ativo"] = "P11"

def filtrar_por_filial(df_base, nome_filial, col_f):
    nome_curto = nome_filial.split("-")[-1].strip()
    if "Sao Joao da Boa Vista" in nome_filial:
        return df_base[df_base[col_f].str.strip() == "MINASSAL LTDA"]
    else:
        return df_base[df_base[col_f].str.contains(nome_curto, case=False, na=False)]

def carregar_dados(caminho_arquivo, periodo_sel):
    try:
        if not os.path.exists(caminho_arquivo):
            return None
        if os.path.getsize(caminho_arquivo) == 0:
            return pd.DataFrame()

        df = pd.read_csv(caminho_arquivo, sep=";", encoding="latin1")
        if len(df.columns) < 2:
            df = pd.read_csv(caminho_arquivo, sep=",", encoding="utf-8")

        if df.empty or len(df.columns) < 2:
            return pd.DataFrame()

        df.columns = df.columns.str.strip()
        mapa = {
            "Filial": "Distribuidor",
            "Promotor": "Usuario",
            "Pergunta": "Item",
            "Resposta": "Status",
            "Question": "Item",
            "Answer": "Status",
            "User": "Usuario",
            "CnpjCpfPdv": "Cnpj_Documento",
        }
        df.rename(columns=mapa, inplace=True)

        if "Item" in df.columns:
            df["Item"] = df["Item"].astype(str).str.strip()
            df["Item"] = df["Item"].replace({
                "Sheba Cremoso - Leve 2 Pague 1": "Sheba Creamy - Leve 2 Pague 1",
                "Sheba Cremoso - Leve 2 pague 1": "Sheba Creamy - Leve 2 Pague 1",
                "Ação de Petiscos": "Ação de Petiscos",
                "Ponto Extra de Sachês/Petiscos": "Ponto Extra de Sachês/Petiscos",
                "Combos Virtuais Sachês": "Combos Virtuais Sachês",
            })
            df["Item"] = df["Item"].apply(
                lambda x: "Combos Virtuais Sachês" if "Combos Virtuais" in x and "Sach" in x and "Optimum" not in x else x
            )
            df["Item"] = df["Item"].apply(
                lambda x: "Combos Virtuais Sachês (KiteKat, Champ, Optimum)" if "Combos Virtuais" in x and "Kite" in x else x
            )
            df["Item"] = df["Item"].apply(
                lambda x: "Combos Virtuais Petiscos" if "Combos Virtuais" in x and "Petisco" in x else x
            )

        if "Data" in df.columns:
            df["Data_Calc"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True).dt.date
            p_info = PERIODOS_CONFIG.get(periodo_sel, PERIODOS_CONFIG["P11"])
            data_inicio_per = p_info["inicio"]
            data_fim_per = p_info["fim"]

            if date.today() < data_fim_per and date.today() >= data_inicio_per:
                data_fim_per = date.today()

            df = df[(df["Data_Calc"] >= data_inicio_per) & (df["Data_Calc"] <= data_fim_per)].copy()

        return df
    except Exception as e:
        return pd.DataFrame()

def carregar_configuracoes():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("metas", METAS_PADRAO)
        except:
            return METAS_PADRAO
    return METAS_PADRAO

def carregar_precos():
    if os.path.exists(CONFIG_PRECOS_FILE):
        try:
            with open(CONFIG_PRECOS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return REGRAS_PRECOS_PADRAO
    return REGRAS_PRECOS_PADRAO

def salvar_precos(precos):
    try:
        with open(CONFIG_PRECOS_FILE, "w", encoding="utf-8") as f:
            json.dump(precos, f, indent=4)
    except Exception as e:
        st.error(f"Erro ao salvar preços: {e}")

def obter_foto_promotor(nome):
    img_padrao = "https://cdn-icons-png.flaticon.com/512/3135/3135715.png"
    if not isinstance(nome, str):
        return img_padrao

    if nome in st.session_state["mapeamento_manual_fotos"]:
        return st.session_state["mapeamento_manual_fotos"][nome]

    pasta_alvo = st.session_state["pasta_fotos"]
    if not os.path.exists(pasta_alvo):
        return img_padrao

    nome_limpo_orig = nome.strip()
    nome_limpo_baixo = nome_limpo_orig.lower()

    for ext in [".jpg", ".jpeg", ".png", ".JPG", ".PNG"]:
        teste = os.path.join(pasta_alvo, nome_limpo_orig + ext)
        if os.path.exists(teste):
            return Image.open(teste)

    try:
        arquivos = os.listdir(pasta_alvo)
        for arq in arquivos:
            nome_arq_base = os.path.splitext(arq)[0].lower()
            if nome_arq_base == nome_limpo_baixo:
                return Image.open(os.path.join(pasta_alvo, arq))
    except:
        pass

    try:
        nomes_arquivos = {
            os.path.splitext(arq)[0].lower(): arq
            for arq in arquivos
            if arq.lower().endswith((".jpg", ".jpeg", ".png", ".JPG", ".PNG"))
        }
        if nomes_arquivos:
            chaves = list(nomes_arquivos.keys())
            match = difflib.get_close_matches(nome_limpo_baixo, chaves, n=1, cutoff=0.8)
            if match:
                arquivo_encontrado = nomes_arquivos[match[0]]
                return Image.open(os.path.join(pasta_alvo, arquivo_encontrado))
    except:
        pass

    return img_padrao

def obter_codigo_produto(nome_produto):
    n_baixo = str(nome_produto).lower()
    if "champ" in n_baixo and "filhote" in n_baixo:
        return "98980"
    elif "champ" in n_baixo:
        return "98985"
    elif "kitekat" in n_baixo:
        return "98679"
    elif "pedigree" in n_baixo or "ped" in n_baixo:
        if "leite" in n_baixo:
            return "98989"
        elif "2,7" in n_baixo or "2.7" in n_baixo:
            if "frang" in n_baixo:
                return "98933"
            elif "veg" in n_baixo or "rp" in n_baixo:
                return "98914"
            elif "fil" in n_baixo:
                return "98930"
            elif "rg" in n_baixo:
                return "98914"
        else:
            if "rp" in n_baixo or "vegetais" in n_baixo:
                return "98915"
            elif "carne, frango" in n_baixo or "carne&frang" in n_baixo or "frang" in n_baixo:
                return "98934"
            elif "veg" in n_baixo:
                return "98915"
            elif "fil" in n_baixo:
                return "98931"
            elif "ne" in n_baixo and "carne" in n_baixo:
                return "98982"
    elif "whiskas" in n_baixo or "whi" in n_baixo:
        if "0,5" in n_baixo or "0.5" in n_baixo:
            if "carne" in n_baixo and "cast" in n_baixo:
                return "98903"
            elif "carne" in n_baixo and "filhote" in n_baixo:
                return "98944"
            elif "carne" in n_baixo:
                return "98899"
            elif "peixe" in n_baixo:
                return "98937"
        else:
            if "carne" in n_baixo and "cast" in n_baixo:
                return "98902"
            elif "carne" in n_baixo:
                return "98898"
            elif "frango" in n_baixo:
                return "98941"
            elif "peixe" in n_baixo and "cast" in n_baixo:
                return "98946"
            elif "peixe" in n_baixo:
                return "98936"
            elif "filhote" in n_baixo:
                return "98942"
    return None

def obter_foto_produto_por_nome(nome_produto):
    pasta_dir = "fotos"
    if not os.path.exists(pasta_dir):
        return None
    codigo_encontrado = obter_codigo_produto(nome_produto)
    if codigo_encontrado:
        for ext in [".png", ".jpg", ".jpeg", ".webp", ".PNG", ".JPG"]:
            caminho_teste = os.path.join(pasta_dir, codigo_encontrado + ext)
            if os.path.exists(caminho_teste):
                try:
                    return Image.open(caminho_teste)
                except:
                    pass
    return None

def renderizar_barra(percentual, cor):
    try:
        p = float(percentual)
    except:
        p = 0.0
    if p > 100.0:
        p = 100.0
    if p < 0.0:
        p = 0.0
    return f"<div style='width: 100%; background-color: #CBD5E1; border-radius: 4px; margin: 4px 0; height: 14px; position: relative;'><div style='width: {p}%; background-color: {cor}; height: 100%; border-radius: 4px;'></div></div>"

# --- BOTÕES DE SELEÇÃO DE PERÍODO NO TOPO ---
st.markdown("### Seleção de Período Vigente")
col_p1, col_p2, col_p3, col_p4 = st.columns(4)

with col_p1:
    if st.button("P10 (06/09 - 03/10)", use_container_width=True, key="p10"):
        st.session_state["periodo_ativo"] = "P10"
        st.rerun()
with col_p2:
    if st.button("P11 (04/10 - 31/10)", use_container_width=True, key="p11"):
        st.session_state["periodo_ativo"] = "P11"
        st.rerun()
with col_p3:
    if st.button("P12 (01/11 - 28/11)", use_container_width=True, key="p12"):
        st.session_state["periodo_ativo"] = "P12"
        st.rerun()
with col_p4:
    if st.button("P13 (29/11 - 02/01)", use_container_width=True, key="p13"):
        st.session_state["periodo_ativo"] = "P13"
        st.rerun()

periodo_atual = st.session_state["periodo_ativo"]

# --- SELETOR DE TEMA DIRETO NO TOPO ---
st.markdown("---")
col_t1, col_t2 = st.columns([2, 5])
with col_t1:
    st.markdown("<b style='font-size: 14px;'>🎨 Tema do Painel:</b>", unsafe_allow_html=True)
with col_t2:
    nome_tema_selecionado = st.selectbox(
        "Selecione o Tema:",
        list(TEMAS_CORES.keys()),
        index=0,
        label_visibility="collapsed",
    )

t = TEMAS_CORES[nome_tema_selecionado]

# --- APLICAÇÃO DINÂMICA DO TEMA ESCOLHIDO VIA CSS ---
st.markdown(
    f"""
    <style>
        [data-testid="stAppViewContainer"] {{ background-color: {t['bg_app']} !important; }}
        [data-testid="stHeader"] {{ background-color: {t['bg_header']} !important; }}
        [data-testid="stSidebar"] {{ display: none; }}
        h1, h2, h3, h4, p, label, strong, b, span {{ color: {t['texto']} !important; font-family: 'Segoe UI', sans-serif; }}
        input, textarea, select {{
            background-color: {t['card_bg']} !important;
            color: {t['texto']} !important;
            border: 1px solid {t['borda']} !important;
        }}
        div[data-baseweb="select"] > div {{
            background-color: {t['card_bg']} !important;
            color: {t['texto']} !important;
            border: 1px solid {t['borda']} !important;
        }}
        div.stButton > button {{
            background-color: {t['primario']} !important;
            color: #FFFFFF !important;
            border: 1px solid {t['secundario']} !important;
            border-radius: 8px !important;
            font-weight: 700 !important;
            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        }}
        div.stButton > button p {{
            color: #FFFFFF !important;
            font-weight: 700 !important;
        }}
        div.stButton > button:hover {{
            background-color: {t['secundario']} !important;
            border-color: {t['secundario']} !important;
            color: #FFFFFF !important;
        }}
        div.stButton > button:hover p {{
            color: #FFFFFF !important;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f"<div style='text-align: center; font-size: 14px; font-weight: bold; color: {t['secundario']}; margin-top: 5px; margin-bottom: 10px;'>Periodo Ativo no Painel: <u>{periodo_atual}</u></div>",
    unsafe_allow_html=True,
)
st.markdown("---")

caminho_csv = "relatorio_questionarios.csv"
horario_arquivo_str = "Arquivo não encontrado"
if os.path.exists(caminho_csv):
    timestamp_mod = os.path.getmtime(caminho_csv)
    dt_ajustada = datetime.fromtimestamp(timestamp_mod) - timedelta(hours=3)
    horario_arquivo_str = dt_ajustada.strftime("%d/%m/%Y às %H:%M:%S")

df = carregar_dados(caminho_csv, periodo_atual) if os.path.exists(caminho_csv) else None
metas_dict = carregar_configuracoes()
precos_dict = carregar_precos()

col_filial, col_promotor, col_item, col_status = "Distribuidor", "Usuario", "Item", "Status"

with st.expander("Gerenciar e Enviar Fotos dos Promotores", expanded=False):
    st.markdown(f"**Pasta padrão atual:** `{st.session_state['pasta_fotos']}`")
    nova_pasta_input = st.text_input("Alterar pasta padrão de fotos:", value=st.session_state["pasta_fotos"])
    if st.button("Definir Pasta Padrão", key="btn_pasta"):
        if os.path.exists(nova_pasta_input):
            st.session_state["pasta_fotos"] = nova_pasta_input
            st.success("Pasta padrão atualizada com sucesso!")
            time.sleep(0.5)
            st.rerun()
        else:
            st.error("O caminho informado não existe.")

    st.markdown("---")
    st.markdown("### Enviar Foto Direta para um Promotor (Permanente)")
    if df is not None and not df.empty:
        lista_promotores_geral = sorted(df[col_promotor].dropna().unique().tolist())
        promotor_selecionado_upload = st.selectbox("Selecione o Promotor para carregar a foto:", lista_promotores_geral)
        arquivo_enviado = st.file_uploader("Escolha a imagem (PNG, JPG):", type=["png", "jpg", "jpeg"])
        if arquivo_enviado is not None:
            if st.button("Salvar e Memorizar Foto para este Promotor", key="btn_foto_prom"):
                imagem_pil = Image.open(arquivo_enviado)
                st.session_state["mapeamento_manual_fotos"][promotor_selecionado_upload] = imagem_pil
                salvar_mapeamento_fotos(st.session_state["mapeamento_manual_fotos"])
                st.success(f"Foto salva e memorizada com sucesso para {promotor_selecionado_upload}!")
                time.sleep(0.5)
                st.rerun()

# --- BARRA SUPERIOR SINCRONIZADA E HORIZONTAL ---
abas_topo = [
    "📺 Poços de Caldas",
    "📺 São João da Boa Vista",
    "📺 São José do Rio Preto",
    "📺 Juiz de Fora",
    "🏆 Ranking de Leituras",
    "🚨 Registros Reprovados",
    "💰 Auditoria de Preços",
    "🗺️ Mapa de Clientes",
    "📊 Presença por Sabores",
]

col_nav, col_btn1, col_btn2 = st.columns([8, 2, 2])
with col_nav:
    pagina_selecionada = st.radio("Navegação:", abas_topo, horizontal=True, label_visibility="collapsed")

with col_btn1:
    if st.button(f"📥 Baixar Dados ({periodo_atual})", use_container_width=True, key="btn_baixar"):
        with st.spinner(f"Executando automação Playwright para o período {periodo_atual}..."):
            try:
                baixar_planilha_pdvpet(periodo_atual)
                st.success("Dados atualizados com sucesso via automação integrada!")
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"Erro na automação: {e}")

with col_btn2:
    if st.button("🔄 Atualizar Tela", use_container_width=True, key="btn_atualizar"):
        st.rerun()

st.markdown(
    f"<div style='text-align: right; font-size: 11px; color: #64748B; margin-top: -5px; margin-bottom: 10px;'>Última atualização do arquivo: <b>{horario_arquivo_str}</b></div>",
    unsafe_allow_html=True,
)

if df is not None and not df.empty:
    col_preco = "PrecoKg" if "PrecoKg" in df.columns else "Preco" if "Preco" in df.columns else col_status
    df[col_status] = df[col_status].fillna("")

    if "OpcaoEmbalagem" in df.columns and "Embalagem" in df.columns:
        df["Op_Clean"] = df["OpcaoEmbalagem"].fillna("").astype(str).str.strip()
        df["Emb_Clean"] = df["Embalagem"].fillna("").astype(str).str.strip()
        df["Produto_Preco"] = df.apply(
            lambda r: f"{r['Op_Clean']} - {r['Emb_Clean']}" if r["Op_Clean"] and r["Emb_Clean"] and r["Op_Clean"] != "nan" else str(r.get(col_item, "")),
            axis=1,
        )
    else:
        df["Produto_Preco"] = df[col_item]

    def checar_preco_completo(row):
        item = row["Produto_Preco"]
        val_str = str(row[col_preco]).replace(",", ".")
        regra = precos_dict.get(item)
        status_atual = str(row[col_status]).lower()

        if any(c in status_atual for c in ["cancelado", "cancelada"]):
            return "🟡 Preço Corrigido (Cancelado)"
        if not regra or regra.get("nao_participante", False):
            return "🔵 Não Participante"
        try:
            preco_lido = float(val_str)
            if preco_lido >= regra["min"] and preco_lido <= regra["max"]:
                return "🟢 Dentro da Faixa"
            elif preco_lido < regra["min"]:
                return "🟢 Mais Barato (Aprovado)"
            else:
                return "🔴 Fora (Acima do Teto)"
        except:
            return "⚪ Erro/Não Numérico"

    df["Auditoria_Preco"] = df.apply(checar_preco_completo, axis=1)
    df["Preco_Caro_Ativo"] = df["Auditoria_Preco"] == "🔴 Fora (Acima do Teto)"

    def renderizar_tela_tv(filial_selecionada):
        nome_curto = filial_selecionada.split("-")[-1].strip()
        chave_toggle = f"modo_hoje_{filial_selecionada}"
        if chave_toggle not in st.session_state:
            st.session_state[chave_toggle] = False

        col_tit_modo, col_btn_modo = st.columns([5, 2])
        with col_btn_modo:
            ativo_hoje = st.toggle("Leituras de Hoje (Agrupado por Promotor)", value=st.session_state[chave_toggle], key=chave_toggle)

        df_f = filtrar_por_filial(df, filial_selecionada, col_filial)

        if ativo_hoje:
            st.markdown(
                f"<h3 style='color:{t['secundario']}; margin-top: 10px;'>Leituras Feitas Hoje (Agrupadas por Promotor) — {nome_curto.upper()}</h3>",
                unsafe_allow_html=True,
            )
            hoje = date.today()
            df_hoje_filtrado = df_f[df_f["Data_Calc"] == hoje]

            if not df_hoje_filtrado.empty:
                promotores_hoje = sorted(df_hoje_filtrado[col_promotor].dropna().unique().tolist())
                for prom in promotores_hoje:
                    df_p = df_hoje_filtrado[df_hoje_filtrado[col_promotor] == prom]
                    foto_prom = obter_foto_promotor(prom)
                    contagem_itens = df_p[col_item].value_counts()

                    with st.container():
                        st.markdown(
                            f"<div style='background-color: {t['card_bg']}; border: 1px solid {t['borda']}; padding: 16px; border-radius: 12px; margin-bottom: 16px; box-shadow: 0 2px 6px rgba(0,0,0,0.05);'>",
                            unsafe_allow_html=True,
                        )
                        c_img, c_resumo = st.columns([1, 4])
                        with c_img:
                            st.image(foto_prom, width=130)
                            st.markdown(
                                f"<div style='font-size: 15px; font-weight: bold; color: {t['texto']}; text-align: center; margin-top: 6px;'>{prom}</div>",
                                unsafe_allow_html=True,
                            )
                        with c_resumo:
                            st.markdown(
                                f"<div style='font-size: 14px; font-weight: bold; color: {t['secundario']}; margin-bottom: 6px;'>Quantidade de Leituras por Item (KPI):</div>",
                                unsafe_allow_html=True,
                            )
                            badges_html = "<div style='display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 12px;'>"
                            for item_nome, qtd in contagem_itens.items():
                                badges_html += f"<span style='background-color: {t['bg_app']}; border: 1px solid {t['borda']}; color: {t['primario']}; padding: 6px 12px; border-radius: 20px; font-size: 13px; font-weight: 600;'><b>{qtd}x</b> {item_nome}</span>"
                            badges_html += "</div>"
                            st.markdown(badges_html, unsafe_allow_html=True)

                            st.markdown(
                                f"<div style='font-size: 13px; font-weight: bold; color: {t['texto']}; margin-bottom: 4px;'>Lojas, Cidades e Status Registrados:</div>",
                                unsafe_allow_html=True,
                            )
                            for _, r in df_p.iterrows():
                                l_nome = r.get("Pdv", "N/A")
                                c_nome = r.get("Cidade", "N/A")
                                s_nome = r.get(col_status, "N/A")
                                i_nome = r.get(col_item, "N/A")
                                st.markdown(
                                    f"<div style='font-size: 12px; color: {t['texto']}; margin-bottom: 2px;'>• <b>{i_nome}</b> | Loja: <b>{l_nome}</b> | Cidade: <b>{c_nome}</b> | Status: <b>{s_nome}</b></div>",
                                    unsafe_allow_html=True,
                                )
                        st.markdown("</div>", unsafe_allow_html=True)
            else:
                st.info(f"Nenhuma leitura registrada hoje para {nome_curto}.")
            return

        metas_f = metas_dict.get(filial_selecionada, METAS_PADRAO.get(filial_selecionada, {}))
        tot_aprovados_f = len(df_f[df_f[col_status].str.contains("Aprovado|Conforme|Sim", case=False, na=False) & df_f[col_item].isin(TODOS_ITENS)])
        tot_pendentes_f = len(df_f[df_f[col_status].str.contains("Pendente", case=False, na=False) & df_f[col_item].isin(TODOS_ITENS)])
        tot_reprovados_f = len(df_f[df_f[col_status].str.contains("Reprovado|Não|Nao|Recusado", case=False, na=False) & df_f[col_item].isin(TODOS_ITENS)])
        meta_total_f = sum([int(metas_f.get(it, 0)) for it in TODOS_ITENS])

        ck1, ck2, ck3, ck4 = st.columns(4)
        with ck1:
            st.markdown(f"<div style='background: {t['card_bg']}; border: 1px solid {t['borda']}; border-left: 6px solid {t['primario']}; padding: 16px; border-radius: 10px;'><div style='font-size: 12px; color: {t['texto']}; font-weight: bold;'>META TOTAL</div><div style='font-size: 26px; font-weight: bold; color: {t['secundario']}; margin-top: 4px;'>{meta_total_f}</div></div>", unsafe_allow_html=True)
        with ck2:
            st.markdown(f"<div style='background: {t['card_bg']}; border: 1px solid {t['borda']}; border-left: 6px solid #16A34A; padding: 16px; border-radius: 10px;'><div style='font-size: 12px; color: {t['texto']}; font-weight: bold;'>APROVADOS</div><div style='font-size: 26px; font-weight: bold; color: #16A34A; margin-top: 4px;'>{tot_aprovados_f}</div></div>", unsafe_allow_html=True)
        with ck3:
            st.markdown(f"<div style='background: {t['card_bg']}; border: 1px solid {t['borda']}; border-left: 6px solid #D97706; padding: 16px; border-radius: 10px;'><div style='font-size: 12px; color: {t['texto']}; font-weight: bold;'>PENDENTES</div><div style='font-size: 26px; font-weight: bold; color: #D97706; margin-top: 4px;'>{tot_pendentes_f}</div></div>", unsafe_allow_html=True)
        with ck4:
            st.markdown(f"<div style='background: {t['card_bg']}; border: 1px solid {t['borda']}; border-left: 6px solid #DC2626; padding: 16px; border-radius: 10px;'><div style='font-size: 12px; color: {t['texto']}; font-weight: bold;'>REPROVADOS</div><div style='font-size: 26px; font-weight: bold; color: #DC2626; margin-top: 4px;'>{tot_reprovados_f}</div></div>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        col_esq, col_dir = st.columns(2)

        with col_esq:
            st.markdown(f"<h3 style='color:{t['secundario']}; margin-bottom:12px;'>Desempenho de Metas — {nome_curto.upper()}</h3>", unsafe_allow_html=True)
            dados_tabela_filial = []
            for it in TODOS_ITENS:
                meta_it = int(metas_f.get(it, 0))
                df_it = df_f[df_f[col_item] == it]
                ap = len(df_it[df_it[col_status].str.contains("Aprovado|Conforme|Sim", case=False, na=False)])
                pe = len(df_it[df_it[col_status].str.contains("Pendente", case=False, na=False)])
                soma_ap_pe = int(ap + pe)
                falta_it = max(0, meta_it - ap)
                perc_ap_pe = (float(soma_ap_pe) / float(meta_it) * 100.0) if meta_it > 0 else 0.0
                perc_ap = (float(ap) / float(meta_it) * 100.0) if meta_it > 0 else 0.0

                barras_html = f"<div style='text-align: left; font-size: 11px; color:{t['texto']}; font-weight:bold; margin-bottom: 2px;'>Apr+Pend ({perc_ap_pe:.1f}%)</div>{renderizar_barra(perc_ap_pe, '#D97706')}<div style='text-align: left; font-size: 11px; color:{t['texto']}; font-weight:bold; margin-top: 4px; margin-bottom: 2px;'>Aprovados ({perc_ap:.1f}%)</div>{renderizar_barra(perc_ap, '#16A34A')}"
                dados_tabela_filial.append({
                    "Item / Ações": f"<b>{it}</b>",
                    "Meta": f"<b>{meta_it}</b>",
                    "Aprovado": f"<span style='color:#16A34A; font-weight:800; font-size:14px;'>{ap}</span>",
                    "Pendente": f"<span style='color:#D97706; font-weight:800; font-size:14px;'>{pe}</span>",
                    "SOMANDO APROVADO E PENDENTES": f"<b>{soma_ap_pe}</b>",
                    "Falta": f"<b>{falta_it}</b>",
                    "Progresso": barras_html,
                })
            df_tab = pd.DataFrame(dados_tabela_filial)
            html_table = f'<table style="width:100%; text-align:center; border-collapse: collapse; border: 1px solid {t["borda"]}; color: {t["texto"]}; background-color: {t["card_bg"]}; font-size: 13px;">'
            html_table += f'<tr style="background-color: {t["secundario"]}; color: white;">'
            for col_n in ["Item / Ações", "Meta", "Aprovado", "Pendente", "SOMANDO APROVADO E PENDENTES", "Falta", "Progresso"]:
                html_table += f'<th style="text-align:center; padding: 10px; border: 1px solid {t["borda"]}; color: white !important;">{col_n}</th>'
            html_table += "</tr>"
            for _, r in df_tab.iterrows():
                html_table += f'<tr style="background-color: {t["card_bg"]};">'
                html_table += f'<td style="text-align:center; padding: 10px; border: 1px solid {t["borda"]};">{r["Item / Ações"]}</td>'
                html_table += f'<td style="text-align:center; padding: 10px; border: 1px solid {t["borda"]};">{r["Meta"]}</td>'
                html_table += f'<td style="text-align:center; padding: 10px; border: 1px solid {t["borda"]};">{r["Aprovado"]}</td>'
                html_table += f'<td style="text-align:center; padding: 10px; border: 1px solid {t["borda"]};">{r["Pendente"]}</td>'
                html_table += f'<td style="text-align:center; padding: 10px; border: 1px solid {t["borda"]};">{r["SOMANDO APROVADO E PENDENTES"]}</td>'
                html_table += f'<td style="text-align:center; padding: 10px; border: 1px solid {t["borda"]};">{r["Falta"]}</td>'
                html_table += f'<td style="text-align:center; padding: 10px; border: 1px solid {t["borda"]};">{r["Progresso"]}</td>'
                html_table += "</tr>"
            html_table += "</table>"
            st.markdown(html_table, unsafe_allow_html=True)

        with col_dir:
            st.markdown(f"<h3 style='color:{t['secundario']}; margin-bottom:12px;'>Acompanhamento Diário (Hoje)</h3>", unsafe_allow_html=True)
            hoje = date.today()
            df_hoje = df_f[df_f["Data_Calc"] == hoje]
            todos_promotores_filial = sorted(df_f[col_promotor].dropna().unique().tolist())
            resumo_diario = []
            for p in todos_promotores_filial:
                df_p_hoje = df_hoje[df_hoje[col_promotor] == p]
                qtd_hoje = len(df_p_hoje[df_p_hoje[col_item].isin(TODOS_ITENS)])
                resumo_diario.append({"Promotor": p, "Leituras Hoje": qtd_hoje})

            df_diario = pd.DataFrame(resumo_diario)
            if not df_diario.empty:
                df_diario = df_diario.sort_values(by="Leituras Hoje", ascending=False)
                num_proms = len(df_diario)
                cols_promotores = st.columns(max(1, num_proms))
                for i, (_, row) in enumerate(df_diario.iterrows()):
                    tem_leituras = row["Leituras Hoje"] > 0
                    cor_borda = "#16A34A" if tem_leituras else "#DC2626"
                    foto_prom = obter_foto_promotor(row["Promotor"])
                    nome_curto_prom = row["Promotor"].split()[0]
                    if isinstance(foto_prom, Image.Image):
                        buffered = BytesIO()
                        foto_prom.save(buffered, format="PNG")
                        img_b64 = base64.b64encode(buffered.getvalue()).decode("utf-8")
                        src_img = f"data:image/png;base64,{img_b64}"
                    else:
                        src_img = foto_prom

                    with cols_promotores[i % len(cols_promotores)]:
                        st.markdown(
                            f"<div style='background-color: {t['card_bg']}; border: 1px solid {t['borda']}; border-top: 5px solid {cor_borda}; padding: 12px 8px; border-radius: 8px; text-align: center;'><img src='{src_img}' style='width: 75px; height: 75px; border-radius: 50%; object-fit: cover; border: 2px solid {cor_borda}; margin-bottom: 6px;'><div style='font-size: 13px; font-weight: bold;'>{nome_curto_prom}</div><div style='font-size: 11px; font-weight: bold; color: {cor_borda};'>{row['Leituras Hoje']} leituras</div></div>",
                            unsafe_allow_html=True,
                        )

    if pagina_selecionada == "📺 Poços de Caldas":
        renderizar_tela_tv("Minassal Ltda - Pocos de Caldas")
    elif pagina_selecionada == "📺 São João da Boa Vista":
        renderizar_tela_tv("Minassal Ltda - Sao Joao da Boa Vista")
    elif pagina_selecionada == "📺 São José do Rio Preto":
        renderizar_tela_tv("Minassal Ltda - Sao Jose do Rio Preto")
    elif pagina_selecionada == "📺 Juiz de Fora":
        renderizar_tela_tv("Minassal Ltda - Juiz de Fora")

    elif pagina_selecionada == "🏆 Ranking de Leituras":
        st.markdown(f"## Ranking Geral de Leituras da Equipe — {periodo_atual}")
        st.write("Acompanhamento competitivo de todos os promotores para o período selecionado.")

        filial_rank_sel = st.selectbox("Filial para Filtrar Ranking:", ["TODAS"] + FILIAIS)
        df_rank = df[df[col_item].isin(TODOS_ITENS)].copy()

        if filial_rank_sel != "TODAS":
            df_rank = filtrar_por_filial(df_rank, filial_rank_sel, col_filial)

        if not df_rank.empty:
            promotores_todos = sorted(df_rank[col_promotor].dropna().unique().tolist())
            lista_ranking = []
            for prom in promotores_todos:
                df_p = df_rank[df_rank[col_promotor] == prom]
                total_geral = len(df_p)
                total_aprovados = len(df_p[df_p[col_status].str.contains("Aprovado|Conforme|Sim", case=False, na=False)])
                total_apv_pend = len(df_p[df_p[col_status].str.contains("Aprovado|Conforme|Sim|Pendente", case=False, na=False)])
                total_cancelados = len(df_p[df_p[col_status].str.contains("Cancelado|Cancelada", case=False, na=False)])
                total_reprovados = len(df_p[df_p[col_status].str.contains("Reprovado|Não|Nao|Recusado", case=False, na=False)])
                lista_ranking.append({
                    "promotor": prom,
                    "total": total_geral,
                    "aprovados": total_aprovados,
                    "apv_pend": total_apv_pend,
                    "cancelados": total_cancelados,
                    "reprovados": total_reprovados,
                    "df_promotor": df_p,
                })

            lista_ranking = sorted(lista_ranking, key=lambda x: x["total"], reverse=True)

            def gerar_pdf_ranking(dados_ranking, filial_nome, df_geral):
                class PDFRelatorio(FPDF):
                    def footer(self):
                        self.set_y(-10)
                        self.set_font("Helvetica", "I", 8)
                        self.set_text_color(100, 116, 139)
                        self.cell(0, 5, "Desenvolvido por Benedito Bandola", 0, 0, "C")

                pdf = PDFRelatorio(orientation="L", unit="mm", format="A4")
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.add_page()
                pdf.set_font("Helvetica", "B", 16)
                pdf.set_text_color(30, 58, 138)
                pdf.cell(0, 8, f"Leituras Minassal - {periodo_atual}", ln=True, align="C")

                pdf.set_font("Helvetica", "I", 9)
                pdf.set_text_color(100, 116, 139)
                pdf.cell(0, 5, f"Filtro: {filial_nome} | Emissão: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align="C")
                pdf.ln(3)

                pdf.set_draw_color(203, 213, 225)
                pdf.set_line_width(0.5)
                pdf.line(10, pdf.get_y(), 287, pdf.get_y())
                pdf.ln(3)

                pdf.set_fill_color(30, 58, 138)
                pdf.set_text_color(255, 255, 255)
                pdf.set_font("Helvetica", "B", 9)
                pdf.rect(10, pdf.get_y(), 277, 7, style="F")

                pdf.set_xy(12, pdf.get_y() + 1)
                pdf.cell(113, 5, "PROMOTOR / FILIAL", 0, 0, "L")
                pdf.cell(28, 5, "TOTAL", 0, 0, "C")
                pdf.cell(32, 5, "APROVADOS", 0, 0, "C")
                pdf.cell(32, 5, "APV + PENDENTES", 0, 0, "C")
                pdf.cell(32, 5, "CANCELADOS", 0, 0, "C")
                pdf.cell(30, 5, "REPROVADOS", 0, 1, "C")
                pdf.ln(2)

                temp_files_to_delete = []
                rgb_filiais = {
                    "Pocos de Caldas": (240, 253, 244),
                    "Sao Joao da Boa Vista": (253, 242, 248),
                    "Sao Jose do Rio Preto": (239, 246, 255),
                    "Juiz de Fora": (254, 252, 232),
                }

                for idx, item in enumerate(dados_ranking, start=1):
                    y_inicial = pdf.get_y()
                    df_p_item = item["df_promotor"]
                    filial_promotor_str = "N/A"
                    if not df_p_item.empty and col_filial in df_p_item.columns:
                        val_filial_bruto = str(df_p_item[col_filial].iloc[0])
                        filial_promotor_str = val_filial_bruto.split("-")[-1].strip()

                    if filial_promotor_str.upper() == "MINASSAL LTDA":
                        filial_promotor_str = "MINASSAL SAO JOAO DA BOA VISTA"

                    cor_rgb_fundo = (255, 255, 255)
                    for chave_f, rgb_val in rgb_filiais.items():
                        if chave_f.lower() in filial_promotor_str.lower():
                            cor_rgb_fundo = rgb_val
                            break

                    pdf.set_fill_color(cor_rgb_fundo[0], cor_rgb_fundo[1], cor_rgb_fundo[2])
                    pdf.set_draw_color(203, 213, 225)
                    pdf.set_line_width(0.3)
                    pdf.rect(10, y_inicial, 277, 21, style="DF")

                    foto_obj = obter_foto_promotor(item["promotor"])
                    if isinstance(foto_obj, Image.Image):
                        try:
                            if foto_obj.mode in ("RGBA", "P"):
                                foto_obj = foto_obj.convert("RGB")
                            temp_img = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                            foto_obj.save(temp_img, format="PNG")
                            temp_img.close()
                            temp_files_to_delete.append(temp_img.name)
                            pdf.image(temp_img.name, x=14, y=y_inicial + 3.5, w=14, h=14)
                        except:
                            pass

                    pdf.set_xy(34, y_inicial + 3)
                    pdf.set_font("Helvetica", "B", 10)
                    pdf.set_text_color(15, 23, 42)
                    pdf.cell(88, 5, f"#{idx} - {item['promotor']}", ln=0)

                    pdf.set_xy(34, y_inicial + 11.5)
                    pdf.set_font("Helvetica", "I", 8.5)
                    pdf.set_text_color(71, 85, 105)
                    pdf.cell(88, 4, f"Filial: {filial_promotor_str}", ln=0)

                    pdf.set_xy(125, y_inicial + 6)
                    pdf.set_font("Helvetica", "B", 10)
                    pdf.set_text_color(30, 58, 138)
                    pdf.cell(28, 6, str(item["total"]), 0, 0, "C")
                    pdf.set_text_color(22, 163, 74)
                    pdf.cell(32, 6, str(item["aprovados"]), 0, 0, "C")
                    pdf.set_text_color(217, 119, 6)
                    pdf.cell(32, 6, str(item["apv_pend"]), 0, 0, "C")
                    pdf.set_text_color(202, 138, 4)
                    pdf.cell(32, 6, str(item["cancelados"]), 0, 0, "C")
                    pdf.set_text_color(220, 38, 38)
                    pdf.cell(30, 6, str(item["reprovados"]), 0, 1, "C")
                    pdf.ln(4)

                return pdf.output(dest="S").encode("latin1", errors="replace")

            pdf_bytes = gerar_pdf_ranking(lista_ranking, filial_rank_sel, df)
            st.download_button(
                label="Baixar Relatório em PDF Executivo (Ranking + Metas)",
                data=pdf_bytes,
                file_name=f"Ranking_Metas_{periodo_atual}_{filial_rank_sel.replace(' ', '_')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
            st.markdown("<br>", unsafe_allow_html=True)

            for posicao, item_rank in enumerate(lista_ranking, start=1):
                prom = item_rank["promotor"]
                total = item_rank["total"]
                aprovados = item_rank["aprovados"]
                apv_pend = item_rank["apv_pend"]
                cancelados = item_rank["cancelados"]
                reprovados = item_rank["reprovados"]
                df_p = item_rank["df_promotor"]
                foto_prom = obter_foto_promotor(prom)

                cor_medalha = "#94A3B8"
                if posicao == 1:
                    cor_medalha = "#F59E0B"
                elif posicao == 2:
                    cor_medalha = "#64748B"
                elif posicao == 3:
                    cor_medalha = "#B45309"

                with st.container():
                    st.markdown(
                        f"<div style='background-color: {t['card_bg']}; border: 1px solid {t['borda']}; border-left: 8px solid {cor_medalha}; padding: 18px; border-radius: 12px; margin-bottom: 16px;'>",
                        unsafe_allow_html=True,
                    )
                    c_pos, c_img, c_detalhes = st.columns([1, 2.5, 6])
                    with c_pos:
                        st.markdown(f"<div style='font-size: 28px; font-weight: 800; color: {cor_medalha}; text-align: center; padding-top: 35px;'>#{posicao}</div>", unsafe_allow_html=True)
                    with c_img:
                        if isinstance(foto_prom, Image.Image):
                            buffered = BytesIO()
                            foto_prom.save(buffered, format="PNG")
                            img_b64 = base64.b64encode(buffered.getvalue()).decode("utf-8")
                            src_img = f"data:image/png;base64,{img_b64}"
                        else:
                            src_img = foto_prom
                        st.markdown(f"<div style='display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center;'><img src='{src_img}' style='width: 90px; height: 90px; border-radius: 8px; object-fit: cover; border: 1px solid {t['borda']}; margin-bottom: 6px;'><div style='font-size: 14px; font-weight: bold;'>{prom}</div></div>", unsafe_allow_html=True)
                    with c_detalhes:
                        st.markdown(
                            f"<div style='display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 10px;'><div style='background-color: {t['bg_app']}; border: 1px solid {t['borda']}; padding: 8px 12px; border-radius: 8px; text-align: center;'><div style='font-size: 11px; font-weight: bold;'>TOTAL</div><div style='font-size: 16px; color: {t['secundario']}; font-weight: 800;'>{total}</div></div><div style='background-color: #F0FDF4; border: 1px solid #DCFCE7; padding: 8px 12px; border-radius: 8px; text-align: center;'><div style='font-size: 11px; color: #166534; font-weight: bold;'>APROVADOS</div><div style='font-size: 16px; color: #16A34A; font-weight: 800;'>{aprovados}</div></div><div style='background-color: #FEF3C7; border: 1px solid #FDE68A; padding: 8px 12px; border-radius: 8px; text-align: center;'><div style='font-size: 11px; color: #92400E; font-weight: bold;'>APV + PEND</div><div style='font-size: 16px; color: #D97706; font-weight: 800;'>{apv_pend}</div></div><div style='background-color: #FEFCE8; border: 1px solid #FEF08A; padding: 8px 12px; border-radius: 8px; text-align: center;'><div style='font-size: 11px; color: #854D0E; font-weight: bold;'>CANCELADOS</div><div style='font-size: 16px; color: #CA8A04; font-weight: 800;'>{cancelados}</div></div><div style='background-color: #FEF2F2; border: 1px solid #FEE2E2; padding: 8px 12px; border-radius: 8px; text-align: center;'><div style='font-size: 11px; color: #991B1B; font-weight: bold;'>REPROVADOS</div><div style='font-size: 16px; color: #DC2626; font-weight: 800;'>{reprovados}</div></div></div>",
                            unsafe_allow_html=True,
                        )
                        contagem_kpi = df_p[col_item].value_counts()
                        badges_rank = "<div style='display: flex; flex-wrap: wrap; gap: 6px;'>"
                        for item_kpi, qtd_kpi in contagem_kpi.items():
                            badges_rank += f"<span style='background-color: {t['bg_app']}; border: 1px solid {t['borda']}; color: {t['primario']}; padding: 4px 10px; border-radius: 16px; font-size: 12px; font-weight: 600;'><b>{qtd_kpi}x</b> {item_kpi}</span>"
                        badges_rank += "</div>"
                        st.markdown(badges_rank, unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)
        else:
            st.info("Nenhum dado encontrado para o ranking.")

    elif pagina_selecionada == "🚨 Registros Reprovados":
        st.markdown(f"## Central Exclusiva de Registros Reprovados — {periodo_atual}")
        filial_rep_sel = st.selectbox("Filial:", ["TODAS"] + FILIAIS)
        df_reprovados_geral = df[df[col_status].str.contains("Reprovado|Não|Nao|Recusado", case=False, na=False)].copy()

        if filial_rep_sel != "TODAS":
            df_reprovados_geral = filtrar_por_filial(df_reprovados_geral, filial_rep_sel, col_filial)

        if not df_reprovados_geral.empty:
            col_justif = "JustificativaRecusa" if "JustificativaRecusa" in df_reprovados_geral.columns else col_status
            tabela_rep_geral = df_reprovados_geral[[col_filial, col_promotor, "Pdv", col_item, col_status, col_justif]]
            tabela_rep_geral.columns = ["Filial", "Promotor", "Loja (PDV)", "Item", "Status", "Motivo da Recusa"]

            html_rep_geral = tabela_rep_geral.to_html(escape=False, index=False)
            html_rep_geral = html_rep_geral.replace("<table>", f'<table style="width:100%; text-align:center; border-collapse: collapse; border: 1px solid {t["borda"]}; color: {t["texto"]}; background-color: {t["card_bg"]}; font-size: 13px;">')
            html_rep_geral = html_rep_geral.replace("<th>", f'<th style="text-align:center; padding: 10px; border: 1px solid {t["borda"]}; background-color: #DC2626; color: white !important; font-size: 13px;">')
            html_rep_geral = html_rep_geral.replace("<td>", f'<td style="text-align:center; padding: 10px; border: 1px solid {t["borda"]}; vertical-align: middle;">')
            st.markdown(html_rep_geral, unsafe_allow_html=True)
        else:
            st.success("Parabéns! Não há nenhum registro reprovado para os filtros selecionados.")

    elif pagina_selecionada == "💰 Auditoria de Preços":
        st.markdown(f"## Auditoria e Análise de Preços no PDV — {periodo_atual}")
        precos_dict = carregar_precos()

        tipo_visao_preco = st.radio(
            "Selecione a Visão da Auditoria:",
            [
                "🔍 Itens com Preço Controlado (Auditoria de Small Bags)",
                "📋 Demais Itens (Sem Controle de Preço na Planilha)",
            ],
            horizontal=True,
        )

        if tipo_visao_preco == "🔍 Itens com Preço Controlado (Auditoria de Small Bags)":
            with st.expander("Gerenciar Regras de Preços (Small Bags)", expanded=True):
                df_regras = pd.DataFrame([
                    {
                        "Produto": k,
                        "Mínimo (R$)": float(v.get("min", 0.0)),
                        "Máximo (R$)": float(v.get("max", 0.0)),
                        "Não Participante": v.get("nao_participante", False),
                        "Excluir": False,
                    }
                    for k, v in precos_dict.items()
                ])

                df_editado = st.data_editor(
                    df_regras,
                    column_config={
                        "Produto": st.column_config.TextColumn("Produto", disabled=True),
                        "Mínimo (R$)": st.column_config.NumberColumn("Mínimo (R$)", format="%.2f", step=0.1),
                        "Máximo (R$)": st.column_config.NumberColumn("Máximo (R$)", format="%.2f", step=0.1),
                        "Não Participante": st.column_config.CheckboxColumn("Não Participante"),
                        "Excluir": st.column_config.CheckboxColumn("Excluir"),
                    },
                    hide_index=True,
                    use_container_width=True,
                )

                if st.button("Salvar Alterações na Tabela", type="primary", key="btn_salvar_regras"):
                    novas_regras = {}
                    for _, row in df_editado.iterrows():
                        if not row["Excluir"]:
                            novas_regras[row["Produto"]] = {
                                "min": row["Mínimo (R$)"],
                                "max": row["Máximo (R$)"],
                                "nao_participante": row["Não Participante"],
                            }
                    salvar_precos(novas_regras)
                    st.success("Alterações de preços salvas com sucesso!")
                    time.sleep(0.5)
                    st.rerun()

            st.markdown("---")
            f1, f2 = st.columns(2)
            with f1:
                filial_audit = st.selectbox("Filial:", ["TODAS"] + FILIAIS, key="fil_audit")
            with f2:
                promotores_audit = ["TODOS"] + sorted(df[col_promotor].dropna().unique().tolist())
                prom_audit = st.selectbox("Promotor:", promotores_audit, key="prom_audit")

            df_audit = df[df["Produto_Preco"].isin(precos_dict.keys())].copy()
            if filial_audit != "TODAS":
                df_audit = filtrar_por_filial(df_audit, filial_audit, col_filial)
            if prom_audit != "TODOS":
                df_audit = df_audit[df_audit[col_promotor] == prom_audit]

            if not df_audit.empty:
                df_audit["Auditoria"] = df_audit["Auditoria_Preco"]
                tot_dentro = len(df_audit[df_audit["Auditoria"].isin(["🟢 Dentro da Faixa", "🟢 Mais Barato (Aprovado)"])])
                tot_barato = len(df_audit[df_audit["Auditoria"] == "🟢 Mais Barato (Aprovado)"])
                tot_fora = len(df_audit[df_audit["Auditoria"] == "🔴 Fora (Acima do Teto)"])
                tot_corrigido = len(df_audit[df_audit["Auditoria"] == "🟡 Preço Corrigido (Cancelado)"])

                ck1, ck2, ck3, ck4 = st.columns(4)
                ck1.markdown(f"<div style='border-left: 5px solid #16A34A; background:{t['card_bg']}; padding:12px; border-radius:8px; border:1px solid {t['borda']};'><div>Dentro do Preço</div><div style='font-size:20px;font-weight:bold;color:#16A34A'>{tot_dentro}</div></div>", unsafe_allow_html=True)
                ck2.markdown(f"<div style='border-left: 5px solid #10B981; background:{t['card_bg']}; padding:12px; border-radius:8px; border:1px solid {t['borda']};'><div>Mais Barato (Apr.)</div><div style='font-size:20px;font-weight:bold;color:#10B981'>{tot_barato}</div></div>", unsafe_allow_html=True)
                ck3.markdown(f"<div style='border-left: 5px solid #DC2626; background:{t['card_bg']}; padding:12px; border-radius:8px; border:1px solid {t['borda']};'><div>Acima do Range</div><div style='font-size:20px;font-weight:bold;color:#DC2626'>{tot_fora}</div></div>", unsafe_allow_html=True)
                ck4.markdown(f"<div style='border-left: 5px solid #EAB308; background:{t['card_bg']}; padding:12px; border-radius:8px; border:1px solid {t['borda']};'><div>Cancelados</div><div style='font-size:20px;font-weight:bold;color:#EAB308'>{tot_corrigido}</div></div>", unsafe_allow_html=True)

                st.markdown("<br>### Detalhamento da Análise de Preços", unsafe_allow_html=True)
                filtro_auditoria = st.selectbox(
                    "Filtrar Tabela por Condição:",
                    [
                        "⚠️ Divergências e Ressalvas (Padrão)",
                        "🔴 Apenas Acima do Range",
                        "🟡 Apenas Preços Cancelados",
                        "🟢 Apenas Mais Barato",
                        "📋 Mostrar Tudo",
                    ],
                )

                if filtro_auditoria == "⚠️ Divergências e Ressalvas (Padrão)":
                    df_tabela = df_audit[df_audit["Auditoria"].isin(["🔴 Fora (Acima do Teto)", "🟢 Mais Barato (Aprovado)", "🟡 Preço Corrigido (Cancelado)"])].copy()
                elif "Acima do Range" in filtro_auditoria:
                    df_tabela = df_audit[df_audit["Auditoria"] == "🔴 Fora (Acima do Teto)"].copy()
                elif "Cancelados" in filtro_auditoria:
                    df_tabela = df_audit[df_audit["Auditoria"] == "🟡 Preço Corrigido (Cancelado)"].copy()
                elif "Mais Barato" in filtro_auditoria:
                    df_tabela = df_audit[df_audit["Auditoria"] == "🟢 Mais Barato (Aprovado)"].copy()
                else:
                    df_tabela = df_audit.copy()

                if not df_tabela.empty:
                    tabela_div = df_tabela[[col_promotor, "Cidade", "Pdv", "Produto_Preco", col_preco, "Auditoria"]]
                    tabela_div.columns = ["Promotor", "Cidade", "Loja", "Produto", "Preço Digitado (R$)", "Status do Preço"]
                    html_div = tabela_div.to_html(escape=False, index=False)
                    html_div = html_div.replace("<table>", f'<table style="width:100%; text-align:center; border-collapse: collapse; border: 1px solid {t["borda"]}; color: {t["texto"]}; background-color: {t["card_bg"]}; font-size: 13px;">')
                    html_div = html_div.replace("<th>", f'<th style="text-align:center; padding: 10px; border: 1px solid {t["borda"]}; background-color: {t["secundario"]}; color: white !important; font-size: 13px;">')
                    html_div = html_div.replace("<td>", f'<td style="text-align:center; padding: 10px; border: 1px solid {t["borda"]}; vertical-align: middle;">')
                    st.markdown(html_div, unsafe_allow_html=True)
                else:
                    st.success("Nenhum dado encontrado para este filtro.")
            else:
                st.warning("Nenhuma leitura encontrada para os produtos configurados.")
        else:
            st.markdown("### Demais Itens (Sem Controle de Preço na Planilha)")
            df_sem_controle = df[~df["Produto_Preco"].isin(precos_dict.keys())].copy()
            f1, f2 = st.columns(2)
            with f1:
                filial_audit_outros = st.selectbox("Filial:", ["TODAS"] + FILIAIS, key="fil_outros")
            with f2:
                prom_outros = st.selectbox("Promotor:", ["TODOS"] + sorted(df[col_promotor].dropna().unique().tolist()), key="prom_outros")

            if filial_audit_outros != "TODAS":
                df_sem_controle = filtrar_por_filial(df_sem_controle, filial_audit_outros, col_filial)
            if prom_outros != "TODOS":
                df_sem_controle = df_sem_controle[df_sem_controle[col_promotor] == prom_outros]

            if not df_sem_controle.empty:
                tabela_outros = df_sem_controle[[col_promotor, "Cidade", "Pdv", "Produto_Preco", col_preco]]
                tabela_outros.columns = ["Promotor", "Cidade", "Loja", "Produto", "Leitura Digitada"]
                html_outros = tabela_outros.to_html(escape=False, index=False)
                html_outros = html_outros.replace("<table>", f'<table style="width:100%; text-align:center; border-collapse: collapse; border: 1px solid {t["borda"]}; color: {t["texto"]}; background-color: {t["card_bg"]}; font-size: 13px;">')
                html_outros = html_outros.replace("<th>", f'<th style="text-align:center; padding: 10px; border: 1px solid {t["borda"]}; background-color: {t["secundario"]}; color: white !important; font-size: 13px;">')
                html_outros = html_outros.replace("<td>", f'<td style="text-align:center; padding: 10px; border: 1px solid {t["borda"]}; vertical-align: middle;">')
                st.markdown(html_outros, unsafe_allow_html=True)
            else:
                st.success("Não há itens sem controle para exibir.")

    elif pagina_selecionada == "🗺️ Mapa de Clientes":
        st.markdown(f"## 🗺️ Mapeamento Geográfico de Clientes e Leituras — {periodo_atual}")
        st.write("Visualização interativa estilo Leaflet/Folium com ícones personalizados para cada item em campo.")

        if os.path.exists(ARQUIVO_CLIENTES_COORD):
            try:
                df_clientes_coord = pd.read_excel(ARQUIVO_CLIENTES_COORD)
                df_clientes_coord["doc_limpo"] = df_clientes_coord["DOCUMENTO"].astype(str).str.replace(r"\D", "", regex=True)
                col_cnpj_df = "Cnpj_Documento" if "Cnpj_Documento" in df.columns else "CnpjCpfPdv"
                df["doc_limpo"] = df[col_cnpj_df].astype(str).str.replace(r"\D", "", regex=True)

                df_mapa = pd.merge(df, df_clientes_coord[["doc_limpo", "LATITUDE", "LONGITUDE", "NOME", "ENDEREÇO", "CIDADE", "BAIRRO"]], on="doc_limpo", how="inner")

                if not df_mapa.empty:
                    st.markdown("---")
                    c_f1, c_f2, c_f3, c_f4, c_f5 = st.columns(5)
                    with c_f1:
                        filial_mapa = st.selectbox("Filtrar Filial:", ["TODAS"] + FILIAIS, key="map_filial")
                    with c_f2:
                        promotores_disponiveis = sorted(df_mapa[col_promotor].dropna().unique().tolist())
                        prom_mapa = st.selectbox("Promotor:", ["TODOS"] + promotores_disponiveis, key="map_prom")
                    with c_f3:
                        itens_disponiveis = sorted(df_mapa[col_item].dropna().unique().tolist())
                        item_mapa = st.selectbox("Item / KPI:", ["TODOS OS ITENS"] + itens_disponiveis, key="map_item")
                    with c_f4:
                        status_mapa = st.selectbox("Status:", ["TODOS", "Aprovado", "Pendente", "Reprovado"], key="map_status")
                    with c_f5:
                        cidades_disponiveis = sorted(df_mapa["CIDADE"].dropna().astype(str).unique().tolist())
                        cidade_mapa = st.selectbox("Cidade:", ["TODAS AS CIDADES"] + cidades_disponiveis, key="map_cidade")

                    df_mapa_filtrado = df_mapa.copy()
                    if filial_mapa != "TODAS":
                        df_mapa_filtrado = filtrar_por_filial(df_mapa_filtrado, filial_mapa, col_filial)
                    if prom_mapa != "TODOS":
                        df_mapa_filtrado = df_mapa_filtrado[df_mapa_filtrado[col_promotor] == prom_mapa]
                    if item_mapa != "TODOS OS ITENS":
                        df_mapa_filtrado = df_mapa_filtrado[df_mapa_filtrado[col_item] == item_mapa]
                    if status_mapa != "TODOS":
                        df_mapa_filtrado = df_mapa_filtrado[df_mapa_filtrado[col_status].str.lower().str.contains(status_mapa.lower())]
                    if cidade_mapa != "TODAS AS CIDADES":
                        df_mapa_filtrado = df_mapa_filtrado[df_mapa_filtrado["CIDADE"].astype(str) == cidade_mapa]

                    df_mapa_filtrado["LATITUDE"] = df_mapa_filtrado["LATITUDE"].astype(str).str.replace(",", ".", regex=False).astype(float)
                    df_mapa_filtrado["LONGITUDE"] = df_mapa_filtrado["LONGITUDE"].astype(str).str.replace(",", ".", regex=False).astype(float)

                    m1, m2, m3 = st.columns(3)
                    m1.metric("Pontos Exibidos no Mapa", len(df_mapa_filtrado))
                    m2.metric("Promotores Atuantes", df_mapa_filtrado[col_promotor].nunique())
                    m3.metric("Lojas Únicas", df_mapa_filtrado["doc_limpo"].nunique())

                    st.markdown("---")
                    if not df_mapa_filtrado.empty:
                        ponto_central_lat = df_mapa_filtrado["LATITUDE"].mean()
                        ponto_central_lon = df_mapa_filtrado["LONGITUDE"].mean()

                        m = folium.Map(location=[ponto_central_lat, ponto_central_lon], zoom_start=7, tiles="Esri.WorldStreetMap")

                        def obter_estilo_personalizado(item_nome):
                            i_str = str(item_nome).lower()
                            if "sheba" in i_str:
                                return "purple", "cat"
                            elif "ponto extra" in i_str:
                                return "red", "layer-group"
                            elif "sachês" in i_str:
                                return "green", "tags"
                            elif "petiscos" in i_str:
                                return "orange", "bone"
                            else:
                                return "blue", "info-sign"

                        for _, r in df_mapa_filtrado.iterrows():
                            cor_icone, icone_fa = obter_estilo_personalizado(r[col_item])
                            html_popup = f"<b>Cliente:</b> {r['NOME']}<br/><b>Endereço:</b> {r['ENDEREÇO']}<br/><b>Promotor:</b> {r[col_promotor]}<br/><b>Item:</b> {r[col_item]}<br/><b>Status:</b> {r[col_status]}"
                            folium.Marker(
                                location=[r["LATITUDE"], r["LONGITUDE"]],
                                popup=folium.Popup(html_popup, max_width=300),
                                tooltip=f"{r['NOME']} - {r[col_item]}",
                                icon=folium.Icon(color=cor_icone, icon=icone_fa, prefix="fa"),
                            ).add_to(m)

                        st_folium(m, use_container_width=True, height=550)
                    else:
                        st.warning("Nenhum registro encontrado para os filtros escolhidos.")
                else:
                    st.warning("Não foram encontrados cruzamentos de CNPJ entre as leituras e a base de clientes.")
            except Exception as e:
                st.error(f"Erro ao carregar o mapa: {e}")
        else:
            st.error(f"O arquivo `{ARQUIVO_CLIENTES_COORD}` não foi encontrado na pasta.")

    elif pagina_selecionada == "📊 Presença por Sabores":
        st.markdown(f"## 📊 Ranking de Presença por Sabores e Produtos — Small Bags ({periodo_atual})")
        st.write("Análise exclusiva de quais sabores e produtos de Small Bags mais e menos aparecem em campo.")

        c_ps1, c_ps2 = st.columns(2)
        with c_ps1:
            filial_ps = st.selectbox("Filtrar Filial:", ["TODAS"] + FILIAIS, key="ps_filial")
        with c_ps2:
            promotores_ps_list = ["TODOS"] + sorted(df[col_promotor].dropna().unique().tolist())
            prom_ps = st.selectbox("Filtrar Promotor:", promotores_ps_list, key="ps_prom")

        df_ps = df[df[col_item].astype(str).str.contains("Small Bags", case=False, na=False)].copy()
        if filial_ps != "TODAS":
            df_ps = filtrar_por_filial(df_ps, filial_ps, col_filial)
        if prom_ps != "TODOS":
            df_ps = df_ps[df_ps[col_promotor] == prom_ps]

        if not df_ps.empty:
            df_ps["Sabor_Produto"] = df_ps.apply(
                lambda r: f"{r.get('OpcaoEmbalagem', 'N/A')} - {r.get('Embalagem', 'N/A')}" if pd.notna(r.get("OpcaoEmbalagem")) else str(r.get("Item", "N/A")),
                axis=1,
            )
            ranking_sabores = df_ps.groupby("Sabor_Produto").size().reset_index(name="Quantidade")
            ranking_sabores = ranking_sabores.sort_values(by="Quantidade", ascending=False).reset_index(drop=True)

            st.markdown("---")
            for idx, row in ranking_sabores.iterrows():
                pos = idx + 1
                prod_nome = row["Sabor_Produto"]
                qtd = row["Quantidade"]
                img_prod = obter_foto_produto_por_nome(prod_nome)
                codigo_prod = obter_codigo_produto(prod_nome)
                prod_nome_exibicao = f"{prod_nome} ({codigo_prod})" if codigo_prod else prod_nome

                with st.container():
                    col_img, col_info, col_qtd = st.columns([1, 8, 2])
                    with col_img:
                        if img_prod:
                            st.image(img_prod, width=50)
                        else:
                            st.markdown("<div style='font-size:11px; color:#94A3B8; text-align:center; padding-top:10px;'>Sem foto</div>", unsafe_allow_html=True)
                    with col_info:
                        st.markdown(f"<div style='font-size: 15px; font-weight: bold; color: {t['secundario']}; padding-top: 12px;'>#{pos} - {prod_nome_exibicao}</div>", unsafe_allow_html=True)
                    with col_qtd:
                        st.markdown(f"<div style='font-size: 16px; font-weight: 800; color: {t['texto']}; text-align: right; padding-top: 12px;'>{qtd} leituras</div>", unsafe_allow_html=True)
                    st.markdown(f"<hr style='margin: 4px 0; border-color: {t['borda']};'>", unsafe_allow_html=True)
        else:
            st.warning("Nenhum dado encontrado para os filtros selecionados.")

# --- ASSINATURA DISCRETA NO RODAPÉ ---
st.markdown("---")
st.markdown(
    f"<div style='text-align: center; font-size: 12px; color: {t['texto']}; opacity: 0.7; margin-bottom: 20px;'>Desenvolvido por Benedito Bandola</div>",
    unsafe_allow_html=True,
)
