from datetime import datetime, date
from pathlib import Path
from playwright.sync_api import sync_playwright
import time
import sys

# --- DEFINIÇÃO DOS PERÍODOS (P10 a P13) ---
PERIODOS = {
    "P10": {"inicio": date(2026, 9, 6), "fim": date(2026, 10, 3)},
    "P11": {"inicio": date(2026, 10, 4), "fim": date(2026, 10, 31)},
    "P12": {"inicio": date(2026, 11, 1), "fim": date(2026, 11, 28)},
    "P13": {"inicio": date(2026, 11, 29), "fim": date(2027, 1, 2)},
}

def obter_datas_periodo(nome_periodo):
    """
    Retorna a data inicial e a data final ajustada para o período.
    Se estivermos dentro do período vigente, a data final é hoje.
    Se o período já passou, usa a data final fixada do período.
    """
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
    """
    Automatiza o acesso ao portal PDV Pet, define o filtro de datas 
    conforme o período selecionado e baixa a planilha de questionários.
    """
    data_de_str, data_ate_str = obter_datas_periodo(periodo_escolhido)
    
    print(f"Periodo selecionado: {periodo_escolhido} | De: {data_de_str} Ate: {data_ate_str}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, slow_mo=500)
        context = browser.new_context()
        page = context.new_page()
        
        try:
            # 1. Acessar o site do PDV Pet
            print("Acessando o site do PDV Pet...")
            page.goto("https://www.pdvpet.com.br/")
            
            # 2. Preencher as credenciais de acesso
            print("Preenchendo dados de login...")
            page.get_by_placeholder("Digite seu CPF/CNPJ").fill("137.318.006-40")
            page.get_by_placeholder("Digite sua senha").fill("Fizzvini1234@")
            
            page.get_by_role("button", name="Entrar").click()
            
            # 3. Navegar para a seção de questionários
            print("Navegando para a aba de Questionarios...")
            page.get_by_role("link", name="Questionários").click()
            
            # 4. Configurar o período dinâmico
            print(f"Preenchendo datas: {data_de_str} ate {data_ate_str}...")
            page.locator("#DataDe").fill(data_de_str)
            page.locator("#DataAte").fill(data_ate_str)
            
            # 5. Executar a busca
            print("Buscando dados com os filtros aplicados...")
            page.get_by_role("button", name="Buscar").click()
            time.sleep(2)
            
            # 6. Exportar e baixar o relatório
            print("Exportando e baixando o relatorio...")
            with page.expect_download() as download_info:
                page.get_by_role("button", name="Exportar").click()
                page.get_by_role("link", name="Abrir").click()
                
            download = download_info.value
            
            # 7. Salvar na pasta do script
            pasta_script = Path(__file__).parent
            caminho_salvar = pasta_script / "relatorio_questionarios.csv"
            
            download.save_as(caminho_salvar)
            print(f"Sucesso! Planilha salva em: {caminho_salvar.resolve()}")
            
        except Exception as e:
            print(f"Ocorreu um erro durante a automacao: {e}")
            
        finally:
            time.sleep(3)
            browser.close()

if __name__ == "__main__":
    periodo_arg = sys.argv[1] if len(sys.argv) > 1 else "P11"
    baixar_planilha_pdvpet(periodo_arg)
