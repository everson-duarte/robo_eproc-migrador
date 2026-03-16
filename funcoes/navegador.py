# navegador.py
# Funções para controle do navegador e automação web

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from .logger import logger

URL_EPROC = "https://eproc1g.tjsp.jus.br/eproc/"

# Variável global para controle do navegador
driver_global = None

def acessar_eproc():
    """Abre o navegador e acessa o e-Proc"""
    global driver_global
    
    if driver_global is None:
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_experimental_option("detach", True)
        
        driver_global = webdriver.Chrome(options=chrome_options)
        
        # Configurar timeout HTTP para 5 minutos (300 segundos) para operações longas
        driver_global.set_page_load_timeout(300)
        driver_global.command_executor.set_timeout(300)

    driver_global.switch_to.window(driver_global.current_window_handle)

    if URL_EPROC not in driver_global.current_url:
        driver_global.get(URL_EPROC)

    try:
        # Aguarda pelo elemento que indica que o login foi feito e a página carregou
        selector = 'i[title="Meus Localizadores"]'
        logger.info("Aguardando o login e carregamento da página inicial do e-Proc (até 90s)...")
        WebDriverWait(driver_global, 90).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, selector))
        )
        logger.info("Página inicial carregada com sucesso.")
    except TimeoutException:
        logger.warning("Tempo de espera esgotado (90s). Não foi possível validar o login.")
        logger.warning("Por favor, realize o login manualmente na janela do navegador para continuar.")


def fechar_navegador():
    """Fecha o navegador e limpa recursos"""
    global driver_global
    if driver_global is not None:
        driver_global.quit()
        driver_global = None


def obter_driver():
    """Retorna o driver do navegador ou None se não houver"""
    global driver_global
    return driver_global


def minimizar_navegador():
    """Minimiza a janela do navegador, se estiver aberta."""
    global driver_global
    if driver_global is not None:
        try:
            driver_global.minimize_window()
        except Exception:
            # Ignora se o driver/versão não suportar minimizar
            pass