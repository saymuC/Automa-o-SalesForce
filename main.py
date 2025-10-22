import re
import time
import tempfile
import shutil
import os
import signal
import sys

from datetime import datetime, timedelta
import time

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.edge.service import Service
    from selenium.webdriver.edge.options import Options
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException
    from selenium.webdriver import ActionChains

    HAS_SELENIUM = True
except Exception:
    HAS_SELENIUM = False

    class _MissingSelenium:
        def __init__(self, *args, **kwargs):
            raise RuntimeError(
                "Selenium não esta instalado ou não foi importado corretamente.  "
                "Instale usando pip install selenium e rode o codigo novamente."
            )

        def __getattr__(self, name):
            return self.__class__

    webdriver = _MissingSelenium
    By = _MissingSelenium
    Keys = _MissingSelenium
    Service = _MissingSelenium
    Options = _MissingSelenium
    WebDriverWait = _MissingSelenium
    EC = _MissingSelenium
    TimeoutException = Exception
    ActionChains = _MissingSelenium

try:
    import questionary
    HAS_QUESTIONARY = True
except ImportError:
    HAS_QUESTIONARY = False

# ANSI colors
USE_COLOR = sys.stdout.isatty()

def _c(text, code):
    return f"\033[{code}m{text}\033[0m" if USE_COLOR else text

def log_info(msg):
    print(_c(f"[INFO] {msg}", "37"))

def log_ok(msg):
    print(_c(f"[✓] {msg}", "32"))

def log_warn(msg):
    print(_c(f"[⚠] {msg}", "33"))

def log_error(msg):
    print(_c(f"[✗] {msg}", "31"))

def log_debug(msg):
    print(_c(f"[DEBUG] {msg}", "36"))

# Configurações
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EDGE_DRIVER_PATH = os.path.join(BASE_DIR, "msedgedriver.exe")
TIMEOUT_DEFAULT = 12
TIMEOUT_MFA = 300
TIMEOUT_SEARCH = 30

# Variável global para rastrear recursos
_GLOBAL_RESOURCES = {
    'driver': None,
    'temp_dir': None,
    'cliente_url': None  # Armazena URL do cliente atual
}

def input_com_timeout(prompt, timeout=60):
    """Input com timeout para evitar travamentos"""
    import select
    
    print(prompt, end='', flush=True)
    
    if sys.platform == 'win32':
        # Windows não suporta select em stdin, usar input normal
        try:
            return input()
        except EOFError:
            return ''
    else:
        # Unix/Linux
        ready, _, _ = select.select([sys.stdin], [], [], timeout)
        if ready:
            return sys.stdin.readline().rstrip('\n')
        else:
            print(f"\n[TIMEOUT {timeout}s] - Usando valor padrão")
            return ''

def limpar_cpf(texto):
    return re.sub(r"\D", "", texto or "")

def validar_cpf(cpf):
    if len(cpf) != 11 or cpf == cpf[0] * 11:
        return False
    
    soma = sum(int(cpf[i]) * (10 - i) for i in range(9))
    digito1 = (soma * 10 % 11) % 10
    
    soma = sum(int(cpf[i]) * (11 - i) for i in range(10))
    digito2 = (soma * 10 % 11) % 10
    
    return cpf[-2:] == f"{digito1}{digito2}"

def cleanup_all_resources():
    """Limpa TODOS os recursos - chamado sempre ao finalizar"""
    global _GLOBAL_RESOURCES
    
    if _GLOBAL_RESOURCES['driver']:
        try:
            _GLOBAL_RESOURCES['driver'].quit()
            log_ok("Driver encerrado")
        except Exception:
            pass
        _GLOBAL_RESOURCES['driver'] = None
    
    if _GLOBAL_RESOURCES['temp_dir'] and os.path.isdir(_GLOBAL_RESOURCES['temp_dir']):
        try:
            shutil.rmtree(_GLOBAL_RESOURCES['temp_dir'], ignore_errors=True)
            log_ok(f"Perfil temporário removido")
        except Exception as e:
            log_warn(f"Não foi possível remover perfil: {e}")
        _GLOBAL_RESOURCES['temp_dir'] = None

def signal_handler(signum, frame):
    """Handler para Ctrl+C"""
    log_warn("\n\nInterrupção detectada (Ctrl+C)")
    cleanup_all_resources()
    log_info("Recursos limpos. Encerrando...")
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)

def criar_driver(initial_url="https://login.salesforce.com/"):
    global _GLOBAL_RESOURCES
    
    service = Service(EDGE_DRIVER_PATH)
    tmp_profile = None
    
    try:
        tmp_profile = tempfile.mkdtemp(prefix="edge_sf_")
        _GLOBAL_RESOURCES['temp_dir'] = tmp_profile
        
        opts = Options()
        opts.add_argument("--start-maximized")
        opts.add_argument(f"--user-data-dir={tmp_profile}")
        opts.add_argument("--disable-extensions")
        opts.add_argument("--disable-background-networking")
        opts.add_argument("--disable-sync")
        opts.add_argument("--no-first-run")
        opts.add_argument("--disable-dev-shm-usage")
        opts.page_load_strategy = 'eager'
        
        driver = webdriver.Edge(service=service, options=opts)
        _GLOBAL_RESOURCES['driver'] = driver
        driver.implicitly_wait(1)
        
        log_ok(f"Edge iniciado")
        
        try:
            driver.get(initial_url)
        except Exception:
            pass
        
        # Forçar foco no terminal após abrir o navegador
        if sys.platform == 'win32':
            try:
                import ctypes
                hwnd = ctypes.windll.kernel32.GetConsoleWindow()
                if hwnd:
                    ctypes.windll.user32.SetForegroundWindow(hwnd)
            except:
                pass
        
        return driver
        
    except Exception as e:
        log_error(f"Falha ao iniciar Edge: {e}")
        cleanup_all_resources()
        raise

def esperar_mfa(driver, timeout=TIMEOUT_MFA):
    """
    Aguarda a aprovação do MFA automaticamente, verificando continuamente
    se o usuário já passou pela autenticação. NÃO precisa apertar Enter!
    """
    
    wait_start = time.time()
    last_check = 0
    check_interval = 2  # Verifica a cada 2 segundos
    
    # JavaScript para verificar se ainda está na página de MFA
    js_verificar_mfa = """
    const url = window.location.href.toLowerCase();
    const pageText = document.body.innerText.toLowerCase();
    const pageHTML = document.body.innerHTML.toLowerCase();
    
    // Indicadores de que AINDA está no MFA
    const mfaIndicators = [
        url.includes('verification'),
        url.includes('mfa'),
        url.includes('authenticate'),
        url.includes('verify'),
        pageText.includes('verificação'),
        pageText.includes('verification'),
        pageText.includes('autenticação'),
        pageText.includes('authentication'),
        pageText.includes('approve'),
        pageText.includes('aprovar'),
        pageText.includes('código'),
        pageText.includes('code'),
        pageHTML.includes('mfa'),
        pageHTML.includes('verification')
    ];
    
    const stillInMFA = mfaIndicators.some(indicator => indicator === true);
    
    // Indicadores de que JÁ passou do MFA (está logado)
    const loggedInIndicators = [
        url.includes('/lightning/'),
        url.includes('/home'),
        url.includes('/setup'),
        document.querySelector('header.slds-global-header') !== null,
        document.querySelector('[class*="oneHeader"]') !== null,
        document.querySelector('[data-aura-class*="oneHeader"]') !== null
    ];
    
    const isLoggedIn = loggedInIndicators.some(indicator => indicator === true);
    
    return {
        stillInMFA: stillInMFA,
        isLoggedIn: isLoggedIn,
        url: url,
        hasLightning: url.includes('lightning')
    };
    """
    
    while time.time() - wait_start < timeout:
        try:
            current_time = time.time() - wait_start
            
            # Log de progresso a cada intervalo
            if int(current_time) - last_check >= check_interval:
                last_check = int(current_time)
                elapsed_minutes = int(current_time // 60)
                elapsed_seconds = int(current_time % 60)
                log_debug(f"Aguardando MFA... ({elapsed_minutes}m {elapsed_seconds}s)")
            
            # Executar verificação JavaScript
            resultado = executar_js_safe(driver, js_verificar_mfa)
            
            if resultado:
                # Se já está logado (passou do MFA)
                if resultado.get('isLoggedIn'):
                    log_ok("✓ MFA aprovado! Login concluído.")
                    time.sleep(1)  # Aguarda um pouco para garantir carregamento
                    return True
                
                # Se claramente NÃO está mais no MFA
                if not resultado.get('stillInMFA') and resultado.get('hasLightning'):
                    log_ok("✓ MFA concluído! Redirecionado com sucesso.")
                    time.sleep(1)
                    return True
            
            # Verificação adicional pela URL (fallback)
            try:
                url_atual = driver.current_url.lower()
                
                # Se já está no Lightning e não tem indicadores de MFA
                if 'lightning' in url_atual and 'verification' not in url_atual and 'mfa' not in url_atual:
                    # Verificar se há elementos da interface do Salesforce
                    try:
                        header = driver.find_elements(By.XPATH, "//header[contains(@class,'slds-global-header')]")
                        if header:
                            log_ok("✓ Interface Lightning detectada! MFA concluído.")
                            time.sleep(0.5)
                            return True
                    except:
                        pass
            except:
                pass
            
            # Aguardar antes da próxima verificação
            time.sleep(1)
            
        except Exception as e:
            log_debug(f"Erro na verificação MFA: {str(e)[:100]}")
            time.sleep(1)
    
    # Se chegou aqui, deu timeout
    log_error(f"⏱️  Timeout MFA ({timeout}s)")
    log_warn("O usuário não completou a autenticação a tempo.")
    
    return False

def logar(driver, usuario, senha):
    """Função de login corrigida com esperas e JavaScript"""
    
    log_info("Realizando login...")
    
    try:
        # Aguardar a página carregar completamente
        time.sleep(2)
        
        # JavaScript para preencher campos de forma mais confiável
        js_fill_and_submit = """
        function findAndFill(selector, value) {
            const el = document.querySelector(selector);
            if (!el) return false;
            
            // Rolar até o elemento
            el.scrollIntoView({block: 'center'});
            
            // Focar
            el.focus();
            
            // Limpar e preencher
            el.value = '';
            el.value = value;
            
            // Disparar eventos
            el.dispatchEvent(new Event('input', {bubbles: true}));
            el.dispatchEvent(new Event('change', {bubbles: true}));
            
            return true;
        }
        
        const username = arguments[0];
        const password = arguments[1];
        
        // Tentar diferentes seletores para username
        const userSelectors = [
            'input[name="username"]',
            'input[id="username"]',
            'input[type="text"]',
            'input[type="email"]'
        ];
        
        let userFilled = false;
        for (const sel of userSelectors) {
            if (findAndFill(sel, username)) {
                userFilled = true;
                break;
            }
        }
        
        if (!userFilled) return { success: false, error: 'Username field not found' };
        
        // Aguardar um pouco
        await new Promise(r => setTimeout(r, 300));
        
        // Tentar diferentes seletores para password
        const passSelectors = [
            'input[name="pw"]',
            'input[id="password"]',
            'input[type="password"]'
        ];
        
        let passFilled = false;
        for (const sel of passSelectors) {
            if (findAndFill(sel, password)) {
                passFilled = true;
                break;
            }
        }
        
        if (!passFilled) return { success: false, error: 'Password field not found' };
        
        // Aguardar um pouco
        await new Promise(r => setTimeout(r, 300));
        
        // Procurar e clicar no botão de login
        const loginSelectors = [
            'input[name="Login"]',
            'input[id="Login"]',
            'button[type="submit"]',
            'input[type="submit"]'
        ];
        
        for (const sel of loginSelectors) {
            const btn = document.querySelector(sel);
            if (btn) {
                btn.click();
                return { success: true };
            }
        }
        
        return { success: false, error: 'Login button not found' };
        """
        
        # Executar JavaScript para fazer login
        result = driver.execute_script(js_fill_and_submit, usuario, senha)
        
        if result and result.get('success'):
            log_ok("Credenciais preenchidas e login clicado via JavaScript")
            
            # Aguardar redirecionamento
            time.sleep(3)
            
            # Verificar se precisa de MFA
            current_url = driver.current_url
            page_text = driver.page_source.lower()

            esperar_mfa(driver, timeout=TIMEOUT_MFA)
            
            # Verificar se login foi bem-sucedido
            current_url = driver.current_url
            
            if 'lightning' in current_url or 'salesforce' in current_url:
                log_ok("Login realizado com sucesso!")
                return True
            else:
                log_warn("Aguardando redirecionamento...")
                time.sleep(2)
                return True
        else:
            log_error(f"Erro no JavaScript: {result.get('error', 'Desconhecido')}")
            
            # Fallback: Tentar com Selenium tradicional mas com esperas
            log_info("Tentando método alternativo com Selenium...")
            
            wait = WebDriverWait(driver, 10)
            
            # Aguardar e preencher username
            try:
                username_el = wait.until(
                    EC.element_to_be_clickable((By.NAME, "username"))
                )
                
                # Usar JavaScript para preencher (mais confiável que .clear() + .send_keys())
                driver.execute_script("arguments[0].value = arguments[1];", username_el, usuario)
                driver.execute_script("arguments[0].dispatchEvent(new Event('input', {bubbles: true}));", username_el)
                
                log_ok("Username preenchido")
                
            except Exception as e:
                log_error(f"Erro ao preencher username: {e}")
                return False
            
            time.sleep(0.3)
            
            # Aguardar e preencher password
            try:
                password_el = wait.until(
                    EC.element_to_be_clickable((By.NAME, "pw"))
                )
                
                driver.execute_script("arguments[0].value = arguments[1];", password_el, senha)
                driver.execute_script("arguments[0].dispatchEvent(new Event('input', {bubbles: true}));", password_el)
                
                log_ok("Password preenchido")
                
            except Exception as e:
                log_error(f"Erro ao preencher password: {e}")
                return False
            
            time.sleep(0.3)
            
            # Clicar no botão de login
            try:
                login_btn = wait.until(
                    EC.element_to_be_clickable((By.NAME, "Login"))
                )
                
                driver.execute_script("arguments[0].click();", login_btn)
                log_ok("Botão de login clicado")
                
                time.sleep(3)
                return True
                
            except Exception as e:
                log_error(f"Erro ao clicar em login: {e}")
                return False
    
    except Exception as e:
        log_error(f"Erro no login: {e}")
        return False


def logar_salesforce_robusto(driver, usuario, senha, max_tentativas=3):
    """
    Versão ainda mais robusta com múltiplas tentativas
    """
    
    for tentativa in range(1, max_tentativas + 1):
        log_info(f"Tentativa de login {tentativa}/{max_tentativas}...")
        
        try:
            # Atualizar página se não for primeira tentativa
            if tentativa > 1:
                log_info("Recarregando página...")
                driver.refresh()
                time.sleep(3)
            
            # Tentar login
            if logar(driver, usuario, senha):
                return True
            
            # Se falhou, aguardar antes de tentar novamente
            if tentativa < max_tentativas:
                log_warn(f"Falha na tentativa {tentativa}. Aguardando 2s...")
                time.sleep(2)
        
        except Exception as e:
            log_error(f"Erro na tentativa {tentativa}: {e}")
            if tentativa < max_tentativas:
                time.sleep(2)
    
    log_error("Todas as tentativas de login falharam!")
    return False


def verificar_login_salesforce(driver):
    """Verifica se o login foi bem-sucedido"""
    
    time.sleep(2)
    
    current_url = driver.current_url
    page_source = driver.page_source.lower()
    
    # Verificações positivas (está logado)
    if any([
        'lightning' in current_url,
        'internal' in current_url,
        'home' in current_url,
        'setuphome' in page_source,
        'salesforce' in current_url and 'login' not in current_url
    ]):
        log_ok("✓ Login verificado com sucesso!")
        return True
    
    # Verificações negativas (não está logado)
    if any([
        'login' in current_url,
        'error' in page_source,
        'invalid' in page_source
    ]):
        log_error("✗ Login não realizado ou credenciais inválidas")
        return False
    
    log_warn("⚠ Status de login incerto")
    return None

def executar_js_safe(driver, script, *args):
    try:
        return driver.execute_script(script, *args)
    except Exception as e:
        log_debug(f"JS Error: {str(e)[:100]}")
        return None

def verificar_pagina_inicial(driver, timeout=10):
    log_info("Verificando página atual...")
    
    time.sleep(1.2)
    
    js_verificar = """
    const url = window.location.href;
    
    if (url.includes('/lightning/r/Account/') || 
        url.includes('/lightning/r/Contact/') ||
        url.includes('/view')) {
        return { onHome: false, onClient: true, url: url };
    }
    
    if (url.includes('/lightning/page/home')) {
        return { onHome: true, onClient: false, url: url };
    }
    
    const links = document.querySelectorAll('a[href*="/lightning/page/home"]');
    for (const link of links) {
        const text = (link.innerText || link.textContent || '').trim();
        if (text.toLowerCase() === 'início' || text.toLowerCase() === 'inicio') {
            return { onHome: true, onClient: false, url: url };
        }
    }
    
    return { onHome: false, onClient: false, url: url };
    """
    
    for i in range(3):
        time.sleep(1.2)
        
        resultado = executar_js_safe(driver, js_verificar)
        
        if resultado:
            if resultado.get('onClient'):
                log_warn("Está na página do CLIENTE - Navegando para Início...")
                url = resultado.get('url', '')
                
                try:
                    log_info("Navegando para Início...")
                    base_url = url.split('/lightning/')[0] if '/lightning/' in url else url.split('.com')[0] + '.com'
                    driver.get(base_url + '/lightning/page/home')
                    time.sleep(2)
                    
                    for _ in range(2):
                        time.sleep(1)
                        verif = executar_js_safe(driver, js_verificar)
                        if verif and verif.get('onHome'):
                            log_ok("Navegação bem-sucedida")
                            return True
                    
                    log_warn("Ainda não está em Início")
                    return False
                except Exception as e:
                    log_error(f"Erro ao navegar: {e}")
                    return False
            
            if resultado.get('onHome'):
                log_ok("Página Início detectada")
                return True
        
        if i < 2:
            log_debug(f"Verificação {i+1}/3 - aguardando...")
    
    log_warn("Não identificou a página")
    return False

def verificar_notificacao_erro_cpf(driver):
    """Verifica se há notificação de erro de CPF (inválido ou não encontrado)"""
    
    js_verificar_notificacao = """
    const toasts = document.querySelectorAll('.forceToastMessage');
    
    for (const toast of toasts) {
        const isVisible = toast.offsetWidth > 0 && toast.offsetHeight > 0;
        
        if (isVisible) {
            const isError = toast.classList.contains('slds-theme--error');
            const isWarning = toast.classList.contains('slds-theme--warning');
            
            if (isError || isWarning) {
                const messageElement = toast.querySelector('.toastMessage');
                if (messageElement) {
                    const message = (messageElement.innerText || messageElement.textContent || '').trim();
                    
                    if (message.toLowerCase().includes('cpf inválido') || 
                        message.toLowerCase().includes('cpf invalido')) {
                        return { hasError: true, type: 'invalid', message: message };
                    }
                    
                    if (message.toLowerCase().includes('cliente não encontrado') ||
                        message.toLowerCase().includes('cliente nao encontrado') ||
                        message.toLowerCase().includes('não encontrado') ||
                        message.toLowerCase().includes('nao encontrado')) {
                        return { hasError: true, type: 'not_found', message: message };
                    }
                }
            }
        }
    }
    
    return { hasError: false };
    """
    
    try:
        resultado = executar_js_safe(driver, js_verificar_notificacao)
        
        if resultado and resultado.get('hasError'):
            erro_tipo = resultado.get('type')
            mensagem = resultado.get('message')
            
            if erro_tipo == 'invalid':
                log_error(f"❌ {mensagem}")
                return 'invalid'
            elif erro_tipo == 'not_found':
                log_warn(f"⚠️ {mensagem}")
                return 'not_found'
        
        return None
        
    except Exception as e:
        log_debug(f"Erro ao verificar notificação: {str(e)[:100]}")
        return None

def buscar_cpf_automatico(driver, cpf, max_tentativas=3):
    wait = WebDriverWait(driver, TIMEOUT_SEARCH)
    
    log_info(f"Buscando CPF {cpf}...")
    
    try:
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//header[contains(@class,'slds-global-header')]")
        ))
    except TimeoutException:
        log_error("Timeout Lightning")
        return False
    
    log_info("Aguardando carregamento completo da página...")
    time.sleep(0.3)
    
    for tentativa in range(1, max_tentativas + 1):
        log_info(f"Tentativa {tentativa}/{max_tentativas}...")
        
        try:
            log_debug("Localizando input para digitação...")
            
            input_element = None
            
            try:
                wait = WebDriverWait(driver, 1)
                input_element = wait.until(
                    EC.element_to_be_clickable((By.NAME, "inputSearch"))
                )
            except:
                try:
                    input_element = wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='search']"))
                    )
                except:
                    try:
                        input_element = wait.until(
                            EC.element_to_be_clickable((By.XPATH, "//input[contains(@placeholder, 'CPF') or contains(@placeholder, 'CLI')]"))
                        )
                    except:
                        log_warn("Input não encontrado ou não está clicável")
                        time.sleep(0.2)
                        continue
            
            if input_element:
                log_debug("Input encontrado e clicável!")
                
                try:
                    driver.execute_script("""
                        const overlays = document.querySelectorAll('.slds-backdrop, .slds-modal, [role="dialog"]');
                        overlays.forEach(el => {
                            if (el.style) el.style.display = 'none';
                        });
                    """)
                    time.sleep(0.2)
                    
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'instant'});", input_element)
                    time.sleep(0.2)
                    
                    driver.execute_script("""
                        arguments[0].focus();
                        arguments[0].click();
                    """, input_element)
                    time.sleep(0.2)
                    
                    driver.execute_script("arguments[0].value = '';", input_element)
                    time.sleep(0.2)
                    
                    for i, char in enumerate(cpf):
                        driver.execute_script("""
                            const input = arguments[0];
                            const char = arguments[1];
                            
                            input.value += char;
                            
                            input.dispatchEvent(new KeyboardEvent('keydown', {key: char, bubbles: true}));
                            input.dispatchEvent(new KeyboardEvent('keypress', {key: char, bubbles: true}));
                            input.dispatchEvent(new Event('input', {bubbles: true}));
                            input.dispatchEvent(new KeyboardEvent('keyup', {key: char, bubbles: true}));
                        """, input_element, char)
                        time.sleep(0.02)
                    
                    driver.execute_script("""
                        arguments[0].dispatchEvent(new Event('change', {bubbles: true}));
                    """, input_element)
                    
                    time.sleep(0.5)
                    
                    valor_digitado = input_element.get_attribute('value')
                    
                    if valor_digitado and cpf in valor_digitado.replace('-', '').replace('.', ''):
                        log_ok(f"CPF digitado com sucesso: {valor_digitado}")
                        time.sleep(0.2)
                    else:
                        log_warn(f"CPF pode não ter sido digitado corretamente. Valor no campo: '{valor_digitado}'")
                        time.sleep(0.2)
                    
                except Exception as e:
                    log_warn(f"Erro na digitação: {str(e)[:100]}")
                    time.sleep(0.2)
                    continue
            else:
                log_warn("Input não foi localizado")
                time.sleep(0.3)
                continue
            
            script_buscar = """
            let searchButton = null;
            
            const brandButtons = Array.from(document.querySelectorAll('button.slds-button_brand, button[class*="slds-button"]'));
            for (const btn of brandButtons) {
                const text = (btn.innerText || btn.textContent || '').trim().toLowerCase();
                const title = (btn.getAttribute('title') || '').toLowerCase();
                
                if (text === 'buscar' || title === 'submit' || text.includes('buscar')) {
                    searchButton = btn;
                    break;
                }
            }
            
            if (!searchButton) {
                const searchInShadow = (root) => {
                    const buttons = root.querySelectorAll('button');
                    for (const btn of buttons) {
                        const text = (btn.innerText || btn.textContent || '').trim().toLowerCase();
                        const title = (btn.getAttribute('title') || '').toLowerCase();
                        if (text === 'buscar' || text.includes('buscar') || title === 'submit') {
                            return btn;
                        }
                    }
                    
                    const allElements = root.querySelectorAll('*');
                    for (const el of allElements) {
                        if (el.shadowRoot) {
                            const found = searchInShadow(el.shadowRoot);
                            if (found) return found;
                        }
                    }
                    return null;
                };
                
                searchButton = searchInShadow(document);
            }
            
            if (!searchButton) {
                const input = document.querySelector('input[name="inputSearch"]');
                if (input) {
                    const parent = input.closest('form, div, lightning-card') || input.parentElement;
                    if (parent) {
                        const nearButtons = parent.querySelectorAll('button');
                        for (const btn of nearButtons) {
                            const text = (btn.innerText || btn.textContent || '').trim().toLowerCase();
                            if (text && text.length < 20) {
                                searchButton = btn;
                                break;
                            }
                        }
                    }
                }
            }
            
            if (!searchButton) {
                return { success: false, error: 'Botão Buscar não encontrado após 3 métodos' };
            }
            
            searchButton.scrollIntoView({block: 'center'});
            searchButton.click();
            searchButton.dispatchEvent(new MouseEvent('click', {bubbles: true}));
            
            return { success: true, buttonText: searchButton.innerText || searchButton.textContent };
            """
            
            resultado_click = executar_js_safe(driver, script_buscar)
            
            if not resultado_click or not resultado_click.get('success'):
                log_warn(f"Erro ao clicar Buscar: {resultado_click.get('error') if resultado_click else 'sem resposta'}")
                
                log_debug("Tentando via Selenium...")
                try:
                    btn_selenium = driver.find_element(By.XPATH, "//button[contains(text(), 'Buscar') or @title='Submit']")
                    btn_selenium.click()
                    log_ok("Botão Buscar clicado via Selenium")
                except Exception as e:
                    log_debug(f"Selenium também falhou: {str(e)[:60]}")
                    time.sleep(1)
                    continue
            else:
                log_ok(f"Botão Buscar clicado: {resultado_click.get('buttonText', 'Buscar')}")
            
            log_info("Verificando resposta da busca...")
            time.sleep(3)
            
            erro_tipo = verificar_notificacao_erro_cpf(driver)
            
            if erro_tipo:
                return erro_tipo
            
            log_info("Aguardando resultado aparecer (usando Selenium)...")
            
            resultado_apareceu = False
            elemento_resultado = None
            
            seletores_resultado = [
                "//a[contains(@class, 'slds-p-') or contains(@class, 'slds-m-')]",
                "//a[contains(@href, '#') and string-length(text()) > 10]",
                "//a[contains(text(), ' ') and not(contains(text(), 'Pular')) and not(contains(text(), 'Início')) and not(contains(text(), 'Ações'))]",
                "//div[contains(@class, 'search')]//a",
                "//div[contains(@class, 'result')]//a",
                "//lightning-formatted-name//a",
                "//span[contains(@class, 'uiOutputText')]/..//a"
            ]
            
            for tentativa_espera in range(20):
                try:
                    for seletor in seletores_resultado:
                        try:
                            elementos = driver.find_elements(By.XPATH, seletor)
                            
                            for elem in elementos:
                                try:
                                    if not elem.is_displayed():
                                        continue
                                    
                                    texto = elem.text.strip()
                                    
                                    if not texto or len(texto) < 6:
                                        continue
                                    
                                    texto_lower = texto.lower()
                                    
                                    if any(palavra in texto_lower for palavra in ['pular', 'skip', 'navegação', 'navigation', 
                                                                                     'início', 'inicio', 'cases', 'contas',
                                                                                     'configurações', 'home', 'help', 'ajuda',
                                                                                     'ações globais', 'global actions', 'ações']):
                                        continue
                                    
                                    if ' ' in texto and any(c.isupper() for c in texto) and not texto.isdigit():
                                        log_ok(f"✓ Resultado encontrado via Selenium: {texto[:40]}")
                                        elemento_resultado = elem
                                        resultado_apareceu = True
                                        break
                                
                                except Exception:
                                    continue
                            
                            if resultado_apareceu:
                                break
                        
                        except Exception:
                            continue
                    
                    if resultado_apareceu:
                        break
                    
                    if tentativa_espera % 3 == 0:
                        log_info(f"Aguardando resultado... ({tentativa_espera + 1}s)")
                    
                    time.sleep(1)
                
                except Exception as e:
                    log_debug(f"Erro na tentativa {tentativa_espera}: {str(e)[:50]}")
                    time.sleep(1)
            
            if not resultado_apareceu:
                log_warn("Resultado não apareceu após 20s")
                time.sleep(0.5)
                continue
            
            log_info("Clicando no resultado...")
            
            if elemento_resultado:
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento_resultado)
                    time.sleep(0.5)
                    
                    clicado = False
                    
                    try:
                        elemento_resultado.click()
                        clicado = True
                    except Exception:
                        pass
                    
                    if not clicado:
                        try:
                            actions = ActionChains(driver)
                            actions.move_to_element(elemento_resultado).click().perform()
                            clicado = True
                        except Exception:
                            pass
                    
                    if not clicado:
                        try:
                            driver.execute_script("arguments[0].click();", elemento_resultado)
                            clicado = True
                        except Exception:
                            pass
                    
                    if clicado:
                        time.sleep(2.5)
                        
                        try:
                            url_atual = driver.current_url
                            if '/lightning/r/' in url_atual or '/Account/' in url_atual or '/Contact/' in url_atual or '/view' in url_atual:
                                log_ok("✓ Navegação confirmada para página do cliente!")
                                # Armazena a URL do cliente para referência futura
                                _GLOBAL_RESOURCES['cliente_url'] = url_atual
                                return True
                            else:
                                log_warn(f"Ainda na página: {url_atual[:60]}")
                        except:
                            pass
                        
                        return True
                    else:
                        log_error("Não conseguiu clicar no resultado")
                        time.sleep(1)
                        continue
                
                except Exception as e:
                    log_error(f"Erro ao clicar: {str(e)[:100]}")
                    time.sleep(1)
                    continue
            
            log_warn(f"Resultado não encontrado na tentativa {tentativa}")
            time.sleep(1)
            
        except Exception as e:
            log_debug(f"Exceção: {str(e)[:100]}")
            time.sleep(1)
    
    log_error(f"Falha após {max_tentativas} tentativas")
    return False

def selecionar_combobox_melhorado(driver, label, arrow_count, descricao="", max_tentativas=3):
    log_info(f"Selecionando '{label}' (opção {arrow_count})")
    
    js_preparar = """
    const label = arguments[0];
    
    function findInShadow(root, selector, attrCheck) {
        const elements = Array.from(root.querySelectorAll(selector));
        for (const el of elements) {
            if (attrCheck(el)) {
                const rect = el.getBoundingClientRect();
                if (rect.width > 0 && rect.height > 0) return el;
            }
        }
        
        const allElements = root.querySelectorAll('*');
        for (const el of allElements) {
            try {
                if (el.shadowRoot) {
                    const found = findInShadow(el.shadowRoot, selector, attrCheck);
                    if (found) return found;
                }
            } catch(e) {}
        }
        return null;
    }
    
    const button = findInShadow(
        document,
        'button[role="combobox"]',
        (btn) => {
            const ariaLabel = (btn.getAttribute('aria-label') || '').toLowerCase();
            return ariaLabel === label.toLowerCase();
        }
    );
    
    if (!button) {
        return { success: false, error: 'Button not found' };
    }
    
    button.scrollIntoView({block: 'center', behavior: 'instant'});
    
    let w = 0;
    while(w < 30) {
        const s = Date.now();
        while(Date.now() - s < 5) {}
        w += 5;
    }
    
    button.classList.add('sf-automation-target');
    window.__sfAutomationButton = button;
    
    return { success: true, hasClass: true };
    """
    
    js_verificar_aberto = """
    const button = window.__sfAutomationButton || document.querySelector('.sf-automation-target');
    
    if (!button) return { opened: false, error: 'Button lost' };
    
    const dropdownId = button.getAttribute('aria-controls');
    const expanded = button.getAttribute('aria-expanded');
    
    function findDropdown(root, id) {
        if (id) {
            const el = root.getElementById(id);
            if (el) return el;
        }
        
        const all = root.querySelectorAll('*');
        for (const elem of all) {
            try {
                if (elem.shadowRoot) {
                    const found = findDropdown(elem.shadowRoot, id);
                    if (found) return found;
                }
            } catch(e) {}
        }
        return null;
    }
    
    let dropdown = findDropdown(document, dropdownId);
    
    if (!dropdown) {
        const parent = button.closest('lightning-combobox, .slds-combobox') || button.parentElement;
        if (parent) {
            dropdown = parent.querySelector('div[role="listbox"], ul[role="listbox"]');
        }
    }
    
    if (dropdown) {
        const rect = dropdown.getBoundingClientRect();
        const items = dropdown.querySelectorAll('[role="option"]');
        
        if (rect.height > 20 && items.length > 0) {
            return { 
                opened: true, 
                items: items.length,
                expanded: expanded === 'true'
            };
        }
    }
    
    return { 
        opened: false, 
        expanded: expanded === 'true',
        hasDropdownId: !!dropdownId
    };
    """
    
    js_clicar_opcao = """
    const targetIndex = arguments[0];
    
    const button = window.__sfAutomationButton || document.querySelector('.sf-automation-target');
    if (!button) return { success: false, error: 'Button not found' };
    
    const dropdownId = button.getAttribute('aria-controls');
    
    function findDropdown(root, id) {
        if (id) {
            const el = root.getElementById(id);
            if (el) return el;
        }
        const all = root.querySelectorAll('*');
        for (const elem of all) {
            try {
                if (elem.shadowRoot) {
                    const found = findDropdown(elem.shadowRoot, id);
                    if (found) return found;
                }
            } catch(e) {}
        }
        return null;
    }
    
    let dropdown = findDropdown(document, dropdownId);
    
    if (!dropdown) {
        const parent = button.closest('lightning-combobox, .slds-combobox') || button.parentElement;
        if (parent) {
            dropdown = parent.querySelector('div[role="listbox"], ul[role="listbox"]');
        }
    }
    
    if (!dropdown) {
        return { success: false, error: 'Dropdown not found' };
    }
    
    const options = Array.from(dropdown.querySelectorAll('[role="option"]'));
    
    if (options.length === 0) {
        return { success: false, error: 'No options found' };
    }
    
    const validOptions = options.filter(opt => {
        const text = (opt.innerText || opt.textContent || '').trim();
        return text && text !== '--Nenhum--' && text !== 'Nenhum';
    });
    
    if (validOptions.length === 0) {
        return { success: false, error: 'No valid options' };
    }
    
    if (targetIndex > validOptions.length) {
        return { success: false, error: `Index ${targetIndex} out of range (${validOptions.length} options disponíveis)` };
    }
    
    const actualIndex = Math.min(targetIndex, validOptions.length - 1);
    const targetOption = validOptions[actualIndex];
    const optionText = (targetOption.innerText || targetOption.textContent || '').trim();
    
    try {
        const dropdownRect = dropdown.getBoundingClientRect();
        const optionRect = targetOption.getBoundingClientRect();
        
        if (optionRect.bottom > dropdownRect.bottom || optionRect.top < dropdownRect.top) {
            dropdown.scrollTop += (optionRect.top - dropdownRect.top) - (dropdownRect.height / 2);
            
            let w = 0;
            while(w < 20) {
                const s = Date.now();
                while(Date.now() - s < 5) {}
                w += 5;
            }
        }
        
        targetOption.scrollIntoView({block: 'nearest', behavior: 'auto'});
        
        let w = 0;
        while(w < 20) {
            const s = Date.now();
            while(Date.now() - s < 5) {}
            w += 5;
        }
        
        targetOption.click();
        
        targetOption.dispatchEvent(new MouseEvent('mousedown', {bubbles: true}));
        targetOption.dispatchEvent(new MouseEvent('mouseup', {bubbles: true}));
        targetOption.dispatchEvent(new MouseEvent('click', {bubbles: true}));
        
        w = 0;
        while(w < 50) {
            const s = Date.now();
            while(Date.now() - s < 5) {}
            w += 5;
        }
        
        const span = button.querySelector('span.slds-truncate');
        const currentValue = span ? (span.innerText || '').trim() : '';
        
        button.classList.remove('sf-automation-target');
        delete window.__sfAutomationButton;
        
        if (currentValue && currentValue !== '--Nenhum--') {
            return { 
                success: true, 
                value: currentValue,
                targetText: optionText,
                method: 'click'
            };
        }
        
        return { 
            success: true,
            value: optionText,
            targetText: optionText,
            method: 'click'
        };
        
    } catch(e) {
        return { success: false, error: String(e) };
    }
    """
    
    for tentativa in range(1, max_tentativas + 1):
        try:
            prep = executar_js_safe(driver, js_preparar, label)
            if not prep or not prep.get('success'):
                time.sleep(0.01)
                continue
            
            time.sleep(0.01)
            
            try:
                try:
                    button_element = driver.find_element(By.CLASS_NAME, 'sf-automation-target')
                except:
                    button_element = driver.execute_script("return window.__sfAutomationButton;")
                    if not button_element:
                        raise Exception("Button not accessible")
                
                driver.execute_script("arguments[0].focus();", button_element)
                time.sleep(0.01)
                
                actions = ActionChains(driver)
                actions.move_to_element(button_element).pause(0.01).click().pause(0.05).perform()
                
                time.sleep(0.01)
                
            except Exception:
                executar_js_safe(driver, "const b = document.querySelector('.sf-automation-target'); if(b) b.classList.remove('sf-automation-target'); delete window.__sfAutomationButton;")
                time.sleep(0.01)
                continue
            
            verif = executar_js_safe(driver, js_verificar_aberto)
            
            if not verif:
                time.sleep(0.01)
                continue
            
            if not verif.get('opened'):
                if tentativa < max_tentativas:
                    try:
                        button_element.click()
                        time.sleep(0.03)
                        button_element.click()
                        time.sleep(0.1)
                        
                        verif2 = executar_js_safe(driver, js_verificar_aberto)
                        if not verif2 or not verif2.get('opened'):
                            time.sleep(0.01)
                            continue
                    except Exception:
                        time.sleep(0.01)
                        continue
                else:
                    time.sleep(0.01)
                    continue
            
            option_index = arrow_count - 1
            
            resultado = executar_js_safe(driver, js_clicar_opcao, option_index)
            
            if resultado and resultado.get('success'):
                valor = resultado.get('value', descricao)
                log_ok(f"Selecionado: {valor}")
                return True
            else:
                time.sleep(0.01)
            
        except Exception:
            time.sleep(0.01)
    
    log_warn(f"Automação falhou após {max_tentativas} tentativas")
    
    manual = input(f"\nSelecionar '{descricao}' MANUALMENTE? (s/n): ").strip().lower()
    
    if manual == 's':
        input(f"Selecione '{descricao}' e pressione Enter...")
        log_ok(f"Seleção manual: {descricao}")
        return True
    
    log_warn(f"'{label}' não foi selecionado")
    return False

def voltar_para_cliente(driver):
    """Navega de volta para a aba do cliente (Account) após salvar um caso"""
    global _GLOBAL_RESOURCES
    
    log_info("Retornando para a página do cliente...")
    
    # PRIMEIRO: Verificar se JÁ está na página do cliente
    try:
        url_atual = driver.current_url
        log_debug(f"URL atual: {url_atual[:70]}")
        
        if '/lightning/r/Account/' in url_atual or '/lightning/r/Contact/' in url_atual:
            log_ok("✓ Já está na página do cliente!")
            
            # Atualizar a URL armazenada se necessário
            if url_atual != _GLOBAL_RESOURCES.get('cliente_url'):
                _GLOBAL_RESOURCES['cliente_url'] = url_atual
                log_debug("URL do cliente atualizada")
            
            return True
    except Exception as e:
        log_debug(f"Erro ao verificar URL: {str(e)[:60]}")
    
    # MÉTODO 1: Usar a URL armazenada (mais confiável)
    if _GLOBAL_RESOURCES.get('cliente_url'):
        try:
            url_cliente = _GLOBAL_RESOURCES['cliente_url']
            log_debug(f"Tentando URL armazenada: {url_cliente[:50]}...")
            
            driver.get(url_cliente)
            time.sleep(2)
            
            # Verificar se chegou na página correta
            url_nova = driver.current_url
            if '/lightning/r/Account/' in url_nova or '/lightning/r/Contact/' in url_nova:
                log_ok("✓ Retornou para o cliente usando URL direta!")
                return True
            else:
                log_debug(f"URL não parece ser do cliente: {url_nova[:60]}")
        except Exception as e:
            log_debug(f"Erro ao usar URL direta: {str(e)[:60]}")
    
    # MÉTODO 2: Procurar e clicar na aba do cliente
    log_info("Tentando encontrar aba do cliente...")
    
    js_voltar = """
    const clienteUrl = arguments[0];
    
    // Procura todas as abas abertas
    const tabs = document.querySelectorAll('a[role="tab"]');
    
    let targetTab = null;
    let tabInfo = [];
    
    // Coletar informações de todas as abas para debug
    for (const tab of tabs) {
        const href = tab.getAttribute('href') || '';
        const title = tab.getAttribute('title') || '';
        
        tabInfo.push({
            href: href.substring(0, 50),
            title: title
        });
        
        // Verifica se é uma aba de Account/Contact (cliente)
        if (href.includes('/lightning/r/Account/') || 
            href.includes('/lightning/r/Contact/') ||
            (clienteUrl && href && clienteUrl.includes(href))) {
            
            // Verificar se a aba está visível
            const rect = tab.getBoundingClientRect();
            if (rect.width > 0 && rect.height > 0) {
                targetTab = tab;
                break;
            }
        }
    }
    
    if (!targetTab) {
        return { 
            success: false, 
            error: 'Aba do cliente não encontrada',
            tabs: tabInfo
        };
    }
    
    try {
        // Rolar até a aba
        targetTab.scrollIntoView({block: 'center', behavior: 'instant'});
        
        // Aguardar um pouco
        let w = 0;
        while(w < 100) {
            const s = Date.now();
            while(Date.now() - s < 5) {}
            w += 5;
        }
        
        // Clicar na aba
        targetTab.click();
        
        // Disparar eventos adicionais
        targetTab.dispatchEvent(new MouseEvent('mousedown', {bubbles: true}));
        targetTab.dispatchEvent(new MouseEvent('mouseup', {bubbles: true}));
        
        return { 
            success: true, 
            title: targetTab.getAttribute('title') || 'Cliente',
            href: targetTab.getAttribute('href')
        };
    } catch(e) {
        return { success: false, error: String(e) };
    }
    """
    
    try:
        cliente_url = _GLOBAL_RESOURCES.get('cliente_url', '')
        resultado = executar_js_safe(driver, js_voltar, cliente_url)
        
        if resultado and resultado.get('success'):
            log_ok(f"✓ Clicou na aba: {resultado.get('title', 'Cliente')}")
            time.sleep(2)
            
            # Verificar se realmente mudou de página
            url_final = driver.current_url
            if '/lightning/r/Account/' in url_final or '/lightning/r/Contact/' in url_final:
                log_ok("✓ Confirmado na página do cliente!")
                _GLOBAL_RESOURCES['cliente_url'] = url_final
                return True
            else:
                log_warn(f"Aba clicada, mas URL não mudou corretamente: {url_final[:60]}")
        else:
            erro = resultado.get('error', 'sem resposta') if resultado else 'sem resposta'
            log_warn(f"Não encontrou aba do cliente: {erro}")
            
            # Debug: mostrar abas encontradas
            if resultado and resultado.get('tabs'):
                log_debug("Abas encontradas:")
                for tab in resultado.get('tabs')[:3]:
                    log_debug(f"  - {tab.get('title', 'sem título')}: {tab.get('href', 'sem href')}")
                    
    except Exception as e:
        log_error(f"Erro ao procurar aba: {str(e)[:100]}")
    
    # MÉTODO 3: Usar histórico do navegador (voltar página)
    log_info("Tentando voltar pelo histórico do navegador...")
    
    try:
        # Voltar 1 página
        driver.back()
        time.sleep(2)
        
        url_back = driver.current_url
        if '/lightning/r/Account/' in url_back or '/lightning/r/Contact/' in url_back:
            log_ok("✓ Voltou para cliente usando histórico!")
            _GLOBAL_RESOURCES['cliente_url'] = url_back
            return True
        else:
            log_debug(f"Histórico não levou ao cliente: {url_back[:60]}")
            # Tentar voltar mais uma vez
            driver.back()
            time.sleep(2)
            
            url_back2 = driver.current_url
            if '/lightning/r/Account/' in url_back2 or '/lightning/r/Contact/' in url_back2:
                log_ok("✓ Voltou para cliente usando histórico (2 páginas)!")
                _GLOBAL_RESOURCES['cliente_url'] = url_back2
                return True
    except Exception as e:
        log_debug(f"Erro ao usar histórico: {str(e)[:60]}")
    
    # Se chegou aqui, nenhum método funcionou
    log_error("✗ Não conseguiu voltar automaticamente para o cliente")
    return False

def registrar_informacao_automatico(driver):
    log_info("Iniciando registro automático...")
    
    js_click = """
    const query = arguments[0];
    const mode = arguments[1] || 'selector';
    
    function findDeep(root, q, isText) {
        if (!isText) {
            try {
                const el = root.querySelector(q);
                if (el && isVisible(el)) return el;
            } catch(e){}
        } else {
            const tags = ['button', 'a', 'span', 'lightning-button'];
            for (const tag of tags) {
                const els = Array.from(root.querySelectorAll(tag));
                for (const el of els) {
                    if (isVisible(el)) {
                        const text = (el.innerText || '').trim();
                        if (text.toLowerCase().includes(q.toLowerCase())) return el;
                    }
                }
            }
        }
        
        const all = root.querySelectorAll('*');
        for (const el of all) {
            try {
                if (el.shadowRoot) {
                    const found = findDeep(el.shadowRoot, q, isText);
                    if (found) return found;
                }
            } catch(e){}
        }
        return null;
    }
    
    function isVisible(el) {
        try {
            const rect = el.getBoundingClientRect();
            const style = window.getComputedStyle(el);
            return rect.width > 0 && rect.height > 0 && 
                   style.display !== 'none' && style.visibility !== 'hidden';
        } catch(e) {
            return false;
        }
    }
    
    const el = findDeep(document, query, mode === 'text');
    if (!el) return { success: false };
    
    try {
        el.scrollIntoView({block: 'center'});
        el.click();
        return { success: true };
    } catch(e) {
        try {
            ['mousedown', 'click'].forEach(ev => {
                el.dispatchEvent(new MouseEvent(ev, {bubbles: true}));
            });
            return { success: true };
        } catch(e2) {
            return { success: false };
        }
    }
    """
    
    js_fill_textarea = """
    const text = arguments[0];
    
    function findDeep(root) {
        const textareas = Array.from(root.querySelectorAll('textarea'));
        for (const ta of textareas) {
            const rect = ta.getBoundingClientRect();
            if (rect.width > 0 && rect.height > 0) return ta;
        }
        
        const all = root.querySelectorAll('*');
        for (const el of all) {
            try {
                if (el.shadowRoot) {
                    const found = findDeep(el.shadowRoot);
                    if (found) return found;
                }
            } catch(e){}
        }
        return null;
    }
    
    const textarea = findDeep(document);
    if (!textarea) return { success: false };
    
    try {
        textarea.focus();
        textarea.value = text;
        textarea.dispatchEvent(new Event('input', {bubbles: true}));
        textarea.dispatchEvent(new Event('change', {bubbles: true}));
        return { success: true };
    } catch(e) {
        return { success: false };
    }
    """
    
    def click_element(query, mode='selector', tries=3):
        for i in range(tries):
            res = executar_js_safe(driver, js_click, query, mode)
            if res and res.get('success'):
                log_ok(f"Clicado: {query}")
                return True
            time.sleep(0.2)
        log_warn(f"Falha: {query}")
        return False
    
    print("\n" + "="*70)
    print("   REGISTRO AUTOMÁTICO")
    print("="*70 + "\n")
    
    log_info("1. Abrindo Casos...")
    if not click_element('a[data-tab-value="flexipage_tab3"]'):
        click_element('Casos', 'text')
    time.sleep(0.3)
    
    log_info("2. Clicando Criar...")
    if not click_element('button[name="NewCase"]'):
        click_element('Criar', 'text')
    time.sleep(0.5)
    
    log_info("3. Aguardando carregamento do formulário...")
    # AGUARDAR O FORMULÁRIO CARREGAR COMPLETAMENTE
    js_aguardar_radio = """
    let tentativas = 0;
    const maxTentativas = 30; // 3 segundos (30 x 100ms)
    
    function verificarRadios() {
        const labels = Array.from(document.querySelectorAll('label, span'));
        for (const label of labels) {
            const text = (label.innerText || '').toLowerCase();
            if (text.includes('informação') || text.includes('informacao') || 
                text.includes('dúvida') || text.includes('elogio')) {
                const input = label.querySelector('input[type="radio"]') ||
                             document.querySelector(`input[id="${label.getAttribute('for')}"]`);
                if (input) {
                    return true; // Radio encontrado
                }
            }
        }
        return false;
    }
    
    while (tentativas < maxTentativas) {
        if (verificarRadios()) {
            return { ready: true, tentativas: tentativas };
        }
        
        // Aguardar 100ms de forma síncrona
        const start = Date.now();
        while (Date.now() - start < 100) {}
        
        tentativas++;
    }
    
    return { ready: false, tentativas: tentativas };
    """
    
    resultado_espera = executar_js_safe(driver, js_aguardar_radio)
    
    if resultado_espera and resultado_espera.get('ready'):
        log_ok(f"Formulário carregado ({resultado_espera.get('tentativas')*100}ms)")
    else:
        log_warn("Formulário pode não ter carregado completamente")
        time.sleep(0.5)  # Espera adicional de segurança
    
    log_info("4. Selecionando tipo...")
    
    js_radio = """
    const labels = Array.from(document.querySelectorAll('label, span'));
    for (const label of labels) {
        const text = (label.innerText || '').toLowerCase();
        if (text.includes('informação') || text.includes('informacao') || 
            text.includes('dúvida') || text.includes('elogio')) {
            const input = label.querySelector('input[type="radio"]') ||
                         document.querySelector(`input[id="${label.getAttribute('for')}"]`);
            if (input) {
                try {
                    input.checked = true;
                    input.click();
                    input.dispatchEvent(new Event('change', {bubbles: true}));
                    return { success: true };
                } catch(e) {}
            }
        }
    }
    return { success: false };
    """
    
    res_radio = executar_js_safe(driver, js_radio)
    if res_radio and res_radio.get('success'):
        log_ok("Radio selecionado")
    
    time.sleep(0.1)
    
    log_info("5. Avançar...")
    for _ in range(4):
        if click_element('Avançar', 'text', tries=2):
            break
        time.sleep(0.05)
    time.sleep(0.8)
    
    log_info("6. Descrição...")
    try:
        descricao = input("\nDescrição: ").strip()
    except (EOFError, KeyboardInterrupt):
        descricao = ""
    
    if not descricao:
        descricao = "Registro de informação - Cliente solicitou informações"
        log_info("Descrição padrão aplicada")
    
    res_desc = executar_js_safe(driver, js_fill_textarea, descricao)
    if res_desc and res_desc.get('success'):
        log_ok("Descrição preenchida")
    
    log_info("7. Motivo do contato...")
    selecionar_combobox_melhorado(driver, 'Motivo do contato', 3, 'Informação')
    
    log_info("8. Origem do caso (Telefone - 13ª opção)...")
    selecionar_combobox_melhorado(driver, 'Origem do caso', 13, 'Telefone')
    
    log_info("9. Unidade de registro...")
    selecionar_combobox_melhorado(driver, 'Unidade de registro', 1, 'Unidade de registro')
    
    log_info("10. SAC responsável...")
    selecionar_combobox_melhorado(driver, 'SAC responsável', 1, 'SAC responsável')
    
    log_info("11. Status do caso (Concluído)...")
    selecionar_combobox_melhorado(driver, 'Status do caso', 2, 'Concluído')
    
    print("\n" + "="*70)
    log_ok("FORMULÁRIO COMPLETO!")
    print("="*70)
    print(f"\nDescrição: {descricao[:50]}...")
    print(f"Motivo: Informação (3ª opção)")
    print(f"Origem: Telefone (13ª opção)")
    print(f"Unidade: 1ª opção")
    print(f"SAC: 1ª opção")
    print(f"Status: Concluído (2ª opção)")
    print("="*70 + "\n")
    
    salvar = input("SALVAR CASO? (s/n): ").strip().lower()
    
    if salvar == 's' or salvar == '':
        log_info("Salvando...")
        
        if click_element('button[name="SaveEdit"]'):
            log_ok("CASO SALVO COM SUCESSO!")
            time.sleep(0.8)
        elif click_element('Salvar', 'text'):
            log_ok("CASO SALVO COM SUCESSO!")
            time.sleep(0.8)
        else:
            log_warn("Salve manualmente se necessário")
    else:
        log_info("Revise e salve manualmente")
    
    return True

def registrar_conta_bemol_automatico(driver):
    """Nova função para registrar casos de Conta Bemol"""
    log_info("Iniciando registro de Conta Bemol...")
    
    js_click = """
    const query = arguments[0];
    const mode = arguments[1] || 'selector';
    
    function findDeep(root, q, isText) {
        if (!isText) {
            try {
                const el = root.querySelector(q);
                if (el && isVisible(el)) return el;
            } catch(e){}
        } else {
            const tags = ['button', 'a', 'span', 'lightning-button'];
            for (const tag of tags) {
                const els = Array.from(root.querySelectorAll(tag));
                for (const el of els) {
                    if (isVisible(el)) {
                        const text = (el.innerText || '').trim();
                        if (text.toLowerCase().includes(q.toLowerCase())) return el;
                    }
                }
            }
        }
        
        const all = root.querySelectorAll('*');
        for (const el of all) {
            try {
                if (el.shadowRoot) {
                    const found = findDeep(el.shadowRoot, q, isText);
                    if (found) return found;
                }
            } catch(e){}
        }
        return null;
    }
    
    function isVisible(el) {
        try {
            const rect = el.getBoundingClientRect();
            const style = window.getComputedStyle(el);
            return rect.width > 0 && rect.height > 0 && 
                   style.display !== 'none' && style.visibility !== 'hidden';
        } catch(e) {
            return false;
        }
    }
    
    const el = findDeep(document, query, mode === 'text');
    if (!el) return { success: false };
    
    try {
        el.scrollIntoView({block: 'center'});
        el.click();
        return { success: true };
    } catch(e) {
        try {
            ['mousedown', 'click'].forEach(ev => {
                el.dispatchEvent(new MouseEvent(ev, {bubbles: true}));
            });
            return { success: true };
        } catch(e2) {
            return { success: false };
        }
    }
    """
    
    js_fill_input = """
    const selector = arguments[0];
    const value = arguments[1];
    
    function findDeep(root, sel) {
        try {
            const el = root.querySelector(sel);
            if (el && el.offsetWidth > 0) return el;
        } catch(e){}
        
        const all = root.querySelectorAll('*');
        for (const elem of all) {
            try {
                if (elem.shadowRoot) {
                    const found = findDeep(elem.shadowRoot, sel);
                    if (found) return found;
                }
            } catch(e){}
        }
        return null;
    }
    
    const input = findDeep(document, selector);
    if (!input) return { success: false, error: 'Input não encontrado' };
    
    try {
        input.focus();
        input.value = value;
        input.dispatchEvent(new Event('input', {bubbles: true}));
        input.dispatchEvent(new Event('change', {bubbles: true}));
        return { success: true };
    } catch(e) {
        return { success: false, error: String(e) };
    }
    """
    
    js_select_radio_conta_bemol = """
    const labels = Array.from(document.querySelectorAll('label, span'));
    for (const label of labels) {
        const text = (label.innerText || label.textContent || '').trim();
        if (text === 'Conta Bemol' || text.includes('Conta Bemol')) {
            const input = label.querySelector('input[type="radio"]') ||
                         document.querySelector(`input[id="${label.getAttribute('for')}"]`);
            if (input) {
                try {
                    input.checked = true;
                    input.click();
                    input.dispatchEvent(new Event('change', {bubbles: true}));
                    return { success: true };
                } catch(e) {}
            }
        }
    }
    return { success: false };
    """
    
    def click_element(query, mode='selector', tries=3):
        for i in range(tries):
            res = executar_js_safe(driver, js_click, query, mode)
            if res and res.get('success'):
                log_ok(f"Clicado: {query}")
                return True
            time.sleep(0.2)
        log_warn(f"Falha: {query}")
        return False
    
    print("\n" + "="*70)
    print("   REGISTRO CONTA BEMOL")
    print("="*70 + "\n")
    
    # 1. Abrir Casos
    log_info("1. Abrindo Casos...")
    if not click_element('a[data-tab-value="flexipage_tab3"]'):
        click_element('Casos', 'text')
    time.sleep(0.3)
    
    # 2. Criar novo caso
    log_info("2. Clicando Criar...")
    if not click_element('button[name="NewCase"]'):
        click_element('Criar', 'text')
    time.sleep(0.5)
    
    # 3. Aguardar e selecionar radio "Conta Bemol"
    log_info("3. Aguardando formulário e selecionando Conta Bemol...")
    time.sleep(1)
    
    res_radio = executar_js_safe(driver, js_select_radio_conta_bemol)
    if res_radio and res_radio.get('success'):
        log_ok("Radio 'Conta Bemol' selecionado")
    else:
        log_warn("Não conseguiu selecionar 'Conta Bemol' automaticamente")
        input("Selecione 'Conta Bemol' manualmente e pressione Enter...")
    
    time.sleep(0.2)
    
    # 4. Avançar
    log_info("4. Avançar...")
    for _ in range(4):
        if click_element('Avançar', 'text', tries=2):
            break
        time.sleep(0.05)
    time.sleep(0.8)
    
    # 5. Coletar telefone e email
    print("\n" + "="*70)
    try:
        telefone_conta = input("Digite o TELEFONE do cliente: ").strip()
        email_conta = input("Digite o EMAIL do cliente: ").strip()
    except (EOFError, KeyboardInterrupt):
        telefone_conta = ""
        email_conta = ""
    
    if not telefone_conta or not email_conta:
        log_error("Telefone e Email são obrigatórios!")
        return False
    
    # 6. Preencher Assunto (combobox 11)
    log_info("5. Selecionando Assunto (11ª opção)...")
    selecionar_combobox_melhorado(driver, 'Assunto', 11, 'Assunto')
    time.sleep(0.2)
    
    # 7. Preencher Descrição
    log_info("6. Preenchendo Descrição...")
    descricao_texto = f"Cliente em contato solicitou a atualização do seu número de telefone, o mesmo não possui acesso ao antigo.\n\nTEL: {telefone_conta}\nemail: {email_conta}\n\nTodos os dados foram confirmados pelo cliente"
    
    js_fill_textarea = """
    const text = arguments[0];
    
    function findDeep(root) {
        const textareas = Array.from(root.querySelectorAll('textarea'));
        for (const ta of textareas) {
            const rect = ta.getBoundingClientRect();
            if (rect.width > 0 && rect.height > 0) return ta;
        }
        
        const all = root.querySelectorAll('*');
        for (const el of all) {
            try {
                if (el.shadowRoot) {
                    const found = findDeep(el.shadowRoot);
                    if (found) return found;
                }
            } catch(e){}
        }
        return null;
    }
    
    const textarea = findDeep(document);
    if (!textarea) return { success: false };
    
    try {
        textarea.focus();
        textarea.value = text;
        textarea.dispatchEvent(new Event('input', {bubbles: true}));
        textarea.dispatchEvent(new Event('change', {bubbles: true}));
        return { success: true };
    } catch(e) {
        return { success: false };
    }
    """
    
    res_desc = executar_js_safe(driver, js_fill_textarea, descricao_texto)
    if res_desc and res_desc.get('success'):
        log_ok("Descrição preenchida")
    
    # 8. Sistema Operacional (3ª opção)
    log_info("7. Sistema Operacional (3ª opção)...")
    selecionar_combobox_melhorado(driver, 'Sistema Operacional', 3, 'Sistema Operacional')
    
    # 9. Origem do caso (2ª opção)
    log_info("8. Origem do caso (2ª opção)...")
    selecionar_combobox_melhorado(driver, 'Origem do caso', 2, 'Origem do caso')
    
    # 10. Motivo do contato (2ª opção)
    log_info("9. Motivo do contato (2ª opção)...")
    selecionar_combobox_melhorado(driver, 'Motivo do contato', 2, 'Motivo do contato')
    
    # 11. Categoria (5ª opção)
    log_info("10. Categoria (5ª opção)...")
    selecionar_combobox_melhorado(driver, 'Categoria', 5, 'Categoria')
    
    # 12. Subcategoria (5ª opção)
    log_info("11. Subcategoria (5ª opção)...")
    selecionar_combobox_melhorado(driver, 'Subcategoria', 5, 'Subcategoria')
    
    # 13. Verificar em (próximo dia) - CORRIGIDO
    log_info("12. Preenchendo data 'Verificar em' (próximo dia)...")
    
    # Calcular o próximo dia
    tomorrow = datetime.now() + timedelta(days=1)
    data_formatada = tomorrow.strftime("%d/%m/%Y")
    
    js_fill_date = """
    const dateStr = arguments[0];
    
    function findDeep(root, selector) {
        try {
            const el = root.querySelector(selector);
            if (el && el.offsetWidth > 0) return el;
        } catch(e){}
        
        const all = root.querySelectorAll('*');
        for (const elem of all) {
            try {
                if (elem.shadowRoot) {
                    const found = findDeep(elem.shadowRoot, selector);
                    if (found) return found;
                }
            } catch(e){}
        }
        return null;
    }
    
    const input = findDeep(document, 'input[name="CheckIn__c"]');
    if (!input) return { success: false, error: 'Input não encontrado' };
    
    try {
        input.focus();
        input.value = '';
        input.value = dateStr;
        input.dispatchEvent(new Event('input', {bubbles: true}));
        input.dispatchEvent(new Event('change', {bubbles: true}));
        input.dispatchEvent(new Event('blur', {bubbles: true}));
        return { success: true, date: dateStr };
    } catch(e) {
        return { success: false, error: String(e) };
    }
    """
    
    res_date = executar_js_safe(driver, js_fill_date, data_formatada)
    
    if res_date and res_date.get('success'):
        log_ok(f"Data preenchida: {data_formatada}")
    else:
        log_warn(f"Não conseguiu preencher automaticamente")
        print(f"Digite a data manualmente no formato DD/MM/YYYY: {data_formatada}")
        input("Pressione Enter após preencher a data...")
        
    # 14. Desmarcar checkbox "Enviar email de notificação para contato"
    log_info("13. Desmarcando notificação por email...")
    
    js_desmarcar_checkbox_lightning = """
    function findAndUncheckLightningCheckbox() {
        // Função para buscar em Shadow DOM recursivamente
        function findInShadowDOM(root, maxDepth = 10, currentDepth = 0) {
            if (currentDepth > maxDepth) return null;
            
            // Buscar lightning-input com o título correto
            const lightningInputs = Array.from(root.querySelectorAll('lightning-input'));
            for (const input of lightningInputs) {
                const title = input.getAttribute('title') || '';
                const dataName = input.getAttribute('data-option-name') || '';
                
                if (title.includes('Enviar email de notificação') || 
                    dataName === 'triggerOtherEmail') {
                    return input;
                }
            }
            
            // Buscar recursivamente em todos os shadow roots
            const allElements = root.querySelectorAll('*');
            for (const el of allElements) {
                try {
                    if (el.shadowRoot) {
                        const found = findInShadowDOM(el.shadowRoot, maxDepth, currentDepth + 1);
                        if (found) return found;
                    }
                } catch(e) {}
            }
            
            return null;
        }
        
        // Buscar o componente lightning-input
        let lightningInput = findInShadowDOM(document);
        
        if (!lightningInput) {
            // Fallback: buscar diretamente por atributos conhecidos
            const allInputs = document.querySelectorAll('lightning-input');
            for (const input of allInputs) {
                const title = input.getAttribute('title') || '';
                if (title.includes('Enviar email de notificação')) {
                    lightningInput = input;
                    break;
                }
            }
        }
        
        if (!lightningInput) {
            return { success: false, error: 'Lightning-input não encontrado' };
        }
        
        // Verificar estado INICIAL
        const isCheckedBefore = lightningInput.hasAttribute('checked');
        
        if (!isCheckedBefore) {
            return { success: true, action: 'already_unchecked', wasChecked: false };
        }
        
        // Encontrar o input real dentro do shadow DOM
        let realCheckbox = null;
        
        try {
            // Navegar pelo shadow DOM do lightning-input
            if (lightningInput.shadowRoot) {
                const primitiveCheckbox = lightningInput.shadowRoot.querySelector('lightning-primitive-input-checkbox');
                if (primitiveCheckbox && primitiveCheckbox.shadowRoot) {
                    realCheckbox = primitiveCheckbox.shadowRoot.querySelector('input[type="checkbox"]');
                }
            }
        } catch(e) {}
        
        // Se não achou no shadow, buscar por ID baseado no padrão
        if (!realCheckbox) {
            // Tentar encontrar pelo name attribute
            const name = lightningInput.getAttribute('data-option-name');
            if (name) {
                realCheckbox = document.querySelector(`input[name="${name}"]`);
            }
        }
        
        // Estratégia de desmarcação
        try {
            // Rolar até o elemento
            lightningInput.scrollIntoView({block: 'center', behavior: 'instant'});
            
            let w = 0;
            while(w < 50) {
                const s = Date.now();
                while(Date.now() - s < 5) {}
                w += 5;
            }
            
            // Método 1: Clicar no checkbox real se encontrou
            if (realCheckbox) {
                realCheckbox.focus();
                realCheckbox.click();
                
                w = 0;
                while(w < 100) {
                    const s = Date.now();
                    while(Date.now() - s < 5) {}
                    w += 5;
                }
                
                // Verificar se funcionou IMEDIATAMENTE após clicar
                const isStillChecked = lightningInput.hasAttribute('checked');
                if (!isStillChecked) {
                    return { success: true, action: 'unchecked_via_real_checkbox', wasChecked: true, nowChecked: false };
                }
            }
            
            // Método 2: Clicar no próprio lightning-input
            lightningInput.click();
            
            w = 0;
            while(w < 100) {
                const s = Date.now();
                while(Date.now() - s < 5) {}
                w += 5;
            }
            
            // Verificar se funcionou após clicar no lightning-input
            let isStillChecked = lightningInput.hasAttribute('checked');
            if (!isStillChecked) {
                return { success: true, action: 'unchecked_via_lightning_click', wasChecked: true, nowChecked: false };
            }
            
            // Método 3: Remover o atributo checked diretamente E forçar desmarcação
            lightningInput.removeAttribute('checked');
            
            if (realCheckbox) {
                realCheckbox.checked = false;
            }
            
            // Disparar eventos DEPOIS de desmarcar
            lightningInput.dispatchEvent(new Event('change', {bubbles: true}));
            lightningInput.dispatchEvent(new Event('input', {bubbles: true}));
            
            if (realCheckbox) {
                realCheckbox.dispatchEvent(new Event('change', {bubbles: true}));
            }
            
            w = 0;
            while(w < 150) {
                const s = Date.now();
                while(Date.now() - s < 5) {}
                w += 5;
            }
            
            // Verificação final DEPOIS de aplicar todos os métodos
            isStillChecked = lightningInput.hasAttribute('checked');
            
            return { 
                success: true, 
                action: 'forced_uncheck',
                wasChecked: true,
                nowChecked: isStillChecked,
                effectivelyUnchecked: !isStillChecked
            };
            
        } catch(e) {
            return { success: false, error: String(e) };
        }
    }
    
    return findAndUncheckLightningCheckbox();
    """
    
    try:
        resultado = executar_js_safe(driver, js_desmarcar_checkbox_lightning)
        
        if resultado:
            if resultado.get('success'):
                action = resultado.get('action', 'unknown')
                
                if action == 'already_unchecked':
                    log_ok("✓ Checkbox já estava desmarcada")
                elif action == 'unchecked_via_real_checkbox':
                    log_ok("✓ Checkbox desmarcada através do input real!")
                elif action == 'unchecked_via_lightning_click':
                    log_ok("✓ Checkbox desmarcada através do lightning-input!")
                elif action == 'unchecked_via_remove_attribute':
                    log_ok("✓ Checkbox desmarcada removendo atributo!")
                    
                    log_info("Aplicando método super forçado para Lightning...")
                    driver.execute_script("""
                        // Buscar TODOS os lightning-input
                        const allLightningInputs = document.querySelectorAll('lightning-input');
                        
                        for (const input of allLightningInputs) {
                            const title = input.getAttribute('title') || '';
                            const dataName = input.getAttribute('data-option-name') || '';
                            
                            if (title.includes('email') || dataName === 'triggerOtherEmail') {
                                // Remover atributo checked
                                input.removeAttribute('checked');
                                
                                // Tentar acessar o shadowRoot e desmarcar o input real
                                try {
                                    if (input.shadowRoot) {
                                        const primitive = input.shadowRoot.querySelector('lightning-primitive-input-checkbox');
                                        if (primitive && primitive.shadowRoot) {
                                            const realInput = primitive.shadowRoot.querySelector('input[type="checkbox"]');
                                            if (realInput) {
                                                realInput.checked = false;
                                                realInput.dispatchEvent(new Event('change', {bubbles: true}));
                                            }
                                        }
                                    }
                                } catch(e) {}
                                
                                // Disparar eventos no lightning-input
                                input.dispatchEvent(new Event('change', {bubbles: true}));
                                input.dispatchEvent(new Event('input', {bubbles: true}));
                                
                                break;
                            }
                        }
                    """)
                    time.sleep(0.3)
                    log_ok("Método super forçado aplicado")
            else:
                erro = resultado.get('error', 'desconhecido')
                log_warn(f"Erro ao buscar checkbox: {erro}")
                
                log_info("Tentando com Selenium direto...")
                try:
                    lightning_inputs = driver.find_elements(By.TAG_NAME, 'lightning-input')
                    
                    for li in lightning_inputs:
                        try:
                            title = li.get_attribute('title')
                            if title and 'Enviar email de notificação' in title:
                                log_ok(f"Encontrado via Selenium: {title}")
                                
                                # Rolar até elemento
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", li)
                                time.sleep(0.2)
                                
                                # Verificar se está marcado
                                is_checked = li.get_attribute('checked') is not None
                                
                                if is_checked:
                                    # Tentar clicar
                                    try:
                                        li.click()
                                        time.sleep(0.2)
                                        log_ok("Clicado via Selenium")
                                    except:
                                        # Clicar via JavaScript
                                        driver.execute_script("arguments[0].click();", li)
                                        time.sleep(0.2)
                                        log_ok("Clicado via JavaScript")
                                    
                                    # Forçar remoção do atributo
                                    driver.execute_script("arguments[0].removeAttribute('checked');", li)
                                    log_ok("Atributo removido via JavaScript")
                                else:
                                    log_ok("Já estava desmarcado")
                                
                                break
                        except:
                            continue
                    else:
                        log_warn("Não encontrou lightning-input com Selenium")
                        
                except Exception as e_sel:
                    log_error(f"Erro com Selenium: {str(e_sel)[:100]}")
        else:
            log_warn("JavaScript não retornou resposta")
    
    except Exception as e:
        log_error(f"Erro ao processar checkbox: {str(e)[:100]}")
        
    # Verificação final robusta
    log_info("Verificando estado final...")
    
    verificacao = executar_js_safe(driver, """
    const lightningInputs = document.querySelectorAll('lightning-input');
    for (const input of lightningInputs) {
        const title = input.getAttribute('title') || '';
        const dataName = input.getAttribute('data-option-name') || '';
        
        if (title.includes('Enviar email de notificação') || dataName === 'triggerOtherEmail') {
            const hasChecked = input.hasAttribute('checked');
            return { 
                found: true, 
                checked: hasChecked,
                title: title 
            };
        }
    }
    return { found: false };
    """)
    
    print("\n" + "="*70)
    log_ok("FORMULÁRIO COMPLETO!")
    print("="*70)
    print(f"\nTelefone: {telefone_conta}")
    print(f"Email: {email_conta}")
    print(f"Assunto: 11ª opção")
    print(f"Sistema Operacional: 3ª opção")
    print(f"Origem: 2ª opção")
    print(f"Motivo: 2ª opção")
    print(f"Categoria: 5ª opção")
    print(f"Subcategoria: 5ª opção")
    print(f"Data Verificar em: {data_formatada}")
    print("="*70 + "\n")
    
    salvar = input("SALVAR CASO? (s/n): ").strip().lower()
    
    if salvar == 's' or salvar == '':
        log_info("Salvando...")
        
        if click_element('button[name="SaveEdit"]'):
            log_ok("CASO SALVO COM SUCESSO!")
            time.sleep(0.8)
        elif click_element('Salvar', 'text'):
            log_ok("CASO SALVO COM SUCESSO!")
            time.sleep(0.8)
        else:
            log_warn("Salve manualmente se necessário")
    else:
        log_info("Revise e salve manualmente")
    
    return True

def menu_principal():
    if HAS_QUESTIONARY:
        return questionary.select(
            "Escolha uma ação:",
            choices=[
                "Registrar informação", 
                "Registrar Conta Bemol", 
                "Buscar outro CPF",
                "Sair"
            ]
        ).ask()
    else:
        print("\n" + "="*30)
        print("      AUTOMAÇÃO SALESFORCE")
        print("="*30)
        print("1) Registrar informação")
        print("2) Registrar Conta Bemol")
        print("3) Buscar outro CPF")
        print("4) Sair")
        print("="*30)
        escolha = input("Escolha (1/2/3/4): ").strip()
        if escolha == "1":
            return "Registrar informação"
        elif escolha == "2":
            return "Registrar Conta Bemol"
        elif escolha == "3":
            return "Buscar outro CPF"
        else:
            return "Sair"

def buscar_novo_cpf(driver):
    """Função para buscar um novo CPF sem sair do sistema"""
    max_tentativas_cpf = 5
    tentativa_cpf = 0
    cpf_encontrado = False
    
    while tentativa_cpf < max_tentativas_cpf and not cpf_encontrado:
        tentativa_cpf += 1
        
        print("\n" + "="*40)
        print(f"   BUSCA DE CLIENTE (Tentativa {tentativa_cpf}/{max_tentativas_cpf})")
        print("="*40 + "\n")
        
        cpf_raw = input("Digite o CPF (ou 'voltar' para cancelar): ").strip()
        
        # Permitir cancelar a busca
        if cpf_raw.lower() == 'voltar':
            log_info("Busca cancelada pelo usuário")
            return False
        
        cpf = limpar_cpf(cpf_raw)
        
        if not validar_cpf(cpf):
            log_error("CPF inválido localmente (formato incorreto)")
            
            tentar_novamente = input("\nDeseja tentar outro CPF? (s/n): ").strip().lower()
            if tentar_novamente != 's':
                log_info("Operação cancelada pelo usuário")
                return False
            continue
        
        log_ok(f"CPF validado: {cpf[:3]}.***.***-{cpf[-2:]}")
        
        log_info("\nBUSCA DE CLIENTE")
        
        # Primeiro, tentar navegar para a página inicial
        log_info("Navegando para a página inicial...")
        try:
            current_url = driver.current_url
            base_url = current_url.split('/lightning/')[0] if '/lightning/' in current_url else current_url.split('.com')[0] + '.com'
            driver.get(base_url + '/lightning/page/home')
            time.sleep(2)
        except Exception as e:
            log_warn(f"Não conseguiu navegar para início: {str(e)[:60]}")
        
        if not verificar_pagina_inicial(driver):
            log_warn("Não está na página Início. Continuando mesmo assim...")
        
        resultado_busca = buscar_cpf_automatico(driver, cpf, max_tentativas=3)
        
        if resultado_busca == 'invalid':
            log_error("\n❌ CPF INVÁLIDO no Salesforce!")
            log_info("O Salesforce rejeitou este CPF como inválido.")
            
            tentar_novamente = input("\nDeseja tentar outro CPF? (s/n): ").strip().lower()
            if tentar_novamente != 's':
                log_info("Operação cancelada pelo usuário")
                return False
            continue
            
        elif resultado_busca == 'not_found':
            log_warn("\n⚠️ CLIENTE NÃO ENCONTRADO no Salesforce!")
            log_info("Este CPF não está cadastrado no sistema.")
            
            tentar_novamente = input("\nDeseja tentar outro CPF? (s/n): ").strip().lower()
            if tentar_novamente != 's':
                log_info("Operação cancelada pelo usuário")
                return False
            continue
            
        elif resultado_busca == True:
            log_ok("\n✓ Cliente encontrado com sucesso!")
            cpf_encontrado = True
            return True
            
        else:
            log_error("Falha na busca automática")
            
            retry = input("\nTentar buscar este CPF novamente? (s/n): ").strip().lower()
            if retry == 's':
                resultado_retry = buscar_cpf_automatico(driver, cpf, max_tentativas=2)
                
                if resultado_retry == True:
                    log_ok("Busca bem-sucedida!")
                    cpf_encontrado = True
                    return True
                elif resultado_retry == 'invalid':
                    log_error("\n❌ CPF INVÁLIDO no Salesforce!")
                    tentar_novamente = input("\nDeseja tentar outro CPF? (s/n): ").strip().lower()
                    if tentar_novamente != 's':
                        return False
                    continue
                elif resultado_retry == 'not_found':
                    log_warn("\n⚠️ CLIENTE NÃO ENCONTRADO!")
                    tentar_novamente = input("\nDeseja tentar outro CPF? (s/n): ").strip().lower()
                    if tentar_novamente != 's':
                        return False
                    continue
                else:
                    log_warn("Ainda não conseguiu.")
                    continuar = input("\nBuscar manualmente? (s/n): ").strip().lower()
                    if continuar == 's':
                        input("Busque manualmente e pressione Enter quando estiver na página do cliente...")
                        cpf_encontrado = True
                        return True
                    else:
                        tentar_outro = input("\nDeseja tentar outro CPF? (s/n): ").strip().lower()
                        if tentar_outro != 's':
                            log_info("Operação cancelada")
                            return False
                        continue
            else:
                continuar = input("\nBuscar manualmente? (s/n): ").strip().lower()
                if continuar == 's':
                    input("Busque manualmente e pressione Enter quando estiver na página do cliente...")
                    cpf_encontrado = True
                    return True
                else:
                    tentar_outro = input("\nDeseja tentar outro CPF? (s/n): ").strip().lower()
                    if tentar_outro != 's':
                        log_info("Operação cancelada")
                        return False
                    continue
    
    if not cpf_encontrado:
        log_error(f"\nNúmero máximo de tentativas ({max_tentativas_cpf}) atingido.")
        return False
    
    return cpf_encontrado

# Parte modificada da função main():
def main():
    print("\n" + "="*40)
    print("   AUTOMAÇÃO SALESFORCE")
    print("="*40 + "\n")
    
    USUARIO = 'saymoncruz@bemol.com.br'
    SENHA = 'saymonGG00!'
    
    try:
        log_info("\nIniciando navegador Edge...")
        driver = criar_driver()
        
        log_info("\nRealizando login no Salesforce...")
        if logar_salesforce_robusto(driver, USUARIO, SENHA):
            if verificar_login_salesforce(driver):
                log_ok("Pronto para automação!")
            else:
                log_error("Verificação de login falhou")
                return
        else:
            log_error("Login falhou após todas as tentativas")
            return
        log_ok(f"Login realizado com sucesso")
        
        # Busca inicial do CPF
        log_info("\n>>> BUSCA INICIAL DE CLIENTE <<<")
        if not buscar_novo_cpf(driver):
            log_info("Nenhum cliente carregado. Encerrando...")
            return
        
        # Loop principal do menu
        while True:
            escolha = menu_principal()
            
            if escolha == "Registrar informação":
                print("\n" + "="*70)
                print("   INICIANDO REGISTRO DE INFORMAÇÃO")
                print("="*70 + "\n")
                
                if registrar_informacao_automatico(driver):
                    log_ok("\nProcesso concluído com sucesso!")
                else:
                    log_warn("\nProcesso foi concluído, porém retornou algum erro.")
                
                continuar = input("\nDeseja registrar outro caso? (s/n): ").strip().lower()
                if continuar != 's':
                    continue
                else:
                    # Antes de criar outro caso, voltar para a página do cliente
                    log_info("\nPreparando para criar novo caso...")
                    if not voltar_para_cliente(driver):
                        log_warn("Não conseguiu voltar automaticamente. Navegue manualmente para a aba do cliente.")
                        input("Pressione Enter quando estiver na aba do cliente...")
                    else:
                        log_ok("Pronto para criar novo caso!")
                    time.sleep(0.5)
                    
            elif escolha == "Registrar Conta Bemol":
                escolha_conta = questionary.select(
                    "Escolha uma opção abaixo: ",
                    choices=[
                        {"name": "Atualização de número de telefone", "value": "atualizacao_telefone"},
                        {"name": "Voltar", "value": "voltar"},
                    ]
                ).ask()
                
                if escolha_conta == "atualizacao_telefone":
                    print("\n" + "="*70)
                    print("   INICIANDO REGISTRO DE CONTA BEMOL")
                    print("="*70 + "\n")
                    
                    if registrar_conta_bemol_automatico(driver):
                        log_ok("\nProcesso de Conta Bemol concluído com sucesso!")
                    else:
                        log_warn("\nProcesso de Conta Bemol foi concluído, porém retornou algum erro.")
                    
                    continuar = input("\nDeseja registrar outra Conta Bemol? (s/n): ").strip().lower()
                    if continuar != 's':
                        continue
                    else:
                        # Antes de criar outro caso, voltar para a página do cliente
                        log_info("\nPreparando para criar nova Conta Bemol...")
                        if not voltar_para_cliente(driver):
                            log_warn("Não conseguiu voltar automaticamente. Navegue manualmente para a aba do cliente.")
                            input("Pressione Enter quando estiver na aba do cliente...")
                        else:
                            log_ok("Pronto para criar nova Conta Bemol!")
                        time.sleep(0.5)
                else:
                    continue
                    
            elif escolha == "Buscar outro CPF":
                print("\n" + "="*70)
                print("   BUSCAR NOVO CLIENTE")
                print("="*70 + "\n")
                
                if buscar_novo_cpf(driver):
                    log_ok("\n✓ Novo cliente carregado com sucesso!")
                    log_info("Retornando ao menu principal...")
                else:
                    log_warn("\nBusca cancelada ou sem sucesso.")
                    log_info("Retornando ao menu principal...")
                
            else:  # Sair
                log_info("Encerrando automação...")
                break
        
        log_info("\nAutomação finalizada.")
        input("Pressione Enter para fechar o navegador e encerrar o script...")
        
    except KeyboardInterrupt:
        log_warn("\n\nOperação interrompida pelo usuário (Ctrl+C)")
    except Exception as e:
        log_error(f"\nErro durante execução: {e}")
        import traceback
        print("\n" + "="*70)
        print("   DETALHES DO ERRO:")
        print("="*70)
        traceback.print_exc()
        print("="*70)
    finally:
        log_info("\nLimpando recursos...")
        cleanup_all_resources()
        log_ok("Limpeza concluída!")

if __name__ == "__main__":
    main()