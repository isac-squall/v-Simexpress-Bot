# -*- coding: utf-8 -*-
import os
import time
from pathlib import Path
import shutil
import glob
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import argparse

load_dotenv()

USUARIO = os.getenv("SIMEXPRESS_USUARIO")
SENHA = os.getenv("SIMEXPRESS_SENHA")
DOWNLOAD_PATH = os.getenv("DOWNLOAD_PATH", str(Path.cwd() / "downloads"))
EXCEL_PATH = os.getenv("EXCEL_PATH", str(Path(__file__).parent / "pedidos.xlsx"))


def _pedidos_do_env():
    _pedidos_raw = os.getenv("PEDIDOS_LOTE", "123456\\n234567\\n345678")
    if "\\n" in _pedidos_raw and "\n" not in _pedidos_raw:
        return '\n'.join([p.strip() for p in _pedidos_raw.split('\\n') if p.strip()])
    return '\n'.join([p.strip() for p in _pedidos_raw.splitlines() if p.strip()])


def _carregar_pedidos_do_arquivo(caminho_arquivo):
    try:
        if caminho_arquivo.lower().endswith('.csv'):
            df = pd.read_csv(caminho_arquivo)
        else:
            df = pd.read_excel(caminho_arquivo)
        # Colunas de fallback padrão
        colunas_candidatas = [
            "Pedido", "pedido", "Pedidos", "pedidos",
            "Order", "order", "Número do Pedido", "numero_pedido"
        ]
        coluna_escolhida = None
        for c in colunas_candidatas:
            if c in df.columns:
                coluna_escolhida = c
                break

        if coluna_escolhida is None:
            if len(df.columns) == 0:
                raise ValueError("A planilha Excel não possui colunas")
            coluna_escolhida = df.columns[0]

        pedidos = df[coluna_escolhida].dropna().astype(str).str.strip()
        pedidos = [p for p in pedidos if p]
        if not pedidos:
            raise ValueError(f"Nenhum pedido válido encontrado na coluna '{coluna_escolhida}'")

        return '\n'.join(pedidos)
    except FileNotFoundError:
        print(f"Arquivo não encontrado: {caminho_arquivo}. Usando PEDIDOS_LOTE do .env como fallback")
        return _pedidos_do_env()
    except Exception as e:
        print(f"Erro ao ler arquivo {caminho_arquivo}: {e}. Usando PEDIDOS_LOTE do .env como fallback")
        return _pedidos_do_env()


PEDIDOS_LOTE = _pedidos_do_env()

XPATH_USUARIO = os.getenv("SIMEXPRESS_XPATH_USUARIO", "")
XPATH_SENHA = os.getenv("SIMEXPRESS_XPATH_SENHA", "")
XPATH_ENTRAR = os.getenv("SIMEXPRESS_XPATH_ENTRAR", "")

if not USUARIO or not SENHA:
    raise RuntimeError("Defina SIMEXPRESS_USUARIO e SIMEXPRESS_SENHA no .env")

URL = "https://simexpress.com.br/"


def main():
    parser = argparse.ArgumentParser(description="Bot para automação Simexpress")
    parser.add_argument('--pedidos', type=str, help='Caminho para arquivo Excel/CSV com pedidos (padrão: pedidos.xlsx)')
    args = parser.parse_args()

    Path(DOWNLOAD_PATH).mkdir(parents=True, exist_ok=True)

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    prefs = {
        "download.default_directory": DOWNLOAD_PATH,
        "download.prompt_for_download": False,
        "profile.default_content_setting_values.automatic_downloads": 1,
    }
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 40)

    def log(msg):
        print(msg)
        with open(Path(DOWNLOAD_PATH) / "simexpress_log.txt", "a", encoding="utf-8") as f:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {msg}\n")

    # Escolhe pedidos em ordem: CLI > .env (PEDIDOS_LOTE) > EXCEL_PATH
    if args.pedidos:
        pedidos_path = args.pedidos
        if os.path.exists(pedidos_path):
            log(f"Pedidos carregados do arquivo: {pedidos_path}")
            pedidos_text = _carregar_pedidos_do_arquivo(pedidos_path)
        else:
            log(f"Arquivo de pedidos não encontrado: {pedidos_path}")
            log("Pedidos carregados do .env (PEDIDOS_LOTE).")
            pedidos_text = _pedidos_do_env()
    elif os.getenv("PEDIDOS_LOTE"):
        log("Pedidos carregados do .env (PEDIDOS_LOTE).")
        pedidos_text = _pedidos_do_env()
    elif EXCEL_PATH and os.path.exists(EXCEL_PATH):
        log(f"Pedidos carregados do arquivo padrão: {EXCEL_PATH}")
        pedidos_text = _carregar_pedidos_do_arquivo(EXCEL_PATH)
    else:
        log("Nenhum arquivo de pedidos encontrado; usando .env (PEDIDOS_LOTE).")
        pedidos_text = _pedidos_do_env()

    # Valor final para envio ao portal
    global PEDIDOS_LOTE
    PEDIDOS_LOTE = pedidos_text

    def find_first(xpath_list):
        for xp in xpath_list:
            try:
                # Use timeout breve por candidate para não bloquear muito tempo em cada fallback
                el = WebDriverWait(driver, 6).until(EC.element_to_be_clickable((By.XPATH, xp)))
                log(f"Elemento encontrado com XPath: {xp}")
                return el, xp
            except Exception:
                log(f"Não encontrado XPath: {xp}")
                continue
        return None, None

    try:
        driver.get(URL)
        log("[1/8] P�gina inicial aberta")

        try:
            btn_acesso = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Acesso ao Sistema') or contains(., 'Acesso ao Sistema') or contains(@href, 'login') or contains(@class,'acesso')]")))
            btn_acesso.click()
            log("[2/8] Clique em Acesso ao Sistema")
        except Exception as e:
            log("[2/8] Acesso ao Sistema n�o clicado (pode j� estar login): " + str(e))

        if XPATH_USUARIO:
            xpath_usuario = XPATH_USUARIO
            log(f"Usando XPATH_USUARIO personalizado: {xpath_usuario}")
        else:
            campo_usuario_el, xpath_usuario_found = find_first([
                "//input[@name='login' or @id='login' or @name='username' or @id='username' or @name='user' or @id='user']",
                "//input[contains(@placeholder,'Usuário') or contains(@placeholder,'user') or contains(@aria-label,'Usuário') or contains(@aria-label,'user')]",
                "//input[@type='text' or @type='email']",
            ])
            if not campo_usuario_el:
                raise RuntimeError("Campo de usuário não encontrado")
            xpath_usuario = xpath_usuario_found

        if XPATH_SENHA:
            xpath_senha = XPATH_SENHA
            log(f"Usando XPATH_SENHA personalizado: {xpath_senha}")
        else:
            campo_senha_el, xpath_senha_found = find_first([
                "//input[@type='password' or @name='senha' or @id='senha']",
                "//input[contains(@placeholder,'Senha') or contains(@aria-label,'Senha') or contains(@class,'password')]",
            ])
            if not campo_senha_el:
                raise RuntimeError("Campo de senha não encontrado")
            xpath_senha = xpath_senha_found

        if XPATH_ENTRAR:
            xpath_entrar = XPATH_ENTRAR
            log(f"Usando XPATH_ENTRAR personalizado: {xpath_entrar}")
        else:
            botao_entrar_el, xpath_entrar_found = find_first([
                "//button[contains(translate(., 'ENTRAR', 'entrar'), 'entrar') or contains(.,'Entrar') or contains(.,'Login') or @type='submit']",
                "//input[@type='submit' and (contains(@value,'Entrar') or contains(@value,'Login'))]",
                "//button[.//svg or .//i or .//path][contains(@id,'login') or contains(@class,'login') or contains(@aria-label,'Entrar') or contains(@title,'Entrar') or contains(@name,'entrar') or contains(@name,'login') or @type='button']",
                "//form//button[.//svg or .//i or .//path or contains(translate(., 'ENTRAR', 'entrar'), 'entrar')]",
                "//form//a[.//svg or .//i or .//path or contains(translate(., 'ENTRAR', 'entrar'), 'entrar')]",
                "//div[contains(@class, 'login') or contains(@class, 'enter')]//button[1]",
            ])
            if not botao_entrar_el:
                log("Fallback: tentando localizar botão 'Entrar' via JavaScript")
                try:
                    el = driver.execute_script("""
                        var texts = ['entrar','login','acessar'];
                        var nodes = Array.from(document.querySelectorAll('button,a,input,div,span'));
                        for (var i=0;i<nodes.length;i++){
                            var n = nodes[i];
                            var txt = (n.innerText || n.value || '').toString().toLowerCase().trim();
                            if(!txt) continue;
                            for(var j=0;j<texts.length;j++){
                                if(txt.indexOf(texts[j]) !== -1){
                                    n.scrollIntoView();
                                    return n;
                                }
                            }
                        }
                        return null;
                    """)
                    if el:
                        botao_entrar_el = el
                        xpath_entrar_found = None
                        try:
                            tag = botao_entrar_el.tag_name
                            text = (botao_entrar_el.text or '').strip()
                        except Exception:
                            tag = 'unknown'
                            text = ''
                        log(f"Botão encontrado via JS fallback: <{tag}> texto='{text[:80]}'")
                    else:
                        raise RuntimeError("Botão Entrar não encontrado (fallback JS retornou nulo)")
                except Exception as e:
                    raise RuntimeError(f"Botão Entrar não encontrado: {e}")
            xpath_entrar = xpath_entrar_found

        log("[3/8] Seletores de login definidos")

        campo_usuario = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_usuario)))
        campo_senha = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_senha)))
        if botao_entrar_el is not None:
            botao_entrar = botao_entrar_el
        else:
            botao_entrar = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_entrar)))

        campo_usuario.clear()
        campo_usuario.send_keys(USUARIO)
        campo_senha.clear()
        campo_senha.send_keys(SENHA)
        try:
            log("Tentando clicar no botão Entrar (click padrão)")
            botao_entrar.click()
        except Exception as e:
            log(f"Click padrão falhou: {e} — tentando click via JS")
            try:
                driver.execute_script("arguments[0].click();", botao_entrar)
                log("Click via JS executado com sucesso")
            except Exception as e2:
                # salva página e HTML para diagnóstico
                html_path = Path(DOWNLOAD_PATH) / "simexpress_erro_page.html"
                try:
                    with open(html_path, 'w', encoding='utf-8') as hf:
                        hf.write(driver.page_source)
                    log(f"Salvo page_source em: {html_path}")
                except Exception:
                    log("Falha ao salvar page_source")
                raise

        log("[4/8] Login enviado")

        wait.until(EC.url_changes(URL))
        log("[5/8] Login confirmado - URL mudou")
        time.sleep(2)

        # agora clique em Pedidos na tela pós-login
        menu_pedidos = wait.until(EC.element_to_be_clickable((By.XPATH,
            "//a[contains(normalize-space(.),'Pedidos') or contains(translate(normalize-space(.),'PEDIDOS','pedidos'),'pedidos') or contains(@href,'pedidos') or contains(@class,'pedido')]"
        )))
        menu_pedidos.click()
        log("[6/8] Clique em Pedidos")
        time.sleep(2)

        # o clique no menu pai pode apenas abrir submenu; aguardar o link 'Em Lote' aparecer
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(normalize-space(.),'Em Lote') or contains(@href,'pedidos.php') or contains(translate(normalize-space(.),'EM LOTE','em lote'),'em lote')]")))
        except Exception:
            time.sleep(1)

        sub_em_lote = wait.until(EC.element_to_be_clickable((By.XPATH,
            "//a[contains(normalize-space(.),'Em Lote') or contains(translate(normalize-space(.),'EM LOTE','em lote'),'em lote') or contains(@href,'lote') or contains(@class,'lote')]"
        )))
        sub_em_lote.click()
        log("[7/8] Clique em Em Lote")
        time.sleep(2)

        textarea_pedidos = wait.until(EC.visibility_of_element_located((By.XPATH, "//textarea[@id='pedidoLote' or contains(@placeholder,'cada item') or @name='pedidos'] | //textarea")))
        textarea_pedidos.clear()
        textarea_pedidos.send_keys(PEDIDOS_LOTE)
        log(f"Textarea preenchido com: {PEDIDOS_LOTE}")
        time.sleep(2)

        # Clicar no botão Consultar
        botao_consultar = wait.until(EC.element_to_be_clickable((By.ID, "btnFiltrar")))
        # Usar JavaScript click para evitar interceptação
        driver.execute_script("arguments[0].click();", botao_consultar)
        log("Botão Consultar clicado via JavaScript")
        time.sleep(5)  # Aguardar os resultados carregarem

        # Screenshot antes de clicar no botão
        driver.save_screenshot(str(Path(DOWNLOAD_PATH) / "antes_csv.png"))
        log("Screenshot salvo antes de clicar no botão CSV")

        # Salvar page_source antes de clicar
        with open(Path(DOWNLOAD_PATH) / "antes_csv.html", 'w', encoding='utf-8') as f:
            f.write(driver.page_source)
        log("Page source salvo antes de clicar no botão CSV")

        botao_csv = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(translate(., 'CSV', 'csv'), 'csv') or contains(., 'Download') or contains(., 'Excel')] | //a[contains(translate(., 'CSV', 'csv'), 'csv') or contains(., 'Download')]")))
        botao_csv.click()

        # aguardar o arquivo CSV aparecer na pasta de downloads e copiar para workspace/downloads.csv
        try:
            csv_found = None
            for _ in range(30):
                csv_files = list(Path(DOWNLOAD_PATH).glob('*.csv'))
                if csv_files:
                    # escolher o mais recente
                    csv_found = max(csv_files, key=lambda p: p.stat().st_mtime)
                    break
                time.sleep(1)

            dest_dir = Path(__file__).parent / 'downloads.csv'
            dest_dir.mkdir(parents=True, exist_ok=True)

            if csv_found:
                dest_path = dest_dir / csv_found.name
                shutil.copy(str(csv_found), str(dest_path))
                log(f"CSV copiado para: {dest_path}")
            else:
                log("Aviso: nenhum arquivo CSV encontrado no diretório de downloads dentro do tempo esperado")
        except Exception as e:
            log(f"Erro ao copiar CSV para pasta workspace: {e}")

        log("[8/8] CSV acionado e processo conclu�do")
        time.sleep(10)  # Manter navegador aberto por 10 segundos para visualização

    except Exception as ex:
        driver.save_screenshot(str(Path(DOWNLOAD_PATH) / "simexpress_erro.png"))
        try:
            with open(Path(DOWNLOAD_PATH) / "simexpress_erro_page.html", 'w', encoding='utf-8') as hf:
                hf.write(driver.page_source)
            log(f"Salvo page_source em: {Path(DOWNLOAD_PATH) / 'simexpress_erro_page.html'}")
        except Exception:
            log("Falha ao salvar page_source de erro")
        log(f"Erro durante automa��o: {ex}")
        raise

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
