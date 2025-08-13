# 1-inicia_exe.py
import pyautogui
import time
import os
import cv2
import numpy as np
import shutil
import glob
import psutil
import subprocess
from logger_config import setup_logger

# Inicializa o logger
logger = setup_logger()

# Configurações iniciais
PROGRAM_PATH = r"C:\Program Files\Intuit\QuickBooks Enterprise Solutions 24.0\QBWEnterprise.exe"
IMAGES_PATH = "img"
PASSWORD = ""
CHECK_INTERVAL = 5  

# Configurações de pastas para cópia do arquivo .qbw
PASTA_ORIGEM_QBW = r"\\192.168.1.250\Programas\Quick Books\Robo Quickboks"
PASTA_DESTINO_QBW = r"C:\Users\samuca2\Documents\base_robo_quickboks"

def copiar_ultimo_arquivo_qbw():
    """
    Fecha processos QuickBooks, limpa a pasta de destino e copia o arquivo .qbw mais recente
    """
    try:
        logger.info("Iniciando processo de cópia do último arquivo .qbw...")
        
        # ETAPA 0: FECHAR PROCESSOS DO QUICKBOOKS
        logger.info("=== FECHANDO PROCESSOS DO QUICKBOOKS ===")
        fechar_processos_quickbooks()
        
        # PRIMEIRA ETAPA: LIMPAR A PASTA DE DESTINO COM FORÇA
        logger.info("=== LIMPANDO PASTA DE DESTINO ===")
        
        # Verificar se a pasta de destino existe
        if os.path.exists(PASTA_DESTINO_QBW):
            logger.info(f"Pasta de destino encontrada: {PASTA_DESTINO_QBW}")
            
            # Listar todos os arquivos na pasta antes de limpar
            arquivos_existentes = os.listdir(PASTA_DESTINO_QBW)
            if arquivos_existentes:
                logger.info(f"Arquivos encontrados na pasta: {arquivos_existentes}")
                
                # Remover todos os arquivos e subpastas da pasta de destino
                for item in arquivos_existentes:
                    caminho_item = os.path.join(PASTA_DESTINO_QBW, item)
                    try:
                        if os.path.isfile(caminho_item):
                            # Tentar múltiplos métodos para remover o arquivo
                            remover_arquivo_forcado(caminho_item)
                            logger.debug(f"Arquivo removido: {item}")
                        elif os.path.isdir(caminho_item):
                            shutil.rmtree(caminho_item)
                            logger.debug(f"Pasta removida: {item}")
                    except Exception as e:
                        logger.warning(f"Erro ao remover {item}: {e}")
                        # Tentar com comando do sistema como último recurso
                        try:
                            if os.path.isfile(caminho_item):
                                subprocess.run(['del', '/F', '/Q', f'"{caminho_item}"'], shell=True, check=True)
                                logger.info(f"Arquivo removido com comando del: {item}")
                        except Exception as e2:
                            logger.error(f"Falha total ao remover {item}: {e2}")
                
                # Verificar se ainda há arquivos na pasta
                arquivos_restantes = os.listdir(PASTA_DESTINO_QBW)
                if arquivos_restantes:
                    logger.warning(f"Alguns arquivos não puderam ser removidos: {arquivos_restantes}")
                else:
                    logger.info("Pasta de destino limpa com sucesso.")
            else:
                logger.info("Pasta de destino já estava vazia.")
        else:
            logger.info("Pasta de destino não existe. Será criada.")
        
        # Criar a pasta de destino (caso não exista ou tenha sido removida)
        os.makedirs(PASTA_DESTINO_QBW, exist_ok=True)
        logger.info(f"Pasta de destino preparada: {PASTA_DESTINO_QBW}")
        
        # SEGUNDA ETAPA: VERIFICAR PASTA DE ORIGEM
        logger.info("=== VERIFICANDO PASTA DE ORIGEM ===")
        
        # Verificar se a pasta de origem existe e está acessível
        if not os.path.exists(PASTA_ORIGEM_QBW):
            logger.error(f"Pasta de origem não encontrada ou inacessível: {PASTA_ORIGEM_QBW}")
            raise FileNotFoundError(f"Pasta de origem não encontrada: {PASTA_ORIGEM_QBW}")
        
        logger.info(f"Pasta de origem acessível: {PASTA_ORIGEM_QBW}")
        
        # TERCEIRA ETAPA: BUSCAR ARQUIVO MAIS RECENTE
        logger.info("=== BUSCANDO ARQUIVO .QBW MAIS RECENTE ===")
        
        # Buscar todos os arquivos .qbw na pasta de origem
        padrao_qbw = os.path.join(PASTA_ORIGEM_QBW, "*.qbw")
        arquivos_qbw = glob.glob(padrao_qbw)
        
        if not arquivos_qbw:
            logger.error(f"Nenhum arquivo .qbw encontrado na pasta: {PASTA_ORIGEM_QBW}")
            raise FileNotFoundError(f"Nenhum arquivo .qbw encontrado na pasta de origem")
        
        logger.info(f"Encontrados {len(arquivos_qbw)} arquivo(s) .qbw na pasta de origem")
        
        # Listar todos os arquivos encontrados com suas datas
        for arquivo in arquivos_qbw:
            data_modificacao = time.ctime(os.path.getmtime(arquivo))
            logger.debug(f"Arquivo: {os.path.basename(arquivo)} - Data: {data_modificacao}")
        
        # Encontrar o arquivo mais recente
        arquivo_mais_recente = max(arquivos_qbw, key=os.path.getmtime)
        nome_arquivo = os.path.basename(arquivo_mais_recente)
        data_arquivo = time.ctime(os.path.getmtime(arquivo_mais_recente))
        
        logger.info(f"Arquivo .qbw mais recente: {nome_arquivo}")
        logger.info(f"Data de modificação: {data_arquivo}")
        
        # QUARTA ETAPA: COPIAR O ARQUIVO COM RETRY
        logger.info("=== COPIANDO ARQUIVO ===")
        
        # Definir o caminho de destino
        caminho_destino = os.path.join(PASTA_DESTINO_QBW, nome_arquivo)
        
        # Obter tamanho do arquivo para monitoramento
        tamanho_origem = os.path.getsize(arquivo_mais_recente)
        logger.info(f"Tamanho do arquivo: {tamanho_origem / (1024*1024):.2f} MB")
        
        logger.info(f"Copiando de: {arquivo_mais_recente}")
        logger.info(f"Copiando para: {caminho_destino}")
        
        # Tentar copiar com retry
        sucesso_copia = False
        max_tentativas = 3
        
        for tentativa in range(1, max_tentativas + 1):
            try:
                logger.info(f"Tentativa {tentativa} de {max_tentativas}")
                
                # Se o arquivo de destino já existe, tentar removê-lo primeiro
                if os.path.exists(caminho_destino):
                    logger.info("Arquivo de destino já existe. Removendo...")
                    remover_arquivo_forcado(caminho_destino)
                
                inicio_copia = time.time()
                shutil.copy2(arquivo_mais_recente, caminho_destino)
                fim_copia = time.time()
                
                tempo_copia = fim_copia - inicio_copia
                logger.info(f"Cópia concluída em {tempo_copia:.2f} segundos")
                sucesso_copia = True
                break
                
            except Exception as e:
                logger.warning(f"Tentativa {tentativa} falhou: {e}")
                if tentativa < max_tentativas:
                    logger.info("Aguardando antes da próxima tentativa...")
                    time.sleep(2)
                else:
                    logger.error("Todas as tentativas de cópia falharam")
                    raise
        
        if not sucesso_copia:
            raise RuntimeError("Falha em todas as tentativas de cópia")
        
        # QUINTA ETAPA: VERIFICAR A CÓPIA
        logger.info("=== VERIFICANDO CÓPIA ===")
        
        # Verificar se a cópia foi bem-sucedida
        if os.path.exists(caminho_destino):
            tamanho_destino = os.path.getsize(caminho_destino)
            
            if tamanho_origem == tamanho_destino:
                logger.info(f"✓ Arquivo copiado com sucesso!")
                logger.info(f"✓ Tamanhos conferem: {tamanho_destino} bytes")
                logger.info(f"✓ Arquivo disponível em: {caminho_destino}")
                return caminho_destino
            else:
                logger.error(f"✗ Erro: Tamanhos diferentes! Origem: {tamanho_origem}, Destino: {tamanho_destino}")
                raise RuntimeError("Falha na verificação do tamanho do arquivo copiado")
        else:
            logger.error("✗ Arquivo não foi encontrado no destino após a cópia")
            raise RuntimeError("Falha na cópia do arquivo .qbw")
            
    except Exception as e:
        logger.exception(f"Erro durante o processo de cópia do arquivo .qbw: {e}")
        raise

def fechar_processos_quickbooks():
    """
    Fecha todos os processos relacionados ao QuickBooks
    """
    try:
        logger.info("Verificando e fechando processos do QuickBooks...")
        
        processos_qb = ['QBWEnterprise.exe', 'qbw32.exe', 'QBDBMgr.exe', 'QBDBMgrN.exe']
        processos_fechados = 0
        
        for processo_nome in processos_qb:
            try:
                # Usar taskkill para forçar o fechamento
                resultado = subprocess.run(['taskkill', '/F', '/IM', processo_nome], 
                                         capture_output=True, text=True)
                if resultado.returncode == 0:
                    logger.info(f"Processo {processo_nome} fechado com sucesso")
                    processos_fechados += 1
                else:
                    logger.debug(f"Processo {processo_nome} não estava em execução")
            except Exception as e:
                logger.debug(f"Erro ao tentar fechar {processo_nome}: {e}")
        
        if processos_fechados > 0:
            logger.info(f"Total de {processos_fechados} processo(s) do QuickBooks fechado(s)")
            time.sleep(3)  # Aguardar os processos terminarem completamente
        else:
            logger.info("Nenhum processo do QuickBooks estava em execução")
            
    except Exception as e:
        logger.warning(f"Erro ao fechar processos do QuickBooks: {e}")

def remover_arquivo_forcado(caminho_arquivo):
    """
    Remove um arquivo usando múltiplos métodos, incluindo remoção forçada
    """
    try:
        # Método 1: Remover normalmente
        os.remove(caminho_arquivo)
        return True
    except:
        pass
    
    try:
        # Método 2: Remover atributos de somente leitura e tentar novamente
        os.chmod(caminho_arquivo, 0o777)
        os.remove(caminho_arquivo)
        return True
    except:
        pass
    
    try:
        # Método 3: Usar unlink
        os.unlink(caminho_arquivo)
        return True
    except:
        pass
    
    try:
        # Método 4: Usar comando del do Windows
        subprocess.run(['del', '/F', '/Q', f'"{caminho_arquivo}"'], shell=True, check=True)
        return True
    except:
        pass
    
    # Se chegou aqui, não conseguiu remover
    raise RuntimeError(f"Não foi possível remover o arquivo: {caminho_arquivo}")

def focar_programa_quickbooks():
    """
    Força o programa QuickBooks a vir para frente e ficar em foco usando métodos mais diretos
    """
    try:
        logger.info("Tentando focar no programa QuickBooks...")
        
        # Método 1: Usar pyautogui para encontrar e focar na janela (mais confiável)
        try:
            logger.info("Método 1: Buscando janela do QuickBooks...")
            windows = pyautogui.getAllWindows()
            qb_window = None
            
            # Procurar por janelas do QuickBooks com diferentes palavras-chave
            palavras_chave = ['quickbooks', 'qbw', 'intuit', 'enterprise']
            
            for window in windows:
                if window.title:  # Verificar se a janela tem título
                    title_lower = window.title.lower()
                    logger.debug(f"Verificando janela: {window.title}")
                    
                    for palavra in palavras_chave:
                        if palavra in title_lower:
                            qb_window = window
                            logger.info(f"Janela do QuickBooks encontrada: {window.title}")
                            break
                    
                    if qb_window:
                        break
            
            if qb_window:
                try:
                    # Trazer para frente e ativar
                    if hasattr(qb_window, 'restore'):
                        qb_window.restore()  # Restaurar se estiver minimizada
                    
                    qb_window.activate()
                    time.sleep(1)
                    
                    # Verificar se conseguiu ativar
                    if verificar_programa_em_foco():
                        logger.info("✓ QuickBooks focado com sucesso usando pyautogui!")
                        return True
                    
                except Exception as e:
                    logger.warning(f"Erro ao ativar janela: {e}")
                    
        except Exception as e:
            logger.warning(f"Método pyautogui falhou: {e}")
        
        # Método 2: Usar processo e comando nircmd (se disponível)
        try:
            logger.info("Método 2: Usando psutil para encontrar processo...")
            
            for proc in psutil.process_iter(['pid', 'name', 'exe']):
                try:
                    if proc.info['name']:
                        nome_processo = proc.info['name'].lower()
                        if any(palavra in nome_processo for palavra in ['qbw', 'quickbooks']):
                            logger.info(f"Processo QuickBooks encontrado: {proc.info['name']} (PID: {proc.info['pid']})")
                            
                            # Tentar usar nircmd se disponível
                            try:
                                subprocess.run(['nircmd', 'win', 'activate', 'process', proc.info['name']], 
                                             check=True, timeout=5, capture_output=True)
                                logger.info("✓ Janela focada usando nircmd!")
                                time.sleep(2)
                                return True
                            except FileNotFoundError:
                                logger.debug("nircmd não está disponível")
                                break
                            except Exception as e:
                                logger.debug(f"nircmd falhou: {e}")
                                break
                                
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
                    
        except Exception as e:
            logger.warning(f"Método psutil falhou: {e}")
        
        # Método 3: Usar VBScript para focar na janela
        try:
            logger.info("Método 3: Usando VBScript...")
            
            vbs_script = '''
            Set objShell = CreateObject("WScript.Shell")
            Set objWindows = CreateObject("Shell.Application").Windows
            
            For Each objWindow In objWindows
                If InStr(LCase(objWindow.LocationName), "quickbooks") > 0 Or _
                   InStr(LCase(objWindow.LocationName), "qbw") > 0 Or _
                   InStr(LCase(objWindow.LocationName), "intuit") > 0 Then
                    objWindow.Visible = True
                    objShell.AppActivate objWindow.HWND
                    WScript.Quit
                End If
            Next
            
            ' Tentar também por nome do processo
            objShell.AppActivate "QuickBooks"
            '''
            
            # Salvar script temporário
            script_path = os.path.join(os.getcwd(), "temp_focus.vbs")
            with open(script_path, 'w') as f:
                f.write(vbs_script)
            
            # Executar script
            subprocess.run(['cscript', '//nologo', script_path], 
                         check=True, timeout=10, capture_output=True)
            
            # Remover script temporário
            os.remove(script_path)
            
            logger.info("✓ Script VBScript executado!")
            time.sleep(2)
            
            if verificar_programa_em_foco():
                logger.info("✓ QuickBooks focado com VBScript!")
                return True
                
        except Exception as e:
            logger.warning(f"Método VBScript falhou: {e}")
        
        # Método 4: Força bruta - clique na barra de tarefas
        try:
            logger.info("Método 4: Procurando na barra de tarefas...")
            
            # Tomar screenshot e procurar por ícones relacionados ao QuickBooks
            screenshot = pyautogui.screenshot()
            
            # Procurar por possíveis ícones do QuickBooks na barra de tarefas
            # (área inferior da tela)
            screen_width, screen_height = pyautogui.size()
            barra_tarefas_y = screen_height - 50  # Aproximadamente a altura da barra de tarefas
            
            # Clique em algumas posições da barra de tarefas onde o QB pode estar
            posicoes_teste = [
                (100, barra_tarefas_y),
                (150, barra_tarefas_y),
                (200, barra_tarefas_y),
                (250, barra_tarefas_y)
            ]
            
            for x, y in posicoes_teste:
                try:
                    pyautogui.click(x, y)
                    time.sleep(1)
                    
                    if verificar_programa_em_foco():
                        logger.info(f"✓ QuickBooks focado clicando na posição ({x}, {y})!")
                        return True
                        
                except Exception:
                    continue
                    
        except Exception as e:
            logger.warning(f"Método barra de tarefas falhou: {e}")
        
        # Método 5: Simples clique no centro da tela
        try:
            logger.info("Método 5: Clique no centro da tela...")
            screen_width, screen_height = pyautogui.size()
            pyautogui.click(screen_width // 2, screen_height // 2)
            time.sleep(1)
            
            # Verificar se por acaso já estava focado
            if verificar_programa_em_foco():
                logger.info("✓ QuickBooks já estava focado!")
                return True
                
        except Exception as e:
            logger.warning(f"Método clique central falhou: {e}")
        
        logger.warning("Todos os métodos de foco falharam, mas continuando...")
        return False
        
    except Exception as e:
        logger.exception(f"Erro crítico ao focar no programa QuickBooks: {e}")
        return False

def verificar_programa_em_foco():
    """
    Verifica se o programa QuickBooks está em foco de forma mais robusta
    """
    try:
        # Método 1: pyautogui
        try:
            active_window = pyautogui.getActiveWindow()
            if active_window and active_window.title:
                title = active_window.title.lower()
                palavras_chave = ['quickbooks', 'qbw', 'intuit', 'enterprise']
                
                for palavra in palavras_chave:
                    if palavra in title:
                        logger.info(f"✓ QuickBooks está em foco: {active_window.title}")
                        return True
                        
                logger.debug(f"Programa em foco: {active_window.title}")
                return False
        except:
            pass
        
        # Método 2: Usar VBScript para verificar janela ativa
        try:
            vbs_check = '''
            Set objShell = CreateObject("WScript.Shell")
            Set objExec = objShell.Exec("tasklist /fi ""STATUS eq running"" /fo csv")
            WScript.Echo "checking"
            '''
            
            # Verificação simplificada
            result = subprocess.run(['powershell', '-Command', 
                                   'Get-Process | Where-Object {$_.MainWindowTitle -like "*QuickBooks*" -or $_.MainWindowTitle -like "*QBW*"} | Select-Object MainWindowTitle'],
                                   capture_output=True, text=True, timeout=5)
            
            if result.stdout and ('QuickBooks' in result.stdout or 'QBW' in result.stdout):
                logger.info("✓ QuickBooks detectado via PowerShell")
                return True
                
        except:
            pass
        
        logger.debug("QuickBooks não está em foco ou não foi detectado")
        return False
        
    except Exception as e:
        logger.warning(f"Erro ao verificar janela em foco: {e}")
        return False

# Funções auxiliares usando OpenCV
def wait_for_images_cv2(image_names, timeout=300, threshold=0.8):
    """
    Espera até que uma das imagens apareça na tela ou até o tempo limite usando OpenCV.
    Retorna o nome da imagem detectada e a localização.
    """
    start_time = time.time()
    templates = {}
    for name in image_names:
        template = cv2.imread(os.path.join(IMAGES_PATH, name), cv2.IMREAD_GRAYSCALE)
        if template is None:
            logger.error(f"Imagem {name} não encontrada no diretório {IMAGES_PATH}.")
            raise FileNotFoundError(f"Imagem {name} não encontrada no diretório {IMAGES_PATH}.")
        templates[name] = template
        logger.debug(f"Template {name} carregado com sucesso.")

    while time.time() - start_time < timeout:
        logger.debug(f"Verificando as imagens {image_names} com OpenCV...")
        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)

        for name, template in templates.items():
            w, h = template.shape[::-1]
            res = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
            loc = np.where(res >= threshold)
            if len(loc[0]) > 0:
                pt = (loc[1][0], loc[0][0])
                center = (pt[0] + w // 2, pt[1] + h // 2)
                logger.info(f"Imagem {name} detectada em {center}.")
                return name, center
        time.sleep(1)
    logger.warning(f"Nenhuma das imagens {image_names} apareceu dentro do tempo limite de {timeout} segundos.")
    raise TimeoutError(f"Nenhuma das imagens {image_names} apareceu dentro do tempo limite.")

def check_image_with_matches(image_name, min_matches=10):
    """
    Verifica se a imagem aparece na tela com base no número de correspondências.
    """
    logger.debug(f"Verificando a imagem {image_name} com base em correspondências...")
    template = cv2.imread(os.path.join(IMAGES_PATH, image_name), cv2.IMREAD_GRAYSCALE)
    if template is None:
        logger.error(f"Imagem {image_name} não encontrada no diretório {IMAGES_PATH}.")
        raise FileNotFoundError(f"Imagem {image_name} não encontrada no diretório {IMAGES_PATH}.")

    screenshot = pyautogui.screenshot()
    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)

    orb = cv2.ORB_create()
    keypoints1, descriptors1 = orb.detectAndCompute(template, None)
    keypoints2, descriptors2 = orb.detectAndCompute(screenshot, None)

    if descriptors1 is None or descriptors2 is None:
        logger.warning("Descritores não encontrados para uma das imagens.")
        return False

    bf = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
    matches = bf.match(descriptors1, descriptors2)
    matches = sorted(matches, key=lambda x: x.distance)

    logger.info(f"Número de correspondências encontradas: {len(matches)}")
    return len(matches) >= min_matches

def click_image_cv2(image_name):
    """Localiza a imagem na tela usando OpenCV e clica no centro."""
    try:
        _, location = wait_for_images_cv2([image_name])
        pyautogui.click(location)
        logger.info(f"Clicado na imagem {image_name} em {location}.")
        time.sleep(2)  # Aguarda brevemente após o clique
    except TimeoutError as e:
        logger.error(f"Erro ao clicar na imagem {image_name}: {e}")
        raise

def handle_security_warning():
    """
    Monitora a tela de aviso de segurança e realiza os cliques necessários para fechar.
    """
    try:
        logger.info("Monitorando a tela de aviso de segurança...")
        detected_image, location = wait_for_images_cv2(["fechar_seguranca.PNG"], timeout=120)
        logger.info(f"Aviso de segurança detectado na tela: {detected_image}. Clicando para fechar...")

        # Clicar na imagem fechar_seguranca_clicar.PNG
        click_image_cv2("fechar_seguranca_clicar.PNG")
        logger.info("Clicado na opção de fechar o aviso de segurança.")
        time.sleep(2)
        pyautogui.press("enter")

    except TimeoutError:
        logger.warning("Tela de aviso de segurança não foi detectada no tempo limite. Continuando o processo...")
    except Exception as e:
        logger.exception(f"Erro ao lidar com o aviso de segurança: {e}")
        raise

# Etapas do processo
def main():
    try:
        # PRIMEIRA ETAPA: Copiar o último arquivo .qbw
        logger.info("=== INICIANDO CÓPIA DO ARQUIVO .QBW ===")
        arquivo_copiado = copiar_ultimo_arquivo_qbw()
        logger.info(f"Arquivo .qbw copiado com sucesso: {arquivo_copiado}")
        
        # Aguardar um pouco após a cópia
        time.sleep(3)
        
        logger.info("=== INICIANDO O PROGRAMA QUICKBOOKS ===")
        
        # 1. Abrir o programa
        os.startfile(PROGRAM_PATH)
        logger.info("Programa iniciado, aguarde...")
        time.sleep(7)

        # 2. FORÇAR O PROGRAMA A VIR PARA FRENTE
        logger.info("=== FOCANDO NO PROGRAMA QUICKBOOKS ===")
        focar_programa_quickbooks()
        
        # Aguardar um pouco após focar
        time.sleep(3)
        
        # 3. Verificar se o programa está em foco
        if not verificar_programa_em_foco():
            logger.warning("QuickBooks pode não estar em foco. Tentando focar novamente...")
            focar_programa_quickbooks()
            time.sleep(2)

        logger.info("Continuando com o processo...")
        
        # 4. Verificar e lidar com o aviso de segurança
        handle_security_warning()

        # 5. Monitorar até aparecer qualquer uma das telas esperadas
        logger.info("Monitorando as telas iniciais...")
        detected_image, location = wait_for_images_cv2(["inicio.PNG", "senha.PNG"])

        if detected_image == "inicio.PNG":
            logger.info("Tela inicial detectada.")
            logger.info("Clicando no botão inicial...")
            click_image_cv2("open_inicio.PNG")
            logger.info("Aguardando a tela de senha...")
            time.sleep(10)
            
            # Monitorar a tela de senha
            logger.info("Monitorando a tela de senha...")
            wait_for_images_cv2(["senha.PNG"])
            logger.info("Tela de senha detectada.")

        if detected_image == "senha.PNG":
            logger.info("Tela de senha detectada diretamente.")
            logger.info("Clicando na tela de senha...")
            click_image_cv2("senha.PNG")  # Adiciona o clique na tela de senha detectada
            
        # Inserir a senha e clicar no botão OK
        logger.info("Inserindo senha...")
        pyautogui.typewrite(PASSWORD, interval=0.1)
        click_image_cv2("ok_senha.PNG")
        time.sleep(15)

        # Verificar se `fechar1.PNG` aparece
        logger.info("Verificando se a imagem 'fechar1.PNG' aparece...")
        try:
            detected_image, location = wait_for_images_cv2(["fechar1.PNG"], timeout=5)
            logger.info("Imagem 'fechar1.PNG' detectada. Clicando em 'fechar1_clicar.PNG'...")
            click_image_cv2("fechar1_clicar.PNG")
        except TimeoutError:
            logger.info("Imagem 'fechar1.PNG' não detectada. Verificando 'logado_sucesso.PNG'...")

        # Monitorar até aparecer a tela de sucesso
        if check_image_with_matches("logado_sucesso.PNG"):
            logger.info("Logado com sucesso detectado com correspondências.")
        else:
            logger.warning("Logado com sucesso não detectado.")

        # Chamar o próximo arquivo
        logger.info("Executando o arquivo 2-logado.py...")
        os.system("python 2-logado.py")
        logger.info("Processo concluído com sucesso.")

    except Exception as e:
        logger.exception(f"Erro: {e}")

if __name__ == "__main__":
    main()