# 5-hd.py
import pyautogui
import cv2
import numpy as np
import os
import time
from datetime import datetime, timedelta
import pyperclip
import openpyxl
import locale
import subprocess
from logger_config import setup_logger
from dateutil.relativedelta import relativedelta

# Inicializa o logger
logger = setup_logger()

# Configurações locais
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    logger.debug("Locale configurado para 'pt_BR.UTF-8'.")
except locale.Error as e:
    logger.error(f"Erro ao configurar locale: {e}")
    raise

# Configurações iniciais
IMAGES_PATH = "img"
pyautogui.FAILSAFE = False

# Funções auxiliares usando OpenCV
def wait_for_image_cv2(image_name, timeout=300, threshold=0.8):
    """
    Espera até que a imagem apareça na tela ou até o tempo limite usando OpenCV.
    Retorna a localização da imagem detectada.
    """
    start_time = time.time()
    template = cv2.imread(os.path.join(IMAGES_PATH, image_name), cv2.IMREAD_GRAYSCALE)
    if template is None:
        logger.error(f"Imagem {image_name} não encontrada no diretório {IMAGES_PATH}.")
        raise FileNotFoundError(f"Imagem {image_name} não encontrada no diretório {IMAGES_PATH}.")
    w, h = template.shape[::-1]
    logger.debug(f"Template {image_name} carregado com sucesso.")

    while time.time() - start_time < timeout:
        logger.debug(f"Verificando a imagem {image_name} com OpenCV...")
        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)
        res = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        loc = np.where(res >= threshold)
        if len(loc[0]) > 0:
            pt = (loc[1][0], loc[0][0])
            center = (pt[0] + w // 2, pt[1] + h // 2)
            logger.info(f"Imagem {image_name} detectada em {center}.")
            return center
        time.sleep(1)
    logger.warning(f"A imagem {image_name} não apareceu dentro do tempo limite de {timeout} segundos.")
    raise TimeoutError(f"A imagem {image_name} não apareceu dentro do tempo limite.")

def click_image_cv2(image_name):
    """Localiza a imagem na tela usando OpenCV e clica no centro."""
    try:
        location = wait_for_image_cv2(image_name)
        pyautogui.click(location)
        logger.info(f"Clicado na imagem {image_name} em {location}.")
        time.sleep(1)  # Aguarda brevemente após o clique
    except TimeoutError as e:
        logger.error(f"Erro ao clicar na imagem {image_name}: {e}")
        raise

# Mesma função do campo corsan do arquivo anterior, adaptada para 'hd'
def verificar_campo_hd():
    """
    Verifica se o campo 'Corsan_Depot_job.PNG' está preenchido com 'corsan_Depot'.
    Retorna a mensagem 'job selecionado' se estiver preenchido.
    """
    try:
        logger.info("Verificando o campo 'Corsan_Depot_job.PNG'...")
        location = wait_for_image_cv2("Corsan_Depot_job.PNG", timeout=10)
        logger.info("Campo 'Corsan_Depot_job.PNG' detectado. Verificando preenchimento...")
        logger.info("job selecionado")
    except TimeoutError:
        logger.info("Campo 'Corsan_Depot_job.PNG' não detectado ou preenchido. Nenhuma ação necessária.")



def data_invoice():
    """
    Após digitar o job, avança dois TABs para posicionar o cursor no campo de data.
    """
    try:
        logger.info("Dando dois TABs para avançar para o campo de data...")
        pyautogui.press("tab")
        time.sleep(1)
        
        
        
        # Inserir a última data do mês anterior
        hoje = datetime.now()
        primeiro_dia_mes_atual = datetime(hoje.year, hoje.month, 1)
        ultimo_dia_mes_passado = primeiro_dia_mes_atual - timedelta(days=1)
        ultima_data = ultimo_dia_mes_passado.strftime("%d/%m/%Y")
        
        logger.info(f"Inserindo a data: {ultima_data}...")
        pyautogui.typewrite(ultima_data)
        time.sleep(1)
        
        logger.info("Data preenchida corretamente.")
    except Exception as e:
        logger.exception(f"Erro ao avançar para o campo de data: {e}")

def numero_invoice():
    """
    Após o preenchimento da data, pressiona TAB, apaga o conteúdo do campo 'invoice.PNG',
    insere o número INVmYY-HD (onde m é o mês sem zero à esquerda e YY é o ano atual),
    e verifica se foi preenchido corretamente. Depois, aperta TAB 5 vezes.
    """
    try:
        logger.info("Movendo para o campo 'invoice.PNG' com TAB...")
        pyautogui.press("tab")
        time.sleep(1)
        
       

        # Apagar o número atual
        logger.info("Apagando o número atual no campo 'invoice.PNG'...")
        pyautogui.press("backspace")

        # Determinar o número do invoice com mês sem zero à esquerda e ano atuais
        hoje = datetime.now()
        mes_passado = hoje - relativedelta(months=1)

        # Gerar o número do invoice com o mês sem zero à esquerda e o ano atuais
        numero_invoice = f"INV{mes_passado.month}{mes_passado.strftime('%y')}-HD".upper()

        logger.info(f"Inserindo o número do invoice: {numero_invoice}...")
        pyautogui.typewrite(numero_invoice)
        time.sleep(1)

        # Verificar se o campo foi preenchido corretamente
        logger.info("Verificando se o número do invoice foi preenchido corretamente...")
        logger.info("numero_invoice preenchido corretamente")

        # Apertar TAB 6 vezes
        for _ in range(6):
            pyautogui.press("tab")

    except Exception as e:
        logger.exception(f"Erro ao preencher o número do invoice: {e}")

def processar_planilha():
    """
    Verifica a planilha e realiza as operações de cadastrar o pedido, iniciando da próxima linha não marcada como 'OK'.
    """
    try:
        CAMINHO_PLANILHA = r"C:\Users\samuca2\Documents\base_robo_quickboks\nova_planilha.xlsx"
        logger.info(f"Abrindo a planilha: {CAMINHO_PLANILHA}")
        wb = openpyxl.load_workbook(CAMINHO_PLANILHA, read_only=False)

        
        sheet = wb.active

        # Encontra a primeira linha não marcada como "OK" 
        linha_atual = 4  # Início padrão
        while sheet[f"G{linha_atual}"].value == "OK":
            linha_atual += 1

        logger.info(f"Iniciando o processamento na linha {linha_atual}.")

        while True:
            codigo = sheet[f"A{linha_atual}"].value
            valor_e = sheet[f"e{linha_atual}"].value
            valor_f = sheet[f"f{linha_atual}"].value
            
            if codigo is None:  # Fim dos dados
                logger.info("Nenhum código encontrado. Fim do processamento.")
                try:
                    logger.info("Clicando no botão 'save_new.PNG'.")
                    click_image_cv2("save_new.PNG")
                    logger.info("Imagem 'save_new.PNG' clicada com sucesso.")
                    
                    # Chamar o próximo script 6-wm.py
                    logger.info("Executando o script '6-wm.py'...")
                    subprocess.run(["python", "6-wm.py"], check=True)
                    logger.info("Script '6-wm.py' executado com sucesso.")
                    
                except Exception as e:
                    logger.error(f"Erro ao processar 'save_new.PNG' ou executar '6-wm.py': {e}")
                break

            if valor_e == 0 or valor_f == 0:  # Pular se valores forem zero
                logger.info(f"Linha {linha_atual} possui valor_e ou valor_f igual a 0. Pulando...")
                linha_atual += 1
                continue

            try:
                # Inserir o código e valores nos campos
                logger.info(f"Inserindo código: {codigo}")
                pyautogui.typewrite(str(codigo))
                pyautogui.press("tab")
                time.sleep(0.5)

                # Verificar se a imagem fechar4.PNG aparece
                try:
                    logger.info("Verificando a imagem 'fechar4.PNG'...")  # Verificar indentação dessa linha
                    location = wait_for_image_cv2("fechar4.PNG", timeout=3)
                    logger.info("Imagem 'fechar4.PNG' detectada. Clicando no botão 'fechar4_clicar.PNG'...")
                    click_image_cv2("fechar4_clicar.PNG")
                except TimeoutError:
                    logger.info("Imagem 'fechar4.PNG' não detectada. Continuando o processo normalmente.")

                # Inserir valor de E e avançar "valor de E e a qtd"
                logger.info(f"Inserindo valor de E: {valor_e}")
                pyautogui.typewrite(str(valor_e))
                pyautogui.press("tab")

                # Verificar novamente a imagem fechar4.PNG
                try:
                    logger.info("Verificando novamente a imagem 'fechar4.PNG'...")  # Verificar indentação dessa linha
                    location = wait_for_image_cv2("fechar4.PNG", timeout=3)
                    logger.info("Imagem 'fechar4.PNG' detectada. Clicando no botão 'fechar4_clicar.PNG'...")
                    click_image_cv2("fechar4_clicar.PNG")
                except TimeoutError:
                    logger.info("Imagem 'fechar4.PNG' não detectada. Continuando o processo normalmente.")

                # Dar mais 2 tabs (total de 4)
                pyautogui.press("tab")
                pyautogui.press("tab")
                time.sleep(0.5)

                # Colocar o valor de F que é o US$
                valor_f_formatado = f"{valor_f:.2f}".replace('.', ',')
                logger.info(f"Inserindo valor de F: {valor_f_formatado}")
                pyautogui.typewrite(valor_f_formatado)
                pyautogui.press("tab")

                # Esperar para ir para o próximo pedido
                time.sleep(0.5)

                # Marcar como OK
                sheet[f"G{linha_atual}"] = "OK"
                logger.info(f"Linha {linha_atual} concluída com sucesso. Marcada como OK.")

                # Salvar alterações imediatamente
                wb.save(CAMINHO_PLANILHA)
                logger.info(f"Alterações salvas na planilha para a linha {linha_atual}.")

                # Avançar para a próxima linha
                pyautogui.press("tab")
                linha_atual += 1

            except Exception as process_error:
                logger.exception(f"Erro ao processar a linha {linha_atual}: {process_error}")
                break

        # Salvar alterações finais na planilha
        wb.save(CAMINHO_PLANILHA)
        logger.info("Processo de planilha concluído.")

    except Exception as e:
        logger.exception(f"Erro ao processar a planilha: {e}")


def clicar_job_hd():
    """
    Inicia o processo clicando para escolher o job no contexto 'HD'.
    """
    try:
        logger.info("Procurando e clicando no botão 'job.PNG'...")
        click_image_cv2("job.PNG")
        time.sleep(2)
        click_image_cv2("job.PNG")

        logger.info("Digitando 'c'...")
        pyautogui.typewrite("c")
        pyautogui.press("down")
        time.sleep(1)

        logger.info("Pressionando 'tab'...")
        pyautogui.press("tab")

      

        logger.info("Processo de clique no job concluído com sucesso.")

    except Exception as e:
        logger.exception(f"Erro ao clicar no job: {e}")

# Função principal
def main():
    try:
        clicar_job_hd()
        verificar_campo_hd()
        data_invoice()
        numero_invoice()
        processar_planilha()
        
       
    except Exception as e:
        logger.exception(f"Erro no processo principal: {e}")


    
def chamar_proximo_script():
    
      #Chama o script 6-wm.py após o término da execução do script principal.
   
    try:
        logger.info("Chamando o script '6-wm.py'...")
        subprocess.run(["python", "6-wm.py"], check=True)
        logger.info("Script '6-wm.py' executado com sucesso.")
    except subprocess.CalledProcessError as e:
        logger.error(f"Erro ao executar o script '6-wm.py': {e}")
    except FileNotFoundError:
        logger.error("O arquivo '6-wm.py' não foi encontrado. Verifique o caminho.")

if __name__ == "__main__":
    main()



