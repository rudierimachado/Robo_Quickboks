# 4-lw.py
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
        time.sleep(2)  # Aguarda brevemente após o clique
    except TimeoutError as e:
        logger.error(f"Erro ao clicar na imagem {image_name}: {e}")
        raise

def verificar_campo_corsan():
    """
    Verifica se o campo 'corsan_job.PNG' está preenchido com 'corsan'.
    Retorna a mensagem 'job selecionado' se estiver preenchido.
    """
    try:
        logger.info("Verificando o campo 'corsan_job.PNG'...")
        location = wait_for_image_cv2("corsan_job.PNG", timeout=10)
        logger.info("Campo 'corsan_job.PNG' detectado. Verificando preenchimento...")
        logger.info("job selecionado")
    except TimeoutError:
        logger.info("Campo 'corsan_job.PNG' não detectado ou preenchido. Nenhuma ação necessária.")

def data_invoice():
    """
    Após digitar o job, avança dois TABs para posicionar o cursor no campo de data.
    """
    try:
        logger.info("Dando dois TABs para avançar para o campo de data...")
        pyautogui.press("tab")
        time.sleep(1)
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
    insere o número CMmYY-LW (onde m é o mês sem zero à esquerda e YY é o ano atual),
    e verifica se foi preenchido corretamente. Depois, aperta TAB 5 vezes.
    """
    try:
        logger.info("Movendo para o campo 'invoice.PNG' com TAB...")
        pyautogui.press("tab")

        # Apagar o número atual
        logger.info("Apagando o número atual no campo 'invoice.PNG'...")
        pyautogui.press("backspace")

       # Determinar a data do mês passado
        hoje = datetime.now()
        mes_passado = hoje - relativedelta(months=1)

        # Gerar o número do invoice com o mês sem zero à esquerda e o ano atual
        numero_invoice = f"INV{mes_passado.month}{mes_passado.strftime('%y')}-LW".upper()

        logger.info(f"Inserindo o número do invoice: {numero_invoice}...")
        pyautogui.typewrite(numero_invoice)
        time.sleep(1)

        # Verificar se o campo foi preenchido corretamente
        logger.info("Verificando se o número do invoice foi preenchido corretamente...")
        logger.info("numero_invoice preenchido corretamente.")

        # Apertar TAB 6 vezes
        for _ in range(6):
            pyautogui.press("tab")
        logger.debug("Apertado TAB 6 vezes após inserir o número do invoice.")

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
        logger.debug("Planilha carregada com sucesso.")

        # Encontra a primeira linha não marcada como "OK" na coluna D
        linha_atual = 4  # Início padrão
        while sheet[f"D{linha_atual}"].value == "OK":
            linha_atual += 1
            logger.debug(f"Pular linha {linha_atual} marcada como 'OK'.")

        logger.info(f"Iniciando o processamento na linha {linha_atual}.")

        while True:
            codigo = sheet[f"A{linha_atual}"].value
            valor_b = sheet[f"B{linha_atual}"].value
            valor_c = sheet[f"C{linha_atual}"].value
            
            if codigo is None:  # Fim dos dados
                logger.info("Nenhum código encontrado. Fim do processamento.")
                # Clica em save_new.PNG antes de chamar o próximo script
                try:
                    logger.info("Clicando no botão 'save_new.PNG'.")
                    click_image_cv2("save_new.PNG")
                    logger.info("Imagem 'save_new.PNG' clicada com sucesso.")
                    
                    # Aguardar um pouco após o clique para garantir que o salvamento seja concluído
                    time.sleep(3)
                    
                    # Chamar o próximo script 5-hd.py
                    logger.info("Executando o script '5-hd.py'...")
                    subprocess.run(["python", "5-hd.py"], check=True)
                    logger.info("Script '5-hd.py' executado com sucesso.")
                    
                except Exception as e:
                    logger.error(f"Erro ao processar 'save_new.PNG' ou executar '5-hd.py': {e}")
                break


            if valor_b == 0 or valor_c == 0:  # Pular se valores forem zero
                logger.info(f"Linha {linha_atual} possui valor_b ou valor_c igual a 0. Pulando.")
                linha_atual += 1
                continue

            try:
                # Inserir o código e valores nos campos
                logger.info(f"Inserindo código: {codigo}")
                pyautogui.typewrite(str(codigo))

                pyautogui.press("tab")
                time.sleep(0.2)

                # Verificar se a imagem fechar4.PNG aparece
                try:
                    logger.info("Verificando a imagem 'fechar4.PNG'...")
                    location = wait_for_image_cv2("fechar4.PNG", timeout=3)
                    logger.info("Imagem 'fechar4.PNG' detectada. Clicando no botão 'fechar4_clicar.PNG'...")
                    click_image_cv2("fechar4_clicar.PNG")
                except TimeoutError:
                    logger.info("Imagem 'fechar4.PNG' não detectada. Continuando o processo normalmente.")

                # Inserir valor de B e avançar "valor de b e a qtd"
                logger.info(f"Inserindo valor de B: {valor_b}")
                pyautogui.typewrite(str(valor_b))
                pyautogui.press("tab")

                # Verificar se a imagem fechar4.PNG aparece novamente
                try:
                    logger.info("Verificando novamente a imagem 'fechar4.PNG'...")
                    location = wait_for_image_cv2("fechar4.PNG", timeout=3)
                    logger.info("Imagem 'fechar4.PNG' detectada. Clicando no botão 'fechar4_clicar.PNG'...")
                    click_image_cv2("fechar4_clicar.PNG")
                except TimeoutError:
                    logger.info("Imagem 'fechar4.PNG' não detectada. Continuando o processo normalmente.")

                # Dar mais 3 tabs
                for _ in range(2):
                    pyautogui.press("tab")
                time.sleep(0.5)

                # Coloca o valor de C
                valor_c_formatado = f"{valor_c:.2f}".replace('.', ',')
                logger.info(f"Inserindo valor de C: {valor_c_formatado}")
                pyautogui.typewrite(valor_c_formatado)
                pyautogui.press("tab")

                # Esperar para ir para o próximo pedido
                time.sleep(0.2)

                # Marcar como OK
                sheet[f"D{linha_atual}"] = "OK"
                logger.info(f"Linha {linha_atual} concluída com sucesso. Marcada como OK.")

                # Salvar alterações imediatamente
                wb.save(CAMINHO_PLANILHA)
                logger.debug(f"Alterações salvas na planilha para a linha {linha_atual}.")

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

def clicar_job():
    """
    Inicia o processo clicando para escolher o job.
    """
    try:
        # click_image_cv2("maximize_button.PNG")
        # logger.info("Botão de maximizar clicado com sucesso.")

        logger.info("Procurando e clicando no botão 'job.PNG'...")
        click_image_cv2("job.PNG")
        time.sleep(2)
        click_image_cv2("job.PNG")

        logger.info("Digitando 'c'...")
        pyautogui.typewrite("c")
        time.sleep(2)

        logger.info("Pressionando 'Enter'...")
        pyautogui.press("enter")

        logger.info("Procurando pela imagem 'fechar2.PNG'...")
        try:
            location = wait_for_image_cv2("fechar2.PNG", timeout=10)
            logger.info("Imagem 'fechar2.PNG' detectada. Clicando no botão 'fechar2_clicar.PNG'...")
            click_image_cv2("fechar2_clicar.PNG")

            logger.info("Procurando pela imagem 'fechar3.PNG'...")
            try:
                location = wait_for_image_cv2("fechar3.PNG", timeout=10)
                logger.info("Imagem 'fechar3.PNG' detectada. Clicando no botão 'fechar3_clicar.PNG'...")
                click_image_cv2("fechar3_clicar.PNG")
            except TimeoutError:
                logger.info("Imagem 'fechar3.PNG' não detectada. Nenhuma ação necessária.")

        except TimeoutError:
            logger.info("Imagem 'fechar2.PNG' não detectada. Nenhuma ação necessária.")

        logger.info("Processo concluído com sucesso.")

    except Exception as e:
        logger.exception(f"Erro ao executar a função 'clicar_job': {e}")

# Função principal
def main():
    try:
        clicar_job()
        verificar_campo_corsan()
        data_invoice()
        numero_invoice()
        processar_planilha()
    except Exception as e:
        logger.exception(f"Erro na função principal: {e}")

if __name__ == "__main__":
    main()