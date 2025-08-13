# 2-logado.py
import pyautogui
import cv2
import numpy as np
import os
import time
import subprocess
from logger_config import setup_logger

# Inicializa o logger
logger = setup_logger()

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

# Etapas do processo
def main():
    try:
        logger.info("Monitorando a tela para o botão 'create_invoices.PNG'...")
        click_image_cv2("create_invoices.PNG")
        logger.info("Botão 'create_invoices.PNG' detectado e clicado com sucesso.")
        
        # Aguarda 5 segundos (o comentário original mencionava 10 segundos, mas o código usa 5)
        logger.info("Aguardando 5 segundos antes de chamar o próximo script...")
        time.sleep(5)
        
        # Chama o arquivo 3-planilha.py
        logger.info("Executando o script '3-planilha.py'...")
        subprocess.run(["python", "3-planilha.py"], check=True)
        logger.info("Script '3-planilha.py' executado com sucesso.")
        
    except Exception as e:
        logger.exception(f"Erro: {e}")

if __name__ == "__main__":
    main()
