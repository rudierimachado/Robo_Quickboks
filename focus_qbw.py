# foco_qbw.py
import win32gui
import win32con
import win32process
import psutil
import time
import pyautogui
from logger_config import setup_logger

# Inicializa o logger
logger = setup_logger()

def find_window_by_process_name(process_name):
    """
    Encontra todas as janelas pertencentes a um processo específico.
    Retorna uma lista de handles de janelas.
    """
    target_windows = []
    
    def callback(hwnd, extra):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            _, process_id = win32process.GetWindowThreadProcessId(hwnd)
            try:
                process = psutil.Process(process_id)
                if process_name.lower() in process.name().lower():
                    window_text = win32gui.GetWindowText(hwnd)
                    logger.debug(f"Janela encontrada: {window_text} (hwnd: {hwnd}) do processo {process.name()}")
                    target_windows.append(hwnd)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        return True
    
    win32gui.EnumWindows(callback, None)
    return target_windows

def focus_quickbooks_window():
    """
    Localiza e ativa a janela do QuickBooks Enterprise.
    Retorna True se conseguir focar a janela, False caso contrário.
    """
    logger.info("Tentando localizar e focar na janela do QuickBooks...")
    
    # Lista de possíveis nomes de processos do QuickBooks
    qb_process_names = ["QBWEnterprise.exe", "qbw.exe", "QuickBooks.exe"]
    
    for process_name in qb_process_names:
        windows = find_window_by_process_name(process_name)
        
        if windows:
            logger.info(f"Encontradas {len(windows)} janelas do processo {process_name}")
            
            # Tenta ativar cada janela encontrada, começando pela última (geralmente a mais recente)
            for hwnd in reversed(windows):
                try:
                    window_text = win32gui.GetWindowText(hwnd)
                    logger.info(f"Tentando ativar janela: {window_text} (hwnd: {hwnd})")
                    
                    # Verifica se a janela não está minimizada
                    if win32gui.IsIconic(hwnd):
                        logger.info("Janela está minimizada. Restaurando...")
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        time.sleep(0.5)
                    
                    # Traz a janela para o primeiro plano
                    win32gui.SetForegroundWindow(hwnd)
                    
                    # Maximiza a janela
                    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                    
                    logger.info(f"Janela do QuickBooks ativada com sucesso: {window_text}")
                    
                    # Pequena pausa para garantir que a janela esteja realmente em foco
                    time.sleep(0.5)
                    
                    # Verificação adicional se está em foco
                    active_window = pyautogui.getActiveWindow()
                    if active_window and "QuickBooks" in active_window.title:
                        logger.info(f"Confirmado: janela ativa é {active_window.title}")
                    else:
                        logger.warning(f"Alerta: janela ativa parece ser {active_window.title if active_window else 'Nenhuma'}")
                    
                    return True
                    
                except Exception as e:
                    logger.warning(f"Erro ao ativar janela {hwnd}: {e}")
                    continue
    
    logger.error("Não foi possível encontrar ou ativar nenhuma janela do QuickBooks")
    return False

# Para uso como módulo ou script standalone
if __name__ == "__main__":
    focus_quickbooks_window()