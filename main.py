import configparser
import logging
import sys
import time
from pathlib import Path

import requests

# --- Environment Detection (GUI vs CLI) ---
USE_GUI = False

if sys.platform == 'win32':
    try:
        import tkinter as tk
        from tkinter import messagebox
        USE_GUI = True
    except ImportError:
        USE_GUI = False
else:
    # 非 Windows 環境 (Linux/PXE) 直接使用 CLI 模式，避免 import 錯誤
    USE_GUI = False


def get_executable_version() -> str:
    """Reads the version from the executable's properties."""
    try:
        from win32api import GetFileVersionInfo, LOWORD, HIWORD
        info = GetFileVersionInfo(sys.executable, "\\")
        ms = info['FileVersionMS']
        ls = info['FileVersionLS']
        return f"{HIWORD(ms)}.{LOWORD(ms)}.{HIWORD(ls)}.{LOWORD(ls)}"
    except Exception:
        return f"Unknown"

# --- Path Handling for PyInstaller ---
def get_resource_path(relative_path: str) -> Path:
    """ Get absolute path to resource, works for dev and for PyInstaller """
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        # The .exe file is in the parent directory of that temp folder
        base_path = Path(sys.executable).parent
    else:
        # Not running in a bundle, use the script's directory
        base_path = Path(__file__).parent

    return base_path / relative_path

# --- Logger Setup ---
def setup_logging(log_file_path: str):
    """Sets up the logging configuration to file and console."""
    logging.basicConfig(
        level=logging.DEBUG,
        format='[%(asctime)s] [%(levelname)s] %(message)s',
        handlers=[logging.FileHandler(log_file_path, mode='w', encoding='utf-8'),
                  logging.StreamHandler(sys.stdout)]
    )

# --- Configuration Handling ---
def load_config():
    """
    Loads configuration from the .ini file.
    Returns a dictionary with config values or None if file is not found.
    """
    config_path = get_resource_path('config.ini')
    if not config_path.is_file():
        logging.error(f"Configuration file '{config_path}' not found.")
        return None

    config = configparser.ConfigParser()
    config.read(config_path, encoding='utf-8')
    
    try:
        settings = {
            'mes_server': config.get('Global', 'MES_Server').strip('"\' '),
            'mes_api': config.get('Global', 'MES_API').strip('"\' '),
            'mb_sn_path': config.get('Global', 'MB_SN_PATH').strip('"\' '),
            'retry_count': config.getint('Global', 'RETRY_COUNT', fallback=3),
            'retry_delay': config.getint('Global', 'RETRY_DELAY_SECONDS', fallback=5),
            'output_path': config.get('Global', 'OUTPUT_PATH', fallback='MES.txt').strip('"\' '),
            'log_path': config.get('Global', 'LOG_PATH', fallback='./log/').strip('"\' '),
            'request_timeout': config.getint('Global', 'REQUEST_TIMEOUT_SECONDS', fallback=10)
        }
        logging.info("Configuration loaded successfully.")
        return settings
    except configparser.NoOptionError as e:
        logging.error(f"Missing required option in config file: {e}")
        return None

# --- Serial Number Handling ---
def get_mb_sn(file_path: str) -> str | None:
    """
    Reads the Motherboard Serial Number from the specified file.
    """
    sn_path = Path(file_path)
    if not sn_path.is_file():
        logging.error(f"SN file '{file_path}' not found.")
        return None
    
    try:
        sn = sn_path.read_text().strip()
        if not sn:
            logging.error(f"SN file '{file_path}' is empty.")
            return None
        logging.info(f"Successfully read SN: {sn}")
        return sn
    except Exception as e:
        logging.error(f"Error reading SN file '{file_path}': {e}")
        return None

# --- Error Handling (GUI or CLI) ---
def show_error_and_exit(message: str):
    """Displays an error message (GUI box or Console) and exits with a non-zero code."""
    logging.error(f"Displaying error and exiting: {message}")
    
    if USE_GUI:
        try:
            root = tk.Tk()
            root.withdraw() # Hide the main window
            messagebox.showerror("Connection Failed", message)
        except Exception as e:
            logging.error(f"Failed to create GUI window: {e}")
            print(f"\n[ERROR] {message}\n", file=sys.stderr)
    else:
        # 非 Windows 或無 GUI 環境，僅輸出到 Console
        print(f"\n[ERROR] {message}\n", file=sys.stderr)

    sys.exit(1) # Exit with a non-zero code

# --- Main Application Logic ---
def main():
    """Main function to run the MES tool.""" # 1. Load configuration
    version = get_executable_version()
    logging.info(f"--- MES Tool Version: {version} ---")
    
    if not USE_GUI:
        logging.info("Running in CLI mode (No GUI).")
    else:
        logging.info("Running in GUI mode.")

    config = load_config()    
    if not config:
        show_error_and_exit("Failed to load configuration, please check the log.")
    # 1-1. Read Serial Number
    mb_sn = get_mb_sn(config['mb_sn_path'])
    if not mb_sn:
        show_error_and_exit("Failed to load SN configuration, please check the log.")

    # 2. Construct API URL and attempt to connect
    api_url = f"{config['mes_server'].rstrip('/')}/{config['mes_api'].lstrip('/')}{mb_sn}"
    logging.info(f"Preparing to connect to MES API: {api_url}")

    response = None
    for attempt in range(config['retry_count']):
        try:
            logging.info(f"Connection attempt {attempt + 1}/{config['retry_count']}...")
            response = requests.get(api_url, timeout=config['request_timeout'])
            # Log status code and raw text for every attempt for debugging
            logging.debug(f"Response Status: {response.status_code}, Response Body: {response.text}")
            
            if response.status_code == 200:
                logging.info("Successfully retrieved information from MES (HTTP 200 OK).")
                try:
                    data = response.json()
                    # logging.info(data) # 視需求開啟詳細 JSON log
                    is_successful = data.get('success') # This will be True, False, or None
                    if is_successful is True:
                        logging.info(f"MES business logic success ('success': {is_successful}).")
                        break # Real success, exit the loop
                    else:
                        # HTTP status code is 200, but MES business logic returned an error code
                        error_message = data.get('message', 'No message provided.')
                        logging.warning(f"MES business logic failed ('success': {is_successful}). Message: {error_message}")
                        response = None # Mark as failed to trigger retry
                except requests.exceptions.JSONDecodeError:
                    logging.error("Failed to parse MES response as JSON.")
                    response = None # Mark as failed to trigger retry
            else:
                logging.warning(f"Connection failed, status code: {response.status_code}. Response: {response.text}")
                response = None # Mark as failed
        
        except requests.exceptions.RequestException as e:
            logging.error(f"An HTTP Request exception occurred: {e}")
            response = None # Mark as failed

        # If not the last attempt, wait and retry
        if response is None and attempt < config['retry_count'] - 1:
            logging.info(f"Retrying in {config['retry_delay']} seconds...")
            time.sleep(config['retry_delay'])

    # 3. Check final connection result
    if response is None:
        show_error_and_exit(f"Could not connect to MES system.\nURL: {api_url}\nPlease check the network connection or contact IT personnel.")

    # 4. Generate file with the received information
    output_file_path = get_resource_path(config['output_path'])
    try:
        # Ensure the output directory exists before writing the file
        output_file_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_file_path, 'w', encoding='utf-8') as f:
            f.write(response.text)
        logging.info(f"Successfully wrote MES information to '{output_file_path}'.")
    except IOError as e:
        logging.error(f"Failed to write to file '{output_file_path}': {e}")
        show_error_and_exit(f"Could not write to output file '{output_file_path}'.")

    logging.info("Tool execution finished.")
    sys.exit(0) # Exit with success code


if __name__ == '__main__':
    # We must configure logging before making any calls to the logger.
    # We'll create a temporary config object just to get the log path.
    config_file_path = get_resource_path('config.ini')
    temp_config = configparser.ConfigParser()
    temp_config.read(config_file_path, encoding='utf-8')
    log_path = temp_config.get('Global', 'LOG_PATH', fallback='./log/').strip('"\' ')

    log_dir = get_resource_path(log_path)
    try:
        log_dir.mkdir(parents=True, exist_ok=True)  # Ensure log directory exists
    except Exception as e:
        # Fallback if cannot create log dir (e.g. permission issue)
        print(f"Warning: Could not create log directory {log_dir}: {e}")
    
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"debug_{timestamp}.log"
    setup_logging(str(log_file))

    try:
        main()
    except KeyboardInterrupt:
        logging.warning("Program interrupted by user (Ctrl+C). Exiting gracefully.")
        sys.exit(130) # Standard exit code for command-line interruption
