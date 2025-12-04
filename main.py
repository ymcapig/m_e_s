import configparser
import logging
import sys
import time
import json
from pathlib import Path
from datetime import datetime  # [NEW] 新增 datetime 引用

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
    # 非 Windows 環境 (Linux/PXE) 直接使用 CLI 模式
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
        base_path = Path(sys.executable).parent
    else:
        base_path = Path(__file__).parent

    return base_path / relative_path

# --- Logger Setup ---
def setup_logging(log_file_path: str):
    logging.basicConfig(
        level=logging.DEBUG,
        format='[%(asctime)s] [%(levelname)s] %(message)s',
        handlers=[logging.FileHandler(log_file_path, mode='w', encoding='utf-8'),
                  logging.StreamHandler(sys.stdout)]
    )

# --- Configuration Handling ---
def load_config():
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
            'template_path': config.get('Global', 'TEMPLATE_PATH', fallback='mes_template.txt').strip('"\' '),
            'output_path': config.get('Global', 'OUTPUT_PATH', fallback='MES.txt').strip('"\' '),
            'raw_output_path': config.get('Global', 'RAW_OUTPUT_PATH', fallback='MES_Raw.json').strip('"\' '),
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

# --- Template Processing Logic ---
def process_mes_template(template_path: Path, mes_data: dict) -> list[str]:
    """
    Reads the template, inserts a timestamp at the first line,
    fills in matching keys from mes_data,
    and inserts unused keys before the final '##' marker if it exists.
    """
    lines = []
    
    if template_path.is_file():
        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
        except Exception as e:
            logging.error(f"Error reading template file: {e}")
    else:
        logging.warning(f"Template file '{template_path}' not found. Generating default key-value list.")

    remaining_keys = list(mes_data.keys())
    new_content = []

    # [NEW] 1. 產生並插入 Timestamp (格式: YYYY-MM-DD HH:MM:SS.ff)
    now = datetime.now()
    # strftime('%f') 取得微秒 (如 345678)，取前兩位 [:2] 變成 34
    timestamp_str = now.strftime("%Y-%m-%d %H:%M:%S") + f".{now.strftime('%f')[:2]}"
    new_content.append(f"{timestamp_str}\n") # 插入為第一行

    # 2. 遍歷模板填入數值
    for line in lines:
        line_stripped = line.strip()
        if not line.endswith('\n'):
            line += '\n'

        if "##" in line_stripped and ("=" in line_stripped or ":" in line_stripped):
            sep = "=" if "=" in line_stripped else ":"
            try:
                start_idx = line.find("##")
                sep_idx = line.find(sep, start_idx)
                
                if sep_idx > start_idx:
                    key_in_template = line[start_idx+2:sep_idx].strip()
                    
                    if key_in_template in mes_data:
                        value = mes_data[key_in_template]
                        prefix = line[:sep_idx+1]
                        new_line = f"{prefix}{value}\n"
                        new_content.append(new_line)
                        
                        if key_in_template in remaining_keys:
                            remaining_keys.remove(key_in_template)
                        continue
            except Exception as e:
                logging.warning(f"Failed to parse template line '{line_stripped}': {e}")

        new_content.append(line)

    # 3. 處理新增的 Key (Template 裡沒有的)
    if remaining_keys:
        new_keys_content = []
        for key in remaining_keys:
            new_keys_content.append(f"##{key}={mes_data[key]}\n")

        # 檢查最後一行是否為 "##" 標記
        if new_content and new_content[-1].strip() == "##":
            last_line = new_content.pop() # 取出最後一行
            new_content.extend(new_keys_content) # 加入新 Key
            new_content.append(last_line) # 放回最後一行
        else:
            new_content.extend(new_keys_content)

    return new_content

# --- Error Handling ---
def show_error_and_exit(message: str):
    logging.error(f"Displaying error and exiting: {message}")
    if USE_GUI:
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Connection Failed", message)
        except Exception as e:
            print(f"\n[ERROR] {message}\n", file=sys.stderr)
    else:
        print(f"\n[ERROR] {message}\n", file=sys.stderr)
    sys.exit(1)

# --- Main Application Logic ---
def main():
    version = get_executable_version()
    logging.info(f"--- MES Tool Version: {version} ---")
    
    if not USE_GUI:
        logging.info("Running in CLI mode (No GUI).")
    else:
        logging.info("Running in GUI mode.")

    config = load_config()    
    if not config:
        show_error_and_exit("Failed to load configuration.")

    mb_sn = get_mb_sn(config['mb_sn_path'])
    if not mb_sn:
        show_error_and_exit("Failed to load SN configuration.")

    api_url = f"{config['mes_server'].rstrip('/')}/{config['mes_api'].lstrip('/')}{mb_sn}"
    logging.info(f"Preparing to connect to MES API: {api_url}")

    response = None
    mes_data_content = None

    for attempt in range(config['retry_count']):
        try:
            logging.info(f"Connection attempt {attempt + 1}/{config['retry_count']}...")
            response = requests.get(api_url, timeout=config['request_timeout'])
            logging.debug(f"Response Status: {response.status_code}")
            
            if response.status_code == 200:
                logging.info("Successfully retrieved information from MES (HTTP 200 OK).")
                try:
                    resp_json = response.json()
                    is_successful = resp_json.get('success')
                    
                    if is_successful is True:
                        logging.info(f"MES business logic success ('success': {is_successful}).")
                        
                        data_dict = resp_json.get('data', {})
                        template_path = get_resource_path(config['template_path'])
                        logging.info(f"Processing template: {template_path}")
                        
                        mes_data_content = process_mes_template(template_path, data_dict)
                        
                        break
                    else:
                        error_message = resp_json.get('message', 'No message provided.')
                        logging.warning(f"MES business logic failed. Message: {error_message}")
                        response = None
                except requests.exceptions.JSONDecodeError:
                    logging.error("Failed to parse MES response as JSON.")
                    response = None
            else:
                logging.warning(f"Connection failed, status code: {response.status_code}.")
                response = None
        
        except requests.exceptions.RequestException as e:
            logging.error(f"An HTTP Request exception occurred: {e}")
            response = None

        if response is None and attempt < config['retry_count'] - 1:
            logging.info(f"Retrying in {config['retry_delay']} seconds...")
            time.sleep(config['retry_delay'])

    if mes_data_content is None or response is None:
        show_error_and_exit(f"Could not connect to MES system or retrieve valid data.\nURL: {api_url}")

    # 4-1. Generate PROCESSED file
    output_file_path = get_resource_path(config['output_path'])
    try:
        output_file_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_file_path, 'w', encoding='utf-8') as f:
            f.writelines(mes_data_content)
        logging.info(f"Successfully wrote Processed MES information to '{output_file_path}'.")
    except IOError as e:
        logging.error(f"Failed to write to file '{output_file_path}': {e}")
        show_error_and_exit(f"Could not write to output file '{output_file_path}'.")

    # 4-2. Generate RAW JSON file
    raw_output_path = get_resource_path(config['raw_output_path'])
    try:
        raw_output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(raw_output_path, 'w', encoding='utf-8') as f:
            try:
                json.dump(response.json(), f, ensure_ascii=False, indent=4)
            except:
                f.write(response.text)
        logging.info(f"Successfully wrote Raw JSON information to '{raw_output_path}'.")
    except IOError as e:
        logging.error(f"Failed to write to raw file '{raw_output_path}': {e}")

    logging.info("Tool execution finished.")
    sys.exit(0)

if __name__ == '__main__':
    config_file_path = get_resource_path('config.ini')
    temp_config = configparser.ConfigParser()
    temp_config.read(config_file_path, encoding='utf-8')
    log_path = temp_config.get('Global', 'LOG_PATH', fallback='./log/').strip('"\' ')
    log_dir = get_resource_path(log_path)
    
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"Warning: Could not create log directory {log_dir}: {e}")
    
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"debug_{timestamp}.log"
    setup_logging(str(log_file))

    try:
        main()
    except KeyboardInterrupt:
        logging.warning("Program interrupted by user. Exiting.")
        sys.exit(130) 
