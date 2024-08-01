import os
from pathlib import Path
import ctypes
from dotenv import load_dotenv


def get_env_from_file():
    env_file_path = Path(__file__).parent / '.env'
    
    if not env_file_path.exists():
        raise FileNotFoundError(f".env file not found at {env_file_path}")

    env_variables = {}
    with env_file_path.open('r', encoding='utf-8') as file:
        for line in file:
            line = line.strip()
            if line and not line.startswith('#'):
                key, *value = line.split('=')
                if key:
                    env_variables[key.strip()] = '='.join(value).strip()
                    
    return env_variables



def update_env_file(updates):
    env_file_path = Path(__file__).parent / '.env'
    
    if not env_file_path.exists():
        raise FileNotFoundError(f".env file not found at {env_file_path}")

    env_variables = {}
    with env_file_path.open('r', encoding='utf-8') as file:
        for line in file:
            line = line.strip()
            if line and not line.startswith('#'):
                key, *value = line.split('=')
                if key:
                    env_variables[key.strip()] = '='.join(value).strip()

    # Atualiza com novos valores
    for key, value in updates.items():
        env_variables[key] = value

    # Escreve as variáveis atualizadas de volta no arquivo
    new_content = '\n'.join(f"{key}={value}" for key, value in env_variables.items())
    
    with env_file_path.open('w', encoding='utf-8') as file:
        file.write(new_content + '\n')

def get_system_theme():
    try:
        # Abre a chave do registro
        user32 = ctypes.windll.user32
        user32.SetProcessDPIAware()
        
        key = ctypes.c_void_p()
        if ctypes.windll.advapi32.RegOpenKeyExW(0x80000001, 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize', 0, 0x20019, ctypes.byref(key)) == 0:
            # Lê o valor AppsUseLightTheme
            data = ctypes.c_ulong()
            data_size = ctypes.c_ulong(4)
            if ctypes.windll.advapi32.RegQueryValueExW(key, 'AppsUseLightTheme', 0, None, ctypes.byref(data), ctypes.byref(data_size)) == 0:
                return 'light' if data.value == 1 else 'dark'
            ctypes.windll.advapi32.RegCloseKey(key)
    except Exception as e:
        print(f"Erro ao verificar o tema do sistema: {e}")
    return 'light'

def get_view_driver():
    load_dotenv()       
    return os.getenv('Ver')
