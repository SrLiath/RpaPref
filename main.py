import os
import webview
from config import ( get_system_theme, get_view_driver)
from api import MyJSAPI
from dotenv import load_dotenv
import schedule
import time
import threading
from datetime import datetime
load_dotenv()
page_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), 'Views'))
index_file = f'file://{os.path.join(page_dir, "index.html")}'
js_api = MyJSAPI()
task_file = 'task.tsk'
last_modified_time = os.path.getmtime(task_file)

def read_task_file():
    global last_modified_time
    current_modified_time = os.path.getmtime(task_file)
    
    if current_modified_time != last_modified_time:
        last_modified_time = current_modified_time
        with open(task_file, 'r') as file:
            lines = file.readlines()

        schedule.clear()

        for line in lines:
            path, dt_str = line.strip().split('||')
            dt = datetime.strptime(dt_str, '%d %H:%M')
            schedule_task(path, dt)

def schedule_task(path, dt):
    now = datetime.now()
    run_time = now.replace(day=dt.day, hour=dt.hour, minute=dt.minute, second=0, microsecond=0)
    
    if run_time < now:
        if run_time.month == 12:
            run_time = run_time.replace(year=run_time.year + 1, month=1)
        else:
            run_time = run_time.replace(month=run_time.month + 1)
    
    interval = (run_time - now).total_seconds()
    
    print(f'Agendando a execução da função para {run_time}')
    schedule.every(interval).seconds.do(execute_function, path)

def execute_function(path):

def task_loop():
    while True:
        read_task_file()
        schedule.run_pending()
        time.sleep(60)  # Espera 1 minuto entre verificações

task_thread = threading.Thread(target=task_loop)
task_thread.daemon = True  # Para que a thread se feche quando o programa principal terminar
task_thread.start()

def set_theme(theme):
    webview.windows[0].evaluate_js(f"window.onThemeChanged('{theme}')")
def set_config(on):
    webview.windows[0].evaluate_js(f"window.ver({on})")
    
def setup(window):
    theme = get_system_theme()
    switch = get_view_driver()
    set_theme(theme)
    set_config(switch)


def start_window():
    window = webview.create_window('Polibot', index_file, js_api=js_api, frameless=True)
    webview.start(setup, window, debug=True)

if __name__ == '__main__':
    start_window()
