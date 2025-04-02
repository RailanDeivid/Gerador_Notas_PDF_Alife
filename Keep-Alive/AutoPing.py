import requests
import threading
import time

def keep_awake():
    while True:
        try:
            requests.get("https://geradornotas.streamlit.app/")  
        except Exception as e:
            print(f"Erro ao tentar manter ativo: {e}")
        time.sleep(600)  

threading.Thread(target=keep_awake, daemon=True).start()
