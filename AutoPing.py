import time
import requests
url = "https://geradornotas.streamlit.app/"  
while True:
    try:
        response = requests.get(url)
        print(f"[{time.strftime('%H:%M:%S')}] App acessado! Status: {response.status_code}")

    except Exception as e:
        print(f"Erro ao acessar o app: {e}")
    time.sleep(10)  
