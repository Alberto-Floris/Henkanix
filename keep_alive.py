import os
import requests

def run():
    url = os.getenv("APP_URL", "https://henkanix.streamlit.app/")
    
    try:
        print(f"Ping a {url}")
        response = requests.get(url, timeout=60)

        print("Status code:", response.status_code)

        if response.status_code == 200:
            print("App attiva o risvegliata con successo!")
        else:
            print("Risposta ricevuta, ma status diverso da 200")

    except Exception as e:
        print("Errore:", e)

if __name__ == "__main__":
    run()