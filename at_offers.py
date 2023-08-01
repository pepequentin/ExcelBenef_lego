import requests
from bs4 import BeautifulSoup

def check_page_response(url):
    user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36'

    headers = {
        'User-Agent': user_agent
    }

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        print(f"La page {url} répond avec succès (code d'état HTTP 200).")
        return True
    else:
        print(f"La page {url} a rencontré une erreur (code d'état HTTP {response.status_code}).")
        return False

if __name__ == "__main__":
    base_url = "https://spiel-und-modellbau.com/?suche=lego+star+wars&seite="
    page_number = 1
    max_page_number = 4  # Définir un nombre maximum de pages à visiter pour éviter une boucle infinie

    while page_number <= max_page_number:
        url = base_url + str(page_number)
        if not check_page_response(url):
            print(f"La recherche n'a pas besoin de continuer. Dernière page testée : {page_number}")
            break
        page_number += 1