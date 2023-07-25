import pandas as pd
import requests
from bs4 import BeautifulSoup
import re

def check_prices(file_path):
    df = pd.read_excel(file_path)
    for index, row in df.iterrows():
        if index < 2:
            continue
        lien = row[1]
        prix_achat = row[5]

        if pd.notna(lien) and isinstance(lien, str):
            # En-tête User-Agent
            user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36'

            # En-têtes de la requête
            headers = {
            'User-Agent': user_agent
            }
        
            # Faire une requête GET en utilisant le proxy et l'en-tête User-Agent
            response = requests.get(lien, headers=headers)

            # Get HTML source code
            html_source_code = response.text

            # Parsing HTML
            soup = BeautifulSoup(html_source_code, "html.parser")
            # Trouver la balise contenant le prix
            # prix_balise = soup.find("oopStage-priceRangePrice")
            html_span = soup.find_all('span', {'class' : 'oopStage-priceRangePrice'})

            # Utiliser une expression régulière pour extraire le prix
            prix_pattern = re.compile(r'\d+,\d+')

            for span in html_span:
                prix_trouve = prix_pattern.search(span.text)
                if prix_trouve:
                    prix = float(prix_trouve.group().replace(",", "."))
                    # Comparer le prix trouvé avec le prix d'achat
                    if prix:
                        # Now let's compare the found price with the purchase price
                        if prix_achat != 0:
                            pourcentage_benef = ((prix - prix_achat) / prix_achat) * 100
                        else:
                            pourcentage_benef = ((prix - prix_achat) / 1) * 100

                        print(f"Lien: {lien.strip():<130} Bénéfice: {pourcentage_benef:.2f}%")
                    else:
                        print("Prix non trouvé pour le lien", lien)
                else:
                    print("Prix non trouvé dans la balise.")

# Utilisation du script avec le fichier 'Achat_lego.xlsx'
check_prices('Achat_lego.xlsx')