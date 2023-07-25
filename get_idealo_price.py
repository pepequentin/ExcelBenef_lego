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



def scrape_idealo():
    user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36'

    # En-têtes de la requête
    headers = {
    'User-Agent': user_agent
    }
    url = "https://www.idealo.fr/cat/9552F774905oE0oJ4/lego.html"
    # Faire une requête GET pour obtenir le contenu HTML de la page
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        html_source_code = response.text
        soup = BeautifulSoup(html_source_code, "html.parser")

        # Trouver la balise qui contient les résultats
        result_list = soup.find("div", class_="sr-resultList resultList--GRID")

        if result_list:
            # Trouver tous les éléments de la liste
            items = result_list.find_all("div", class_="sr-resultList__item")

            data = []  # Liste pour stocker les informations des éléments

            for item in items:
                # Trouver le lien de l'élément
                link = item.find("div", class_="sr-resultItemLink sr-resultItemTile__link")
                if link:
                    link_url = link.a["href"]

                # Trouver le prix de l'élément
                price = item.find("div", class_="sr-detailedPriceInfo__price")
                if price:
                    # Utiliser une expression régulière pour extraire le prix du texte
                    prix_pattern = re.compile(r'\d+,\d+')
                    prix_trouve = prix_pattern.search(price.text)
                    if prix_trouve:
                        prix = prix_trouve.group()

                # Trouver le pourcentage dans la balise de réduction
                saving_badge = item.find("div", class_="sr-bargainBadge__savingBadge")
                if saving_badge:
                    saving_percentage = saving_badge.text.strip()
                    # Convertir le pourcentage en nombre pour le tri
                    saving_percentage_value = float(saving_percentage[:-1]) if saving_percentage[-1] == "%" else 0

                # Ajouter les informations dans la liste des données
                data.append((link_url.strip(), prix.strip(), saving_percentage_value))

            # Tri des données en fonction du pourcentage de réduction
            sorted_data = sorted(data, key=lambda x: x[2], reverse=True)

            for item in sorted_data:
                print(f"Lien : {item[0]:<110} Prix : {item[1]:<7} €    Réduction : {item[2]:.2f}%")
        else:
            print("La balise contenant les résultats n'a pas été trouvée.")
    else:
        print("Impossible d'accéder à la page web.")

# Appeler la fonction pour récupérer les informations depuis la page web
scrape_idealo()