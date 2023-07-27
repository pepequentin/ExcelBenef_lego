from math import nan
import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import os
from datetime import date, timedelta
from win10toast import ToastNotifier
import time


# Fonction pour afficher la notification sans durée
def show_notification(title, message):
    toaster = ToastNotifier()
    toaster.show_toast(title, message, duration=None, threaded=True)

# Fonction pour charger les liens depuis un fichier
def load_links(file_path):
    if os.path.exists(file_path):
        with open(file_path, "r") as file:
            links = file.read().splitlines()
            return links
    else:
        return []

# Fonction pour sauvegarder les liens dans un fichier
def save_links(file_path, links):
    with open(file_path, "w") as file:
        for link in links:
            file.write(link + "\n")

# Fonction pour trouver le fichier du jour le plus proche
def find_closest_file():
    today = date.today()
    closest_file = None
    closest_diff = timedelta(days=365)  # Initial value set to one year

    for file_name in os.listdir("liens"):
        if file_name.endswith("_links.txt"):
            file_date_str = file_name[:10]
            file_date = date.fromisoformat(file_date_str)
            diff = abs(today - file_date)
            if diff < closest_diff:
                closest_diff = diff
                closest_file = file_name

    return closest_file


def set_cell_color(ws, row, column, color):
    cell = ws.cell(row=row, column=column)
    cell.fill = PatternFill(start_color=color, end_color=color, fill_type='solid')

def check_prices(file_path):
    df = pd.read_excel(file_path)
    for index, row in df.iterrows():


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

                        # Store the values in the DataFrame with '%' symbol
                        df.at[index, 'Prix actuel idéalo'] = f"{prix:.2f} €"
                        df.at[index, 'Dénéfice potentiel'] = f"{pourcentage_benef:.2f}%"

    # Save the updated DataFrame to a new Excel file
    output_file = 'Achat_lego_temp.xlsx'
    df.to_excel(output_file, index=False)

    # Load the workbook to apply color formatting
    wb = load_workbook(output_file)
    ws = wb.active

    # Apply color formatting based on 'Dénéfice potentiel' values
    for index, row in df.iterrows():
        if row['Dénéfice potentiel'] != nan:
            pourcentage_benef = row['Dénéfice potentiel']
            if pd.notna(pourcentage_benef) and isinstance(pourcentage_benef, str):  # Check if it's a string and not NaN
                pourcentage_benef = float(pourcentage_benef[:-1])  # Remove the '%' symbol and convert to float
                if pourcentage_benef > 0:
                    set_cell_color(ws, index+2, 10, '00FF00')  # Green
                elif pourcentage_benef < 0:
                    set_cell_color(ws, index+2, 10, 'FF0000')  # Red
                else:
                    set_cell_color(ws, index+2, 10, 'C0C0C0')  # Gray

    # Save the final Excel file with color formatting
    final_output_file = 'Achat_lego_updated.xlsx'
    wb.save(final_output_file)
    print("Data updated and saved to", final_output_file)

# Utilisation du script avec le fichier 'Achat_lego.xlsx'
check_prices('Achat_lego.xlsx')
print()
print()


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

            # Obtenir la date du jour au format YYYY-MM-DD
            today = date.today().strftime("%Y-%m-%d")

            # Trouver le fichier du jour le plus proche
            closest_file = find_closest_file()

            if closest_file:
                # Charger les liens depuis le fichier du jour le plus proche
                existing_links = load_links(os.path.join("liens", closest_file))
            else:
                existing_links = []

            # Vérifier les nouveaux liens et les mettre en "vert" lors de l'affichage
            new_links = []
            mail_to_send = ""
            for item in sorted_data:
                link_url = item[0].strip()
                if link_url not in existing_links:
                    new_links.append(link_url)
                    mail_to_send += "Lien: {link_url:<130} Prix: {item[1]:<7} €    Réduction: {item[2]:.2f}%\n"
                    print("\033[92m" + f"Lien: {link_url:<130} Prix: {item[1]:<7} €    Réduction: {item[2]:.2f}%" + "\033[0m")
                else:
                    print(f"Lien: {link_url:<130} Prix: {item[1]:<7} €    Réduction: {item[2]:.2f}%")
            notification_title = "Notification Title"

            show_notification(notification_title, mail_to_send)
            # Sauvegarder les liens mis à jour dans le fichier du jour
            updated_links = existing_links + new_links
            save_links(os.path.join("liens", today + "_links.txt"), updated_links)
        else:
            print("La balise contenant les résultats n'a pas été trouvée.")
    else:
        print("Impossible d'accéder à la page web.")
if __name__ == "__main__":
    # Boucle pour appeler la fonction scrape_idealo() toutes les 2 minutes
    while True:
        # Appeler la fonction pour récupérer les informations depuis la page web
        scrape_idealo()
        time.sleep(5)  # Attendre 2 minutes (120 secondes) avant de rappeler la fonction
