from math import nan
import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import os
from datetime import date, timedelta
from plyer import notification
import webbrowser
import time
import pyglet
import pyglet.media as media
import random
import openpyxl
from openpyxl.styles import PatternFill, Color
import xlwings as xw

c3po_ico = 'c3po.ico'
sound_to_play = "R2D2-hey-you.wav"
##
#   Fonction pour gérer l'événement de fin de lecture du son
##
def on_player_eos():
    pyglet.app.exit()

##
#   Fonction pour jouer le son en arrière-plan
##
def play_sound():
    src = media.load(sound_to_play)
    player = media.Player()
    player.queue(src)
    player.volume = 1.0
    player.play()

    # Attacher la fonction on_player_eos à l'événement on_eos
    player.push_handlers(on_eos=on_player_eos)
    try:
        pyglet.app.run()
    except KeyboardInterrupt:
        player.next()

##
#   Fonction pour ouvrir une page web et jouer un son en arrière-plan
##
def open_webpage(link):
    # Ouvrir la page web
    webbrowser.open(link)
    # Jouer le son en arrière-plan
    play_sound()

##
#   Fonction pour afficher la notification avec un lien cliquable
##
def show_notification(title, message):
    notification.notify(title=title, message=message, app_icon=c3po_ico, timeout=None)

##
#   Fonction pour charger les liens depuis un fichier
##
def load_links(file_path):
    if os.path.exists(file_path):
        with open(file_path, "r") as file:
            links = file.read().splitlines()
            return links
    else:
        return []

##
#   Fonction pour sauvegarder les liens dans un fichier
##
def save_links(file_path, links):
    with open(file_path, "w") as file:
        for link in links:
            file.write(link + "\n")

##
#   Fonction pour trouver le fichier du jour le plus proche
##
def find_closest_file_idealo():
    today = date.today()
    closest_file = None
    closest_diff = timedelta(days=365)  # Valeur initiale définie à un an

    for file_name in os.listdir("liens"):
        if file_name.endswith("_idealo_links.txt"):
            file_date_str = file_name[:10]
            file_date = date.fromisoformat(file_date_str)
            diff = abs(today - file_date)
            if diff < closest_diff:
                closest_diff = diff
                closest_file = file_name

    return closest_file


##
#   Fonction pour trouver le fichier du jour le plus proche
##
def find_closest_file_german():
    today = date.today()
    closest_file = None
    closest_diff = timedelta(days=365)  # Valeur initiale définie à un an

    for file_name in os.listdir("liens"):
        if file_name.endswith("_german_links.txt"):
            file_date_str = file_name[:10]
            file_date = date.fromisoformat(file_date_str)
            diff = abs(today - file_date)
            if diff < closest_diff:
                closest_diff = diff
                closest_file = file_name

    return closest_file


##
#   Fonction pour définir la couleur d'une cellule dans le fichier Excel
##
def set_cell_color(ws, row, column, color):
    cell = ws.cell(row=row, column=column)
    cell.fill = PatternFill(start_color=color, end_color=color, fill_type='solid')


##
#   Utilisation du script avec le fichier 'Achat_lego.xlsx'
##
def check_prices(file_path):
    df = pd.read_excel(file_path)
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    vente_total = 0
    for index, row in df.iterrows():
        lien = row[1]
        prix_achat = row[6]
        nb_exemplaires = row[12]

        if pd.notna(lien) and isinstance(lien, str) and pd.notna(prix_achat) and nb_exemplaires > 0:
            user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36'
            headers = {'User-Agent': user_agent}
            response = requests.get(lien, headers=headers)
            html_source_code = response.text
            soup = BeautifulSoup(html_source_code, "html.parser")
            html_span = soup.find_all('span', {'class': 'oopStage-priceRangePrice'})
            pattern = re.compile(r'<span class="oopStage-conditionButton-wrapper-text-price-prefix">\s*\(non disponible\)\s*</span>')
            html_source_code = response.text
            html_span_for_no_news = pattern.findall(html_source_code)
            prix_pattern = re.compile(r'\d+,\d+')
            prix_achat_par_exemplaire = prix_achat / nb_exemplaires
            # Condition pour les lego qui non pas de prix idealo
            if not html_span:
                prix_trouve = row[8]
                if prix_trouve:
                    prix = prix_trouve
                    if prix:
                        if prix_achat != 0:
                            pourcentage_benef = ((prix - prix_achat_par_exemplaire) / prix_achat_par_exemplaire) * 100
                        else:
                            pourcentage_benef = ((prix - prix_achat_par_exemplaire) / 1) * 100


                        # Prix mul by the number of product in a temp var
                        tmp_price = prix * nb_exemplaires
                        vente_total += tmp_price
                        df.at[index, 'Prix actuel idéalo'] = f"{prix:.2f}"
                        df.at[index, 'Bénéfice potentiel'] = f"{pourcentage_benef * nb_exemplaires:.2f}"

                        # Appliquer le formatage des couleurs en fonction des valeurs de 'Bénéfice potentiel'
                        if pd.notna(row['Bénéfice potentiel']):
                            pourcentage_benef = float(row['Bénéfice potentiel'][:-1])
                            if pourcentage_benef > 0:
                                ws.cell(row=index + 2, column=11).fill = openpyxl.styles.PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                            elif pourcentage_benef < 0:
                                ws.cell(row=index + 2, column=11).fill = openpyxl.styles.PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                            else:
                                ws.cell(row=index + 2, column=11).fill = openpyxl.styles.PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
            else:
                for span in html_span:
                    if html_span_for_no_news:
                        prix_trouve = prix_pattern.search(span.text)
                        if prix_trouve:
                            prix = float(prix_trouve.group().replace(",", "."))
                            if prix:
                                prix *= 4
                                if prix_achat != 0:
                                    pourcentage_benef = ((prix - prix_achat_par_exemplaire) / prix_achat_par_exemplaire) * 100
                                else:
                                    pourcentage_benef = ((prix - prix_achat_par_exemplaire) / 1) * 100


                                # Prix mul by the number of product in a temp var
                                tmp_price = prix * nb_exemplaires
                                vente_total += tmp_price
                                df.at[index, 'Prix actuel idéalo'] = f"{prix:.2f}"
                                df.at[index, 'Bénéfice potentiel'] = f"{pourcentage_benef * nb_exemplaires:.2f}"

                                # Appliquer le formatage des couleurs en fonction des valeurs de 'Bénéfice potentiel'
                                if pd.notna(row['Bénéfice potentiel']):
                                    pourcentage_benef = float(row['Bénéfice potentiel'][:-1])
                                    if pourcentage_benef > 0:
                                        ws.cell(row=index + 2, column=11).fill = openpyxl.styles.PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                                    elif pourcentage_benef < 0:
                                        ws.cell(row=index + 2, column=11).fill = openpyxl.styles.PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                                    else:
                                        ws.cell(row=index + 2, column=11).fill = openpyxl.styles.PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")

                    else:
                        prix_trouve = prix_pattern.search(span.text)
                        if prix_trouve:
                            prix = float(prix_trouve.group().replace(",", "."))
                            if prix:
                                if prix_achat != 0:
                                    pourcentage_benef = ((prix - prix_achat_par_exemplaire) / prix_achat_par_exemplaire) * 100
                                else:
                                    pourcentage_benef = ((prix - prix_achat_par_exemplaire) / 1) * 100


                                # Prix mul by the number of product in a temp var
                                tmp_price = prix * nb_exemplaires
                                vente_total += tmp_price
                                df.at[index, 'Prix actuel idéalo'] = f"{prix:.2f}"
                                df.at[index, 'Bénéfice potentiel'] = f"{pourcentage_benef * nb_exemplaires:.2f}"

                                # Appliquer le formatage des couleurs en fonction des valeurs de 'Bénéfice potentiel'
                                if pd.notna(row['Bénéfice potentiel']):
                                    pourcentage_benef = float(row['Bénéfice potentiel'][:-1])
                                    if pourcentage_benef > 0:
                                        ws.cell(row=index + 2, column=11).fill = openpyxl.styles.PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                                    elif pourcentage_benef < 0:
                                        ws.cell(row=index + 2, column=11).fill = openpyxl.styles.PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                                    else:
                                        ws.cell(row=index + 2, column=11).fill = openpyxl.styles.PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")

    final_output_file = 'Achat_lego_updated.xlsx'
    wb.save(final_output_file)
    df.to_excel(final_output_file, index=False)

    # Charger le classeur pour appliquer le formatage des couleurs
    new_wb = openpyxl.load_workbook(final_output_file)
    ws = new_wb.active

    # Appliquer le formatage des couleurs en fonction des valeurs de 'Bénéfice potentiel'
    for index, row in df.iterrows():
        if pd.notna(row['Bénéfice potentiel']):
            pourcentage_benef = float(row['Bénéfice potentiel'][:-1])
            if pourcentage_benef > 0:
                ws.cell(row=index + 2, column=11).fill = openpyxl.styles.PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
            elif pourcentage_benef < 0:
                ws.cell(row=index + 2, column=11).fill = openpyxl.styles.PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            else:
                ws.cell(row=index + 2, column=11).fill = openpyxl.styles.PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")

    # Calculer le coût total
    cout_total = df['Prix d\'achat'].sum()
    ws.append([])
    ws.append([None, None, "Cout total", cout_total])

    # Calculer la vente total
    ws.append([None, None, "Vente total", vente_total])

    # Calculer le potentiel bénéficiaire
    ws.append([None, None, "Potentiel benef", vente_total - cout_total])


    final_output_file = 'Achat_lego_updated.xlsx'
    new_wb.save(final_output_file)

    # Ouvrir les fichiers Excel
    wb1 = openpyxl.load_workbook(file_path)
    wb2 = openpyxl.load_workbook('Achat_lego_updated.xlsx')

    # Sélectionner les feuilles actives des deux fichiers
    ws1 = wb1.active
    ws2 = wb2.active

    # Spécifier les colonnes à copier
    col_to_copy = 'H'

    for row_num, (row1, row2) in enumerate(zip(ws1.iter_rows(min_row=2, values_only=True), ws2.iter_rows(min_row=2, values_only=True)), start=2):
        color1 = ws1[f'{col_to_copy}{row_num}'].fill.start_color
        ws2[f'{col_to_copy}{row_num}'].fill = openpyxl.styles.PatternFill(start_color=color1, end_color=color1, fill_type="solid")
   
    col_to_copy = 'Q'

    for row_num, (row1, row2) in enumerate(zip(ws1.iter_rows(min_row=2, values_only=True), ws2.iter_rows(min_row=2, max_row=5, values_only=True)), start=2):
        color1 = ws1[f'{col_to_copy}{row_num}'].fill.start_color
        ws2[f'{col_to_copy}{row_num}'].fill = openpyxl.styles.PatternFill(start_color=color1, end_color=color1, fill_type="solid")

    # Copier la largeur de colonne du fichier1 sur le fichier2
    for column in ws1.column_dimensions:
        ws2.column_dimensions[column].width = ws1.column_dimensions[column].width

    # Sauvegarder les modifications dans le fichier2
    wb2.save('Achat_lego_updated.xlsx')
    print("Données mises à jour et sauvegardées dans", final_output_file)

##
#   Fonction principale pour récupérer les informations depuis la page web
##
def scrape_idealo(url):
    # User agent pour la requête HTTP
    user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36'

    # En-têtes de la requête
    headers = {
        'User-Agent': user_agent
    }

    # Faire une requête GET pour obtenir le contenu HTML de la page
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        # Obtenir le contenu HTML de la page
        html_source_code = response.text

        # Parser le contenu HTML avec BeautifulSoup
        soup = BeautifulSoup(html_source_code, "html.parser")

        # Trouver la balise qui contient les résultats
        result_list = soup.find("div", class_="offerList wishlist offerList--tileview")

        if result_list:
            # Trouver tous les éléments de la liste
            items = result_list.find_all("div", class_="offerList-item")

            # Liste pour stocker les informations des éléments
            data = []

            for item in items:
                # Trouver le lien de l'élément
                link = item.find("a", class_="offerList-itemWrapper")["href"]
                link_temp = "https://www.idealo.fr"
                link_temp += link
                link = link_temp
                # Trouver le prix de l'élément
                price = item.find("div", class_="offerList-item-priceMin")
                if price:
                    # Utiliser une expression régulière pour extraire le prix du texte
                    prix_pattern = re.compile(r'\d+,\d+')
                    prix_trouve = prix_pattern.search(price.text)
                    if prix_trouve:
                        prix = prix_trouve.group()

                # Trouver le pourcentage de réduction
                saving_badge = item.find("div", class_="sales-badge")
                if saving_badge:
                    saving_percentage = saving_badge.text.strip()
                    # Convertir le pourcentage en nombre pour le tri
                    saving_percentage_value = float(saving_percentage[:-1]) if saving_percentage[-1] == "%" else 0

                # Ajouter les informations dans la liste des données
                data.append((link, prix.strip(), saving_percentage_value))
            # Tri des données en fonction du pourcentage de réduction
            sorted_data = sorted(data, key=lambda x: x[2], reverse=True)

            # Obtenir la date du jour au format YYYY-MM-DD
            today = date.today().strftime("%Y-%m-%d")

            # Trouver le fichier du jour le plus proche
            closest_file = find_closest_file_idealo()

            if closest_file:
                # Charger les liens depuis le fichier du jour le plus proche
                existing_links = load_links(os.path.join("liens", closest_file))
            else:
                existing_links = []

            # Vérifier les nouveaux liens et les mettre en "vert" lors de l'affichage
            new_links = []
            # pop_up_message = ""
            for item in sorted_data:
                link_url = item[0]
                if link_url not in existing_links:
                    new_links.append(link_url)
                    # pop_up_message = f"Lien: {link_url:<130} Prix: {item[1]:<7} €    Réduction: {item[2]:.2f}%\n"
                    print("\033[92m" + f"Lien: {link_url:<130} Prix: {item[1]:<7} €    Réduction: {item[2]:.2f}%" + "\033[0m")
                    # show_notification("Nouvelle offre à vérifier", pop_up_message)
                    open_webpage(link_url)

            # Sauvegarder les liens mis à jour dans le fichier du jour
            updated_links = existing_links + new_links
            save_links(os.path.join("liens", today + "_idealo_links.txt"), updated_links)
        else:
            # Trouver la balise qui contient les résultats
            result_list = soup.find("div", class_="sr-resultList resultList--GRID")

            if result_list:
                # Trouver tous les éléments de la liste
                items = result_list.find_all("div", class_="sr-resultList__item")

                # Liste pour stocker les informations des éléments
                data = []

                for item in items:
                    # Trouver le lien de l'élément
                    link = item.find("div", class_="sr-resultItemTile__link")
                    if link:
                        link_url = link.a["href"]

                    # Trouver le prix de l'élément
                    price = item.find("div", class_="sr-detailedPriceInfo__price")
                    if price:
                        # Utiliser une expression régulière pour extraire le prix du texte
                        prix_pattern = re.compile(r'\d+,\d+')
                        prix_trouve = prix_pattern.search(price.text)
                        if prix_trouve:
                            prix = float(prix_trouve.group().replace(",", "."))

                    # Trouver le pourcentage dans la balise de réduction
                    saving_badge = item.find("div", class_="sr-bargainBadge__savingBadge")
                    if saving_badge:
                        saving_percentage = saving_badge.text.strip()
                        # Convertir le pourcentage en nombre pour le tri
                        saving_percentage_value = float(saving_percentage[:-1]) if saving_percentage[-1] == "%" else 0

                    # Ajouter les informations dans la liste des données
                data.append((link_url.strip(), prix, saving_percentage_value))
                # Tri des données en fonction du pourcentage de réduction
                sorted_data = sorted(data, key=lambda x: x[2], reverse=True)

                # Obtenir la date du jour au format YYYY-MM-DD
                today = date.today().strftime("%Y-%m-%d")

                # Trouver le fichier du jour le plus proche
                closest_file = find_closest_file_idealo()

                if closest_file:
                    # Charger les liens depuis le fichier du jour le plus proche
                    existing_links = load_links(os.path.join("liens", closest_file))
                else:
                    existing_links = []

                # Vérifier les nouveaux liens et les mettre en "vert" lors de l'affichage
                new_links = []
                # pop_up_message = ""
                for item in sorted_data:
                    link_url = item[0]
                    if link_url not in existing_links:
                        new_links.append(link_url)
                        # pop_up_message = f"Lien: {link_url:<130} Prix: {item[1]:<7} €    Réduction: {item[2]:.2f}%\n"
                        print("\033[92m" + f"Lien: {link_url:<130} Prix: {item[1]:<7} €    Réduction: {item[2]:.2f}%" + "\033[0m")
                        # show_notification("Nouvelle offre à vérifier", pop_up_message)
                        open_webpage(link_url)

                # Sauvegarder les liens mis à jour dans le fichier du jour
                updated_links = existing_links + new_links
                save_links(os.path.join("liens", today + "_idealo_links.txt"), updated_links)
            else:
                print("V1 et V2 La balise contenant les résultats n'a pas été trouvée.")
    else:
        print("Impossible d'accéder à la page web.")




def scrape_german(url):
    user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36'

    headers = {
        'User-Agent': user_agent
    }

    response = requests.get(url, headers=headers)

    soup = BeautifulSoup(response.content, 'html.parser')

    # Trouver toutes les balises 'div' avec la classe 'productbox'
    product_boxes = soup.find_all('div', class_='productbox')

    # Liste pour stocker les informations des produits
    data = []

    for product_box in product_boxes:
        # Vérifier si la balise 'div' avec la classe 'ribbon' contient le texte 'Ausverkauft'
        ribbon = product_box.find('div', class_='ribbon')
        if ribbon and 'Ausverkauft' in ribbon.text:
            # Le produit est épuisé, nous le sautons
            continue

        # Extraire le titre du produit
        title = product_box.find('a', class_='text-clamp-2').text.strip()

        # Extraire l'URL du produit
        product_url = product_box.find('a', class_='text-clamp-2')['href']

        # Extraire le prix du produit et enlever le symbole "*" et le signe "€"
        price_with_symbol = product_box.find('div', class_='productbox-price').text.strip()
        price = price_with_symbol.replace(' *', '').replace(' €', '').replace(',', '.').replace(' ', '')

        # Ajouter les informations dans la liste des données
        data.append((title, product_url, price))

    # Tri des données en fonction du prix
    sorted_data = sorted(data, key=lambda x: float(x[2]))

    # Obtenir la date du jour au format YYYY-MM-DD
    today = date.today().strftime("%Y-%m-%d")

    # Trouver le fichier du jour le plus proche
    closest_file = find_closest_file_german()

    if closest_file:
        # Charger les liens depuis le fichier du jour le plus proche
        existing_links = load_links(os.path.join("liens", closest_file))
    else:
        existing_links = []

    # Vérifier les nouveaux liens et les mettre en "vert" lors de l'affichage
    new_links = []
    for item in sorted_data:
        link_url = item[1]
        if link_url not in existing_links:
            new_links.append(link_url)
            print("\033[92m" + f"Url: {item[1]:<130} Prix: {item[2]} €" + "\033[0m")
            open_webpage(link_url)

    # Sauvegarder les liens mis à jour dans le fichier du jour
    updated_links = existing_links + new_links
    save_links(os.path.join("liens", today + "_german_links.txt"), updated_links)



if __name__ == "__main__":
    # Utilisation du script avec le fichier 'Achat_lego.xlsx'
    check_prices('Achat_lego.xlsx')
    # Boucle pour appeler la fonction scrape_idealo() toutes les 2 minutes
    # while True:
    #     base_url = "https://spiel-und-modellbau.com/?suche=lego+star+wars&seite="
    #     for i in range(1, 3):  # Boucle de 1 à 8 inclus
    #         url_to_scrape = base_url + str(i)
    #         scrape_german(url_to_scrape)
    #     scrape_idealo("https://www.idealo.fr/cat/9552F774905oE0oJ4/lego.html")
    #     # Générer un nombre aléatoire entre 4 et 12 (exclus) pour time.sleep()
    #     random_sleep_time = random.randrange(20, 60)
    #     time.sleep(random_sleep_time)