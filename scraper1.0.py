import requests
from bs4 import BeautifulSoup
import random
import time
import pandas as pd
from urllib.parse import urljoin
import sys
import subprocess
import os
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

visited_pages = []  # To keep track of visited pages with title and URL

def is_valid_wiki_link(href):
    """Check if the link is a valid Wikipedia article link."""
    if href is None:
        return False
    if href.startswith("/wiki/"):
        # Exclude links to non-article pages
        excluded_prefixes = ['/wiki/Special:', '/wiki/Help:', '/wiki/File:', '/wiki/Template:', '/wiki/Talk:',
                             '/wiki/Category:', '/wiki/Portal:', '/wiki/Main_Page', '/wiki/User:', '/wiki/Wikipedia:']
        return not any(href.startswith(prefix) for prefix in excluded_prefixes)
    return False

def scrapeWikiArticle(url):
    # Check if the URL has been visited
    if any(page['URL'] == url for page in visited_pages):
        return

    try:
        response = requests.get(url)
        response.raise_for_status()  # Ensure we raise an exception for HTTP errors
        soup = BeautifulSoup(response.content, 'html.parser')

        title_element = soup.find(id="firstHeading")
        if title_element:
            article_title = title_element.text.strip()
            print(f"Scraping: {article_title}")
            print(f"URL: {url}")

            # Store the title and URL
            visited_pages.append({'Title': article_title, 'URL': url})

            allLinks = soup.find(id="bodyContent").find_all("a")
            random.shuffle(allLinks)
            linkToScrape = None

            for link in allLinks:
                href = link.get('href', '')
                # We are only interested in valid wiki article links
                if is_valid_wiki_link(href):
                    linkToScrape = link
                    break

            # Check if we found a valid link to scrape
            if linkToScrape:
                next_url = urljoin("https://en.wikipedia.org", linkToScrape['href'])
                # Continue scraping the next URL
                scrapeWikiArticle(next_url)

    except requests.exceptions.RequestException as e:
        print(f"An error occurred: {e}")

    # Wait a few seconds before making the next request to avoid hitting the server too hard
    time.sleep(random.uniform(1, 3))

def get_random_wikipedia_page():
    """Fetch a random Wikipedia page URL."""
    try:
        response = requests.get("https://en.wikipedia.org/wiki/Special:Random")
        response.raise_for_status()
        return response.url
    except requests.exceptions.RequestException as e:
        print(f"Failed to get random Wikipedia page: {e}")
        return None

def save_to_files(filename_base='wiki_urls'):
    """Save the data to both CSV and Excel files, and adjust column widths in the Excel file."""
    # Convert the list of visited pages to a pandas DataFrame
    df = pd.DataFrame(visited_pages)
    
    # Sort the DataFrame by the 'Title' column in alphabetical order
    df_sorted = df.sort_values(by='Title')
    
    # Save to CSV
    csv_filename = f'{filename_base}.csv'
    df_sorted.to_csv(csv_filename, index=False)
    print(f"Data saved to {csv_filename}")
    
    # Save to Excel
    xlsx_filename = f'{filename_base}.xlsx'
    df_sorted.to_excel(xlsx_filename, index=False, engine='openpyxl')
    print(f"Data saved to {xlsx_filename}")
    
    # Adjust column widths for the Excel file
    wb = load_workbook(xlsx_filename)
    ws = wb.active
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter  # Get the column letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column_letter].width = adjusted_width
    wb.save(xlsx_filename)

def print_completion_message():
    print("\n\n\n")
    print("                                                                                                    ")
    print("                                                                                                    ")
    print("                                                                                                    ")
    print("                    ::-::--                                         =   ==---=                      ")
    print("                  :-::::-=-                    -:::--        -----------===----===                  ")
    print("               :::--:---=-                    -:::::-----------------------+========                ")
    print("             .:-==::----                     :=-::::::::::::::------------*+*+*#*====               ")
    print("            ::.:=....:=-                :::::..::::+-:-::::::::::::-------=#=**+*-====              ")
    print("           :...:-...:==                 ::......:++=::++:-*::::::-=-=------=#+=*=--======-          ")
    print("         :..:......:-=-                 ::...::..:*:::*-:=+::::::--:---------==-----=======         ")
    print("       ::..:...:..:-=-                   :::----:....--:::=-::::::--::-----------------=====        ")
    print("      ::..........:--:-       +*###      +=+++++:...::.::::::::::::---:---------------=======       ")
    print("     ::.....:+:......:--     =+****###*****-:::......::::::::::::::::-----:---------=-========      ")
    print("     :.:=.::==.......:=-     *************+.........:.....::::::::--:::----------------*%=====+     ")
    print("    :.:*+=+-:...::---==- ++==+************+.........::....::::::::::::::---------------***====++    ")
    print("   .:.....--...:=+****#********************:..........::....::::::::::::------------=--=*#+=====   ")
    print("  .::..::......-=**************************=.......::::......::---:::----------------*+==**======   ")
    print("  :...:..:....:=+********************+++===-:..........::-=--*::*::::-:::::-----------+###+====++   ")
    print(" :....:.......:=+**************++-:.........:.......:=#*:.+*-=::+::::-:::::--------------========+  ")
    print(" :::.::..:::::::::........=****:............::.........=#:.:*+::=:::::--:-=-------------=========+  ")
    print(" ...:..........:.........=******-............:..........:*+:+=#+-::::::----:----------===-=======++ ")
    print(":...-=-........:.........+***+=-.......::::.::...........:=*-.:+::::::::::::--------------=======++ ")
    print(":..:-=+=.......:......................:....................::..:::::::::::::--------------=======++ ")
    print(":...:=-==..:..:.............:::......:........................:::--::--:::::-------------=====+*+++ ")
    print("....-+*-..:..............*=...-#+....:.......................-::::::::-:::----------------===+**+++ ")
    print(":...=-:=..:.............#=.....-#=....-..::..:::............::.::::::---:--:-----------=--==++**+++ ")
    print(":...:++:..:..::........:#+.....:#=.....::......::...........::::::::::::::::-------------====+*+*+++ ")
    print(":..::::....:..:.........-#-...:+=.-.............-:...........::::::::::::--:-------------===+*+*+++ ")
    print("::.:..:........:........-:=*:.=+***.............::........::::::::::::-=#*:------------======++++++ ")
    print(" :.:..::.......:........=*+=:...................::::::::::::.::::=*+:::-##-----------==========++++ ")
    print(" :....::.......:......................::::..:...::...........::::-%+:::+=%=---------=--======+++++  ")
    print(" ::.....:::::..:::::::::::.......::::...........::.........:::::::*#::+-:#*---------=========+++++  ")
    print("  ::.+=:.......:...........:.....::.............::......::::::::::=%=+-::+%=--------=========++++   ")
    print("  .:.=:--......:..........::.......:.............-:...:::::::::::::##-::-++=---------=======+++++   ")
    print("   .::++==:....:..........:....::::..............:::::-:::::-::::::**=-:------------========++++    ")
    print("    ::-:*+:....::...........:..........................:::::-:::::::::------=------========+**+     ")
    print("     ::--=+.....::::::::......++:.:#-=+..............::::::::::::::---------=----=========+#%++     ")
    print("     :::::--...........::...:==-*-*=-#=:...........:::::::::::::::-:-----------==========+***+      ")
    print("      ::::::.....::::..::....:++==:*-*+:........:::-::::::::::::::-------------=========++**+       ")
    print("       ::::::::..::::::::....-*+--:===#=:.....:::::-::::::::::::--------------=========++**+        ")
    print("         :::::::::::..........=-*:=:*:*=+=::::::::::-:::::::---::-----%%%%*-========+=++++          ")
    print("          ::::::::-:::::...::.::::.::::::::::::::::::-:::::::::-------===**====+=++++++++           ")
    print("           :::::::::::::::::::::::::::::::::::::::::::--::::-------------**=======++++++            ")
    print("             ::=-::::::::::::::::::::::::::::::::::::::----------------==%======++=+++               ")
    print("               --==:::-::::::::::::::::::-::::::::::::---------------====*======+=+++                ")
    print("                 :==-:::-:::::::::::::::::::::::------------------============++++                  ")
    print("                   --+=--:::::::::::=+++=-:::------------==-----============++=+                    ")
    print("                     ==---------------=-----=----=====----==-==========++++++=                      ")
    print("                        -------------===--=--=*------------=-=======+*+++=++                        ")
    print("                           -----------==++==+++-------=-==========+++===                            ")
    print("                                ----=--------=--====================                                ")
    print("                                                                                                    ")
    print("                                                                                                    ")
    print("                                                                                                    ")

def open_files():
    """Prompt the user to open the CSV and/or Excel files."""
    if input("Would you like to open the CSV file now? (y/n): ").strip().lower() == 'y':
        open_csv('wiki_urls.csv')
    if input("Would you like to open the Excel file now? (y/n): ").strip().lower() == 'y':
        open_csv('wiki_urls.xlsx')

def open_csv(filename):
    """Open the file using the default application based on the operating system."""
    try:
        if sys.platform == "win32":
            os.startfile(filename)
        elif sys.platform == "darwin":
            subprocess.call(["open", filename])
        elif sys.platform == "linux" or sys.platform == "linux2":
            subprocess.call(["xdg-open", filename])
        else:
            print("Unsupported OS. Please open the file manually.")
    except Exception as e:
        print(f"Failed to open file: {e}")

def main():
    try:
        # Prompt the user for the initial URL
        initial_url = input("Enter the initial Wikipedia URL to start scraping (press Enter to start on a random page): ").strip()
        
        # If no URL is provided, start on a random Wikipedia page
        if not initial_url:
            initial_url = get_random_wikipedia_page()
            if initial_url is None:
                print("Failed to get a random Wikipedia page. Exiting...")
                return

        # Start scraping from the provided or random URL
        scrapeWikiArticle(initial_url)
    except KeyboardInterrupt:
        print("\nProcess interrupted. Exiting...")
    finally:
        # Ensure cleanup functions are called
        save_to_files()
        print_completion_message()
        # Prompt the user to open the CSV and/or Excel files
        open_files()

if __name__ == "__main__":
    main()
