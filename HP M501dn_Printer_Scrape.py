import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import logging

# Initialize logging
logging.basicConfig(filename='printer_metrics.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Initialize a dictionary to map printer names to IDs
printer_id_map = {}

# Function to get printer ID or assign a new ID if it's a new IP Address
def get_printer_id(printer_ip_address):
    global printer_id_map
    if printer_name not in printer_id_map:
        # Assign a new ID for the printer name
        printer_id_map[printer_ip_address] = len(printer_id_map) + 1
    return printer_id_map[printer_ip_address]

def get_printer_model(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        data = soup.find('td', class_='itemFont')
        if data:
            page_count = data.text
            return page_count
        else:
            logging.error("Tag with class 'itemFont' not found.")
            return None
    except requests.exceptions.RequestException:
        logging.error("Issue gathering Page Count...")
        return None

def get_printer_name(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        h3_element = soup.find('h3', class_='subTitle', string='Identification réseau')
        if h3_element:
            data = h3_element.find_next('td', class_='itemFont')
            if data:
                name = data.text.strip()
                return name  
            else:
                logging.error("No 'itemFont' element found after 'Impressions' subtitle.")
                return None
        else:
            logging.error("SubTitle 'Impressions' not found.")
            return None
    except requests.exceptions.RequestException as e:
        logging.error("Error:", e)
        return None   

def get_model_ip_address(url):
    try:
        response = requests.get(url)
        response.raise_for_status()  
        soup = BeautifulSoup(response.text, 'html.parser')
        item_fonts = soup.find_all('td', class_='itemFont')
        if len(item_fonts) >= 2:
            return item_fonts[1].text.strip()  # Return the text content of the second matching element
        else:
            logging.error("Second 'itemFont' element not found.")
            return None
    except requests.exceptions.RequestException as e:
        logging.error("Error:", e)
        return None

def get_page_count(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Find the h3 element with class "subTitle" containing "Impressions"
        h3_element = soup.find('h3', class_='subTitle', string='Impressions')
        
        if h3_element:
            # Find the sibling td element with class "itemFont"
            data = h3_element.find_next('td', class_='itemFont')
            if data:
                page_count = data.text.strip().replace(',', '')  # Remove commas from the string
                return int(page_count)  # Convert the result to an integer
            else:
                logging.error("No 'itemFont' element found after 'Impressions' subtitle.")
                return None
        else:
            logging.error("SubTitle 'Impressions' not found.")
            return None
    except requests.exceptions.RequestException as e:
        logging.error("Error:", e)
        return None

try:
    logging.info("Running code...")
    # Read IP addresses from the text file
    with open('M501dn.txt', 'r') as file:
        ip_addresses = file.read().splitlines()

    # Attempt to load the previous data from Excel if it exists
    try:
        prev_df = pd.read_excel('Printer_Metrics.xlsx')
    except FileNotFoundError:
        prev_df = pd.DataFrame()

    max_id = prev_df['Printer ID'].max() + 1 if not prev_df.empty else 1

    # Update the printer_id_map with existing data
    for index, row in prev_df.iterrows():
        printer_id_map[row['IP Address']] = row['Printer ID']
    
    # Printers
    for ip_address in ip_addresses:
        try:
            logging.info("Running code for printer at IP address: " + ip_address)

            # URLs for the Usage Metrics
            printer_pcount_url = f'http://{ip_address}/info_configuration.html?tab=Home&menu=DevConfig'

            # URLs for the consumable levels
            printer_device_status = f'http://{ip_address}/info_config_network.html?tab=Home&menu=NetConfig'

            printer_device_name = f'http://{ip_address}/info_config_network.html?tab=Networking&menu=NetConfig'

            # Check if the printer URLs are accessible
            for url in [printer_pcount_url, printer_device_status, printer_device_name]:
                try:
                    response = requests.get(url)
                    response.raise_for_status()  # Raise an exception for no responses
                except requests.exceptions.HTTPError as http_err:
                    logging.error(f"HTTP error occurred while checking URL {url} for printer at IP address {ip_address}: {http_err}")
                    raise  # Raise the error to skip further processing for this printer

            # Page Counts for the printers   
            printer_page_count = get_page_count(printer_pcount_url)

            # Toner levels for each printer.
            model_ip_address = get_model_ip_address(printer_device_status)

            # Impressions values from the page
            printer_model = get_printer_model(printer_pcount_url) 
            # Printer name
            printer_name = get_printer_name(printer_device_name)

            printer_id = get_printer_id(ip_address)

            logging.info("Done scraping." + ip_address)
            
            

            # Get the Current Date
            now = datetime.now()
            date_string = now.strftime("%Y-%m-%d")

            # Attempt to load the previous data from Excel if it exists
            try:
                prev_df = pd.read_excel('Printer_Metrics.xlsx')
            except FileNotFoundError:
                prev_df = pd.DataFrame()

            max_id = prev_df['Printer ID'].max() + 1 if not prev_df.empty else 1
            # Update the printer_id_map with existing data
            for index, row in prev_df.iterrows():
                printer_id_map[row['IP Address']] = row['Printer ID']

            # Get or assign printer ID for the current printer
            if model_ip_address in printer_id_map:
                printer_id = printer_id_map[model_ip_address]
            else:
                # Assign a new ID for the printer name
                printer_id = max_id
                printer_id_map[model_ip_address] = printer_id
                max_id += 1  # Increment max_id for the next new printer
            

            # Create the Pandas DataFrame for the Excel export.
            df = pd.DataFrame({
                'Printer ID': [printer_id],
                'Date': [date_string] ,
                'Printer model': [printer_model],
                'Printer name': [printer_name],
                'IP Address': [ model_ip_address],
                'A4 page': [0],
                'A5 page': [printer_page_count],  # Add impressions total to the DataFrame # Add impressions values to the DataFrame
                

            })


            # Append new data to the previous data
            df = pd.concat([prev_df, df], ignore_index=True)

            # Write to Excel 
            df.to_excel('Printer_Metrics.xlsx', sheet_name='printers', index=False)

            # Define the data for the pages table
            pages_data = {
                'Page ID': [1, 2],
                'Page Size': ['A4', 'A5'],
                'Cost': [0.07, 0.07]
            }

            # Create a DataFrame for the pages table
            pages_df = pd.DataFrame(pages_data)

            # Write to Excel in a separate sheet named 'pages'
            with pd.ExcelWriter('Printer_Metrics.xlsx', engine='openpyxl', mode='a') as writer:
                pages_df.to_excel(writer, sheet_name='pages', index=False)

            logging.info("Data added to Excel file.")

            

        except Exception as e:
            logging.error(f"An error occurred: {str(e)}")
except Exception as e:
    logging.error(f"An error occurred: {str(e)}")

