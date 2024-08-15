import requests
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime
import logging
import subprocess

# Initialize logging
logging.basicConfig(filename='printer_metrics.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Initialize a dictionary to map printer IP addresses to IDs
printer_id_map = {}

# Function to get printer ID or assign a new ID if it's a new IP Address
def get_printer_id(printer_ip_address):
    global printer_id_map
    if printer_ip_address not in printer_id_map:
        # Assign a new ID for the printer IP address
        max_id = max(printer_id_map.values()) if printer_id_map else 0
        printer_id_map[printer_ip_address] = max_id + 1
    return printer_id_map[printer_ip_address]



try:
    logging.info("Running code...")
    
    # Read IP addresses from the text file
    with open('IP Address.txt', 'r') as file:
        ip_addresses = file.read().splitlines()
    
    for ip_address in ip_addresses:
        try:
            logging.info("Running code for printer at IP address: " + ip_address)
            
            # Append IP address to URLs
            printer_device_model = f'http://{ip_address}/hp/device/DeviceInformation/View'
            printer_pcount_url = f'http://{ip_address}/hp/device/InternalPages/Index?id=UsagePage'
            printer_device_status = f'http://{ip_address}/hp/device/DeviceStatus/Index'
            printer_device_name = f'http://{ip_address}/network_id.htm'
            ipaddress = 'HomeDeviceIp'
            model = 'DeviceName'
            A4 = 'UsagePage.ImpressionsByMediaSizeTable.Print.A4.Total'
            A5 = 'UsagePage.ImpressionsByMediaSizeTable.Print.A5.Total'

            # Check if the printer URLs are accessible
            for url in [printer_device_model,printer_pcount_url, printer_device_status, printer_device_name]:
                try:
                    response = requests.get(url)
                    response.raise_for_status()  # Raise an exception for no responses
                except requests.exceptions.HTTPError as http_err:
                    logging.error(f"HTTP error occurred while checking URL {url} for printer at IP address {ip_address}: {http_err}")
                    raise  # Raise the error to skip further processing for this printer

            logging.info("Done scraping." + ip_address)

            def get_ip_address(url):
                try:
                    response = requests.get(url)
                    soup = BeautifulSoup(response.text, 'html.parser')
                    data = soup.find('p', {'id': ipaddress}).text
                    return data
                except requests.exceptions.RequestException:
                    logging.error("Issue gathering IP address...")
                    return None

            def get_printer_model(url):
                try:
                    response = requests.get(url)
                    soup = BeautifulSoup(response.text, 'html.parser')
                    data = soup.find('p', {'id': model}).text
                    return data
                except requests.exceptions.RequestException:
                    logging.error("Issue gathering Printer Name...")
                    return None
                
            def get_printer_name(url):
                try:
                    response = requests.get(url)
                    soup = BeautifulSoup(response.text, 'html.parser')
                    hostname_input = soup.find('input', {'id': 'IPv4_HostName'})
                    hostname = hostname_input.get('value')
                    return hostname 
                except requests.exceptions.RequestException as e:
                        logging.error(f"Issue gathering Printer Name: {e}")
                        return None
                
            def get_page_A4(url):
                try:
                    response = requests.get(url)
                    soup = BeautifulSoup(response.text, 'html.parser')
                    data = soup.find('td', {'id': A4}).text
                    data = int(data.replace(',', ''))
                    return data
                except requests.exceptions.RequestException:
                    logging.error("Issue gathering Page Count A4...")
                    return None
                
            def get_page_A5(url):
                try:
                    response = requests.get(url)
                    soup = BeautifulSoup(response.text, 'html.parser')
                    data = soup.find('td', {'id': A5}).text
                    data = int(data.replace(',', ''))
                    return data
                except AttributeError:
                    logging.warning("A5 attribute not found, setting to 0 by default.")
                    return 0
                except requests.exceptions.RequestException:
                    logging.error("Issue gathering Page Count A5...")
                    return None

            # Page Counts for the printers
            printer_page_count4 = get_page_A4(printer_pcount_url)
            printer_page_count5 = get_page_A5(printer_pcount_url)
            # Toner level for the printer
            printer_ip_address = get_ip_address(printer_device_status)
            # Printer model
            printer_model = get_printer_model(printer_device_model)
            # Printer name
            printer_name = get_printer_name(printer_device_name)
            # Get or assign printer ID
            printer_id = get_printer_id(ip_address)
            
            # Get the current date
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
            if printer_ip_address in printer_id_map:
                printer_id = printer_id_map[printer_ip_address]
            else:
                # Assign a new ID for the printer name
                printer_id = max_id
                printer_id_map[printer_ip_address] = printer_id
                max_id += 1  # Increment max_id for the next new printer
            
            
            # Create the Pandas DataFrame for the Excel export
            df = pd.DataFrame({
                'Printer ID': [printer_id],  # Incremental ID column
                'Date': [date_string],
                'Printer model': [printer_model],
                'Printer name': [printer_name],
                'IP Address': [ printer_ip_address ],
                'A4 page': [printer_page_count4 ],
                'A5 page': [printer_page_count5 if printer_page_count5 is not None else 0],
            })
           


            # Append new data to the previous data
            df = pd.concat([prev_df, df], ignore_index=True)

            # Write to Excel
            df.to_excel('Printer_Metrics.xlsx', sheet_name='printers', index=False)

            logging.info("Data added to Excel file.")

        except Exception as e:
            logging.error(f"An error occurred for printer at IP address {ip_address}: {str(e)}")
            continue  # Move to the next IP address if there's an error

        
except Exception as e:
    logging.error(f"An error occurred: {str(e)}")




# Execute the command using subprocess
excel_file= 'Printer_Metrics.xlsx'
sheet='printers'
subprocess.run(['python', 'HP M501dn_Printer_Scrape.py', excel_file, sheet])
subprocess.run(['python', 'printer_processing.py', excel_file, sheet])
subprocess.run(['python', 'merge tables.py', excel_file, sheet])
