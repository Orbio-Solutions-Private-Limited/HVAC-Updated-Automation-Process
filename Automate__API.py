from lxml import html
import requests
import pandas as pd
from sodapy import Socrata
from tqdm import tqdm
import logging
from bs4 import BeautifulSoup
import os
import re 
import warnings

# Ignore future warnings
warnings.filterwarnings("ignore", category=FutureWarning)

# Importing custom modules
from Automate__formatted import format_main as formatting_main
from Automate_JSON import creation_json as creation_json_file
from Comparing_excel_up import format_upExcel as formatting_upExcel

# Load configuration from config.json
with open('config.txt') as f:
    content = f.read()
    print("Content:", content)

# Extract values from the config
config = {}
for line in content.split('\n'):
    if line.strip():  # Ignore empty lines
        key, value = line.strip().split('=')
        config[key.strip()] = value.strip()

# Remove double quotes from file paths
for key, value in config.items():
    if value.startswith('"') and value.endswith('"'):
        config[key] = value[1:-1]

# Assign config values to variables
file_path_to_output = config['file_path_to_output']
brands_mapping_location = config['brands_mapping_location']
map_api_location = config['map_api_location']
map_location = config['map_location']
last_month_file_location = config['last_month_file_location']

# Set the logging level to suppress the warning
logging.getLogger().setLevel(logging.ERROR)

# Function to scrape all pages of a website
def scrape_all_pages(url):
    all_data = pd.DataFrame()
    page_num = 0  # Start with the first page
    while True:
        page_url = f"{url}&page_number={page_num}"
        df = scrape_central_air(page_url)
        if df.empty:  # Check if there's no data on this page (e.g., last page)
            break
        all_data = pd.concat([all_data, df], ignore_index=True)  # Concatenate only if data exists
        page_num += 1  # Move to the next page
    return all_data

# Function to scrape data for central air
def scrape_central_air(url):
    headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36",
    "Accept-Encoding": "gzip, deflate",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "DNT": "1",
    "Connection": "close",
    "Upgrade-Insecure-Requests": "1"
}
    # Send a GET request to the URL
    webpage = requests.get(url, headers=headers)
    soup = BeautifulSoup(webpage.content, "html.parser")
    heat_list = []

    # Find all rows in the webpage
    rowheading = soup.find_all('div', attrs={'class': 'row'})
    for item in rowheading:
        heat_dict = {}
        title = item.find('div', class_='title')
        if title:
            # Extract brand and model information from the title
            name = ' '.join(title.text.strip().split())
            cleaned_text = ' '.join(name.split())
            match = re.match(r'^(.*?)\s*-\s*(.*)$', cleaned_text)

            if match:
                brand = match.group(1).strip()
                model = match.group(2).strip()
            else:
                print("Unable to extract brand and model.")

            heat_dict['brand-name'] = brand
            heat_dict['name'] = model

            head = ""
            field = item.find_all('div', attrs={'class': 'field'})
            additional_feature_values = []
            default_head_values = []
            for r in field:
                label = r.find_all('div', attrs={'class': 'label'})
                value = r.find_all('div', attrs={'class': 'value'})
                if label:
                    for lab in label:
                        head = lab.text.replace("\n", "").strip()
                        value_txt = [val.text.replace("\n", "").strip() for val in value]
                        value_txt = [re.sub(r'\s+', ' ', txt) for txt in value_txt]
                        if head == str("Additional Features :"): 
                            additional_feature_values.extend(value_txt)
                        elif head == "Default Head":
                            default_head_values.extend(value_txt)
                        if len(value_txt) > 1:
                            heat_dict[head] = value_txt
                        else:
                            heat_dict[head] = value_txt[0]
                else:
                    for val in value:
                        value_txt = val.text.replace("\n", "").strip()
                        value_txt = re.sub(r'\s+', ' ', value_txt)
                        heat_dict['Default Head'] = value_txt

        if heat_dict:
            heat_list.append(heat_dict)  # Append data from each row

    df = pd.DataFrame.from_dict(heat_list)
    df = df.apply(lambda x: x.strip() if isinstance(x, str) else x)
    df.drop_duplicates(inplace=True)  # Remove duplicates within the scraped data for this page
    return df

# Main function to execute the scraping and processing tasks
def main(file_path_to_output, brands_mapping_location, map_api_location, map_location, last_month_file_location):

    # Create a directory for storing raw data
    os.mkdir(file_path_to_output + "\\RAW")

    # Find all elements under 'Heating & Cooling' category
    page = requests.get('https://www.energystar.gov/productfinder/advanced')
    tree = html.fromstring(page.content)

    # Find all elements under 'Heating & Cooling' category
    span_elements = tree.xpath("//div//h3[text()='Heating & Cooling']/../ul//li//div//span")
    print("files are present under the 'Heating & Cooling' on webpage:" )
    for span in tqdm(span_elements, total=len(span_elements)):
        file = span.text.strip()
        if ('Heat Pumps (Ducted)' in file) or ('Central Air Conditioners (Ducted)' in file):
            if 'Central Air Conditioners (Ducted)' in file:
                # Process for Central Air Conditioners
                url_central_air = "https://www.energystar.gov/productfinder/product/certified-central-air-conditioners/?formId=72-98-4771-954-840392&scrollTo=4350.66650390625&search_text=&product_filter=&outdoorunitbrandname_isopen=0&zip_code_filter=&product_types=Select+a+Product+Category&sort_by=max_seer2&sort_direction=desc"
                num_pages_central_air = 25
                req_html = requests.get(url_central_air)  # Sending a GET request to the URL
                df = scrape_all_pages(url_central_air)
                soup_html = BeautifulSoup(req_html.content, features='lxml')
                site_count = int(soup_html.find_all('div', class_="records-found-small")[0].get_text().strip().replace("\xa0Records Found", ""))
                print("\n", df.shape[0], site_count, file) 
            elif 'Heat Pumps (Ducted)' in file:
                # Process for Heat Pumps
                url_heat_pump = "https://www.energystar.gov/productfinder/product/certified-central-heat-pumps/?formId=515716-9089-47-902-1339357&scrollTo=5342.66650390625&search_text=&product_filter=&outdoorunitbrandname_isopen=0&is_most_efficient_filter=0&zip_code_filter=&product_types=Select+a+Product+Categorym_pages_heat_pump = 30"
                df = scrape_all_pages(url_heat_pump)
                req_html = requests.get(url_heat_pump) 
                soup_html = BeautifulSoup(req_html.content, features='lxml')
                site_count = int(soup_html.find_all('div', class_="records-found-small")[0].get_text().strip().replace("\xa0Records Found", ""))
                print("\n", df.shape[0], site_count, file) 

            if not df.empty:
                merged_values = []
                for index, row in df.iterrows():
                    if pd.isna(row[str("Additional Features :")]):
                        merged_values.append(row['Default Head'])
                    elif pd.isna(row['Default Head']):
                        merged_values.append(row[str("Additional Features :")])
                    else:
                        merged_values.append(row[str("Additional Features :")] + '; ' + row['Default Head'])
                df['Additional Feature:'] = merged_values
                # Drop the original columns if needed
                df.drop([str("Additional Features :"), 'Default Head'], axis=1, inplace=True)
            df.to_excel(f"{file_path_to_output}\\RAW\\{file}-Raw.xlsx")
            
        
        else:
            for a_tag in span.findall("a"):
                if "Finder" in a_tag.text:
                    finder_url = "https://www.energystar.gov/productfinder" + \
                        a_tag.get('href')[1:]
                if "API" in a_tag.text:
                    gotten_url = a_tag.get('href')
                    url = gotten_url.replace("https://dev.socrata.com/foundry/" + "data.energystar.gov/", "")

                    client = Socrata("data.energystar.gov", None)
                    results = client.get(url, limit=10000)

                    df = pd.DataFrame.from_records(results)
                    # Removing markets that don't contain the U.S
                    df = df[df['markets'].str.contains('United States', na=False)]

                if "Finder" in a_tag.text:
                    finder_url = "https://www.energystar.gov/productfinder" + \
                        a_tag.get('href')[1:]
            req_html = requests.get(finder_url, timeout=10)
            if req_html.status_code == 200:
                soup_html = BeautifulSoup(req_html.content, features='lxml')
                site_count = int(soup_html.find_all('div', class_="records-found-small")[0].get_text().strip().replace("\xa0Records Found", ""))

                print("\n", df.shape[0], site_count, file) 
            df.to_excel(f"{file_path_to_output}/RAW/{file}-Raw.xlsx")            

    print("\nRAW FILES DONE\n")

    
main(file_path_to_output, brands_mapping_location, map_api_location, map_location, last_month_file_location)
formatting_main(file_path_to_output, brands_mapping_location, map_api_location, map_location, last_month_file_location)
creation_json_file(file_path_to_output, brands_mapping_location, map_api_location, map_location, last_month_file_location)
formatting_upExcel(file_path_to_output, brands_mapping_location, map_api_location, map_location, last_month_file_location)

