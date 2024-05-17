import pandas as pd
import re
import os
import openpyxl
from datetime import datetime
import warnings
warnings.filterwarnings("ignore", category=FutureWarning)


# Define function for formatting main data
def format_main(file_path_to_output, brands_mapping_location, map_api_location, map_location, last_month_file_location):
    
    # Create directory for formatted files
    os.mkdir(file_path_to_output + "/FORMATTED")
    
    # List all files in the RAW directory
    files = [x for x in os.listdir(file_path_to_output +"/RAW")]
    print("FORMATTED UP FILES ARE STARTED")
    
    # Iterate through each file for formatting
    for file in files:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path_to_output + "/RAW/" + file , sheet_name="Sheet1", keep_default_na=False)

        # Check if the file is related to Heat Pumps or Central Air Conditioners
        if ('Heat Pumps (Ducted)' in file) or ('Central Air Conditioners (Ducted)' in file):
            # Data cleaning and standardization
            # Define functions for cleaning and transformation
            
            # Replace NA values with None
            
            df = df.where(pd.notnull(df), None)
            na_pattern = re.compile(r"^N/A$|^N\\?A$|^None$", re.I)
            
            # Function to remove NA values
            def remove_na(text):
                if text is None or re.search(na_pattern, str(text)):
                    return None
                elif text == '[]':
                    return None
                else:
                    return text
            
            def skip_none(f):
                def wrapped(text):
                    if text:
                        return f(text)
                    else:
                        return None
                return wrapped
            
            # Function to remove special characters
            def remove(text):
                chars = ['\\n', '\\r', '\\t', u'\u2022', u'\\u0040', u'\\xa0', '\\V']
                quotes = ['“', '”', '″']
                for ch in chars:
                    text = str(text).replace(ch, ' ')
                for q in quotes:
                    text = str(text).replace(q, '"')    
                return text

            # Function to trim extra spaces
            def trim_space(text):
                text = re.sub('<[^<]+?>', '', str(text))
                text = re.sub("\s+", " ", str(text).strip())
                return text
            
            # Apply cleaning functions to each column
            for column in list(df.columns):
                df[column] = df[column].map(remove_na).map(remove).map(trim_space)
                
            # Rename columns based on mapping
            column_name_map=pd.read_excel(map_location,sheet_name="Sheet1", na_filter=False) 
            column_names={}
            for index, row in column_name_map.iterrows():
                column_names[row['orig_column_name']] = row['rename'] 
            df = df.rename(column_names, axis=1)
            
            # Standardize brand names using brands mapping
            brand_map=pd.read_excel(brands_mapping_location, sheet_name="brands-mapping", na_filter=False) 
            brand_names = {}
            for index, row in brand_map.iterrows():
                brand_names[row['brand_name']] = row['brand_rename']
                
            df['brand-name'].replace(brand_names, inplace=True)

            # Add additional columns and fill values
            df['category'] = ''
            df['energy-star-model-name']=''	
            df['energy-star-model-number']=''

            # Update category, energy-star-model-name, and energy-star-model-number
            def add_new(row):
                row['category'] = ''
                row['category'] = "HVAC/Mechanical"
                row['energy-star-model-name'] = str(row['name'])
                row['energy-star-model-number'] = str(row['name'])
                return row
            df = df.apply(add_new, axis=1)

            #add timestamp
            df['timestamp'] = ''
            
            #update date
            def add_timestamp(row):
                row['timestamp'] = ''
                row['timestamp'] = "2024-04-03T00:00:00Z"
                return row
            df = df.apply(add_timestamp, axis=1)
           
            # Generate SKU based on certain conditions
            df['sku'] = ''
            def sku_add(row):
                row['sku'] = ''
                if row['name']:
                    name = str(row['name']).replace(" ", "-")
                    error=re.sub(r"[\r\n\t\x07\x0b\xa0\u0040]", " ", name).lower()
                    nama_single = error.replace("'", "")
                    nama_double = nama_single.replace('"','')
                    removeSpecialChars = nama_double.translate ({ord(c): "-" for c in "!@#$）%^（&*()•–＃—”“’‘[]{}';%:,./<>?\|`~-=_+"})
                    remove_space =  re.sub(' +','-',removeSpecialChars)                                  
                    remove_char = re.sub("™|®|", "", remove_space) 
                    mult_hyph = re.sub('-+','-',remove_char) 
                    row['sku'] = mult_hyph.rstrip('-').lstrip('-')          
                return row
            df = df.apply(sku_add, axis=1)

            #trim space
            def trim_space(text):
                text = re.sub('<[^<]+?>', '', str(text))
                text = re.sub("\s+", " ", str(text).strip())
                return text
            
            #remove duplicate  
            df = df.sort_values(['brand-name', 'name', 'sku'],ascending=True)
            df.drop_duplicates(subset=['brand-name', 'name'], keep='last', inplace = True)
            df['dup_rank'] = (df.groupby(['brand-name', 'sku']).cumcount().add(1)).astype(int)
            rep_df = df.groupby(['brand-name','sku'])['dup_rank'].agg(max) > 1
            dup_sku = [i for i in rep_df.index if rep_df[i]]
            
            
            def distinct_sku(row):
                sku_bra = row['sku']
                bra =row['brand-name']
                str_brand_sku=str(bra+sku_bra)
                for p in dup_sku:
                    join_list=[''.join(p)]
                    if str_brand_sku in join_list:  
                        row['sku'] = ('-').join([row['sku'],str(row['dup_rank'])]) 
                return row
            df = df.apply(distinct_sku, axis=1)

            def add_new1(row, file_name):
                    if 'sku' not in row:
                        row['sku'] = ''
                    if file_name == "Central Air Conditioners (Ducted)-Raw.xlsx":
                        row['subcategory'] = "Air Conditioning, Central"
                        row['sku'] = "cac-" + str(row['sku'])
                    elif file_name == "Heat Pumps (Ducted)-Raw.xlsx":
                        row['subcategory'] = "Heat Pumps, Air-Source"
                        row['sku'] = "hpd-" + str(row['sku'])
                    return row
            df = df.apply(lambda row: add_new1(row, file), axis=1)
            
            df.drop(columns=['dup_rank'], inplace=True)
            df.drop(columns=['Unnamed: 0'], inplace=True)
            
            # Update DataFrame structure
            df = df.reindex(columns=(['timestamp', 'brand-name', 'sku', 'name', 'energy-star-model-name', 'energy-star-model-number', 'category', 'subcategory', 'type', 'cooling-capacity-range', 'seer-range-btu-wh', 'eer-range-btu-wh', 'additional-features'] + 
            [a for a in df.columns if a not in ['timestamp', 'brand-name', 'sku', 'energy-star-model-name', 'name', 'energy-star-model-number', 'category', 'subcategory', 'type', 'cooling-capacity-range',  'seer-range-btu-wh', 'eer-range-btu-wh', 'additional-features']]))
            
            today_date = datetime.today().strftime("%Y-%m-%d")
            df.to_excel(file_path_to_output + "/FORMATTED/" + f"{file.replace('-Raw.xlsx', '')}-{today_date}.xlsx", index=False)

        else: 
            df = df.where(pd.notnull(df), None)
            na_pattern = re.compile(r"^N/A$|^N\\?A$|^None$", re.I)

            def remove_na(text):
                if text is None or re.search(na_pattern, str(text)):
                    return None
                elif text == '[]':
                    return None
                else:
                    return text

            def trim_space(text):
                text = re.sub('<[^<]+?>', '', str(text))
                text = re.sub("\s+", " ", str(text).strip())
                return text

            for column in list(df.columns):
                df[column] = df[column].map(remove_na).map(trim_space)

            column_name_map = pd.read_excel(map_api_location, sheet_name="Sheet1", na_filter=False)
            column_names = {}
            for index, row in column_name_map.iterrows():
                column_names[row['orig_column_name']] = row['rename']
            df = df.rename(column_names, axis=1)

            brand_map = pd.read_excel(brands_mapping_location, sheet_name="brands-mapping", na_filter=False)
            brand_names = {}
            for index, row in brand_map.iterrows():
                brand_names[row['brand_name']] = row['brand_rename']
            df['brand-name'].replace(brand_names, inplace=True)

            def sku_add(row):
                row['sku'] = ''
                if row['energy-star-model-number']:
                    name = str(row['energy-star-model-number']).replace(" ", "-")
                    error = re.sub(r"[\r\n\t\x07\x0b\xa0\u0040]", " ", name).lower()
                    nama_single = error.replace("'", "")
                    nama_double = nama_single.replace('"', '')
                    removeSpecialChars = nama_double.translate({ord(c): "-" for c in "!@#$）%^（&*()•–＃—”“’‘[]{}';%:,./<>?\|`~-=_+"})
                    remove_space = re.sub(' +', '-', removeSpecialChars)
                    remove_char = re.sub("™|®|", "", remove_space)
                    mult_hyph = re.sub('-+', '-', remove_char)
                    row['sku'] = mult_hyph.rstrip('-').lstrip('-')
                return row
            df = df.apply(sku_add, axis=1) 

            def add_new(row, file_name):
                row['category'] = "HVAC/Mechanical"

                if file_name == "Mini-Split Air Conditioners-Raw.xlsx":
                        row['type'] = "Mini-Split AC"
                        row['subcategory'] = "Ductless Heating and Cooling"
                        row['sku'] = "msac-" + row['sku']
                        
                elif file_name == "Room Air Conditioners-Raw.xlsx":
                    row['subcategory'] = "Air Conditioning, Room"
                    
                elif file_name == "Boilers-Raw.xlsx":
                    row['type'] = "Residential"
                    row['subcategory'] = "Boilers"
                    
                elif file_name == "Commercial Boilers-Raw.xlsx":
                    row['subcategory'] = "Boilers"
                    row['type'] = 'Commercial'
                    
                elif file_name == "Geothermal Heat Pumps-Raw.xlsx":
                    row['subcategory'] = "Heat Pumps, Geothermal or Ground-Source"
                    
                elif file_name == "Ventilating Fans-Raw.xlsx":
                    row['subcategory'] = "Ventilation Fans"
                    
                elif file_name == "Furnaces-Raw.xlsx":
                    row['subcategory'] = "Furnaces"
                    
                elif file_name == "Smart Thermostats-Raw.xlsx":
                    row['subcategory'] = "Smart Thermostats"
                    
                elif file_name == "Heat Pumps (Mini-Split)-Raw.xlsx":
                    row['subcategory'] = "Heat Pumps (Mini-Split)"
                    row['type'] = "HP - Mini or Multi Split"
                    
                elif file_name == "Light Commercial HVAC-Raw.xlsx":
                    row['subcategory'] = "Light Commercial Heating and Cooling"
                    row['type'] = "Small or Large AC"

                return row
            df = df.apply(lambda row: add_new(row, file), axis=1)
            df['url'] = df.apply(lambda row: f"https://www.energystar.gov/productfinder/product/certified-{file.split('-Raw.xlsx')[0].lower().replace(' ', '-')}/details/{row['energy-star-id']}", axis=1)
            def add_bool(df):
                for i, row in df.iterrows():
                    if "energy-star-lamp-included" in row:
                       if row['energy-star-lamp-included'] == "Yes":
                           df.at[i, 'energy-star-lamp-included'] = True
                       elif row['energy-star-lamp-included'] == "No":
                           df.at[i, 'energy-star-lamp-included'] = ""
                
                    if "meets-most-efficient-criteria-2024" in row:
                       if row['meets-most-efficient-criteria-2024'] == "Yes":
                           df.at[i, 'meets-most-efficient-criteria-2024'] = True
                       elif row['meets-most-efficient-criteria-2024'] == "No":
                           df.at[i, 'meets-most-efficient-criteria-2024'] = ""
        
                    if "low-noise" in row:
                       if row['low-noise']=="Yes":
                           df.at[i,'low-noise']=True
                       if row['low-noise']=="No":
                        df.at[i,'low-noise']=""   
                            
                    if "variable-speed-compressor" in row:
                       if row['variable-speed-compressor']=="Yes":
                           df.at[i,'variable-speed-compressor']=True
                       if row['variable-speed-compressor']=="No":
                           df.at[i,'variable-speed-compressor']=""                   
        
                    if "date-available-on-market" in row:
                       date_ava = row['date-available-on-market']
                       df.at[i,'date-available-on-market']=date_ava.replace("T00:00:00.000", "T00:00:00Z")
        
                    if "starts" in row:
                       date_ava = row['starts']
                       df.at[i,'starts']=date_ava.replace("T00:00:00.000", "T00:00:00Z")          
                return df
            df = add_bool(df)

            rem_brand2=['TeK','Ace','Toshiba','Philips','PHILIPS','/', 'SAVIN', 'RICOH', 'LANIER', 'HPE', 'Acer','Compumax','dynabook', '®','-','?','Dell EMC','EMC','HP']

            def name_add(row):
                row['name'] = ''
                if row['energy-star-model-name']:
                    name = str(row['energy-star-model-name'])
                    row['name'] = name
                return row

            if 'energy-star-model-name' in df.columns:
                df = df.apply(name_add, axis=1)
            else:
                df['energy-star-model-name'] = df['energy-star-model-number']
                df = df.apply(name_add, axis=1)
            df = df.apply(sku_add, axis=1)  # Apply sku_add function here
            df = df.apply(lambda row: add_new(row, file), axis=1)
            df = add_bool(df)  

            def remv_brand2(df):
                if 'name' in df.columns:
                    for i, row in df.iterrows():
                        if row['brand-name']:
                            for b in rem_brand2:
                                if b in row['name']:
                                    row['name'] = row['name'].lstrip(b).strip()
                                if b in row['energy-star-model-number']:
                                    row['energy-star-model-number'] = row['energy-star-model-number'].lstrip(b).strip()
                return df

            df = remv_brand2(df)

            def add_timestamp(row):
                row['timestamp'] = ''
                row['timestamp'] = "2024-03-04T00:00:00Z"
                return row
            df = df.apply(add_timestamp, axis=1)

            df = df.reindex(columns=(['timestamp', 'brand-name', 'sku', 'name', 'energy-star-model-name', 'energy-star-model-number', 'category', 'subcategory', 'type', 'energy-star-id', 'additional-model-information', 'energy-star-partner', 'markets'] +
                                      [a for a in df.columns if a not in ['timestamp', 'brand-name', 'sku', 'energy-star-model-name', 'name', 'energy-star-model-number', 'category', 'subcategory', 'type', 'energy-star-id', 'additional-model-information', 'energy-star-partner', 'markets']]))

            df = df.sort_values(['brand-name', 'name', 'sku'], ascending=True)
            df.drop_duplicates(subset=['brand-name', 'energy-star-model-name', 'sku', 'energy-star-id'], keep='last', inplace=True)

            df['dup_rank'] = df.groupby(['brand-name', 'sku']).cumcount().fillna(0).astype(int).add(1)
            rep_df = df.groupby(['brand-name', 'sku'])['dup_rank'].agg(max) > 1
            dup_sku = [i for i in rep_df.index if rep_df[i]]

            def distinct_sku(row):
                sku_bra = str(row['sku'])  # Ensure sku is converted to string
                bra = row['brand-name']
                str_brand_sku = str(bra + sku_bra)
                if str_brand_sku in dup_sku:
                    if pd.notnull(row['dup_rank']):  # Check if dup_rank is not NaN
                        row['sku'] = '-'.join([sku_bra, str(int(row['dup_rank']))])
                    else:
                        row['sku'] = sku_bra  # If dup_rank is NaN, retain original sku
                return row

            df = df.apply(distinct_sku, axis=1)
            df.drop(columns=['dup_rank'], inplace=True)
            df.drop(columns=['Unnamed: 0'], inplace=True)
            
            today_date = datetime.today().strftime("%Y-%m-%d")
            df.to_excel(file_path_to_output + "/FORMATTED/" + f"{file.replace('-Raw.xlsx', '')}-{today_date}.xlsx", index=False)

        if file.endswith(".xlsx"):
                file_path = os.path.join(file_path_to_output + "/FORMATTED/" + f"{file.replace('-Raw.xlsx', '')}-{today_date}.xlsx")
                wb = openpyxl.load_workbook(file_path) # Open the Excel 
                ws = wb.worksheets[0]
                worksheet = wb.active # Get the active worksheet
                worksheet.freeze_panes = "A2"  # Freezes everything above row 2 (including row 1)
                worksheet.auto_filter.ref = worksheet.dimensions  # Applies filter to the entire
                # Left-align the text in the first row (assuming you want this for all columns)
                for cell in worksheet[1]:
                    alignment = openpyxl.styles.Alignment(horizontal="left")
                wb.save(os.path.join(file_path_to_output + "/FORMATTED/" + f"{file.replace('-Raw.xlsx', '')}-{today_date}.xlsx"))

            
    print("\n FORMATTED UP FILES DONE\n")
    
    
    
    
    