import pandas as pd
import json,jsonlines
import os
from datetime import datetime
import xlrd 
import sys
from tqdm import tqdm
import os

from Comparing_excel_up import format_upExcel as formatting_upExcel

def creation_json(file_path_to_output, brands_mapping_location, map_api_location, map_location, last_month_file_location):
    print("\n  CREATION OF JSON AND JSONL FILES: ")
    os.makedirs(os.path.join(file_path_to_output, "JSON"), exist_ok=True)

    files = [x for x in os.listdir(file_path_to_output + "/FORMATTED")]

    for file_name in tqdm(files, total=len(files)):
        df = pd.read_excel(os.path.join(file_path_to_output + "/FORMATTED", file_name), sheet_name="Sheet1", keep_default_na=False)
        print("list of files:", file_name)
        data_dict = df.to_dict()
        df = df.where(pd.notnull(df), None)

        
        if ('Heat Pumps (Ducted)' in file_name) or ('Central Air Conditioners (Ducted)' in file_name):
            data = df.to_dict()
            data_list = []
            data_dict = {}
            for i in range(len(data['name'])):
                data_dict = {}
                data_dict['disabled'] = False 
                data_dict['energy-star'] = True  
                if data['timestamp'][i]:
                    data_dict['timestamp'] = str(data['timestamp'][i]) 
                if data['brand-name'][i]:
                    data_dict['brand-name'] = str(data['brand-name'][i])
                if data['sku'][i]:
                    data_dict['sku'] = str(data['sku'][i])  
                if data['name'][i]:
                    data_dict['name'] = str(data['name'][i])         
                if data['energy-star-model-name'][i]:
                    data_dict['energy-star-model-name'] = str(data['energy-star-model-name'][i])   
                if data['energy-star-model-number'][i]:
                    data_dict['energy-star-model-number'] = str(data['energy-star-model-number'][i])        
                if 'category' in data.keys():
                    if data['category'][i]:    
                        data_dict['category'] = [data['category'][i]] 
                if 'subcategory' in data.keys():
                    if data['subcategory'][i]:    
                        data_dict['subcategory'] = [data['subcategory'][i]] 
                if 'type' in data.keys():
                    if data['type'][i]:    
                        data_dict['type'] = str(data['type'][i])

                x_str=['cooling-capacity','cooling-capacity-range','heating-capacity','seer-range-btu-wh','eer-range-btu-wh','hspf','heating-capacity-at-47-f','heating-capacity-at-5-f','indoor-unit-model-number']
                for j in x_str:
                    if j in data.keys():
                        if data[j][i]:
                            data_dict[j]  = str(data[j][i])

                sort_dict=dict(sorted(data_dict.items()))
                data_list.append(sort_dict)


            with open(os.path.join(file_path_to_output, "JSON", file_name.replace(".xlsx", ".json")), 'w', encoding='utf-8') as fout:
                for j in data_list:
                    json.dump(j , fout, indent=1, ensure_ascii=False, default=str)
                    fout.write('\n \n')                

            jsonl_folder_path = os.path.join(file_path_to_output, "JSONL")
            os.makedirs(jsonl_folder_path, exist_ok=True)
            jsonl_filename = os.path.join(file_path_to_output, "JSONL", file_name.replace(".xlsx", ".jsonl"))
            with jsonlines.open(jsonl_filename, mode='w') as writer:
                for j in data_list:
                    writer.write(j)

        else:
            data = df.to_dict()
            data_list = []
            
            for i in range(len(data['name'])):
                data_dict = {}
                data_dict['disabled'] = False 
                # data_dict['energy-star'] = True  
                if data['timestamp'][i]:
                    data_dict['timestamp'] = str(data['timestamp'][i]) 
                if data['brand-name'][i]:
                    data_dict['brand-name'] = str(data['brand-name'][i])         
                if data['sku'][i]:
                    data_dict['sku'] = str(data['sku'][i])  
                if data['name'][i]:
                    data_dict['name'] = str(data['name'][i])         
                if data['energy-star-model-name'][i]:
                    data_dict['energy-star-model-name'] = str(data['energy-star-model-name'][i])   
                if data['energy-star-model-number'][i]:
                    data_dict['energy-star-model-number'] = str(data['energy-star-model-number'][i])        
                if 'category' in data.keys():
                    if data['category'][i]:    
                        data_dict['category'] = [data['category'][i]] 
                if 'subcategory' in data.keys():
                    if data['subcategory'][i]:    
                        data_dict['subcategory'] = [data['subcategory'][i]] 
                if 'type' in data.keys():
                    if data['type'][i]:    
                        data_dict['type'] = str(data['type'][i])
                if data['energy-star-id'][i]:
                    data_dict['energy-star-id'] = str(data['energy-star-id'][i])            


                data_dict["energy-star-certificate"] = []
                cert ={}
                if data['energy-star-id'][i]:
                    cert['id'] = str(data['energy-star-id'][i]) 
                if data['starts'][i]:
                    cert['starts'] = str(data['starts'][i])  
                if data['url'][i]:
                    cert['url'] = str(data['url'][i])         
                data_dict["energy-star-certificate"].append(cert)  

                x_str=['additional-ct-device-model-numbers', 'additional-model-information', 'ahri-reference-number', 'alternate-energy-star-lamps-esuid', 'boiler-application', 'boiler-control-type', 'broadband-connection-needed-for-demand-response', 'can-integrate-hot-water-heating', 'capable-of-two-way-communication', 'casement-window', 'cold-climate', 'communication-hardware-architecture', 'communication-method-other', 'communication-standard-application-layer', 'compressor-staging', 'connected-capability', 'connected-capable', 'connected-functionality', 'connect-using', 'cooling-capacity', 'cooling-capacity-range', 'cop-at-5f', 'correlated-color-temperature-kelvin', 'ct-device-brand-name', 'ct-device-brand-owner', 'ct-device-communication-method', 'ct-device-model-name', 'ct-device-model-number', 'ct-product-heating-and-cooling-control-features', 'demand-response-product-variations', 'demand-response-summary', 'depth-inches', 'direct-on-premises-open-standard-based-interconnection', 'dr-protocol', 'duct-size', 'eer-range-btu-wh', 'efficiency-afue', 'energy-star-lamp-partner', 'energy-star-model-identifier', 'energy-star-partner', 'family-id', 'fan-lamp-model-number', 'features', 'fuel-type', 'furnace-is-energy-star-certified-in', 'heating-capacity', 'heating-capacity-at-17-f-btu-h', 'heating-capacity-at-47-f', 'heating-capacity-at-47-f-btu-h', 'heating-capacity-at-5-f', 'heating-capacity-at-5-f-btu-h', 'heating-mode', 'height-inches', 'hspf', 'hspf-range-btu-wh', 'indoor-unit-model-number', 'installation-capabilities', 'installation-mounting-type', 'lighting', 'lighting-technology-used', 'meets-peak-cooling-requirements', 'network-security-standards', 'notes', 'number-of-speeds', 'other-heating-and-cooling-control-features', 'outdoor-unit-brand-name', 'primary-communication-module-device-brand-name-and-model-number', 'product-class', 'refrigerant-type', 'refrigerant-type-gwp', 'reverse-cycle', 'seer-range-btu-wh', 'sound-level-sones', 'special-features-dimming-motion-sensing-etc', 'support-bracket', 'tax-credit-eligible', 'tax-credit-eligible-cac-national', 'tax-credit-eligible-heat-pumps-north', 'tax-credit-eligible-heat-pumps-south', 'voltage-volts', 'weight-lbs', 'width-inches']
                
                for j in x_str:
                    if j in data.keys():
                        if data[j][i]:
                            data_dict[j]  = str(data[j][i]) 
                
                x_int=['aeu', 'airflow-1-cfm', 'airflow-2-cfm', 'airflow-3-cfm', 'bathroom-utility-room-airflow-at-025-in-wg', 'boiler-full-load-input-rate', 'boiler-turndown-ratio', 'color-rendering-index-cri', 'combined-energy-efficiency-ratio-ceer', 'cool-cap', 'cooling-capacity-kbtu-h', 'cop-rating', 'cop-rating-at-17-degrees', 'cop-rating-at-47-degrees', 'eer2-rating-btu-wh', 'eer-rating-btu-wh', 'efficacy-1-cfm-w', 'efficacy-2-cfm-w', 'efficacy-3-cfm-w', 'energy-efficiency-ratio-eer', 'energy-star-lamp-esuid', 'ieer-rating', 'le-measured', 'light-out', 'light-source-life-hours', 'merv-of-in-line-fan-filter', 'network-standby-average-power-consumption', 'percent-less-energy-use-than-us-fed-standard', 'power-factor', 'seer2-rating-btu-wh', 'seer-rating-btu-wh', 'static-temperature-accuracy', 'thermal-efficiency-te']
                
                for k in x_int:
                    if k in data.keys():
                        if data[k][i]:
                            data_dict[k]  = int(float(data[k][i]))
                            
                if 'markets' in data.keys():
                    if data['markets'][i]:
                        if ',' in data['markets'][i]:
                            mark = data['markets'][i].split(",")
                            data_dict['markets'] = [x.strip(' ') for x in mark]  
                        else:
                            data_dict['markets'] = str(data['markets'][i])
                            
                if 'upc' in data.keys():
                    if data['upc'][i]:
                        if ';' in data['upc'][i]:
                            upc_str = str(data['upc'][i]).split(';')
                            upc_trim = [x.strip(' ') for x in upc_str]       
                            data_dict['upc'] = upc_trim
                        else:
                            data_dict['upc'] = str(data['upc'][i])
                            
                if 'date-available-on-market' in data.keys():
                    if data['date-available-on-market'][i]:    
                        data_dict['date-available-on-market'] = str(data['date-available-on-market'][i])

                if 'meets-most-efficient-criteria-2024' in data.keys():
                    if data['meets-most-efficient-criteria-2024'][i]:
                        data_dict['meets-most-efficient-criteria-2024'] = bool(data['meets-most-efficient-criteria-2024'][i])
                        
                if 'low-noise' in data.keys():
                    if data['low-noise'][i]:
                        data_dict['low-noise'] = bool(data['low-noise'][i])
                        
                if 'variable-speed-compressor' in data.keys():
                    if data['variable-speed-compressor'][i]:
                        data_dict['variable-speed-compressor'] = bool(data['variable-speed-compressor'][i])
                        
                if 'energy-star-lamp-included' in data.keys():
                    if data['energy-star-lamp-included'][i]:
                        data_dict['energy-star-lamp-included'] = bool(data['energy-star-lamp-included'][i])
                
                sort_dict=dict(sorted(data_dict.items()))
                data_list.append(sort_dict)

                with open(os.path.join(file_path_to_output, "JSON", file_name.replace(".xlsx", ".json")), 'w', encoding='utf-8') as fout:
                    for j in data_list:
                        json.dump(j , fout, indent=1, ensure_ascii=False, default=str)
                        fout.write('\n \n')   
                 
              
             
            jsonl_folder_path = os.path.join(file_path_to_output, "JSONL")
            os.makedirs(jsonl_folder_path, exist_ok=True)
            jsonl_filename = os.path.join(file_path_to_output, "JSONL", file_name.replace(".xlsx", ".jsonl"))
            
            with jsonlines.open(jsonl_filename, mode='w') as writer:
                for j in data_list:
                    writer.write(j)
    print("JSON AND JSONL FILES ARE DONE")
    creation_json(file_path_to_output, brands_mapping_location, map_api_location, map_location, last_month_file_location)