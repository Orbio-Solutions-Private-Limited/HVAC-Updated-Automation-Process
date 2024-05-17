# HVAC-Updated-Automation-Process

________________URL:  https://www.energystar.gov/productfinder/advanced ________________

Process Document: Automated Data Scraping, Mapping, and Comparison

1. Objective:
The objective of this process is to automate the extraction of data from the ENERGY STAR Certified Product Data Sets and APIs webpage, map the raw data according to customer requirements, create formatted UP files, convert them into JSON files, and compare them with the previous month's UP files to identify any new columns.
2. Steps:
Step 1: Data Scraping:
Utilize Python and relevant libraries (e.g., BeautifulSoup, requests) to scrape data from the ENERGY STAR webpage.
Identify and scrape data under the "Heating & Cooling" header and its 12 subcategories.
Handle cases where direct scraping is not possible by using HTML code.
Store the scraped data in Excel files into the RAW folder which is created by code.

Step 2: Data Mapping
Based on customer requirements, map the raw Excel files to create formatted UP files.
Also map the raw excel file by sing brands mapping and also map-api to create formatted files 
Theses formatted files are stored in the FORMATTED folder. 
Perform necessary data manipulation, cleaning, and formatting during the mapping process.

Step 3: JSON File Creation
Convert the formatted UP files into JSON files using Python.
Ensure the JSON files adhere to the specified structure and contain all necessary data fields.

Step 4: Comparison with Previous Month
Retrieve the UP files from the previous month.
Automatically compare the current formatted UP files with the previous month's UP files.
Identify any new columns added in the current files compared to the previous month.

Step 5: Automated Execution
Develop a Python script to automate the entire process.
Use spyder IDE to execute the script, providing all required file paths in the config.txt file as input 
Implement error handling and logging mechanisms to ensure smooth execution and troubleshooting.

5. Quality Assurance:
Implement quality control checks at each step of the process to ensure accuracy, completeness, and consistency of the data.
Conduct regular reviews and audits to identify and address any issues or discrepancies.
