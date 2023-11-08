from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from openpyxl import load_workbook
import pandas as pd
import pickle
import xlsxwriter
import warnings
import time
import os
import shutil
warnings.filterwarnings("ignore", category=DeprecationWarning) 

class WebScrape():
    def __init__(self, page_url):
        """ 
        constructor initialization..
        """
        self.page_url = page_url
        self.chrome_options = Options()
        self.chrome_options.add_experimental_option("prefs", {
        "download.default_directory": os.getcwd(),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
        })
        self.driver = webdriver.Chrome(options = self.chrome_options)
        #self.driver.set_window_size(1920, 1080)
        self.res = pd.DataFrame()
        self.categories = None
        self.file = 'metadata_file.xlsx'
        self.parent_directory = os.getcwd()
        self.df_dic = {}
        self.file_dic = {}

    def load_page(self):
        """ 
        Load the page..
        """
        # self.driver = webdriver.Chrome(options = chrome_options)
        self.driver.get(self.page_url)
        time.sleep(5)
        #self.driver.execute_script("window.scrollBy(0,500)","")
        print('Page Loaded Successfully')

    def close_page(self):
        """ 
        close the driver
        """
        self.driver.quit()
        print('Driver is closed')

    def find_link_text(self):
        """ 
        find link_text for all 10 categories in web page
        """
        ul = self.driver.find_element(By.XPATH, "//div[@id='data-categories']//ul[@class='pqdc-icon-list']")
        options = ul.find_elements(By.TAG_NAME, "li")
        link_text = []
        for option in options:
            links = option.find_elements(By.TAG_NAME, "a")
            text = links[0].find_elements(By.TAG_NAME, "div")[0].text
            link_text.append(text)
        self.categories = link_text

    def create_metadata_file(self):
        """
        create a metadata file if not exist in directory..
        """
        if not os.path.isfile(self.file):
            df = pd.DataFrame()

            # Create an ExcelWriter to save the DataFrame to an Excel file
            with pd.ExcelWriter('metadata_file.xlsx', engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Sheet1', index=False)
        else:
            print("File already exist!!!")

    def create_folder(self, new_folder):
        """
        create folder for each sub-category if not exist..
        """
        path = os.path.join(self.parent_directory, new_folder)
        if os.path.exists(path):
            print('Folder already exist!!!')
            return path
        else:
            os.makedirs(path)
            print("Folder Created Successfully !!!")
            return path

    def create_current_legacy(self, category_folder, folder_name):
        if os.path.exists(category_folder):
            category_folder_path = os.path.join(category_folder, folder_name)
            os.makedirs(category_folder_path)
            current_path = os.path.join(category_folder_path, "current")
            legacy_path = os.path.join(category_folder_path, "legacy")
            os.makedirs(current_path)
            os.makedirs(legacy_path)
            return category_folder_path, current_path, legacy_path


    def is_directory_empty(self, folder):
        contents = os.listdir(folder)
        return len(contents) == 0

    def check_folder_exist(self, category_folder, folder):
        path = os.path.join(category_folder, folder)
        return os.path.exists(path)

    def move_csv_files(self, source, destination, file):
        files = os.listdir(source)
        if file in files:
            source_path = os.path.join(source, file)
            destination_path = os.path.join(destination, file)
            try:
                shutil.move(source_path, destination_path)
                print(print(f"Moved {file} to {destination}"))
            except FileNotFoundError:
                print(f"File not found in the source directory: {file}")


    def newest(self, path):
        files = os.listdir(path)
        return max(files, key=os.path.getctime)

    def pickle_dictionary(self):
        file_name = 'filename_dictionary'
        try:
            with open(file_name, 'rb') as file:
                existing_data = pickle.load(file)
        except FileNotFoundError:
            # If the file doesn't exist, create a new dictionary
            existing_data = {}
        existing_data.update(self.file_dic)

        # Write the updated dictionary back to the pickle file
        with open(file_name, 'wb') as file:
            pickle.dump(existing_data, file)

    def download(self, i, name):
        wait = WebDriverWait(self.driver, 10)
        tag = i.find_element(By.LINK_TEXT, 'Download CSV')
        download = wait.until(EC.element_to_be_clickable(tag))
        self.driver.execute_script("arguments[0].click();", download)
        print('Downloading file...')
        time.sleep(3)
        button = self.driver.find_element(By.XPATH, "/html//div[@id='main-content']/div//ol[@class='search-list']//dialog/div[@role='document']/main[@role='main']//button[.='Yes, download']")
        button.click()
        print('Download Success')
        if name == "National Downloadable File":
            time.sleep(180)
        else:
            time.sleep(30)
    
    def last_date_same_asin_metadata(self, subcategory,dataset_name,updated_date):
        # Provide the file path of your Excel file
        file_path = 'metadata_file.xlsx'

        # Use pandas to read the Excel file and get all sheet names
        excel_file = pd.ExcelFile(file_path, engine='openpyxl')
        sheet_names = excel_file.sheet_names

        if subcategory in sheet_names:
            df = pd.read_excel(excel_file, sheet_name=subcategory)
            
            if dataset_name in df['dataset_name'].values:
                    
                if not df[(df['dataset_name'] == dataset_name) & (df['last_update_date'] == updated_date)].empty:
                    return True # dates are same
                else:
                    return False # dates do not match 



    def scrape_each_category(self):
        """ 
        scrape data and store in meta data file for 
        each sub-category..
        """
        for cat in self.categories:
            source = os.getcwd()
            category_folder = self.create_folder(cat)
            self.load_page()
            scroll_distance = 600
            self.driver.execute_script(f"window.scrollBy(0, {scroll_distance});")
            element = self.driver.find_element(By.LINK_TEXT, cat)
            self.driver.execute_script("arguments[0].click();", element)
            time.sleep(5)
            print("Scrapping Data in " + cat + " page")
            file_dictionary = {}
            while True:
                opt = self.driver.find_element(By.XPATH, "/html//div[@id='main-content']/div//ol[@class='search-list']")
                li = opt.find_elements(By.CLASS_NAME,"search-list-item")
                for i in li:
                    dic = {}
                    h2 = i.find_element(By.TAG_NAME, "h2")
                    description = i.find_element(By.TAG_NAME, "p").text
                    name = h2.find_element(By.TAG_NAME, "a").text
                    date = i.find_element(By.CLASS_NAME,"dataset-date")
                    date_list = date.find_elements(By.CLASS_NAME, "dataset-date-item")
                    last_update = date_list[0].text.split(': ')[1]
                    released_date = date_list[1].text.split(': ')[1]
                    if ':' in name:
                        name = name.replace(':','')
                    dic['dataset_name'] = name
                    dic['dataset_description'] = description
                    dic['last_update_date'] = last_update
                    dic['released_date'] = released_date
                    new_row = pd.DataFrame([dic])
                    self.res = pd.concat([self.res, new_row], ignore_index = True)

                    if not self.check_folder_exist(category_folder, name):
                        category_folder_path, current_path, legacy_path = self.create_current_legacy(category_folder, name)
                        self.download(i, name)
                        file_name = self.newest(source)
                        file_dictionary[name] = file_name
                        self.move_csv_files(source, current_path, file_name)
                    else:
                        updated_date_status = self.last_date_same_asin_metadata(cat, name, last_update)
                        if updated_date_status:
                            pass
                        else:
                            self.download(i, name)
                            file_name = self.newest(source)
                            file_dictionary[name] = file_name
                            current = os.path.join(cat, name, "current")
                            legacy = os.path.join(cat, name, "legacy")
                            self.move_csv_files(current, legacy, file_name)
                            self.move_csv_files(source, current, file_name)


                try:
                    next_button = self.driver.find_element(By.XPATH, "/html//div[@id='main-content']/div/div[@class='row']/div[1]/div[@class='pagination-wrapper']/div/div[@class='pagination-container']/nav[@class='ds-c-pagination']/button[2]")
                    if next_button.get_property('disabled'): 
                        break
                    else:
                        wait = WebDriverWait(self.driver, 10)
                        wait.until(EC.element_to_be_clickable((By.XPATH, "/html//div[@id='main-content']/div/div[@class='row']/div[1]/div[@class='pagination-wrapper']/div/div[@class='pagination-container']/nav[@class='ds-c-pagination']/button[2]")))

                        # Scroll to the "Next" button
                        self.driver.execute_script("arguments[0].scrollIntoView();", next_button)

                        # Click the "Next" button using JavaScript
                        self.driver.execute_script("arguments[0].click();", next_button)
                        time.sleep(10)
                except NoSuchElementException:
                    break
            self.df_dic[cat] = self.res
            self.file_dic[cat] = file_dictionary
            self.res = pd.DataFrame()
            print("Scrapping Successful in " + cat + " page")
            

    def dataframe_to_excel(self):
        book = load_workbook("metadata_file.xlsx")
        with pd.ExcelWriter("metadata_file.xlsx", engine='openpyxl', mode='a') as writer:
            for sheet_name, df in self.df_dic.items():
                if sheet_name in writer.sheets:
                    sheet = book[sheet_name]
                    book.remove(sheet)
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    df.to_excel(writer, sheet_name=sheet_name, index=False) 
