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
        self.driver.set_window_size(1920, 1080)
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
        self.driver.execute_script("window.scrollBy(0,500)","")
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
        else:
            os.makedirs(path)
            current = os.path.join(path, "current_folder")
            legacy = os.path.join(path, "legacy_folder")
            os.makedirs(current)
            os.makedirs(legacy)
            print("Folder Created Successfully..")
            return current

    def is_directory_empty(self, folder):
        contents = os.listdir(folder)
        return len(contents) == 0

    def move_csv_files(self, source, destination):
        files = os.listdir(source)
        for file in files:
            if file.endswith(".csv"):
                source_path = os.path.join(source, file)
                destination_path = os.path.join(destination, file)
                shutil.move(source_path, destination_path)
                print(f"Moved {file} to {destination}")

    def newest(self, path):
        files = os.listdir(path)
        return max(files, key=os.path.getctime)

    def pickle_dictionary(self):
        with open('filename_dictionary', 'wb') as file:
            pickle.dump(self.file_dic, file)


    def scrape_each_category(self):
        """ 
        scrape data and store in meta data file for 
        each sub-category..
        """
        for cat in self.categories:
            source = os.getcwd()
            destination = self.create_folder(cat)
            self.load_page()
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
                    dic['dataset_name'] = name
                    dic['dataset_description'] = description
                    dic['last_update_date'] = last_update
                    dic['released_date'] = released_date
                    # new_row = pd.DataFrame([dic])
                    # self.res = pd.concat([self.res, new_row], ignore_index = True)
                    # if self.is_directory_empty(destination):
                    #     wait = WebDriverWait(self.driver, 10)
                    #     tag = i.find_element(By.LINK_TEXT, 'Download CSV')
                    #     download = wait.until(EC.element_to_be_clickable(tag))
                    #     self.driver.execute_script("arguments[0].click();", download)
                    #     print('Downloading file...')
                    #     time.sleep(3)
                    #     button = self.driver.find_element(By.XPATH, "/html//div[@id='main-content']/div//ol[@class='search-list']//dialog/div[@role='document']/main[@role='main']//button[.='Yes, download']")
                    #     button.click()
                    #     print('Download Success')
                    #     time.sleep(30)
                    #     file_dictionary[name] = self.newest(source) 
                    new_row = pd.DataFrame([dic])
                    self.res = pd.concat([self.res, new_row], ignore_index = True)

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
            self.res = pd.DataFrame()
            print("Scrapping Successful in " + cat + " page")
            if self.is_directory_empty(destination):
                self.move_csv_files(source, destination)
                self.pickle_dictionary()



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
