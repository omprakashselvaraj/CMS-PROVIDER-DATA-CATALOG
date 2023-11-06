from scraping import WebScrape

if __name__ == '__main__':
    base_link = 'https://data.cms.gov/provider-data/'
    ws = WebScrape(base_link)
    ws.create_metadata_file()
    ws.load_page()
    ws.find_link_text()
    ws.scrape_each_category()
    ws.dataframe_to_excel()
    ws.pickle_dictionary()
    print(ws.df_dic.keys())