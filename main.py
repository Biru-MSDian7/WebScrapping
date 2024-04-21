from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import docx
import time




def extract_links_from_docx(file_path):
    doc = docx.Document(file_path)
    links_dict = {}
    title = None
    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            if paragraph.text.startswith("PFB the title wise reference links"):
                continue
            if paragraph.text.startswith("Next"):
                title = paragraph.text.strip()
                links_dict[title] = []
            elif title:
                links_dict[title].append(paragraph.text.strip())
    return links_dict

def scrape_data(links_dict):
    driver = webdriver.Chrome()  
    data = []
    for title, links in links_dict.items():
        for link in links:
            driver.get(link)
            time.sleep(2)  # Add a delay to ensure the page fully loads

            text = driver.find_element(By.TAG_NAME, 'p').text
            data.append({'Title': title, 'Link': link, 'Text': text})
    driver.quit()
    return data




docx_file_path = 'Assignment2.docx'
links_dict = extract_links_from_docx(docx_file_path)
scraped_data = scrape_data(links_dict)

df = pd.DataFrame(scraped_data)
df.to_excel('scraped_data.xlsx', index=False)
