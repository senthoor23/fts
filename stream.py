import streamlit as st
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from datetime import datetime
import os

def scrape_data(excel_file, output_file, username, password):
    df = pd.read_excel(excel_file, engine='openpyxl')
    sedol_values = df[['Sedol', 'Ex date']].dropna().values.tolist()

    edge_options = Options()
    edge_options.use_chromium = True
    service = Service(r"C:\path\to\msedgedriver.exe")
    driver = webdriver.Edge(service=service, options=edge_options)

    wb = load_workbook(output_file)
    ws_screenshots = wb["WebScrape Screenshots"]

    try:
        for index, (sedol_value, ex_date) in enumerate(sedol_values):
            try:
                if isinstance(ex_date, pd.Timestamp):
                    ex_date = ex_date.strftime('%d-%b-%y')
                ex_date_obj = datetime.strptime(ex_date, '%d-%b-%y')
                ex_date_formatted = ex_date_obj.strftime('%d %B %Y')

                driver.get(f'https://{username}:{password}@ids.interactivedata.com/ftsbin/w_client')
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

                input_box = driver.find_element(By.NAME, 'CCODE')
                input_box.send_keys(sedol_value)

                locate_button = driver.find_element(By.NAME, 'IMAGE')
                locate_button.click()

                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
                expand_all_link = driver.find_element(By.LINK_TEXT, 'Expand all')
                expand_all_link.click()

                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'sectionname')))

                try:
                    income_section = driver.find_element(By.XPATH, '//td[@id="nameselected"]/a[contains(text(), "Income")]')
                    print(sedol_value)
                except Exception:
                    print(f"Income section not found for SEDOL {sedol_value}.")
                    ws_screenshots[f'Q{2 + index * 35}'] = f"Income section not found for SEDOL {sedol_value}."
                    continue

                rows = driver.find_elements(By.XPATH, '//td[@id="subtable"]//tr')
                ex_date_found = False
                for row in rows:
                    try:
                        ex_date_text = row.find_element(By.XPATH, './td[contains(@class, "results")][7]').text.strip()
                        if ex_date_text == ex_date_formatted:
                            ex_date_found = True
                            dividend_detail_link = row.find_element(By.XPATH, './td[contains(@class, "results")][1]/a')
                            dividend_detail_link.click()

                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
                            screenshot_path = f'dividend_detail_{sedol_value}.png'
                            driver.save_screenshot(screenshot_path)

                            if os.path.exists(screenshot_path):
                                img = Image(screenshot_path)
                                img.width = 900
                                img.height = 600
                                ws_screenshots[f'Q{2 + index * 35}'] = sedol_value
                                ws_screenshots.add_image(img, f'Q{3 + index * 35}')
                            else:
                                print(f"Screenshot not found for SEDOL {sedol_value}.")
                            break
                    except Exception:
                        continue

                if not ex_date_found:
                    print(f"SEDOL {sedol_value} does not contain the Ex Date {ex_date_formatted}.")
                    ws_screenshots[f'Q{2 + index * 35}'] = f"SEDOL {sedol_value} does not contain the Ex Date {ex_date_formatted}."

            except Exception as e:
                print(f"An error occurred for SEDOL {sedol_value}: {e}")
                ws_screenshots[f'Q{2 + index * 35}'] = f"An error occurred for SEDOL {sedol_value}: {e}"
                continue

    finally:
        driver.quit()
        wb.save(output_file)

    print("Screenshots saved and pasted into output Excel file.")

st.title('Web Scraping with Selenium and Streamlit')
uploaded_file = st.file_uploader('Upload Excel File', type=['xlsx'])
output_file = st.file_uploader('Upload Excel File', type=['xlsx'])
username = st.text_input('Username')
password = st.text_input('Password', type='password')

if uploaded_file is not None:
    with open('uploaded_file.xlsx', 'wb') as f:
        f.write(uploaded_file.getbuffer())
    st.success('File uploaded successfully!')

if st.button('Scrape Data'):
    if uploaded_file is not None:
        scrape_data('uploaded_file.xlsx', output_file, username, password)
        st.success('Data scraped and saved successfully!')
    else:
        st.error('Please upload an Excel file first.')
