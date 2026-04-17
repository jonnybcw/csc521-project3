from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
import base64
import re

url = 'https://www.fire.ca.gov/incidents/2025/1/7/palisades-fire/updates'

options = Options()
# options.add_argument("--headless") # Runs browser in the background
driver = webdriver.Chrome(options=options)

driver.get("https://www.fire.ca.gov/incidents/2025/1/7/palisades-fire/updates")

# Give the page 3 seconds to load the dynamic content
time.sleep(3) 

# Now the 'detail-page' content will be there
html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')
detail = soup.find('div', class_='detail-page')
updates = detail.find_all('a')
updates.reverse()

print_options = {
    'landscape': False,
    'displayHeaderFooter': False,
    'printBackground': True,
    'preferCSSPageSize': True,
}

for update in updates:
    link = update['href']
    driver2 = webdriver.Chrome(options=options)
    driver2.get('https://www.fire.ca.gov' + link)
    pagename = driver2.title
    pagename = re.sub(r'[:/|]', '_', pagename)
    time.sleep(3)
    result = driver2.execute_cdp_cmd("Page.printToPDF", print_options)
    with open(f"reports/{pagename}.pdf", "wb") as file:
        file.write(base64.b64decode(result['data']))

    driver2.quit()


driver.quit()
        