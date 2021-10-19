#Import all the packages
import os
import pandas as pd
from selenium.webdriver.chrome.options import Options
import selenium.webdriver as webdriver
import time
from email.message import EmailMessage
import smtplib
from conf import query, num_page, receiver, sender_email, sender_password

query_link = f"https://www.semanticscholar.org/search?q={query}&sort=relevance&page="


# Create working paths
work_dir = os.path.dirname(os.path.realpath(__file__))
folder = os.path.join(work_dir, "articles")
webdrive = os.path.join(work_dir, "chromedriver")  


# Chek the existance of an articles directory 
if not os.path.isdir(folder):
    os.mkdir(folder)

# Webdriver
cd = Options()
preferences = {"download.default_directory": folder, "download.prompt_for_download": False}
cd.add_experimental_option('prefs', preferences)
os.environ["webdriver.chrome.driver"] = webdrive   


#Create links to follow
links_list = [query_link + str(page+1) for page in range(num_page)] 

driver = webdriver.Chrome(executable_path=webdrive, options=cd)

final_info = []  

for search_link in links_list:
    # Get all links to articles from the page
    iterator = 0 
    driver.get(search_link)
    time.sleep(5)
    articles = driver.find_elements_by_xpath("//*[@data-selenium-selector='title-link']")
    articles_links = []
    articles_dates = []
    temp_dates = driver.find_elements_by_class_name('cl-paper-pubdates')
    for data in temp_dates:
        articles_dates.append(data.text)

    for article in articles:
        try:
            link = article.get_attribute("href")
            articles_links.append(link)

        except:
            pass

    for link in articles_links:
        # get info of each article 
        tmp_info = {}

        driver.get(link)

        title = driver.find_elements_by_xpath("//*[@data-selenium-selector='paper-detail-title']")[0].text
        author = driver.find_elements_by_class_name('paper-meta-item')[0].text
        tmp_info.update({
                        'title': title,
                        'date' : articles_dates[iterator],
                        'authors': author
                        })

        # Download the article's document
        try:
            initial_dir = os.listdir(folder)
            driver.find_element_by_xpath("//*[@class='alternate-sources__dropdown-wrapper']").click()
            time.sleep(5)

            current_dir = os.listdir(folder)
            filename = list(set(current_dir) - set(initial_dir))[0]
            full_path = os.path.join(folder, filename)

        except Exception as e:
            full_path = None

        tmp_info.update({'path_to_file':full_path})

        final_info.append(tmp_info.copy())
        time.sleep(2)

        iterator +=1

driver.quit()

# Create an excel file
df = pd.DataFrame(final_info)
excel_path = os.path.join(work_dir, "articles.xlsx")
df.to_excel(excel_path, index=False)

# Create an email and add excel file
login, password = sender_email, sender_password   
mail = EmailMessage()
mail['From'] = login
mail['To'] = receiver
mail['Subject'] = "Topics analysis"
mail.set_content("Dear Alexandra,\n\nHere is your homework:)\n\n Good luck!")

with open(excel_path, 'rb') as f:
    file_data = f.read()
    file_name = f'articles_info.xlsx'
mail.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

# Send email
server = smtplib.SMTP('smtp.office365.com')  
server.starttls()  
server.login(login, password)    
server.send_message(mail)      
server.quit()      
