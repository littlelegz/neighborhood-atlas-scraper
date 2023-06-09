import pandas as pd
import io, re, time
import msoffcrypto, getpass
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By

# Asking for user input
xlsx_file = input("Please enter the name of the excel file:\n")
xlsx_file = re.sub('.xlsx','', xlsx_file) #removing the .xlsx at the end of the name (if entered)

# Get password
password = getpass.getpass(prompt="Please enter the password of the file\n")

decrypted_workbook = io.BytesIO()

print("Opening " + xlsx_file + "...\n")

try:
    with open(f'{xlsx_file}.xlsx', 'rb') as file:
        office_file = msoffcrypto.OfficeFile(file)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted_workbook)
    df = pd.read_excel(decrypted_workbook)
except:
    print("Error reading file")
    exit(1)
print(f"Read {df.shape[0]} lines!\n")

addr_col = input("Please enter the col with addresses:\n")

addr_lst = df[addr_col]
print(f"Found {len(addr_lst)} addresses, first address: {addr_lst[0]}\n")

# Setting the html
url = "https://www.neighborhoodatlas.medicine.wisc.edu/mapping"

# Setting parameters for selenium to work
path = r'/home/jzhao128/bias_autopopulator/chromedriver.exe'
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--remote-debugging-port=9222") # Had a "it works on my computer moment"
driver = webdriver.Chrome(path, options=options)
driver.get(url)
# executable_path is deprecated, use Service object instead. Will proceed with this

# Click to acknowledge cookies
cookie_btn = driver.find_element(By.ID, 'cookieconsent-dismiss')
cookie_btn.click()

# Selecting State
state_select = Select(driver.find_element(By.NAME, 'map-state-init'))
state_select.select_by_visible_text('Maryland')
driver.implicitly_wait(10)

def fulfill_form(addr):
    # Targeting input field and search button
    input_addr = driver.find_element(By.XPATH, '/html/body/div[6]/div/div[4]/div[2]/input') # This is janky as hell, should use relative path
    submit = driver.find_element(By.CLASS_NAME, 'leaflet-control-search-button')

    #input the values and hold a bit for the next action
    input_addr.clear()
    input_addr.send_keys(addr)
    time.sleep(1)
    submit.click()
    time.sleep(1)

    # Grab output
    ret = []
    outputs = driver.find_elements(By.CLASS_NAME, 'pop-up-content')
    for output in outputs:
        ret.append(output.text)

    return ret

# Setting outputs
for idx, addr in enumerate(addr_lst):
    output = pd.NA
    try:
        output = fulfill_form(addr)
    except: 
        print(f"Form failed for row {idx + 1}")
    if output != pd.NA:
        df.at[idx, 'Decile']=output[0]
        df.at[idx, 'Percentile']=output[1]
    else:
        df.at[idx, 'Decile']=output
        df.at[idx, 'Percentile']=output
df.to_excel("output.xlsx")  


print("End!")