# Import Libraries
import time
import sys
import os
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By

file_drop_path = r'C:\Users\allensworthc\Desktop\scraping' # Create variable for destination file path
our_files = Path(file_drop_path)

# Define function to wait for file to download before proceeding
def download_wait(path_to_downloads): 
    seconds = 60
    dl_wait = True
    while dl_wait:
        time.sleep(seconds)
        dl_wait = False
        for fname in os.listdir(path_to_downloads):
            if fname.endswith('.crdownload'):
                dl_wait = True
        seconds = 5
    return "File download complete!"

# Define function to find and store the name of the most recently downloaded file
def latest_download_file(path_to_downloads):
    path = path_to_downloads
    os.chdir(path)
    files = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)
    newest = files[-1]
    return newest

# Set Chrome OptionsFY2
chrome_options = webdriver.ChromeOptions() # Creates variable chrome_options to store custom options into webdriver
chrome_options.add_argument('--ignore-certificate-errors') # Ignore unnecessary logging errors 
chrome_options.add_argument('--ignore-ssl-errors') # Ignore unnecessary logging errors
chrome_options.add_argument('--start-maximized') # Ignore unnecessary logging errors
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging']) # Ignore unnecessary logging errors
prefs = {"download.default_directory" : file_drop_path} # Creates variable 'prefs' that saves files to shared drive in correct folder
chrome_options.add_experimental_option("prefs", prefs) # Calls prefs into webdriver

# Enter Cycle Date
fy = str(input("Enter Cycle Fiscal Year (ex. FY22): ")) # Asking user to input what fiscal year the report is for
cm = str(input("Enter Cycle Month and Year (ex. March 2022): ")) # Asking user to input what calendar month the report is for
sd = str(input("Enter Cycle Start Date (ex. 02/22/2022): ")) # Asking user for start date input
ed = str(input("Enter Cycle End Date (ex. 03/21/2022): ")) # Asking user for end date input


# Log In Steps
driver = webdriver.Chrome(r'C:\Users\allensworthc\Downloads\chromedriver_win32\chromedriver', options=chrome_options) # Using Chrome to access web
driver.get('https://www.access.usbank.com/cpsApp1/AxolPreAuthServlet?requestCmdId=login') # Open the website
org_box = driver.find_element(By.NAME, 'RLTSHIPSHORTNAME').send_keys('removed') # Enter Org Short Name
id_box = driver.find_element(By.NAME, 'USERID').send_keys('removed') # Enter User Name
pass_box = driver.find_element(By.NAME, 'PSWRD').send_keys('removed') # Enter password
login_button = driver.find_element(By.NAME, 'Login').click() # Click login
time.sleep(5) # Allow time for browser to load
rep_button = driver.find_element(By.LINK_TEXT, 'Reporting').click() # Click reporting button
time.sleep(5) # Allow time for browser to load


# Export Transaction Details Report
fm_button = driver.find_element(By.LINK_TEXT, 'Financial Management').click() # Select the 'Financial Management' hyperlink
time.sleep(2) # Allow time for browser to load
td_button = driver.find_element(By.LINK_TEXT, 'Transaction Detail').click() # Select the 'Transaction Detail' hyperlink
time.sleep(2) # Allow time for browser to load
dt_radio = driver.find_element(By.ID, 'DATE_TYPE2').click() # Select Posting Date Range
start_date = driver.find_element(By.NAME, 'BEGIN_DATE~D~~N').clear() # Remove default begin date
start_date = driver.find_element(By.NAME, 'BEGIN_DATE~D~~N').send_keys(sd) # Enter current cycle begin date
end_date = driver.find_element(By.NAME, 'END_DATE_UI~D~~N').clear() # Remove default end date
end_date = driver.find_element(By.NAME, 'END_DATE_UI~D~~N').send_keys(ed) # Enter current cycle end date
pay_radio = driver.find_element(By.ID, 'PAYMENTS_IND1').click() # Include Payments
fee_radio = driver.find_element(By.ID, 'FEES_IND1').click() # Include Fees
dtcf = driver.find_element(By.ID, 'TRANS_CUST').click() # Display Transaction Custom Fields
dtc = driver.find_element(By.ID, 'TRANS_CMT').click() # Display Transaction Comments
alloc = driver.find_element(By.ID, 'SHOWALLOC').click() # Display Allocation Detail
merc = driver.find_element(By.ID, 'MERCH_DATA').click() # Display Merchant Data
win_handle_before = str(driver.current_window_handle) # Capture window handle of current browser window
rr_trans = driver.find_element(By.NAME, "RunReport").click() # Run Report
print(download_wait(file_drop_path)) # Call the dowload_wait function to stall script until file is downloaded
trans_file = latest_download_file(file_drop_path) # Call the latest_download_file function to capture the name of the file in 'trans_file'

# Move file to correct folder and rename file
for file in our_files.iterdir():
    tem_file = f'{file.stem}{file.suffix}'
    if tem_file == trans_file:
        directory = file.parent
        extension = file.suffix
        new_name = f'{"Transaction Detail"} - {file.stem}{extension}'
        new_fy_path = our_files.joinpath(fy)
        if not new_fy_path.exists():
            new_fy_path.mkdir()
        new_path = new_fy_path.joinpath(cm)
        if not new_path.exists():
            new_path.mkdir()
        new_file_path = new_path.joinpath(new_name)
        file.replace(new_file_path)
time.sleep(2)

# Export Merchant Details Report
driver.switch_to.window(win_handle_before) # Switch back to original window
rep_button = driver.find_element(By.LINK_TEXT, 'Reporting').click() # Click reporting button
time.sleep(2) # Allow time for browser to load
sm_button = driver.find_element(By.LINK_TEXT, 'Supplier Management').click() # Click supplier management button
msali_button = driver.find_element(By.LINK_TEXT, 'Merchant Spend Analysis by Line Item').click() # Click Merchant Spend Analysis by Line Item button
dt_radio = driver.find_element(By.ID, 'DATE_TYPE2').click() # Select Posting Date Range
start_date = driver.find_element(By.NAME, 'BEGIN_DATE~D~~N').clear() # Remove default begin date
start_date = driver.find_element(By.NAME, 'BEGIN_DATE~D~~N').send_keys(sd) # Enter current cycle begin date
end_date = driver.find_element(By.NAME, 'END_DATE_UI~D~~N').clear() # Remove default end date
end_date = driver.find_element(By.NAME, 'END_DATE_UI~D~~N').send_keys(ed) # Enter current cycle end date
merc = driver.find_element(By.ID, 'MERCH_DATA').click() # Include Additional Merchant Data
search_for_button = driver.find_element(By.LINK_TEXT, 'Search for Position or Add Multiple').click() # Click Search For Button
time.sleep(2) # Allow time for browser to load
search_button = driver.find_element(By.NAME, 'SearchButton').click() # Click Search Button
time.sleep(2) # Allow time for browser to load
check_all = driver.find_element(By.LINK_TEXT, 'Check All Shown').click() # Click check all boxes button
select_pos_button = driver.find_element(By.NAME, "SelectMultipleProcessingButton").click() # Click Select Postion Button

for i in range(2,7): # Loop through pages 2-6
    click_button = driver.find_element(By.LINK_TEXT, str(i)).click() # Change page
    time.sleep(2) # Allow time for browser to load
    check_all = driver.find_element(By.LINK_TEXT, 'Check All Shown').click() # Click check all boxes button
    select_pos_button = driver.find_element(By.NAME, "SelectMultipleProcessingButton").click() # Click Select Position Button

acc_hier = driver.find_element(By.NAME, "AcceptHierarchyButton").click() # Click Accep Hierarchy Button
rr_trans = driver.find_element(By.NAME, "RunReport").click() # Run Report
print(download_wait(file_drop_path)) # Call the dowload_wait function to stall script until file is downloaded
merch_file = latest_download_file(file_drop_path) # Call the latest_download_file function to capture the name of the file in 'merch_file'

# Move file to correct folder and rename file
for file in our_files.iterdir():
    tem_file = f'{file.stem}{file.suffix}'
    if tem_file == merch_file:
        directory = file.parent
        extension = file.suffix
        new_name = f'{"Merchant Detail"} - {file.stem}{extension}'
        new_fy_path = our_files.joinpath(fy)
        if not new_fy_path.exists():
            new_fy_path.mkdir()
        new_path = new_fy_path.joinpath(cm)
        if not new_path.exists():
            new_path.mkdir()
        new_file_path = new_path.joinpath(new_name)
        file.replace(new_file_path)

sys.exit() # Closes Script