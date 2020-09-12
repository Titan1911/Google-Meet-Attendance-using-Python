# web driver imports
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
# excel file imports
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
# time import
from time import sleep

gmail_id = input('Enter your gmail id: ')
gmail_password = input('Enter your gmail password: ')
meet_link = input('Paste the link of the Google Meet: ')

# Some necessary things for automation with google driver
opt = Options()
opt.add_argument("--disable-infobars")
opt.add_argument("start-maximized")
opt.add_argument("--disable-extensions")

# This part allows the notifications, mic and camera permissions
# Pass the argument 1 to allow and 2 to block
opt.add_experimental_option("prefs", {
    "profile.default_content_setting_values.media_stream_mic": 1,
    "profile.default_content_setting_values.media_stream_camera": 1,
    "profile.default_content_setting_values.geolocation": 1,
    "profile.default_content_setting_values.notifications": 1
  })

# Sign in to google through stackoverflow
driver = webdriver.Chrome(options=opt, executable_path=r"Enter your path here")
driver.get('https://stackoverflow.com/users/signup?ssrc=head&returnurl=%2fusers%2fstory%2fcurrent%27')
sleep(2)
driver.find_element_by_xpath('//*[@id="openid-buttons"]/button[1]').click()
driver.find_element_by_xpath('//input[@type="email"]').send_keys(gmail_id)
driver.find_element_by_xpath('//*[@id="identifierNext"]').click()
sleep(2)
driver.find_element_by_xpath('//input[@type="password"]').send_keys(gmail_password)
driver.find_element_by_xpath('//*[@id="passwordNext"]').click()
sleep(2)

driver.get('https://meet.google.com/')

# Enter the meeting
# Case when logged in with personal gmail account
try:
    driver.find_element_by_xpath('//*[@id="i3"]').send_keys(meet_link)  # Enter a code or link
    sleep(1)
    driver.find_element_by_xpath('//*[@id="yDmH0d"]/c-wiz/div/div[2]/div/div[1]/div[3]/div[2]/div[2]/button').click()  # join
    sleep(3)

    try:
        driver.find_element_by_xpath('//*[@id="yDmH0d"]/c-wiz/div/div/div[5]/div[3]/div/div[2]/div/div/div[1]/div[1]/div[3]/div[1]/div/div/div').click()  # camera
        driver.find_element_by_xpath('//*[@id="yDmH0d"]/c-wiz/div/div/div[5]/div[3]/div/div[2]/div/div/div[1]/div[1]/div[3]/div[2]/div/div').click()  # mic
        driver.find_element_by_xpath('//*[@id="yDmH0d"]/c-wiz/div/div/div[5]/div[3]/div/div[2]/div/div/div[2]/div/div[2]/div/div[1]/div[1]').click()  # join now
        sleep(2)
        driver.find_element_by_xpath('//*[@id="ow3"]/div[1]/div/div[5]/div[3]/div[6]/div[3]/div/div[2]/div[1]').click()  # participant list
        sleep(2)
        n = int(driver.find_element_by_xpath('//*[@id="ow3"]/div[1]/div/div[5]/div[3]/div[3]/div/div[2]/div[2]/div[1]/div[1]/span/div/span[2]').text.strip(')').strip('('))
        names = []
        for e in range(2, n+1):
            name = driver.find_element_by_xpath(f'//*[@id="ow3"]/div[1]/div/div[5]/div[3]/div[3]/div/div[2]/div[2]/div[2]/span[1]/div[2]/div[{e}]/div/div/div[2]/div[1]/div').text.lower()  # names of the participants
            names.append(name)

    except:
        driver.find_element_by_xpath('//*[@id="yDmH0d"]/c-wiz/div/div/div[4]/div[3]/div/div[2]/div/div/div[1]/div[1]/div[3]/div[1]/div/div/div').click()  # camera
        driver.find_element_by_xpath('//*[@id="yDmH0d"]/c-wiz/div/div/div[4]/div[3]/div/div[2]/div/div/div[1]/div[1]/div[3]/div[2]/div/div').click()  # mic
        driver.find_element_by_xpath('//*[@id="yDmH0d"]/c-wiz/div/div/div[4]/div[3]/div/div[2]/div/div/div[2]/div/div[2]/div/div[1]/div[1]').click()  # join now
        sleep(2)
        driver.find_element_by_xpath('//*[@id="ow3"]/div[1]/div/div[4]/div[3]/div[6]/div[3]/div/div[2]/div[1]').click()  # participant list
        sleep(2)
        n = int(driver.find_element_by_xpath('//*[@id="ow3"]/div[1]/div/div[4]/div[3]/div[3]/div/div[2]/div[2]/div[1]/div[1]/span/div/span[2]').text.strip(')').strip('('))  # no. of participants
        names = []
        for e in range(2, n+1):
            name = driver.find_element_by_xpath(f'//*[@id="ow3"]/div[1]/div/div[4]/div[3]/div[3]/div/div[2]/div[2]/div[2]/span[1]/div[2]/div[{e}]/div/div/div[2]/div[1]/div').text.lower()  # names of the participants
            names.append(name)

# Case when logged in with some organisation gmail account
except:
    driver.find_element_by_xpath('//*[@id="yDmH0d"]/c-wiz/div/div/div/div[2]/div[2]/div[2]/div/c-wiz/div[1]/div/div/div[1]').click()  # join or start the meeting
    sleep(1)
    driver.find_element_by_xpath('//*[@id="yDmH0d"]/div[3]/div/div[2]/span/div/div[2]/div[1]/div[1]/input').send_keys('https://meet.google.com/yjs-njdf-fyd')  # entering the link
    sleep(0.7)
    driver.find_element_by_xpath('//*[@id="yDmH0d"]/div[3]/div/div[2]/span/div/div[4]/div[2]/div').click()  # continue
    sleep(3)

    try:
        driver.find_element_by_xpath('//*[@id="yDmH0d"]/c-wiz/div/div/div[5]/div[3]/div/div[2]/div/div/div[1]/div[1]/div[3]/div[1]/div/div/div').click()  # camera
        driver.find_element_by_xpath('//*[@id="yDmH0d"]/c-wiz/div/div/div[5]/div[3]/div/div[2]/div/div/div[1]/div[1]/div[3]/div[2]/div/div').click()  # mic
        driver.find_element_by_xpath('//*[@id="yDmH0d"]/c-wiz/div/div/div[5]/div[3]/div/div[2]/div/div/div[2]/div/div[2]/div/div[1]/div[1]').click()  # join now
        sleep(2)
        driver.find_element_by_xpath('//*[@id="ow3"]/div[1]/div/div[5]/div[3]/div[6]/div[3]/div/div[2]/div[1]').click()  # participant list
        sleep(2)
        n = int(driver.find_element_by_xpath('//*[@id="ow3"]/div[1]/div/div[5]/div[3]/div[3]/div/div[2]/div[2]/div[1]/div[1]/span/div/span[2]').text.strip(')').strip('('))  # no. of participants
        names = []
        for e in range(2, n+1):
            name = driver.find_element_by_xpath(f'//*[@id="ow3"]/div[1]/div/div[5]/div[3]/div[3]/div/div[2]/div[2]/div[2]/span[1]/div/div[{e}]/div/div/div[2]/div[1]/div').text.lower()  # names of the participants
            names.append(name)

    except:
        driver.find_element_by_xpath('//*[@id="yDmH0d"]/c-wiz/div/div/div[4]/div[3]/div/div[2]/div/div/div[1]/div[1]/div[3]/div[1]/div/div/div').click()  # camera
        driver.find_element_by_xpath('//*[@id="yDmH0d"]/c-wiz/div/div/div[4]/div[3]/div/div[2]/div/div/div[1]/div[1]/div[3]/div[2]/div/div').click()  # mic
        driver.find_element_by_xpath('//*[@id="yDmH0d"]/c-wiz/div/div/div[4]/div[3]/div/div[2]/div/div/div[2]/div/div[2]/div/div[1]/div[1]').click()  # join now
        sleep(2)
        driver.find_element_by_xpath('//*[@id="ow3"]/div[1]/div/div[4]/div[3]/div[6]/div[3]/div/div[2]/div[1]').click()  # participant list
        sleep(2)
        n = int(driver.find_element_by_xpath('//*[@id="ow3"]/div[1]/div/div[4]/div[3]/div[3]/div/div[2]/div[2]/div[1]/div[1]/span/div/span[2]').text.strip(')').strip('('))  # no. of participants
        names = []
        for e in range(2, n+1):
            name = driver.find_element_by_xpath(f'//*[@id="ow3"]/div[1]/div/div[4]/div[3]/div[3]/div/div[2]/div[2]/div[2]/span[1]/div/div[{e}]/div/div/div[2]/div[1]/div').text.lower()  # names of the participants
            names.append(name)

# Marking attendance in excel sheet
wb = load_workbook(f'Google_Attendance.xlsx')
sheet = wb['Attendance Sheet']

for name in names:
    for box in range(2, sheet.max_row+1):
        if name == (sheet.cell(row=box, column=column_index_from_string('A')).value.lower()):
            sheet[f'B{box}'] = 'Present'
        else:
            sheet[f'B{box}'] = 'Absent'

wb.save('Google_Attendance.xlsx')
print(f'No. of participants : {n-1}')
print('Attendance taken successfully!')

