#! python3
# workflow.py - Automate a workflow entry

import webbrowser, time, pyautogui
from datetime import datetime


username = input('Username: ')
#username = 'OLV300264'
password = input('Password: ')
#password = 'NYO855th'
inv_amount = '51.17'
NYO_amount = '40.94'
DT_amount = '10.23'



# Open browser to workflow website
url = 'http://flow.fairtrade.com.tw:8080/WebAgenda/'
chrome_path = '"C:\Program Files\Google\Chrome\Application\chrome.exe" %s'
webbrowser.get(chrome_path).open(url)
time.sleep(5)

# Enter Username/Password
pyautogui.write(username)
pyautogui.write('\t')
pyautogui.write(password)
pyautogui.write('\t')
pyautogui.write('\t')
pyautogui.write('\n')
time.sleep(5)

pyautogui.moveTo(168, 198)
pyautogui.click(194, 320)

time.sleep(5)

for i in range(0, 15):
    pyautogui.write('\t')
    i += 1

pyautogui.press('down')
pyautogui.press('down')
pyautogui.write('\t')

# Stores current month in string variable
now = datetime.now()
month = now.strftime('%B')   # Displays the current month’s name

pyautogui.write('NYO Drinking Water Rental ' + month)

pyautogui.write('\t')
pyautogui.write('\t')
pyautogui.write('\t')
pyautogui.write('\t')
pyautogui.write('\n')
time.sleep(5)

# Click payment option
pyautogui.click(728, 559)

# Open dropdown menu and choose NY
pyautogui.click(808, 583)
pyautogui.press('down')
pyautogui.press('down')
pyautogui.press('down')
pyautogui.press('down')
pyautogui.write('\t')
time.sleep(5)

# Open dropdown menu and choose NYO
pyautogui.click(912, 583)
pyautogui.press('down')
pyautogui.write('\t')

# Enter percentage for NYO
pyautogui.click(1124, 585)
pyautogui.write('80%')


# Enter NYO amount
pyautogui.click(1230, 583)
pyautogui.write(NYO_amount)


# Enter notes for NYO
pyautogui.click(788, 613)
pyautogui.write('NYO Drinking Water Rental ' + month)


# Enter business unit
pyautogui.click(1136, 609)
pyautogui.press('down')
pyautogui.press('\t')


#Add NYO
pyautogui.click(1238, 611)
time.sleep(5)


# Open dropdown menu and choose Other Location
pyautogui.click(808, 583)
pyautogui.press('down')
pyautogui.press('down')
pyautogui.press('down')
pyautogui.write('\t')
time.sleep(5)

# Open dropdown menu and choose Design Team
pyautogui.click(912, 583)
for i in range(0,8):
    pyautogui.press('down')
    i += 1
pyautogui.write('\t')

# Enter percentage for DT
pyautogui.click(1124, 585)
pyautogui.write('20%')


# Enter DT amount
pyautogui.click(1230, 583)
pyautogui.write(DT_amount)


# Enter notes for DT
pyautogui.click(788, 613)
pyautogui.write('NYO Drinking Water Rental ' + month)


# Enter business unit
pyautogui.click(1136, 609)
pyautogui.press('up')
pyautogui.press('\t')


#Add DT
pyautogui.click(1238, 611)
time.sleep(3)

# Enter Company name
pyautogui.click(668, 957)
pyautogui.write('Corporate Coffee Systems')

# Enter address
pyautogui.click(670, 991)
pyautogui.write('745 Summa Avenue\nWestbury\nNY 11590')

# Enter Due Date
pyautogui.click(1170, 1027)
time.sleep(3)
pyautogui.click(1018, 657)

# Scroll down
pyautogui.scroll(-1100)

# Enter NYO Itemized
pyautogui.click(684, 345)
pyautogui.write('NYO Drinking Water Rental (NYO 80%)')

# Enter NYO Unit Price
pyautogui.click(992, 343)
pyautogui.hotkey('ctrl', 'a')
pyautogui.write(NYO_amount)

# Enter NYO quantity
pyautogui.click(1088, 343)
pyautogui.hotkey('ctrl', 'a')
pyautogui.write('1')

# Add NYO entry
pyautogui.click(1230, 295)
time.sleep(3)

# Enter DT Itemized
pyautogui.click(684, 345)
pyautogui.write('NYO Drinking Water Rental (DT 20%) ')

# Enter DT Unit Price
pyautogui.click(992, 343)
pyautogui.hotkey('ctrl', 'a')
pyautogui.write(DT_amount)

# Enter DT quantity
pyautogui.click(1088, 343)
pyautogui.hotkey('ctrl', 'a')
pyautogui.write('1')

# Add DT entry
pyautogui.click(1230, 295)
time.sleep(3)

# Enter Title
pyautogui.click(688, 637)
pyautogui.write('NYO Drinking Water Rental ' + month)


