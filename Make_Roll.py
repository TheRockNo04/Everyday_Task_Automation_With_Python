import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl import load_workbook


def main():
    login()

def login():
    browser = webdriver.ChromeOptions()
    browser.add_experimental_option('useAutomationExtension', False)
    browser = webdriver.Chrome(options=browser)

    browser.get('https://dopagent.indiapost.gov.in/corp/AuthenticationController?FORMSGROUP_ID__=AuthenticationFG&__START_TRAN_FLAG__=Y&__FG_BUTTONS__=LOAD&ACTION.LOAD=Y&AuthenticationFG.LOGIN_FLAG=3&BANK_ID=DOP&AGENT_FLAG=Y')
    time.sleep(2)
        
    AgentID = browser.find_element(By.XPATH, '//*[@id="AuthenticationFG.USER_PRINCIPAL"]')
    AgentID.send_keys('dop.mi3630020200029')

    Password = browser.find_element(By.XPATH, '//*[@id="AuthenticationFG.ACCESS_CODE"]')
    Password.send_keys('Narnari@36')

    captcha = browser.find_element(By.XPATH, '//*[@id="AuthenticationFG.VERIFICATION_CODE"]')
    captcha.send_keys(input())

    login = browser.find_element(By.XPATH, '//*[@id="VALIDATE_RM_PLUS_CREDENTIALS_CATCHA_DISABLED"]')
    login.click()

    Accounts = browser.find_element(By.XPATH, '//*[@id="Accounts"]')
    Accounts.click()

    Aeus = browser.find_element(By.XPATH, '//*[@id="Agent Enquire & Update Screen"]')
    Aeus.click()
    
    time.sleep(5)

#def slct_ac():
#    pass

#def save():
#    pass

#def selecting():
#    pass

if __name__ == "__main__":
    main()
