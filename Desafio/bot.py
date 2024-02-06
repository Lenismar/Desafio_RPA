
from botcity.web import WebBot, Browser, By
from cv2 import displayOverlay
import pandas as pd
from botcity.web.util import element_as_select




from botcity.maestro import *


BotMaestroSDK.RAISE_NOT_CONNECTED = False


def main():
   
    maestro = BotMaestroSDK.from_sys_args()
    
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot = WebBot()

    bot.headless = False

    
    bot.browser = Browser.CHROME

    bot.driver_path = "C:\Projetos\DesafioRPA\chromedriver-win64\chromedriver-win64\chromedriver.exe"
    #dados do desafio 
    
    tabela = pd.read_excel('C:\Projetos\DesafioRPA\challenge.xlsx')
    

    tabela = tabela.reindex(columns=['First Name', 'Last Name ', 'Company Name', 'Role in Company',
       'Address', 'Email', 'Phone Number'])
    
   
    bot.browse("https://rpachallenge.com/")

    bt_star = bot.find_element ("/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button",By.XPATH)
    bt_star.click()

   

    for i in range(len(tabela)):

        #Labels do foumulario
        last_name = bot.find_element ("//input[@ng-reflect-name='labelLastName']",By.XPATH )
        address = bot.find_element("//input[@ng-reflect-name='labelAddress']", By.XPATH)
        first_name = bot.find_element("//input[@ng-reflect-name='labelFirstName']", By.XPATH)
        role = bot.find_element("//input[@ng-reflect-name='labelRole']", By.XPATH)
        phone = bot.find_element("//input[@ng-reflect-name='labelPhone']", By.XPATH)
        company_name = bot.find_element("//input[@ng-reflect-name='labelCompanyName']", By.XPATH)
        email = bot.find_element("//input[@ng-reflect-name='labelEmail']", By.XPATH)
        botao = bot.find_element("//input[@type='submit']", By.XPATH)


        #Interções da tabela
        first_name.send_keys(tabela.loc[i,'First Name'])
        last_name.send_keys(tabela.loc[i,'Last Name '])
        company_name.send_keys(tabela.loc[i,'Company Name'])
        role.send_keys(tabela.loc[i,'Role in Company'])
        address.send_keys(tabela.loc[i,'Address'])
        email.send_keys(tabela.loc[i,'Email'])
        phone.send_keys(str(tabela.loc[i, 'Phone Number']))

        botao.click()
        bot.wait(1000)

    bot.wait(3000)

    bot.stop_browser()

def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
