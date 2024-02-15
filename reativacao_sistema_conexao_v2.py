from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import pandas as pd
from datetime import datetime
import os
import warnings
import smtplib as smtp

warnings.filterwarnings('ignore')
cwd = os.getcwd()
data = datetime.today().strftime("%d%m")
month = datetime.today().strftime("%m")

base = pd.read_excel(cwd + f'\\sheets\\{month}\\{data}\\dashboard_equipe_{data}2024_INATIVOS.xlsx')
base_json = base.to_dict(orient='records')

class Reativacao():
    def __init__(self):
        self.driver_path = cwd + r'\driver\chromedriver.exe'
        self.options = webdriver.ChromeOptions()
        print('ok')
        self.options.add_argument('--log-level=3')
        self.options.add_experimental_option('debuggerAddress', 'localhost:8051')
        self.driver = webdriver.Chrome(options=self.options)
        self.driver.implicitly_wait(4)

    def fill_cpf(self, d):
        self.d = d
        self.cpf_input = self.driver.find_element(by=By.ID, value='cpf')
        self.cpf_input.clear()
        sleep(1)
        if len(str(self.d)) == 10:
            self.cpf_input.send_keys(str(0) + str(self.d))

        elif len(str(self.d)) == 9:
            self.cpf_input.send_keys(str(0) + str(0) + str(self.d))

        else:
            self.cpf_input.send_keys(str(self.d))
        
        sleep(1)
        
    def buscar_cpf(self):
        self.buscar_btn = self.driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div[3]/main/section/div/div[3]/div[1]/form/div[2]/div[1]/button')
        self.buscar_btn.click()
        sleep(1)
        
    def scroll_down(self):
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        sleep(1)
        
    def get_cfg(self):
        self.cfg_btn = self.driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div[3]/main/section/div/div[3]/div[2]/div[1]/div/div/table/tbody/tr/td[9]/div/button')
        self.cfg_btn.click()
        sleep(1)

    def ativar_usuario(self):
        self.ativar_btn = self.driver.find_element(by=By.XPATH, value='/html/body/div[3]/div[3]/ul/li[2]')
        self.ativar_btn.click()
        sleep(1)
        
    def confirmar(self):
        self.confirmar_btn = self.driver.find_element(by=By.XPATH, value='/html/body/div[3]/div[3]/div/div[3]/button[2]')
        self.confirmar_btn.click()
        sleep(1)
        
    def check_status(self):
        sleep(1)
        self.status = self.driver.find_element(by=By.XPATH, value='/html/body/div[1]/div[1]/div[3]/main/section/div/div[3]/div[2]/div[1]/div/div/table/tbody/tr/td[8]/div/p')
        return self.status.text

    def refresh_page(self):
        self.driver.refresh()
        sleep(1)

    def export(self):
        base.to_excel(cwd + f'\\sheets\\{month}\\{data}\\dashboard_equipe_{data}2023_INATIVOS.xlsx', index=False)

    def email_connection(self):
        self.connection = smtp.SMTP_SSL('smtp.gmail.com', 465)    
        self.email_addr = 'bona.ifsul@gmail.com'
        self.email_passwd = 'tlexrnmbaewuoxsh'
        self.connection.login(self.email_addr, self.email_passwd)
        self.connection.sendmail(
            from_addr=self.email_addr,
            to_addrs='bona.notifipy@gmail.com',
            msg=f'Script "{os.path.basename(__file__)}" finalizado com sucesso !!!')
        self.connection.close()

def run():
    bot = Reativacao()
    for i in range(len(base_json)):
        print("Reativando {} de {}".format(i + 1, len(base_json)))
        print("{} - {}".format(base_json[i]['PROMOTOR'], base_json[i]['CPF']))
        bot.fill_cpf(base_json[i]['CPF'])
        bot.buscar_cpf()
        bot.scroll_down()

        try:
            status = bot.check_status()
            if status == 'AT':
                base['STATUS'][i] = 'JÁ ATIVO'
                print('USUÁRIO JÁ ATIVO\n')
                bot.export()
                continue

            else:
                bot.get_cfg()
                bot.ativar_usuario()
                bot.confirmar()

                status = bot.check_status()
                if status == 'AT':
                    base['STATUS'][i] = 'REATIVADO'
                    print('USUÁRIO REATIVADO COM SUCESSO\n')
                    bot.export()
                    
                else:
                    base['STATUS'][i] = 'ERRO'
                    print('ERRO AO REATIVAR USUÁRIO\n')
                    bot.export()

        except:
            base['STATUS'][i] = 'CADASTRADO POR OUTRO AGENTE'
            print('USUÁRIO CADASTRADO POR OUTRO AGENTE\n')
            bot.export()
            continue

    bot.email_connection()

run()