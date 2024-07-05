import time
import os
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait  # available since 2.4.0
from selenium.webdriver.support import expected_conditions as EC  # available since 2.26.0
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options

from datetime import datetime
from random import randint

# Global settings for pandas to display all rows/cols
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)


class WO:
    def __init__(self, proj, amt, invnum, po, contact, url):
        self.proj = proj
        self.amt = amt
        self.invnum = invnum
        self.po = po
        self.contact = contact
        self.url = url


class PortalSettings:
    def __init__(self):
        self.main_url = 'www.clientportalexample.com'
        self.po_url = 'www.clientportalexample.com/po/'
        self.username = 'username'
        self.password = 'supersecret'


class INVBOT:
    def __init__(self):

        # CONSTANTS
        self.mdir = os.path.abspath(os.path.dirname(__name__))  # C:\Users\vctrs\PY\port_projs\invoice_uploader
        self.portal_data =  PortalSettings()

        # MAIN DIR ENVS
        self.data_dir = os.path.join(self.mdir,'data')
        self.inv_dir = os.path.join(self.mdir,'data/invoices')  # Local Path for pdf invoices to be uploaded
        self.wo_dir = os.path.join(self.mdir,'data/wo.csv')
        

        # WEBDRIVER ENVS
        self.geckodriver_fpath = os.path.join(self.mdir,'webdriver/geckodriver.exe')
        self.firefox_binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'

        # ARIBA ENVS
        self.url = self.portal_data.main_url
        self.po_url = self.portal_data.po_url
        self.username = self.portal_data.username
        self.senha = self.portal_data.password

        self.wo = pd.DataFrame()
        self.wo_df = pd.read_csv(self.wo_dir)

        self.s_date = '03/01/2024'
        self.e_date = '03/31/2024'

        self.result_colnames = [
                                'PROJNUM',
                                'INVNUM',
                                'PO',
                                'STATUS',
                                'TIMESTAMP',
                                ]

        # main menu selection dict
        # Its time to choose
        self.menu_choices = {
            'LAUNCH BOT': self.upload_invoices,
            'VIEW QUEUE DETAIL': self.show_wo,
            'RECORDS': self.see_records,
            'CHANGE START/END DATE' : self.change_start_end_date,
            'REFRESH WO' : self.refresh_wo,
            'CHECK INV EXISTS' : self.check_invoice_exists,
            'CLEAR WO' : self.clear_wo,
            'QUIT\n': exit,
        }

        # mapping of DOC elements for XPATH
        '''
        webdriver >
        
        launches firefox with the initual client portal url login page>>
        
        client portal login page>
            'username_textbox'  enters username
            'password_textbox'  enters password
            'login_button'      clicks on login to enter client portal
        enters client portal>>
        
        client portal>
            *tab directly opens based on po/url mapping*
            'createinv_button' Create button (drowndown arrow)
            'standardinv_button' Standard Invoice
        creates invoice form page>>
        
        INV FORM PAGE>
            'addtoheader_button'
            'attachment_button'
            
            
        upload_step_num: 1
        upload_step_num: 2
        upload_step_num: 3
        upload_step_num: 4
        
        
        '''
        
        self.upload_steps = {
            1 : 'Creating invoice form in PO',
            2 : 'Adding Attachment into Form and Adding Services',
            3 : 'Filling the form with main values & adding uploaded attachment',
            4 : 'Finalizing invoice',
        }
        
        self.xpath_mapping = {

            # Main Login MAPPING -------
            'username_textbox': '//*[@id="userid"]',
            'loginnext_button': '/html/body/div[5]/table/tbody/tr[2]/td/div/table/tbody/tr/td[1]/form/div[2]/table/tbody/tr[3]/td/a/div/span',                                
            'password_textbox': '//*[@id="Password"]',
            'signin_button': '/html/body/div[5]/table/tbody/tr[2]/td/div/table/tbody/tr/td[1]/form/div[2]/table/tbody/tr[3]/td/input',
            
            # 'login_button': '/html/body/div[5]/table/tbody/tr[2]/td/div/table/tbody/tr/td[1]/form/' \
            #                 'div[2]/table/tbody/tr[3]/td/input',

            # PO HUB MAPPING -------
            'createinv_button': '/html/body/div[5]/div[2]/table/tbody/tr/td/table/tbody/tr[2]/' \
                                'td/table/tbody/tr/td/div/div/div/div[2]/form[1]/div[1]/table/' \
                                'tbody/tr[1]/td/table/tbody/tr/td[3]/div/button/span/span',  # PO form > down arrow

            'standardinv_button': '//*[@id="_nydmcb"]',  # PO form > standard invoice
            
            
            # INVOICE FORM MAPPING -------
            
            'addtoheader_button': '/html/body/div[5]/div[2]/table/tbody/tr/td/table/tbody/tr[1]/' \
                                  'td/table/tbody/tr/td/div/div/div/table/tbody/tr/td/span/form/' \
                                  'table/tbody/tr[3]/td/div/div/table[1]/tbody/tr[3]/td/table/' \
                                  'tbody/tr/td/span/div/table/tbody/tr[1]/td/table/tbody/tr/td[3]/' \
                                  'div/table/tbody/tr/td[2]/table/tbody/tr/td/div[1]/button/span/span',
            
            # 'attachment_button': '//*[@id="_ih6noc"]', 
            'attachment_button': '//*[@id="_xskhfd"]',
            
            'po_checkbox' : '/html/body/div[5]/div[2]/table/tbody/tr/td/table/tbody/tr[1]/td/table/' \
                            'tbody/tr/td/div/div/div/table/tbody/tr/td/span/form/table/tbody/tr[3]/' \
                            'td/div/div/table[1]/tbody/tr[4]/td/div[1]/table/tbody/tr[2]/td/table/' \
                            'tbody/tr[2]/td[1]/div',
            
            'create_button' : '/html/body/div[5]/div[2]/table/tbody/tr/td/table/tbody/tr[1]/td/table/' \
                              'tbody/tr/td/div/div/div/table/tbody/tr/td/span/form/table/tbody/tr[3]/' \
                              'td/div/div/table[1]/tbody/tr[4]/td/div[1]/table/tbody/tr[3]/td/table/' \
                              'tbody/tr/td[1]/div[1]/button/span/span',
            
            'service_button' : '//*[@id="_fr6ktd"]',            
            'invamt_textbox' : '//*[@id="_xjmh3d"]',
            'invcreate_button' : '//*[@id="_3m4pwc"]',
            # 'invoice_textbox': '//*[@id="_vhqewb"]',  # inv form >  invoice text field (old)
            'invoice_textbox': '//*[@id="_nygzrc"]',  # inv form >  invoice text field
            # 'taxrate_textbox': '//*[@id="_ptrrj"]',  # inv form > tax rate text field (old)
            'taxrate_textbox': '//*[@id="_9sxg_c"]',  # inv form > tax rate text field
            # 'taxamt_textbox': '//*[@id="_2eedyc"]',  # inv form > tax amt text field (old)
            'taxamt_textbox': '//*[@id="_ifikv"]',  # inv form > tax amt text field
            # 'startdate_textbox': '//*[@id="DF_d$zu1d"]',  # inv form > start date field (old)
            'startdate_textbox': '//*[@id="DF_3y3cbc"]',  # inv form > start date field
            # 'enddate_textbox': '//*[@id="DF_lin7h"]',  # inv form > end date date field (old)
            'enddate_textbox': '//*[@id="DF_pvjxpb"]',  # inv form > end date date field
            # 'contact_textbox': '//*[@id="_tgpdp"]',  # inv form > contact start text field (old)
            'contact_textbox': '//*[@id="_sv04ad"]',  # inv form > contact start text field
                                    
            'attch_browse_button' : '/html/body/div[5]/div[2]/table/tbody/tr/td/table/tbody/tr[1]/td/' \
                                    'table/tbody/tr/td/div/div/div/table/tbody/tr/td/span/form/table/'  \
                                    'tbody/tr[3]/td/div/div/table[1]/tbody/tr[3]/td/table/tbody/tr/td/' \
                                    'span/div/table/tbody/tr[2]/td/table/tbody/tr[24]/td[1]/table/tbody/'  \
                                    'tr[2]/td/span/input',
                                    
            # 'addattch_button' : '//*[@id="__xtv3d"]', #old
            'addattch_button' : '//*[@id="_7mvpxd"]',
            
            # 'next_button' : '//*[@id="_w_slkb"]', # (old)
            'next_button' : '//*[@id="_w_slkb"]',
            'submit_button' : '//*[@id="_jsl7tb"]',
            'exit_button' : '//button[@_a="exit"][0]',
        }

        # MUTABLES -------------------------------
        self.driver = None


    def welcome_message(self):

        print('------------------------------------')
        print(f'Hello I am INVBOT')
        print(f'Today is: {datetime.now().strftime("%Y/%m/%d %H:%M:%S")}')
        print(f'Total Work Order in queue: {len(self.wo_df)} invoice(s):')
        print(f'Total BATCH INV AMT: {"${:,.2f}".format(self.wo_df["AMT"].sum())}')
        print(f'Current PER set to: {self.s_date} - {self.e_date}\n')

    def main_menu(self):

        while True:
            drawn_num = randint(0,1)
            choose_message = ''

            if drawn_num:
                choose_message = 'It\'s time to choose..'
            else:
                choose_message = 'TIme...too TCHooOOoOSee...'
                
            self.welcome_message()

            print('---------MAIN MENU ---------')
            for i, v in enumerate(self.menu_choices.keys()):
                print(f'{i} - {v}')

            user_input = input(f'{choose_message}\n')

            try:
                self.menu_choices[list(self.menu_choices.keys())[int(user_input)]]()

            except Exception as e:
                print(e)

    def show_wo(self):
        print(self.wo_df)
        return

    def see_records(self):
        
        df_results = pd.read_csv(os.path.join(self.data_dir,'upload_records.csv'),
                            usecols=[*range(1,5)], #removed 5
                            header=None,
                            names=self.result_colnames,
                            )
        print(df_results)
        print('\n----------STATS-----------')
        print(f'Total upload attempts: {len(df_results)}')
        
        return

    def change_start_end_date(self):
        
        new_start_date = input('Please enter start date MM/DD/YYY')
        new_end_date = input('Please enter end date MM/DD/YYY')
        self.s_date = new_start_date
        self.e_date = new_end_date
        
        return

    def refresh_wo(self):
        print('Refreshing workorder data...')
        self.wo_df = pd.read_csv(self.wo_dir)
        time.sleep(1)
        print('Done!')
        return

    def check_invoice_exists(self):

        inv_path_tocheck = self.inv_dir
    
        result_list = self.wo_df['INV'].tolist()
        
        if len(result_list) == 0:
            print('No Workorder loaded. No invoice path to check\nReturning to main menu...\n')
            
        else:
            for x in result_list:
                print(f'INVOICE {x}: ', os.path.isfile(fr'{inv_path_tocheck}\{x}.pdf'), os.path.getsize(fr'{inv_path_tocheck}\{x}.pdf'))
        
        return

    def clear_wo(self):
        print('cleaning workorder data...')
        self.wo_df = self.wo_df[0:0]
        time.sleep(0.6)
    
        return
    
    def upload_invoices(self, submit_invoice=True, local_inv=True):

        # initiates webdriver
        self.launch_webdriver()

        # logs into ariba portal
        self.login()

        result_list = pd.DataFrame(columns=self.result_colnames)

        wo_count = 0
        wo_total = len(self.wo_df)

        # loop for work order
        for i in self.wo_df.index:
            wo = WO(
                self.wo_df.iloc[i][0],  # proj
                self.wo_df.iloc[i][1],  # amt
                self.wo_df.iloc[i][2],  # invnum
                self.wo_df.iloc[i][3],  # po
                self.wo_df.iloc[i][4],  # client contact
                self.wo_df.iloc[i][5],  # url
            )
            wo_count += 1
            upload_step_num = 0

            print('Uploading {} to PO: {}\n'.format(wo.invnum, wo.po))
            print(f'WO: {wo_count} out of {wo_total} ------------------')
            print('General Info:\nPROJ:{}\nINVNUM:{}\nAMT:{}\nCONTACT:{}\n'.format(wo.proj, wo.invnum, wo.amt, wo.contact))
            try:
                upload_step_num += 1
                print(f'Upload step {upload_step_num}: {self.upload_steps[upload_step_num]}')
                self.driver.execute_script('window.open("{}","PO_window");'.format(f'{self.po_url}{wo.url}'))
                time.sleep(5)
                self.driver.implicitly_wait(15)

                # switch to the PO tab
                self.driver.switch_to.window(self.driver.window_handles[2])
                time.sleep(2)
                # waits for help icon to be loaded (breathing room)
                self.driver.implicitly_wait(15)
                elementpage = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "enable-now-webassistant")))
                time.sleep(1)

                # creates a single invoice from arrowdown in PO page
                self.driver.find_element(By.XPATH, self.xpath_mapping['createinv_button']).click()
                time.sleep(1)

                self.driver.find_element(By.XPATH, self.xpath_mapping['standardinv_button']).click()  # correct one
                # self.driver.find_element_by_xpath('//*[@id="_nyddfsfsfdsf"]').click()  #wrong one

                self.driver.implicitly_wait(10)
                time.sleep(4)


            except Exception as e:
                print(e)
                time.sleep(2)
                self.driver.close()
                time.sleep(3)
                self.driver.switch_to.window(self.driver.window_handles[1])
                time.sleep(3)
                self.driver.refresh()
                time.sleep(5)

                print(f'INVOICE UPLOAD FAILED: {wo.invnum}')
                result_list = result_list.append({'PROJNUM': wo.proj,
                                                  'INVNUM': wo.invnum,
                                                  'PO': wo.po,
                                                  'STATUS': 'ERROR: 1',
                                                  'TIMESTAMP': datetime.now().strftime('%Y%m%d%H%M%S')}
                                                 , ignore_index=True, )
                continue


            try:
                upload_step_num += 1
                print(f'Upload step {upload_step_num}: {self.upload_steps[upload_step_num]}')
                elementpage = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "enable-now-webassistant")))

                # adds attachment component into invoice form created from single invoice dropdown
                self.driver.find_element(By.XPATH, self.xpath_mapping['addtoheader_button']).click()
                time.sleep(1)
                
                self.driver.find_element(By.XPATH, self.xpath_mapping['attachment_button']).click()
                time.sleep(2)

                # goes bottom, click on PO checkbox
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                self.driver.find_element(By.XPATH, self.xpath_mapping['po_checkbox']).click()
                time.sleep(1)

                # then creates service
                self.driver.find_element(By.XPATH, self.xpath_mapping['create_button']).click()
                time.sleep(1)

                self.driver.find_element(By.XPATH, self.xpath_mapping['service_button']).click()
                time.sleep(1)

                # created service popup
                self.driver.find_element(By.XPATH, self.xpath_mapping['invamt_textbox']).clear()
                self.driver.find_element(By.XPATH, self.xpath_mapping['invamt_textbox']).send_keys(str(wo.amt))
                time.sleep(1)
                self.driver.find_element(By.XPATH, self.xpath_mapping['invcreate_button']).click()
                time.sleep(2)


            except Exception as e:
                print(e)
                time.sleep(2)
                self.driver.close()
                time.sleep(3)
                self.driver.switch_to.window(self.driver.window_handles[1])
                time.sleep(3)
                self.driver.refresh()
                time.sleep(5)

                print(f'INVOICE UPLOAD FAILED: {wo.invnum}')
                result_list = result_list.append({'PROJNUM': wo.proj,
                                                  'INVNUM': wo.invnum,
                                                  'PO': wo.po,
                                                  'STATUS': 'ERROR: 2',
                                                  'TIMESTAMP': datetime.now().strftime('%Y%m%d%H%M%S')}
                                                 , ignore_index=True, )
                continue

            try:
                upload_step_num += 1
                print(f'Upload step {upload_step_num}: {self.upload_steps[upload_step_num]}')
                # fill up required textbox fields
                # invoice number, tax rate, tax amount, labor start date, labor end date, client_po contact
                self.driver.find_element(By.XPATH, self.xpath_mapping['invoice_textbox']).send_keys(str(wo.invnum))
                time.sleep(0.3)
                self.driver.find_element(By.XPATH, self.xpath_mapping['taxrate_textbox']).send_keys(0)
                time.sleep(0.3)
                self.driver.find_element(By.XPATH, self.xpath_mapping['taxamt_textbox']).send_keys(0)
                time.sleep(0.3)
                self.driver.find_element(By.XPATH, self.xpath_mapping['startdate_textbox']).send_keys(self.s_date)
                time.sleep(0.3)
                self.driver.find_element(By.XPATH, self.xpath_mapping['enddate_textbox']).send_keys(self.e_date)
                time.sleep(0.3)
                self.driver.find_element(By.XPATH, self.xpath_mapping['contact_textbox']).send_keys(wo.contact)
                time.sleep(1)

                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                # sends attachment dir values and click on Add Attachment button
                
                if local_inv:
                    self.driver.find_element(By.XPATH, self.xpath_mapping['attch_browse_button']).send_keys(fr'{self.local_invoice_rep}\{wo.invnum}.pdf')
                    inv_file_size = os.path.getsize(fr'{self.local_invoice_rep}\{wo.invnum}.pdf')
                else:
                    self.driver.find_element(By.XPATH, self.xpath_mapping['attch_browse_button']).send_keys(fr'{self.invoice_rep}\{wo.invnum}.pdf')
                    inv_file_size = os.path.getsize(fr'{self.local_invoice_rep}\{wo.invnum}.pdf')
                
                print(f'INVOICE FILESIZE: {inv_file_size}')
                self.driver.implicitly_wait(10)
                time.sleep(3)
                
                # add attachment after file into payload
                self.driver.find_element(By.XPATH, self.xpath_mapping['addattch_button']).click()
                if inv_file_size > 100000:
                    time.sleep(15)
                else:
                    time.sleep(3)

            except Exception as e:
                print(e)
                time.sleep(2)
                self.driver.close()
                time.sleep(3)
                self.driver.switch_to.window(self.driver.window_handles[1])
                time.sleep(3)
                self.driver.refresh()
                time.sleep(5)

                print(f'INVOICE UPLOAD FAILED: {wo.invnum}')
                result_list = result_list.append({'PROJNUM': wo.proj,
                                                  'INVNUM': wo.invnum,
                                                  'PO': wo.po,
                                                  'STATUS': 'ERROR: 3',
                                                  'TIMESTAMP': datetime.now().strftime('%Y%m%d%H%M%S')}
                                                 , ignore_index=True, )
                continue


            try:
                upload_step_num += 1
                print(f'Upload step {upload_step_num}: {self.upload_steps[upload_step_num]}')
                # clicks on next to processed to confirmation window,
                # it proceeds once the question mark icon image is loaded
                self.driver.find_element(By.XPATH, self.xpath_mapping['next_button']).click()
                self.driver.implicitly_wait(10)
                
                # elementpage = WebDriverWait(self.driver, 10).until(
                #     EC.presence_of_element_located((By.ID, "enable-now-webassistant")))
                time.sleep(3)
                if submit_invoice:
                    # clicks on submit. and it takes you to the receipt page
                    time.sleep(2)
                    self.driver.find_element(By.XPATH, self.xpath_mapping['submit_button']).click()
                    elementpage = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.ID, "enable-now-webassistant")))
                    self.driver.implicitly_wait(10)
                    time.sleep(2)

                else:
                    # clicks on exit (for testing purposes)
                    self.driver.find_element(By.XPATH, self.xpath_mapping['exit_button']).click()
                    time.sleep(2)
                    elementpage = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.ID, "enable-now-webassistant")))


                # closes tab, then reload main login page... then start the loop once again
                # closes po tab
                time.sleep(5)
                self.driver.close()
                time.sleep(3)
                print('inv {} has been successfully uploaded!\n'.format(wo.invnum))
                self.driver.switch_to.window(self.driver.window_handles[1])
                time.sleep(3)
                self.driver.refresh()
                time.sleep(5)


                result_list = result_list.append({
                    'PROJNUM': wo.proj,
                    'INVNUM': wo.invnum,
                    'PO': wo.po,
                    'STATUS': 'SUCCESS',
                    'TIMESTAMP': datetime.now().strftime('%Y%m%d%H%M%S')
                },
                    ignore_index=True,
                )
                


            except Exception as e:
                print(e)
                time.sleep(2)
                self.driver.close()
                time.sleep(3)
                self.driver.switch_to.window(self.driver.window_handles[1])
                time.sleep(3)
                self.driver.refresh()
                time.sleep(5)

                print(f'INVOICE UPLOAD FAILED: {wo.invnum}')
                result_list = result_list.append({'PROJNUM': wo.proj,
                                                  'INVNUM': wo.invnum,
                                                  'PO': wo.po,
                                                  'STATUS': 'ERROR: 4',
                                                  'TIMESTAMP': datetime.now().strftime('%Y%m%d%H%M%S')}
                                                 , ignore_index=True, )
                continue

        print('Job Done!\n')
        time.sleep(1)
        print(result_list)
        result_list.to_csv(os.path.join(self.data_dir,'upload_records.csv'), mode='a', header=False)

        self.driver.close()

        while True:
            user_input = input('Clear WO? (y/n)')
            
            if user_input == 'y':
                self.clear_wo()
                self.main_menu()
                
            elif user_input == 'n':
                print('returning to main menu...')
                time.sleep(0.5)
                self.main_menu()
                
            else:
                print('please select y or n')

    def launch_webdriver(self):

        print('Launching webdriver...\n')
        time.sleep(1)

        # launches firefox
        o = Options()
        s = Service(self.geckodriver_fpath)
        
        o.binary_location = self.firefox_binary_location
        self.driver = webdriver.Firefox(service=s, options=o)
        time.sleep(4)

        # changes active tab to newly opened tab (index 1)
        self.driver.switch_to.window(self.driver.window_handles[1])

        return

    def login(self):
        self.driver.get(self.url)

        print('Logging into CLIENT PORTAL...')
        time.sleep(5)
        # get DOM element using xpath to get username textbox
        self.driver.find_element(By.XPATH,self.xpath_mapping['username_textbox']).send_keys(self.username)
        time.sleep(1.5)
        
        #initial login needs time to load the JScript as well since the username > next > password login screen update
        # self.driver.find_element(By.XPATH,self.xpath_mapping['loginnext_button']).send_keys(f'\n')
        self.driver.find_element(By.XPATH,self.xpath_mapping['loginnext_button']).click()
        time.sleep(2)
        
        # get DOM element using xpath to get password textbox
        self.driver.find_element(By.XPATH, self.xpath_mapping['password_textbox']).send_keys(self.senha)
        time.sleep(0.5)
        
        # get DOM element using xpath to click on login button
        self.driver.find_element(By.XPATH, self.xpath_mapping['signin_button']).click()
        time.sleep(3)
        # logged.switch_to.window(logged.window_handles[2])
        return


def main():
    ab = INVBOT()
    ab.main_menu()

if __name__ == '__main__':
    main()
