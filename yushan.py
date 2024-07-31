from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException


class Panel:
    def __init__(self, url):
        options = Options()
        options.add_experimental_option("useAutomationExtension", False)
        options.add_experimental_option("excludeSwitches", ['enable-automation'])
        self.__driver = webdriver.Chrome(options=options)
        self.__driver.get(url)
        delay = 5  # seconds
        myElem = WebDriverWait(self.__driver, delay).until(EC.presence_of_element_located((By.ID, 'ContentPlaceHolder1_ucNewsBack1_hlLocate1')))

    def moveTo(self, element):
        js_code = "arguments[0].scrollIntoView();"
        self.__driver.execute_script(js_code, element)

    def clickElement(self, id: str, delay: float = 0):
        '''It can be used for checkbox, button or radiobox
        '''
        element = self.__driver.find_element(By.ID, id)
        self.moveTo(element)
        element.click()
        time.sleep(delay)

    def fillTextbox(self, id: str, text: str):
        textbox = self.__driver.find_element(By.ID, id)
        textbox.send_keys(text)

    def clearTextbox(self, id: str):
        textbox = self.__driver.find_element(By.ID, id)
        textbox.clear()

    def selectByText(self, id: str, text: str, delay: float = 0):
        select = Select(self.__driver.find_element(By.ID, id))
        select.select_by_visible_text(text)
        time.sleep(delay)

    def removeReadonly(self, id: str):
        '''It can be used to fill dates without interacting with the date selection dialog
        '''
        element = self.__driver.find_element(By.ID, id)
        self.__driver.execute_script("arguments[0].removeAttribute('readonly','readonly')", element)

    def findElement(self, strategy, name: str):
        return self.__driver.find_element(strategy, name)


def main():
    url = "https://npm.nps.gov.tw/apply_1_2.aspx?unit=c951cdcd-b75a-46b9-8002-8ef952ec95fd"
    panel = Panel(url)

    # Check safety notifications and go to the next page
    check_button_ids = ['chk[]0', 'chk[]1', 'chk[]15', 'chk[]16', 'chk[]']
    for id in check_button_ids:
        panel.clickElement(id)
    panel.clickElement('ContentPlaceHolder1_btnagree')

    # fill team name
    panel.fillTextbox('ContentPlaceHolder1_teams_name', "不吃香菜Ping")
    panel.selectByText('ContentPlaceHolder1_climblinemain', "玉山線", delay=1)
    panel.selectByText('ContentPlaceHolder1_climbline', "2~5天(塔塔加 - 玉山線 - 塔塔加)", delay=0.5)
    panel.selectByText('ContentPlaceHolder1_sumday', "共2天", delay=1)
    panel.selectByText('ContentPlaceHolder1_applystart', "2024-09-15", delay=0.5)

    # 2天玉山行程
    schedule_indices = [[0, 0, 0], [2, 0, 7, 6]]
    for day_indices in schedule_indices:
        for index in day_indices:
            panel.clickElement(f"ContentPlaceHolder1_rblNode_{index}", delay=0.2)
        panel.clickElement('ContentPlaceHolder1_btnover', delay=0.2)
    panel.selectByText('ContentPlaceHolder1_seminar', "網路線上學習")
    panel.clickElement('ContentPlaceHolder1_btnToStep22', delay=0.1)

    # Fill applicant information
    panel.clickElement('headingOne', delay=0.3)
    panel.clickElement('ContentPlaceHolder1_applycheck', delay=0.2)
    panel.fillTextbox('ContentPlaceHolder1_apply_name', "黃釋平")
    panel.fillTextbox('ContentPlaceHolder1_apply_tel', "0920511676")
    panel.selectByText('ContentPlaceHolder1_ddlapply_country', "台北市", delay=0.2)
    panel.selectByText('ContentPlaceHolder1_ddlapply_city', "內湖區")
    panel.fillTextbox('ContentPlaceHolder1_apply_addr', "行善路268號9樓之6")
    panel.fillTextbox('ContentPlaceHolder1_apply_mobile', "0920511676")
    panel.fillTextbox('ContentPlaceHolder1_apply_email', "mpskingdom@gmail.com")
    panel.selectByText('ContentPlaceHolder1_apply_nation', "中華民國", delay=0.1)
    panel.fillTextbox('ContentPlaceHolder1_apply_sid', "L124546914")
    panel.selectByText('ContentPlaceHolder1_apply_sex', "男")
    panel.removeReadonly('ContentPlaceHolder1_apply_birthday')
    panel.fillTextbox('ContentPlaceHolder1_apply_birthday', "1995-03-22")
    panel.clickElement('ContentPlaceHolder1_apply_mobile')  # click any other place to close the pop-out date selector
    panel.fillTextbox('ContentPlaceHolder1_apply_contactname', "蔡端慧")
    panel.fillTextbox('ContentPlaceHolder1_apply_contacttel', "0920303078")

    # Fill leader information
    panel.clickElement('headingTwo', delay=0.3)
    panel.clickElement('ContentPlaceHolder1_copyapply', delay=0.2)

    # Fill member information
    # Clicking Heading three doesn't work because its span is not empty (請填個人手機號碼...)
    # We have to specify the h4 element to click
    element = panel.findElement(By.ID, 'headingThree')
    title = element.find_element(By.CLASS_NAME, 'panel-title')
    title.click()
    time.sleep(0.3)
    panel.clickElement('ContentPlaceHolder1_lbInsMember', delay=0.2)

    # TODO: replace the tail zeros to the index of members
    panel.fillTextbox('ContentPlaceHolder1_lisMem_member_name_0', "黃睿文")
    panel.fillTextbox('ContentPlaceHolder1_lisMem_member_tel_0', "0912600934")
    panel.selectByText('ContentPlaceHolder1_lisMem_ddlmember_country_0', "台北市", delay=0.2)
    panel.selectByText('ContentPlaceHolder1_lisMem_ddlmember_city_0', "內湖區")
    panel.fillTextbox('ContentPlaceHolder1_lisMem_member_addr_0', "行善路268號9樓之6")
    panel.fillTextbox('ContentPlaceHolder1_lisMem_member_mobile_0', "0912600934")
    panel.fillTextbox('ContentPlaceHolder1_lisMem_member_email_0', "s498600177@gmail.com")
    panel.selectByText('ContentPlaceHolder1_lisMem_member_nation_0', "中華民國", delay=0.1)
    panel.fillTextbox('ContentPlaceHolder1_lisMem_member_sid_0', "H223897565")
    panel.selectByText('ContentPlaceHolder1_lisMem_member_sex_0', "女")
    panel.removeReadonly('ContentPlaceHolder1_lisMem_member_birthday_0')
    panel.fillTextbox('ContentPlaceHolder1_lisMem_member_birthday_0', "1991-07-17")
    # click any other place to close the pop-out date selector
    panel.clickElement('ContentPlaceHolder1_lisMem_member_mobile_0')
    panel.fillTextbox('ContentPlaceHolder1_lisMem_member_contactname_0', "黃琇雅")
    panel.fillTextbox('ContentPlaceHolder1_lisMem_member_contacttel_0', "0912600914")

    # Fill stay person information and go to the next page
    panel.clickElement('headingFour', delay=0.3)
    panel.fillTextbox('ContentPlaceHolder1_stay_name', "黃睿睿")
    panel.fillTextbox('ContentPlaceHolder1_stay_mobile', "0912600934")
    '''The following data are not necessary to fill'''
    # fillTextbox('ContentPlaceHolder1_stay_email', "s498600177@gmail.com")
    # selectByText('ContentPlaceHolder1_stay_nation', "中華民國", delay=0.1)
    # fillTextbox('ContentPlaceHolder1_stay_sid', "H223897565")
    # removeReadonly('ContentPlaceHolder1_stay_birthday')
    # fillTextbox('ContentPlaceHolder1_stay_birthday', "1991-07-17")
    # clickElement('ContentPlaceHolder1_stay_tel') # click any other place to close the pop-out date selector
    time.sleep(1000)


if __name__ == '__main__':
    main()
