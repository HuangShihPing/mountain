from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, ElementNotInteractableException
import pandas as pd
import traceback

# Wait for the option to appear in the dropdown


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

    def clickElement(self, id: str):
        '''It can be used for checkbox, button or radiobox
        '''
        element = self.findElement(By.ID, id)
        self.moveTo(element)
        while True:
            try:
                element.click()
                break
            except ElementNotInteractableException:
                time.sleep(0.1)
        time.sleep(0.2)

    def fillTextbox(self, id: str, text: str):
        textbox = self.findElement(By.ID, id)
        textbox.send_keys(text)
        time.sleep(0.2)

    def clearTextbox(self, id: str):
        textbox = self.findElement(By.ID, id)
        textbox.clear()
        time.sleep(0.2)

    def selectByText(self, id: str, text: str):
        def optionPresentInDropdown(option_text):
            try:
                dropdown_element = self.findElement(By.ID, id)
                dropdown = Select(dropdown_element)
                return any(option.text == option_text for option in dropdown.options)
            except StaleElementReferenceException:
                # Handle the stale element reference exception by returning False
                return False

        # wait until options appear
        WebDriverWait(self.__driver, 5).until(lambda driver: optionPresentInDropdown(text))
        select = Select(self.findElement(By.ID, id))
        select.select_by_visible_text(text)
        time.sleep(0.2)

    def removeReadonly(self, id: str):
        '''It can be used to fill dates without interacting with the date selection dialog
        '''
        element = self.findElement(By.ID, id)
        self.__driver.execute_script("arguments[0].removeAttribute('readonly','readonly')", element)

    def findElement(self, strategy, name: str, timeout: float = 5):
        return WebDriverWait(self.__driver, timeout).until(EC.presence_of_element_located((strategy, name)))

    def findElementByText(self, text):
        # Locate the label by its text
        label = self.findElement(By.XPATH, f"//label[text()='{text}']")

        # Use the 'for' attribute of the label to find the checkbox
        element_id = label.get_attribute("for")
        element = self.__driver.find_element(By.ID, element_id)
        return element

    def findElementByValue(self, text):
        # Locate the label by its text
        label = self.findElement(By.XPATH, f"//label[value()='{text}']")

        # Use the 'for' attribute of the label to find the checkbox
        element_id = label.get_attribute("for")
        element = self.__driver.find_element(By.ID, element_id)
        return element


class MemberHandler:
    def addMembers(self, panel: Panel, excel: str, participates: list[str]):
        df = pd.read_excel(excel, dtype={"電話": str, "手機": str, "緊急聯絡人電話": str})

        i = 0
        for _, data in df.iterrows():
            if data["姓名"] not in participates:
                continue
            panel.clickElement('ContentPlaceHolder1_lbInsMember')  # 新增隊員
            time.sleep(1)
            panel.fillTextbox(f'ContentPlaceHolder1_lisMem_member_name_{i}', data["姓名"])
            panel.fillTextbox(f'ContentPlaceHolder1_lisMem_member_tel_{i}', data["電話"])
            panel.selectByText(f'ContentPlaceHolder1_lisMem_ddlmember_country_{i}', data["城市"])
            panel.selectByText(f'ContentPlaceHolder1_lisMem_ddlmember_city_{i}', data["區"])
            panel.fillTextbox(f'ContentPlaceHolder1_lisMem_member_addr_{i}', data["地址"])
            panel.fillTextbox(f'ContentPlaceHolder1_lisMem_member_mobile_{i}', data["手機"])
            panel.fillTextbox(f'ContentPlaceHolder1_lisMem_member_email_{i}', data["email"])
            panel.selectByText(f'ContentPlaceHolder1_lisMem_member_nation_{i}', data["國籍"])
            panel.fillTextbox(f'ContentPlaceHolder1_lisMem_member_sid_{i}', data["身份證"])
            panel.selectByText(f'ContentPlaceHolder1_lisMem_member_sex_{i}', data["性別"])
            panel.removeReadonly(f'ContentPlaceHolder1_lisMem_member_birthday_{i}')
            panel.fillTextbox(f'ContentPlaceHolder1_lisMem_member_birthday_{i}', data["生日"])
            # click any other place to close the pop-out date selector
            panel.clickElement(f'ContentPlaceHolder1_lisMem_member_mobile_{i}')
            panel.fillTextbox(f'ContentPlaceHolder1_lisMem_member_contactname_{i}', data["緊急聯絡人姓名"])  # 緊急聯絡人姓名
            panel.fillTextbox(f'ContentPlaceHolder1_lisMem_member_contacttel_{i}', data["緊急聯絡人電話"])  # 緊急連絡人電話
            i += 1


def main():

    url = "https://npm.nps.gov.tw/apply_1_2.aspx?unit=c951cdcd-b75a-46b9-8002-8ef952ec95fd"
    panel = Panel(url)

    # Check safety notifications and go to the next page
    check_button_ids = ['chk[]0', 'chk[]1', 'chk[]15', 'chk[]16', 'chk[]']
    for id in check_button_ids:
        panel.clickElement(id)
    panel.clickElement('ContentPlaceHolder1_btnagree')

    # fill team name
    panel.fillTextbox('ContentPlaceHolder1_teams_name', "中秋百岳團")
    panel.selectByText('ContentPlaceHolder1_climblinemain', "玉山線")
    panel.selectByText('ContentPlaceHolder1_climbline', "2~5天(塔塔加 - 玉山線 - 塔塔加)")
    panel.selectByText('ContentPlaceHolder1_sumday', "共2天")
    panel.selectByText('ContentPlaceHolder1_applystart', "2024-09-21")

    # 2天玉山行程
    # schedule_indices = [[0, 0, 0], [2, 0, 7, 6]]
    destinations = [["排雲登山服務中心", "塔塔加登山口", "排雲山莊"], ["玉山主峰", "排雲山莊", "塔塔加登山口", "排雲登山服務中心"]]
    for dests in destinations:
        for dest in dests:
            checkbox = panel.findElementByText(dest)
            checkbox.click()
        time.sleep(0.5)
        panel.clickElement('ContentPlaceHolder1_btnover')  # 完成今日路線按鈕
        time.sleep(0.5)
    panel.selectByText('ContentPlaceHolder1_seminar', "網路線上學習")
    panel.clickElement('ContentPlaceHolder1_btnToStep22')  # Next step

    # Fill applicant information
    panel.clickElement('headingOne')
    panel.clickElement('ContentPlaceHolder1_applycheck')
    panel.fillTextbox('ContentPlaceHolder1_apply_name', "黃釋平")
    panel.fillTextbox('ContentPlaceHolder1_apply_tel', "0920511676")
    panel.selectByText('ContentPlaceHolder1_ddlapply_country', "台北市")
    panel.selectByText('ContentPlaceHolder1_ddlapply_city', "內湖區")
    panel.fillTextbox('ContentPlaceHolder1_apply_addr', "行善路268號9樓之6")
    panel.fillTextbox('ContentPlaceHolder1_apply_mobile', "0920511676")
    panel.fillTextbox('ContentPlaceHolder1_apply_email', "mpskingdom@gmail.com")
    panel.selectByText('ContentPlaceHolder1_apply_nation', "中華民國")
    panel.fillTextbox('ContentPlaceHolder1_apply_sid', "L124546914")
    panel.selectByText('ContentPlaceHolder1_apply_sex', "男")
    panel.removeReadonly('ContentPlaceHolder1_apply_birthday')
    panel.fillTextbox('ContentPlaceHolder1_apply_birthday', "1995-03-22")
    panel.clickElement('ContentPlaceHolder1_apply_mobile')  # click any other place to close the pop-out date selector
    panel.fillTextbox('ContentPlaceHolder1_apply_contactname', "蔡端慧")
    panel.fillTextbox('ContentPlaceHolder1_apply_contacttel', "0920303078")

    # Fill leader information
    panel.clickElement('headingTwo')
    panel.clickElement('ContentPlaceHolder1_copyapply')

    # Fill member information
    # Clicking Heading three doesn't work because its span is not empty (請填個人手機號碼...)
    # We have to specify the h4 element to click
    time.sleep(1)
    element = panel.findElement(By.ID, 'headingThree')
    title = element.find_element(By.CLASS_NAME, 'panel-title')
    title.click()

    mem_handler = MemberHandler()
    mem_handler.addMembers(panel, "D:\\登山\\登山個資.xlsx", ["黃睿文", "黃宣錡", "劉育君", "張仕岱"])

    # Fill stay person information and go to the next page
    panel.clickElement('headingFour')
    panel.fillTextbox('ContentPlaceHolder1_stay_name', "范媛婷")
    panel.fillTextbox('ContentPlaceHolder1_stay_mobile', "0927345819")
    '''The following data are not necessary to fill'''
    panel.fillTextbox('ContentPlaceHolder1_stay_email', "yuanyare@gmail.com")
    # selectByText('ContentPlaceHolder1_stay_nation', "中華民國", delay=0.1)
    # fillTextbox('ContentPlaceHolder1_stay_sid', "H223897565")
    # removeReadonly('ContentPlaceHolder1_stay_birthday')
    # fillTextbox('ContentPlaceHolder1_stay_birthday', "1991-07-17")
    # clickElement('ContentPlaceHolder1_stay_tel') # click any other place to close the pop-out date selector
    time.sleep(10000)


if __name__ == '__main__':
    try:
        main()
    except:
        traceback.print_exc()
        time.sleep(10000)
