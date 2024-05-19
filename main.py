from excel_utils import excel_utils
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from selenium.webdriver.firefox.service import Service as FirefoxService
from webdriver_manager.firefox import GeckoDriverManager
from datetime import datetime
import time
import sys
file = "data.xlsx"
class TestCreateQuiz():
    def __init__(self):
      
        self.driver = webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()))
        self.driver.implicitly_wait(10)
        self.driver.get("https://school.moodledemo.net/my/")
        self.driver.maximize_window()
        
  
    def teardown_method(self):
        
        self.driver.quit()
    def put_test(self, name, o_day, o_month, o_year, o_hour, o_minute,
                 c_day, c_month, c_year, c_hour, c_minute):
            print(o_day, o_month, o_year, o_hour, o_minute, c_day, c_month,c_year, c_hour, c_minute)
            time.sleep(3)
            self.driver.find_element(By.XPATH, f"//select[@name='timeopen[day]']/option[@value='{o_day}']").click() 
            self.driver.find_element(By.XPATH, f"//select[@name='timeopen[month]']/option[@value='{o_month}']").click() 
            self.driver.find_element(By.XPATH, f"//select[@name='timeopen[year]']/option[@value='{o_year}']").click()
            self.driver.find_element(By.XPATH, f"//select[@name='timeopen[hour]']/option[@value='{o_hour}']").click()
            self.driver.find_element(By.XPATH, f"//select[@name='timeopen[minute]']/option[@value='{o_minute}']").click()

            
            self.driver.find_element(By.XPATH, f"//select[@name='timeclose[day]']/option[@value='{c_day}']").click()
            self.driver.find_element(By.XPATH, f"//select[@name='timeclose[month]']/option[@value='{c_month}']").click() 
            self.driver.find_element(By.XPATH, f"//select[@name='timeclose[year]']/option[@value='{c_year}']").click()
            self.driver.find_element(By.XPATH, f"//select[@name='timeclose[hour]']/option[@value='{c_hour}']").click()
            self.driver.find_element(By.XPATH, f"//select[@name='timeclose[minute]']/option[@value='{c_minute}']").click()
            time.sleep(3)
    def checkDateValid(self, year, month, day):
        try:
            datetime(int(year), int(month), int(day))
        except ValueError:
            return False
        return True
    def test_cQBasicFlow(self, sheetname):
        
        excel = excel_utils(file,sheetname)
        rows = excel.get_row_count()

        self.driver.set_window_size(1552, 832)

        self.driver.find_element(By.LINK_TEXT, "Log in").click()

        self.driver.find_element(By.ID, "username").click()

        self.driver.find_element(By.ID, "username").send_keys("teacher")

        self.driver.find_element(By.ID, "password").click()

        self.driver.find_element(By.ID, "password").send_keys("moodle")

        self.driver.find_element(By.ID, "loginbtn").click()

        self.driver.find_element(By.LINK_TEXT, "My courses").click()

        WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[@href='https://school.moodledemo.net/course/view.php?id=51']"))).click()
        try:

            element = WebDriverWait(self.driver, 10).until(
        EC.presence_of_element_located((By.NAME, "setmode"))
    )
            element.click()
        except:

            element = WebDriverWait(self.driver, 10).until(
        EC.presence_of_element_located((By.ID, "yui_3_18_1_1_1714615616721_224"))
    )
            element.click()

        element = self.driver.find_element(By.ID, "action-menu-toggle-5")
        actions = ActionChains(self.driver)
        actions.move_to_element(element).perform()

        element = self.driver.find_element(By.CSS_SELECTOR, "body")
        actions = ActionChains(self.driver)
        actions.move_to_element(element).perform()

        self.driver.execute_script("window.scrollTo(0,49.599998474121094)")

        self.driver.find_element(By.CSS_SELECTOR, "#coursecontentcollapse0 .activity-add-text").click()

        self.driver.find_element(By.CSS_SELECTOR, ".option:nth-child(16) .optionicon > .icon").click()
        

        self.driver.find_element(By.ID, "collapseElement-1").click()      

        self.driver.find_element(By.ID, "id_timeopen_enabled").click()
        self.driver.find_element(By.ID, "id_timeclose_enabled").click()
        for r in range(2, rows + 1):
        
            name = excel.read_data(r, 1)
            o_day = excel.read_data(r, 2)
            o_month = excel.read_data(r, 3)
            o_year = excel.read_data(r, 4)
            o_hour = excel.read_data(r, 5)
            o_minute = excel.read_data(r, 6)
            c_day = excel.read_data(r, 7)
            c_month = excel.read_data(r, 8)
            c_year = excel.read_data(r, 9)
            c_hour = excel.read_data(r, 10)
            c_minute = excel.read_data(r, 11)

            
            self.put_test( name, o_day, o_month, o_year, o_hour, o_minute, 
                          c_day, c_month, c_year, c_hour, c_minute)
            actual_o_day = Select(self.driver.find_element(By.ID, "id_timeopen_day")).first_selected_option.get_attribute('value')
            actual_o_month = Select(self.driver.find_element(By.ID, "id_timeopen_month")).first_selected_option.get_attribute('value')
            actual_o_year = Select(self.driver.find_element(By.ID, "id_timeopen_year")).first_selected_option.get_attribute('value')
            actual_o_hour = Select(self.driver.find_element(By.ID, "id_timeopen_hour")).first_selected_option.get_attribute('value')
            actual_o_minute = Select(self.driver.find_element(By.ID, "id_timeopen_minute")).first_selected_option.get_attribute('value')
            actual_c_day = Select(self.driver.find_element(By.ID, "id_timeclose_day")).first_selected_option.get_attribute('value')
            actual_c_month = Select(self.driver.find_element(By.ID, "id_timeclose_month")).first_selected_option.get_attribute('value')
            actual_c_year = Select(self.driver.find_element(By.ID, "id_timeclose_year")).first_selected_option.get_attribute('value')
            actual_c_hour = Select(self.driver.find_element(By.ID, "id_timeclose_hour")).first_selected_option.get_attribute('value')
            actual_c_minute = Select(self.driver.find_element(By.ID, "id_timeclose_minute")).first_selected_option.get_attribute('value')
            print(actual_o_day, actual_o_month, actual_o_year, actual_o_hour, actual_o_minute, 
                  actual_c_day, actual_c_month,actual_c_year, actual_c_hour, actual_c_minute)
            time.sleep(3)

            if self.checkDateValid(o_year, o_month, o_day) and self.checkDateValid(c_year, c_month, c_day):
                if (datetime(int(actual_o_year), int(actual_o_month), int(actual_o_day), int(actual_o_hour), int(actual_o_minute)) == datetime(int(o_year), int(o_month), int(o_day), int(o_hour), int(o_minute))) and (datetime(int(actual_c_year), int(actual_c_month), int(actual_c_day), int(actual_c_hour), int(actual_c_minute)) == datetime(int(c_year), int(c_month), int(c_day), int(c_hour), int(c_minute))):
                    if datetime(int(actual_c_year), int(actual_c_month), int(actual_c_day), int(actual_c_hour), int(actual_c_minute)) > datetime(int(actual_o_year), int(actual_o_month), int(actual_o_day), int(actual_o_hour), int(actual_o_minute)):
                        excel.write_data(r, 13, "Passed")
                        excel.fill_Color(r, 13, "99CCFF")
                    else: 
                        excel.write_data(r, 13, "Failed")
                        excel.fill_Color(r, 13, "FB584B")
                else:
                    print("Error: inputting value to the web")
            else: 
                excel.write_data(r, 13, "ERROR")
                excel.fill_Color(r, 13, "00B050")    
              
        self.driver.find_element(By.ID, "id_submitbutton").click()

        self.teardown_method()
class TestEditQuestionQuiz():
    def __init__(self):
      
        self.driver = webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()))
        self.driver.implicitly_wait(10)
        self.driver.get("https://school.moodledemo.net/my/")
        self.driver.maximize_window()
        
  
    def teardown_method(self):
        
        self.driver.quit()
    def test_EQQBasicFlow(self, sheetname):
        excel = excel_utils(file,sheetname)
        rows = excel.get_row_count()

        self.driver.set_window_size(1552, 832)

        self.driver.find_element(By.LINK_TEXT, "Log in").click()

        self.driver.find_element(By.ID, "username").click()

        self.driver.find_element(By.ID, "username").send_keys("teacher")

        self.driver.find_element(By.ID, "password").click()

        self.driver.find_element(By.ID, "password").send_keys("moodle")

        self.driver.find_element(By.ID, "loginbtn").click()

        self.driver.find_element(By.LINK_TEXT, "My courses").click()

        WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[@href='https://school.moodledemo.net/course/view.php?id=51']"))).click()
        try:

            element = WebDriverWait(self.driver, 10).until(
        EC.presence_of_element_located((By.NAME, "setmode"))
    )
            element.click()
        except:

            element = WebDriverWait(self.driver, 10).until(
        EC.presence_of_element_located((By.ID, "yui_3_18_1_1_1714615616721_224"))
    )
            element.click()
        self.driver.find_element(By.LINK_TEXT, "Quiz: Test your knowledge on Alpinism").click()
        self.driver.find_element(By.LINK_TEXT, "Questions").click()
        
        for r in range(2, rows + 1):
            data = excel.read_data(r, 1)
            self.driver.find_element(By.NAME, "maxgrade").clear()
            self.driver.find_element(By.NAME, "maxgrade").send_keys(data)
            self.driver.find_element(By.XPATH, '//input[@name="savechanges"]').click()
            actual_data = self.driver.find_element(By.NAME, "maxgrade").get_attribute('value')
            print(int(data), int(actual_data))
            if (int(data) <= 0) and (int(actual_data) <= 0):
                excel.write_data(r, 3, "Failed")
                excel.fill_Color(r, 3, "FB584B")
            else: 
                excel.write_data(r, 3, "Passed")
                excel.fill_Color(r, 3, "99CCFF")
            time.sleep(3)
        self.teardown_method()
        
def myfunc(argv):
    try:
        sheetname = sys.argv[1]
        test = sys.argv[2]
    except:
        print("Please follow the format: \n \"python main.py Sheet1 cq\" (for testing CreateQuiz) \n and \n \"python main.py Sheet2 eqq\" (for testing EditQuestionQuiz)")
        sys.exit(2)
    if test == "cq":
        run = TestCreateQuiz()
        run.test_cQBasicFlow(sheetname)
        run.teardown_method()
    elif test == "eqq":
        run = TestEditQuestionQuiz()
        run.test_EQQBasicFlow(sheetname)
        run.teardown_method()
if __name__ == "__main__":
    myfunc(sys.argv)