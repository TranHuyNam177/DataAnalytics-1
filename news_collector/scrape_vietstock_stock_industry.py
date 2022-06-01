from request.stock import *


class PageFailToLoad(Exception):
    pass


class vietstock:
    def __init__(self):
        self.PATH = join(dirname(dirname(realpath(__file__))),'dependency','chromedriver')
        self.ignored_exceptions = (
            ValueError,
            IndexError,
            NoSuchElementException,
            StaleElementReferenceException,
            TimeoutException,
            ElementNotInteractableException,
            PageFailToLoad,
        )
        self.options = Options()

    def run(self):
        url = 'https://finance.vietstock.vn/doanh-nghiep-a-z?languageid=2&page=1'
        driver = webdriver.Chrome(service=Service(self.PATH), options=self.options)
        driver.get(url)

        # Đăng nhập
        login_element = driver.find_element(By.XPATH, '/html/body/div[2]/div[6]/div/div[2]/div[2]/a[3]')
        login_element.click()
        email_element = driver.find_element(By.XPATH, '//*[@id="txtEmailLogin"]')
        email_element.clear()
        email_element.send_keys('namtran@phs.vn')
        password_element = driver.find_element(By.XPATH, '//*[@id="txtPassword"]')
        password_element.clear()
        password_element.send_keys('123456789')
        login_element = driver.find_element(By.XPATH, '//*[@id="btnLoginAccount"]')
        login_element.click()
        time.sleep(1)
        page = driver.find_element(By.XPATH, '//*[@id="az-container"]/div[3]/div[2]/div/span[1]/span[2]').text
        i = 1
        stock_lst = []
        industry_lst = []
        while True:
            stock_elems = driver.find_elements(By.XPATH, '//*[@id="az-container"]/div[2]/table[1]/tbody/tr/td[2]')
            industry_elems = driver.find_elements(By.XPATH, '//*[@id="az-container"]/div[2]/table[1]/tbody/tr/td[4]')

            stock_lst.extend([s.text for s in stock_elems])
            industry_lst.extend([idt.text for idt in industry_elems])
            i += 1
            try:
                btn_next = driver.find_element(By.XPATH, '//*[@id="btn-page-next"]')
                btn_next.click()
                time.sleep(1)
            except self.ignored_exceptions:
                continue
            if i > int(page):
                break
        df = pd.DataFrame({
            'stock': stock_lst,
            'industry': industry_lst
        })
        return df


vietstock = vietstock().run()




