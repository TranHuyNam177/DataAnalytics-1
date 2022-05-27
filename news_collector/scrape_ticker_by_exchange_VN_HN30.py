from request.stock import *


def run(
    hide_window=False
) -> pd.DataFrame:

    PATH = join(dirname(dirname(realpath(__file__))),'dependency','chromedriver')
    ignored_exceptions = (
        ValueError,
        IndexError,
        NoSuchElementException,
        StaleElementReferenceException,
        TimeoutException,
        ElementNotInteractableException
    )
    options = Options()
    if hide_window:
        options.headless = True

    URL = 'https://priceboard.vcbs.com.vn/Priceboard/'
    driver = webdriver.Chrome(service=Service(PATH),options=options)
    wait = WebDriverWait(driver,70,ignored_exceptions=ignored_exceptions)
    driver.get(URL)

    # HOSE
    print('Getting tickers in HOSE-VN30')
    action = ActionChains(driver)
    action.move_to_element(driver.find_element(By.XPATH,'//*[text()="HOSE"]'))
    time.sleep(2)
    action.click(driver.find_element(By.XPATH,'//*[text()="Bảng giá VN30"]'))
    action.perform()
    time.sleep(3)
    ticker_elems_hose = wait.until(
        EC.presence_of_all_elements_located((By.XPATH,'//tbody/*[@name!=""]'))
    )[1:]
    tickers_hose = list(map(lambda x:x.get_attribute('name'),ticker_elems_hose))
    table_hose = pd.DataFrame(index=pd.Index(tickers_hose,name='ticker'))
    table_hose['exchange'] = 'VN30'

    # HNX
    print('Getting tickers in HNX-HNX30')
    action = ActionChains(driver)
    action.move_to_element(driver.find_element(By.XPATH,'//*[text()="HNX"]'))
    time.sleep(2)
    action.click(driver.find_element(By.XPATH,'//*[text()="Bảng giá HNX30"]'))
    action.perform()
    time.sleep(3)
    ticker_elems_hnx = wait.until(
        EC.presence_of_all_elements_located((By.XPATH,'//tbody/*[@name!=""]'))
    )[1:]
    tickers_hnx = list(map(lambda x:x.get_attribute('name'),ticker_elems_hnx))
    table_hnx = pd.DataFrame(index=pd.Index(tickers_hnx,name='ticker'))
    table_hnx['exchange'] = 'HN30'

    driver.quit()
    result = pd.concat([table_hose,table_hnx])

    return result
