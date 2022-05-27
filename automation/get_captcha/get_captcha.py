from automation.get_captcha import *


class NoNewsFound(Exception):
    pass


class PageFailToLoad(Exception):
    pass


ignored_exceptions = (
    ValueError,
    IndexError,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    ElementNotInteractableException,
    PageFailToLoad,
    ElementClickInterceptedException
)

PATH = 'D:/DataAnalytics-1/dependency/chromedriver.exe'
chrome_options = Options()
driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
wait = WebDriverWait(driver, 5, ignored_exceptions=ignored_exceptions)


def read_image(path):
    try:

        return pytesseract.image_to_string(path).replace('\n', '')
    except:
        return "[ERROR] Unable to process file: {0}".format(path)


def tool_ocr(img_path):
    error_word = ['A', 'C', 'G', 'J', 'T', 'Z', 'a', 'c', 'g', 'i', 'j', 'q', 'z']
    error_num = [1, 4, 5, 7, 8, 9]
    error_char = ['[', ']', '.', ',', ' ', '/', '|']
    error_str = 'ACEFGIJLOSTVWXYZacdegijloqsuvwxyz145789[].,/|>¢)'
    res = read_image(img_path)

    check_1 = any(word in res for word in error_word)
    check_2 = any(str(num) in res for num in error_num)
    check_3 = any(char in res for char in error_char)
    check_4 = len(res) != 6
    if check_1 | check_2 | check_3 | check_4:
        return True
    else:
        return False


img = r'C:\Users\namtran\Share Folder\Get Captcha\training dataset 2\captcha_100.png'
read_image(img)


# test get captcha, loc it based on condition and put it in the input box on web
def run():
    i = 1
    while True:
        driver.get('https://www.bidv.vn/iBank/MainEB.html')

        captcha = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="idImageCap"]')))
        img_path = fr'D:\DataAnalytics\news_analysis\captcha_data\test_dataset\captcha_{i}.png'
        time.sleep(1)
        captcha.screenshot(img_path)

        if tool_ocr(img_path):
            refr_btn = driver.find_element(
                By.XPATH, '//*[@id="authform"]/div[2]/div[3]/div/button'
            )
            refr_btn.click()
            time.sleep(1)
            i += 1
        else:
            print('Captcha: ', read_image(img_path))
            captcha_box = driver.find_element(By.XPATH, '//*[@id="captcha"]')
            captcha_box.clear()
            captcha_box.send_keys(read_image(img_path))
            break


# Code crawl captcha image data on web
def crawl_captcha():
    i = 1
    while True:
        driver.get('https://www.bidv.vn/iBank/MainEB.html')
        img_path = fr'C:\Users\namtran\Share Folder\Get Captcha\BIDV dataset\captcha_{i}.png'
        captcha = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="idImageCap"]')))
        captcha.screenshot(img_path)
        time.sleep(2)
        refr_btn = driver.find_element(
            By.XPATH, '//*[@id="authform"]/div[2]/div[3]/button/img'
        )
        refr_btn.click()
        i += 1
        if i > 100:
            break

    driver.close()


def process_captcha():
    for i in range(1, 101):
        imgPATH = fr'C:\Users\namtran\Share Folder\Get Captcha\BIDV dataset\captcha_{i}.png'
        image = cv2.imread(imgPATH)
        crop = image[5:35, 5:90, :]
        cv2.imwrite(imgPATH, crop)

        predictedCAPTCHA = pytesseract.image_to_string(imgPATH).replace('\n', '').replace(' ', '')
        print(predictedCAPTCHA)


