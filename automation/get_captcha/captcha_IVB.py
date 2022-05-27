from automation.get_captcha import *

input_path = r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\dataset'
input_path_2 = r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\dataset_2'
input_bw_path = r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\black white dataset'
fixed_path = r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\fixed dataset'
black_white_path = r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\black white dataset'

########################################################################################################


# Mr Hiep's code
def ivb_mr_hiep(i):
    img = Image.open(join(input_path_2, fr'captcha_{i}.png'))
    data = np.array(img)
    data = data[:, :, :3]
    data[np.sum(data, axis=2) > 400] = 0
    img = Image.fromarray(data)
    img = img.resize((720, 200))
    img.save(join(r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\fixed dataset\captcha_27.png'))
    predictedCAPTCHA = pytesseract.image_to_string(
        r'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\fixed dataset\captcha_27.png'
    )
    return predictedCAPTCHA


########################################################################################################

# Nam's code
# delete black line and read text using pytesseract
def ivb_captcha(i):
    imagePATH = cv2.imread(join(input_path_2, fr'captcha_{i}.png'))
    gray = cv2.cvtColor(imagePATH, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_OTSU + cv2.THRESH_BINARY_INV)[1]

    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (19, 1))
    detected_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel)

    cnts = cv2.findContours(detected_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = cnts[0] if len(cnts) == 2 else cnts[1]

    for c in cnts:
        cv2.drawContours(imagePATH, [c], -2, (255, 255, 255), -1)

    repair_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 3))
    result = 255 - cv2.morphologyEx(255 - imagePATH, cv2.MORPH_CLOSE, repair_kernel, iterations=1)

    cv2.imwrite(join(fixed_path, fr'captcha_{i}.png'), result)

    predictedCAPTCHA = pytesseract.image_to_string(
        join(fixed_path, fr'captcha_{i}.png')
    )
    return predictedCAPTCHA


########################################################################################################

# Converting an image to black and white
def ivb_bw(i):
    originalImage = cv2.imread(join(input_path_2, fr'captcha_{i}.png'))
    grayImage = cv2.cvtColor(originalImage, cv2.COLOR_BGR2GRAY)

    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 127, 255, cv2.THRESH_BINARY)

    cv2.imwrite(join(black_white_path, fr'captcha_{i}.png'), blackAndWhiteImage)


########################################################################################################

# Increase or decrease the alpha and beta value to get required output
def test(i: int, alpha: float, beta: int):
    img = cv2.imread(join(input_path, fr'captcha_{i}.png'))

    new = alpha * img + beta
    new = np.clip(new, 0, 255).astype(np.uint8)
    cv2.imwrite(join(fixed_path, fr'captcha_{i}.png'), new)

    predictedCAPTCHA = pytesseract.image_to_string(
        join(fixed_path, fr'captcha_{i}.png')
    )
    return predictedCAPTCHA


########################################################################################################

# Delete black line
def process_image(i):
    # Load image, grayscale, and Otsu's threshold
    image = cv2.imread(join(input_path, fr'captcha_{i}.png'))
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

    # Morph open to remove noise
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    opening = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel, iterations=1)

    # Find contours and remove small noise
    cnts = cv2.findContours(opening, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = cnts[0] if len(cnts) == 2 else cnts[1]
    for c in cnts:
        area = cv2.contourArea(c)
        if area < 50:
            cv2.drawContours(opening, [c], -1, 0, -1)

    # Invert and apply slight Gaussian blur
    result = 255 - opening
    result = cv2.GaussianBlur(result, (3, 3), 0)
    cv2.imwrite(join(fixed_path, fr'captcha_{i}.png'), result)


########################################################################################################

# crawl data from web
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

# PATH = join(dirname(dirname(realpath(__file__))), 'dependency', 'chromedriver')
PATH = 'D:/DataAnalytics-1/dependency/chromedriver.exe'
chrome_options = Options()
driver = webdriver.Chrome(options=chrome_options, service=Service(PATH))
wait = WebDriverWait(driver, 5, ignored_exceptions=ignored_exceptions)


def crawl_captcha():
    i = 1
    while True:
        driver.get('https://ebanking.indovinabank.com.vn/corp/Request?&dse_sessionId=zd6WfyskxO58Y2FoxA2AsOb'
                   '&dse_applicationId=-1&dse_pageId=2&dse_operationName=corpIndexProc&dse_errorPage=error_page.jsp'
                   '&dse_processorState=initial&dse_nextEventName=start')

        captcha = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="safecode"]')))
        img_path = fr'C:\Users\namtran\Share Folder\Get Captcha\ivb_dataset\dataset\captcha_{i}.png'
        captcha.screenshot(img_path)

        refr_btn = driver.find_element(
            By.XPATH, '//*[@id="loginForm"]/div/div[4]/div/span/a'
        )
        refr_btn.click()
        time.sleep(2)
        i += 1
        if i > 100:
            break

    driver.close()


########################################################################################################


# Python OpenCV mouse events
def drawCircle(event, x, y, flags, param):
    if event == cv2.EVENT_MOUSEMOVE:
        print('({}, {})'.format(x, y))

        imgCopy = img.copy()
        cv2.circle(imgCopy, (x, y), 10, (255, 0, 0), -1)

        cv2.imshow('image', imgCopy)


img = cv2.imread(join(input_path, fr'captcha_80.png'))
img = cv2.resize(img, (750, 208))
cv2.imshow('image', img)

cv2.setMouseCallback('image', drawCircle)

cv2.waitKey(0)
cv2.destroyAllWindows()

########################################################################################################

# Capturing mouse click events with Python and OpenCV
# Create point matrix get coordinates of mouse click on image
point_matrix = np.zeros((2, 2), np.int)

counter = 0


def mousePoints(event, x, y, flags, params):
    global counter
    # Left button mouse click event <a href="https://thinkinfi.com/basic-python-opencv-tutorial-function/"
    # data-internallinksmanager029f6b8e52c="14" title="OpenCV" target="_blank" rel="noopener">opencv</a>
    if event == cv2.EVENT_LBUTTONDOWN:
        point_matrix[counter] = x, y
        counter = counter + 1


def run():
    img = cv2.imread(join(input_path, fr'captcha_81.png'))
    img = cv2.resize(img, (750, 208))
    while True:
        for x in range(0, 2):
            cv2.circle(img, (point_matrix[x][0], point_matrix[x][1]), 3, (0, 255, 0), cv2.FILLED)

        if counter == 2:
            starting_x = point_matrix[0][0]
            starting_y = point_matrix[0][1]

            ending_x = point_matrix[1][0]
            ending_y = point_matrix[1][1]
            # Draw rectangle for area of interest
            cv2.rectangle(img, (starting_x, starting_y), (ending_x, ending_y), (0, 255, 0), 3)

            # Cropping image
            img_cropped = img[starting_y:ending_y, starting_x:ending_x]
            cv2.imshow("ROI", img_cropped)

        # Showing original image
        cv2.imshow("Original Image ", img)
        # Mouse click event on original image
        cv2.setMouseCallback("Original Image ", mousePoints)
        # Printing updated point matrix
        print(point_matrix)
        # Refreshing window all time
        cv2.waitKey(1)


########################################################################################################

# text detection with OpenCV
img = cv2.imread(join(fixed_path, fr'captcha_10_old.png'))
img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
# print(pytesseract.image_to_string(img))

### Detecting characters ###
hImg,wImg,_ = img.shape
boxes = pytesseract.image_to_boxes(img)
for b in boxes.splitlines():
    b = b.split(' ')
    print(b)
    x,y,w,h = int(b[1]),int(b[2]),int(b[3]),int(b[4])
    cv2.rectangle(img,(x,hImg-y),(w,hImg-h),(0,0,255),1)

# ### Detecting Words ###
# hImg,wImg,_ = img.shape
# boxes = pytesseract.image_to_data(img)
# print(boxes)
# for x,b in enumerate(boxes.splitlines()):
#     if x!=0:
#         b = b.split()
#         print(b)
#         if len(b)==12:
#             x,y,w,h = int(b[6]),int(b[7]),int(b[8]),int(b[9])
#             cv2.rectangle(img,(x,y),(w+x,h+y),(0,0,255),1)

img = cv2.resize(img,(750,208))
cv2.imshow('Result',img)
cv2.waitKey(0)