import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from knw_Chromedriver_manager import Chromedriver_manager
import logging

# --- [전역 변수] ---
LOGIN_URL = "http://label-craft.is.kakaocorp.com/login"
HOME_URL = "http://label-craft.is.kakaocorp.com/tasks/45/"

# --- [로거 설정] ---
logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(asctime)s - %(message)s')
logger = logging.getLogger("prefix_logger")


# 로그인
def label_login(username, password, headless=True):
    driver = None
    try:
        options = Options()
        options.add_argument("--start-maximized")

        if headless:
            logger.info("헤드리스 모드로 실행합니다.")
            options.add_argument("--headless=new")
            options.add_argument("--window-size=1920,1080")
            options.add_argument("--disable-gpu")

        driver_path = Chromedriver_manager.install()
        driver = webdriver.Chrome(service=Service(driver_path), options=options)

        driver.get(LOGIN_URL)

        # 아이디 입력
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "exampleInputEmail"))
        ).send_keys(username)

        # 비밀번호 입력
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "exampleInputPassword"))
        ).send_keys(password + Keys.RETURN)

        # 로그인 성공 대기
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#accordionSidebar > a > img"))
        )
        logger.info("✅ 라벨크래프트 로그인에 성공했습니다.")
        logger.debug(f"크롬드라이버 경로: {driver_path}")

        return driver

    except Exception as e:
        logger.error(f"❌ [ERROR] 라벨크래프트 로그인 실패: {str(e)}")
        if driver:
            driver.quit()
        return None


# 크롬 종료
def close_chrome(driver):
    if driver:
        driver.quit()


# 작업 페이지 이동
def move_to_prefix(driver):
    driver.get(HOME_URL)
    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, "reviewStart"))).click()
    time.sleep(5)


# 메인 실행
if __name__ == "__main__":
    user_id = input("아이디를 입력하세요: ")
    user_pw = input("비밀번호를 입력하세요: ")

    logger.info("✅ 라벨크래프트 로그인 시도 중...")
    my_driver = label_login(user_id, user_pw, headless=False)  # headless=True로 바꾸면 창 안 뜨고 실행됨

    if my_driver:
        move_to_prefix(my_driver)
        logger.info("✅ 모든 작업을 완료했습니다.")
        logger.info("✅ 5초후 자동으로 크롬을 닫습니다.")
        time.sleep(5)
        close_chrome(my_driver)
    else:
        logger.info("✅ 작업에 실패했습니다.")
        logger.info("✅ 라벨크래프트 작업 프로그램을 종료합니다.")
