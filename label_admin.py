import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from knw_Chromedriver_manager import Chromedriver_manager

# --- [전역 변수] ---
LOGIN_URL = "http://label-craft.is.kakaocorp.com/login"
HOME_URL = "http://label-craft.is.kakaocorp.com/tasks/45/"

# --- [로거 설정] ---
# logging.basicConfig는 main.py에서 한 번만 호출합니다.
# 여기서는 main.py에서 설정한 로거를 이름으로 가져옵니다.
logger = logging.getLogger("main_logger")


# 로그인
def label_login(username, password, headless=True):
    """
    라벨크래프트 사이트에 로그인합니다.
    성공 시 드라이버 객체를, 실패 시 None을 반환합니다.
    """
    driver = None
    try:
        options = Options()
        options.add_argument("--start-maximized")

        if headless:
            logger.info("헤드리스 모드로 실행합니다.")
            options.add_argument("--headless=new")
            options.add_argument("--window-size=1920,1080")
            options.add_argument("--disable-gpu")

        # 팝업 차단 해제 (팝업 메뉴 대응)
        options.add_experimental_option("prefs", {
            "profile.default_content_setting_values.popups": 1
        })

        driver_path = Chromedriver_manager.install()
        if not driver_path:
            logger.error("❌ Chromedriver를 설치하거나 찾는 데 실패했습니다.")
            return None

        logger.debug(f"크롬드라이버 경로: {driver_path}")
        driver = webdriver.Chrome(service=Service(driver_path), options=options)

        logger.info(f"{LOGIN_URL} 페이지로 이동합니다...")
        driver.get(LOGIN_URL)

        # 아이디 입력
        logger.debug("아이디 입력란 대기 중...")
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "exampleInputEmail"))
        ).send_keys(username)
        logger.debug("아이디 입력 완료.")

        # 비밀번호 입력
        logger.debug("비밀번호 입력란 대기 중...")
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "exampleInputPassword"))
        ).send_keys(password + Keys.RETURN)
        logger.debug("비밀번호 입력 및 Enter 전송 완료.")

        # 로그인 성공 대기 (사이드바 로고 이미지 확인)
        logger.debug("로그인 성공 대기 중... (사이드바 로고 확인)")
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#accordionSidebar > a > img"))
        )
        logger.info("✅ 라벨크래프트 로그인에 성공했습니다.")

        return driver

    except Exception:
        logger.error(f"❌ [ERROR] 라벨크래프트 로그인 실패", exc_info=True)
        if driver:
            driver.quit()
        return None


# 크롬 종료
def close_chrome(driver):
    """
    웹 드라이버를 종료합니다.
    """
    if driver:
        logger.info("크롬 드라이버를 종료합니다.")
        driver.quit()