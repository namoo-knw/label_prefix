import logging
import time
import os
from datetime import datetime
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

# 로컬 모듈 임포트
from label_admin import label_login, close_chrome, HOME_URL
from prefix_util import load_patterns_from_gsheet, process_page

# --- [로거 설정 (기존과 동일)] ---
LOG_LEVEL = logging.DEBUG
LOG_FORMAT = '[%(levelname)s] (%(name)s) %(asctime)s - %(message)s'
DATE_FORMAT = '%Y-%m-%d %H:%M:%S'
SAVE_FOLDER = "save"
os.makedirs(SAVE_FOLDER, exist_ok=True)
timestamp_str = datetime.now().strftime("%m%d_%H%M%S")
log_file_name = f"automation_{timestamp_str}.log"
LOG_FILENAME = os.path.join(SAVE_FOLDER, log_file_name)
logger = logging.getLogger("main_logger")
logger.info(f"텍스트 로그 파일이 {LOG_FILENAME} 경로에 저장됩니다.")
logger.setLevel(LOG_LEVEL)
formatter = logging.Formatter(fmt=LOG_FORMAT, datefmt=DATE_FORMAT)
if not logger.hasHandlers():
    console_handler = logging.StreamHandler()
    console_handler.setLevel(LOG_LEVEL)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    file_handler = logging.FileHandler(LOG_FILENAME, mode='a', encoding='utf-8')
    file_handler.setLevel(LOG_LEVEL)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)


# --- [로거 설정 끝] ---


# --- [수정된 main_task_loop 함수] ---
def main_task_loop(driver, patterns):
    """
    [수정된 메인 작업 루프]
    1. '작업 시작'을 단 한 번만 클릭.
    2. 새 창으로 단 한 번만 전환.
    3. 새 창 안에서 process_page를 무한 반복 (새 작업이 자동으로 로드된다고 가정).
    """
    try:
        # 1. 메인 페이지(HOME_URL)로 이동 (최초 1회)
        logger.info("🚀 메인 페이지로 이동하여 '작업 시작'을 클릭합니다...")
        driver.get(HOME_URL)
        original_window = driver.current_window_handle

        # 2. '작업 시작' 버튼 클릭 (최초 1회)
        WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "reviewStart"))
        ).click()
        logger.info("✅ '작업 시작' 클릭 완료. 새 창을 대기합니다...")

        # 3. 새 창 대기 및 전환 (최초 1회)
        WebDriverWait(driver, 15).until(EC.number_of_windows_to_be(2))

        all_windows = driver.window_handles
        new_window = None
        for window in all_windows:
            if window != original_window:
                new_window = window
                break

        if not new_window:
            logger.error("❌ 새 작업창을 찾지 못했습니다. 프로그램을 종료합니다.")
            return

        # [핵심] 새 작업창으로 영구적으로 전환합니다.
        driver.switch_to.window(new_window)
        logger.info(f"✅ 새 작업창으로 영구 전환 완료 (Handle: {new_window})")
        logger.info("이제 이 창 안에서 작업이 자동으로 갱신된다고 가정하고 루프를 시작합니다.")

        # 4. [수정] 새 창 안에서 무한 루프 시작
        while True:
            logger.info("=" * 50)
            logger.info("🚀 다음 작업 처리를 시작합니다 (현재 창 갱신 대기)...")

            try:
                # 5. 새 창에서 작업 처리 (prefix_util.py 함수 호출)
                #    (process_page가 끝나면 웹사이트가 자동으로 다음 작업을 로드한다고 가정)
                process_page(driver, patterns)

                # 6. 작업 처리 후, 웹사이트가 다음 작업을 로드할 시간을 줌
                logger.info("✅ 작업 처리 완료. 다음 작업이 로드될 때까지 2초 대기...")
                time.sleep(2)

            except Exception:
                # process_page에서 오류가 나도 루프는 계속되어야 함
                logger.error("❌ 작업 처리 중 오류 발생. 5초 후 다음 작업 시도.", exc_info=True)
                # (오류 시 새로고침 등이 필요하면 여기에 추가)
                # driver.refresh()
                time.sleep(5)

    except KeyboardInterrupt:
        logger.info("🛑 사용자가 Ctrl+C를 눌러 작업을 중단했습니다.")
    except Exception:
        # 새 창을 찾지 못하는 등의 치명적 오류
        logger.error(f"❌ 복구 불가능한 오류 발생. 작업 루프 종료.", exc_info=True)


# --- [수정 끝] ---


if __name__ == "__main__":
    # 사용자 입력
    user_id = input("아이디를 입력하세요: ")
    user_pw = input("비밀번호를 입력하세요: ")

    logger.info("🚀 라벨크래프트 자동 작업 시작")

    # 1. 로그인
    my_driver = label_login(user_id, user_pw, headless=False)  # headless=True → 창 안 뜨고 실행

    if my_driver:
        # 2. 구글시트에서 패턴 로드 (한 번만)
        patterns = load_patterns_from_gsheet()

        if not patterns:
            logger.error("❌ 구글시트에서 패턴을 불러오지 못했습니다. 작업을 종료합니다.")
            input("\n[!] 구글시트 패턴 로드 실패. 로그를 확인하세요.\n엔터 키를 누르면 프로그램을 종료합니다...")
        else:
            # 3. 메인 작업 루프 실행 (수정된 루프로 실행됨)
            main_task_loop(my_driver, patterns)

        # 4. 모든 작업 종료 후 드라이버 닫기
        logger.info("⏳ 5초 후 크롬을 종료합니다...")
        time.sleep(5)
        close_chrome(my_driver)
    else:
        logger.error("❌ 로그인 실패. 프로그램을 종료합니다.")
        input("\n[!] 로그인 실패. 아이디/비밀번호 또는 로그를 확인하세요.\n엔터 키를 누르면 프로그램을 종료합니다...")