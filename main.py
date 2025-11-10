import logging
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

# 로컬 모듈 임포트
from label_admin import label_login, close_chrome, HOME_URL
from prefix_util import load_patterns_from_gsheet, process_page

# --- [로거 설정] ---
# 로그 레벨을 DEBUG로 설정하면 Selenium의 상세 로그까지 볼 수 있습니다.
# INFO로 변경하면 조금 더 간결한 로그를 봅니다.
LOG_LEVEL = logging.DEBUG
LOG_FORMAT = '[%(levelname)s] (%(name)s) %(asctime)s - %(message)s'
DATE_FORMAT = '%Y-%m-%d %H:%M:%S'
LOG_FILENAME = "automation.log"  # 로그 파일 이름

# 1. 로거 가져오기
logger = logging.getLogger("main_logger")
logger.setLevel(LOG_LEVEL)  # 로거의 최소 레벨 설정

# 2. 포맷터 생성
formatter = logging.Formatter(fmt=LOG_FORMAT, datefmt=DATE_FORMAT)

# 3. 핸들러가 이미 설정되었는지 확인 (중복 추가 방지)
if not logger.hasHandlers():
    # 3-1. 콘솔 핸들러 (StreamHandler)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(LOG_LEVEL)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # 3-2. 파일 핸들러 (FileHandler)
    # mode='a' (append, 이어쓰기), encoding='utf-8' (한글 깨짐 방지)
    file_handler = logging.FileHandler(LOG_FILENAME, mode='a', encoding='utf-8')
    file_handler.setLevel(LOG_LEVEL)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)


# (기존 logging.basicConfig 라인 삭제됨)
# (기존 logger = logging.getLogger("main_logger") 라인은 위로 이동됨)


def main_task_loop(driver, patterns):
    """
    메인 작업 루프:
    1. 메인 페이지 이동
    2. '작업 시작' 클릭
    3. 새 창으로 전환
    4. 새 창에서 'process_page' 실행
    5. 새 창 닫기
    6. 메인 창으로 복귀
    7. 반복
    """
    try:
        # 시작 시점의 메인 윈도우 핸들 저장
        original_window = driver.current_window_handle
        logger.debug(f"메인 윈도우 핸들 저장: {original_window}")

        while True:
            logger.info("=" * 50)
            logger.info("🚀 새 작업 시작: 메인 페이지로 이동합니다...")

            try:
                # 1. 메인 페이지(HOME_URL)로 이동
                driver.get(HOME_URL)

                # 2. '작업 시작' 버튼 클릭
                logger.info("'작업 시작' 버튼(#reviewStart)을 찾아 클릭합니다...")
                WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.ID, "reviewStart"))
                ).click()
                logger.info("✅ '작업 시작' 클릭 완료. 새 창을 대기합니다...")

                # 3. 새 창 대기 및 전환 (핵심 수정 사항)
                # 2개의 창이 열릴 때까지 대기
                WebDriverWait(driver, 15).until(EC.number_of_windows_to_be(2))

                all_windows = driver.window_handles
                new_window = None
                for window in all_windows:
                    if window != original_window:
                        new_window = window
                        break

                if new_window:
                    driver.switch_to.window(new_window)
                    logger.info(f"✅ 새 작업창으로 전환 완료 (Handle: {new_window})")
                else:
                    logger.warning("⚠ 새 창을 찾지 못했습니다. 1초 대기 후 루프를 다시 시작합니다.")
                    time.sleep(1)
                    continue  # while 루프 처음으로

                # 4. 새 창에서 작업 처리 (prefix_util.py 함수 호출)
                process_page(driver, patterns)

                # 5. 작업 완료된 새 창 닫기
                logger.info("작업창 처리가 완료되었습니다. 현재 창을 닫습니다...")
                driver.close()

                # 6. 드라이버 포커스를 메인 창으로 복귀
                driver.switch_to.window(original_window)
                logger.info(f"✅ 메인 창으로 복귀 완료 (Handle: {original_window})")

                # 7. (안정성을 위한) 다음 작업 전 짧은 대기
                time.sleep(2)

            except Exception:
                logger.error("❌ 작업 루프 중 오류 발생. 5초 후 다음 루프 시도.", exc_info=True)
                # 메인 창으로 복귀 시도 (오류 발생 시 창 상태가 불명확할 수 있음)
                try:
                    # 현재 창이 2개 이상이면, 메인 창 제외하고 닫기
                    if len(driver.window_handles) > 1:
                        all_windows = driver.window_handles
                        for window in all_windows:
                            if window != original_window:
                                driver.switch_to.window(window)
                                driver.close()
                    driver.switch_to.window(original_window)
                    logger.info("오류 복구: 메인 창으로 강제 복귀")
                except Exception as e_recovery:
                    logger.fatal(f"FATAL: 메인 창 복구 실패. {e_recovery}")
                    raise  # 복구 불가능 시 프로그램 종료
                time.sleep(5)


    except KeyboardInterrupt:
        logger.info("🛑 사용자가 Ctrl+C를 눌러 작업을 중단했습니다.")
    except Exception:
        logger.error(f"❌ 복구 불가능한 오류 발생. 작업 루프 종료.", exc_info=True)


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
        else:
            # 3. 메인 작업 루프 실행
            main_task_loop(my_driver, patterns)

        # 4. 모든 작업 종료 후 드라이버 닫기
        logger.info("⏳ 5초 후 크롬을 종료합니다...")
        time.sleep(5)
        close_chrome(my_driver)
    else:
        logger.error("❌ 로그인 실패. 프로그램을 종료합니다.")