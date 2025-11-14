import logging
import os
import sys  # <--- sys 임포트
from datetime import datetime


# --- [경로 헬퍼 함수] ---
def get_base_path():
    """
    .exe로 실행되었는지 (frozen), 스크립트로 실행되었는지 확인하고
    기준 경로를 반환합니다.
    """
    if getattr(sys, 'frozen', False):
        # .exe로 실행된 경우: .exe 파일이 있는 폴더
        return os.path.dirname(sys.executable)
    else:
        # 스크립트로 실행된 경우: .py 파일이 있는 폴더
        return os.path.abspath(".")


def resource_path(relative_path):
    """
    .exe에 포함된 *읽기 전용* 파일(ui, json)의 실제 경로를 찾습니다.
    """
    try:
        # PyInstaller가 임시 폴더(_MEIPASS)에 파일을 풉니다.
        base_path = sys._MEIPASS
    except Exception:
        # 스크립트로 실행된 경우: 현재 폴더
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# --- [로그 설정 상수] ---
LOG_LEVEL = logging.DEBUG
LOG_FORMAT = '[%(levelname)s] (%(name)s) %(asctime)s - %(message)s'
DATE_FORMAT = '%Y-%m-%d %H:%M:%S'
SAVE_FOLDER = "save"

# [수정] 쓰기 경로는 get_base_path()를 기준으로 합니다.
# .exe 파일이 있는 폴더 하위에 'save' 폴더를 만듭니다.
OUTPUT_DIR = os.path.join(get_base_path(), SAVE_FOLDER)


def setup_logger():
    """
    'main_logger'라는 이름의 로거를 설정하고 반환합니다.
    """

    # [수정] OUTPUT_DIR 사용
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 2. 로그 파일명을 위한 현재 시간 생성 (년월일_시분초)
    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file_name = f"automation_{timestamp_str}.log"

    # [수정] OUTPUT_DIR 사용
    LOG_FILENAME = os.path.join(OUTPUT_DIR, log_file_name)

    # 3. 로거 가져오기 (이하 동일)
    logger = logging.getLogger("main_logger")
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

    return logger, LOG_FILENAME