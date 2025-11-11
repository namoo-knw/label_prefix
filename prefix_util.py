import time
import re
import logging
import gspread
import openpyxl  # 엑셀 로깅을 위해 추가
import os  # 엑셀 로깅을 위해 추가
from datetime import datetime  # 엑셀 로깅을 위해 추가
from urllib.parse import urlparse
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException  # 예외 처리용

# --- 로거 설정 ---
# main_logger를 가져와서 사용 (ui_main.py에서 이미 설정됨)
logger = logging.getLogger("main_logger")

# --- 구글시트 설정 ---
GSHEET_JSON = "indexcell-e71d69f270ca.json"  # 서비스 계정 JSON 파일
GSHEET_NAME = "[RPA] 테스트용"  # 구글시트 문서 이름
SHEET_NAME = "패턴단어"  # 시트 이름
PATTERN_COL_NUM = 3  # 단어가 들어있는 컬럼 (C열 = 3)

# --- 엑셀 로그 설정 ---
SAVE_FOLDER = "save"  # ui_main.py와 동일하게 'save' 폴더 사용
# [수정] 파일명은 스크립트 시작 시 1회만 생성
timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")  # 년도 추가됨
EXCEL_LOG_FILE = os.path.join(SAVE_FOLDER, f"log_{timestamp_str}.xlsx")
EXCEL_HEADER = ["작업시간", "href", "패턴 결과", "작업"]


def log_to_excel(timestamp, href, match_result, action):
    """
    작업 내역을 엑셀 파일에 한 줄씩 기록합니다.
    """
    data_row = [timestamp, href, match_result, action]

    try:
        # 1. 파일 존재 여부 확인
        if not os.path.exists(EXCEL_LOG_FILE):
            logger.info(f"새 엑셀 로그 파일 생성: {EXCEL_LOG_FILE}")
            # 새 워크북(엑셀 파일) 생성 및 헤더 추가
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "AutomationLog"
            sheet.append(EXCEL_HEADER)  # 헤더 추가
        else:
            # 기존 파일 열기
            workbook = openpyxl.load_workbook(EXCEL_LOG_FILE)
            sheet = workbook.active

        # 2. 데이터 행 추가
        sheet.append(data_row)

        # 3. 파일 저장
        workbook.save(EXCEL_LOG_FILE)
        logger.debug(f"엑셀 로그 저장 완료: {timestamp}")

    except Exception:
        logger.error(f"❌ 엑셀 로그 저장 실패. 데이터: {data_row}", exc_info=True)


def load_patterns_from_gsheet():
    """구글시트에서 패턴 단어 불러오기"""
    logger.info("구글시트에서 패턴 단어 불러오는 중...")
    try:
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_name(GSHEET_JSON, scope)
        client = gspread.authorize(creds)
        sheet = client.open(GSHEET_NAME).worksheet(SHEET_NAME)

        # C열 전체 읽기 (첫 번째 행 제외)
        patterns = sheet.col_values(PATTERN_COL_NUM)[1:]
        # 빈 값 제거
        patterns = [p.strip() for p in patterns if p.strip()]

        logger.info(f"✅ 패턴 단어 {len(patterns)}개 불러옴 (예: {patterns[:3]}...)")
        return patterns
    except Exception:
        logger.error("❌ 구글시트 불러오기 실패.", exc_info=True)
        return None


def extract_domain_name(href: str) -> str:
    """URL에서 순수 도메인 이름만 추출 (예: https://salvla1234.tistory.com → salvla1234)"""
    try:
        parsed = urlparse(href)
        netloc = parsed.netloc
        if ".tistory.com" in netloc:
            # .tistory.com 앞부분을 반환
            domain_part = netloc.split(".tistory.com")[0]
            logger.debug(f"Tistory 도메인 파싱: {netloc} -> {domain_part}")
            return domain_part

        # tistory가 아닌 경우 (예: example.com)
        logger.debug(f"일반 도메인 파싱: {netloc}")
        return netloc
    except Exception:
        logger.warning(f"⚠ 도메인 파싱 실패 ({href})", exc_info=True)
        return href


def check_href_match(href, patterns):
    """도메인 부분이 '패턴단어 + 숫자 4자리 이상' 형식에 부합하는지 확인"""
    domain_name = extract_domain_name(href)
    logger.info(f"🔍 비교 대상 도메인: {domain_name}")

    for word in patterns:
        # 정규식 생성: (단어) + (숫자 4자리 이상) + (문자열 끝)
        regex = rf"^{re.escape(word)}\d{{4,}}$"
        logger.debug(f"    [패턴 검사] {domain_name} vs {regex}")

        if re.match(regex, domain_name):
            logger.info(f"✅ 정규식 일치: {regex} ← {domain_name}")
            return True, word  # (일치함, 일치한 단어)

    logger.info("❌ 정규식 불일치.")
    return False, None  # (일치 안 함, None)


def process_page(driver, patterns):
    """
    현재 페이지에서 href 검사 후 동작 실행.
    [수정] UI 스레드에 보낼 값을 반환(return)합니다.
    """
    # 로깅용 변수 초기화
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    href_result = "N/A"
    match_result = "N/A"
    action_taken = "N/A"

    try:
        # 1️⃣ href 추출 단계
        logger.info("👉 [1/4] href 추출 시도 중...")
        try:
            # presence_of_all_elements_located: 해당 CSS를 가진 *모든* 요소를 찾음
            href_elems = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "span.h5 a"))
            )
            logger.debug(f"발견된 href 요소 개수: {len(href_elems)}")
            # 그 중 마지막 요소 (가장 하단의 링크)
            href_elem = href_elems[-1]
            href = href_elem.get_attribute("href")
            href_result = href  # 로깅 변수에 저장
            logger.info(f"🔗 href 추출 성공: {href}")
        except TimeoutException:
            logger.error("❌ [1/4 실패] href 요소를 10초 내에 찾지 못했습니다.")
            match_result = "href 추출 실패"
            action_taken = "오류"
            return  # 'finally' 블록으로 이동
        except Exception:
            logger.error("❌ [1/4 실패] href 추출 중 알 수 없는 오류", exc_info=True)
            match_result = "href 추출 오류"
            action_taken = "오류"
            return  # 'finally' 블록으로 이동

        # 2️⃣ 패턴 일치 여부 확인
        logger.info("👉 [2/4] 패턴 일치 여부 확인 중...")
        try:
            is_match, matched_word = check_href_match(href_result, patterns)
        except Exception:
            logger.error(f"❌ [2/4 실패] 패턴 검사 중 오류", exc_info=True)
            match_result = "패턴 검사 오류"
            action_taken = "오류"
            return  # 'finally' 블록으로 이동

        # 3️⃣ 패턴 일치 시 키보드 입력
        if is_match:
            try:
                logger.info(f"👉 [3/4] 패턴 '{matched_word}' 일치 → 키보드 'E' 입력 시도")
                match_result = f"일치 ({matched_word})"

                actions = ActionChains(driver)
                actions.send_keys('e').perform()

                logger.info("⌨️ 'E' 키 입력 완료")
                action_taken = "E (패턴 일치)"
            except Exception:
                logger.error(f"❌ [3/4 실패] 키보드 'E' 입력 중 오류", exc_info=True)
                action_taken = "E 입력 오류"
                return

        # 4️⃣ 패턴 불일치 시 '작업 미루기' 버튼 클릭
        else:
            try:
                logger.info("👉 [3/4] 패턴 불일치 → '작업 미루기' 버튼 클릭 시도 중...")
                match_result = "불일치"

                postpone_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[text()='작업 미루기']"))
                )
                postpone_btn.click()
                logger.info("✅ '작업 미루기' 버튼 클릭 완료")

                logger.info("👉 [4/4] '아무에게나 미루기' 버튼 클릭 시도 중...")
                assign_any_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[text()='아무에게나 미루기']"))
                )
                assign_any_btn.click()
                logger.info("✅ '아무에게나 미루기' 버튼 클릭 완료")
                action_taken = "작업 미루기"
            except Exception:
                logger.error(f"❌ [4/4 실패] '작업 미루기' 버튼 클릭 중 오류", exc_info=True)
                action_taken = "미루기 오류"
                return

    except Exception:
        logger.error(f"❌ [기타 예외] 페이지 처리 중 알 수 없는 오류", exc_info=True)
        action_taken = "알 수 없는 오류"

    finally:
        # 5️⃣ 엑셀에 결과 로깅
        # (오류가 발생했든 성공했든, 지금까지의 결과를 기록)

        # --- [오류 수정] ---
        # 로깅할 데이터를 먼저 리스트로 만듭니다.
        data_to_log = [timestamp, href_result, match_result, action_taken]

        # [수정] 'data_row' 대신 'data_to_log' 변수를 사용합니다.
        logger.info(f"📋 엑셀 로그 기록 시도: {data_to_log}")
        log_to_excel(timestamp, href_result, match_result, action_taken)

        # 6️⃣ 대기 및 다음 항목 준비 (UI 스레드로 반환)
        logger.debug(f"3초 대기 후 현재 창을 닫습니다...")
        time.sleep(3)

        # ui_main.py의 Worker 스레드가 이 값을 받을 수 있도록 반환(return)합니다.
        return href_result, match_result, action_taken