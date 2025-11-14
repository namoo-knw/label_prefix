import time
import re
import logging
import gspread
import openpyxl
import os
from datetime import datetime
from urllib.parse import urlparse
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException

# [수정] label_log에서 헬퍼 함수들을 임포트
from label_log import resource_path, OUTPUT_DIR

logger = logging.getLogger("main_logger")

# --- 구글시트 설정 ---
# [수정] GSHEET_JSON은 이제 '이름'만 가리킵니다.
GSHEET_JSON = "indexcell-e71d69f270ca.json"
GSHEET_NAME = "[RPA] 테스트용"
SHEET_NAME = "패턴단어"
PATTERN_COL_NUM = 3

# --- 엑셀 로그 설정 ---
# [수정] OUTPUT_DIR을 사용합니다.
timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
EXCEL_LOG_FILE = os.path.join(OUTPUT_DIR, f"log_{timestamp_str}.xlsx")
EXCEL_HEADER = ["작업시간", "href", "패턴 결과", "작업"]


def log_to_excel(timestamp, href, match_result, action):
    # (이 함수 내부는 수정할 필요 없음)
    data_row = [timestamp, href, match_result, action]
    try:
        if not os.path.exists(EXCEL_LOG_FILE):
            logger.info(f"새 엑셀 로그 파일 생성: {EXCEL_LOG_FILE}")
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "AutomationLog"
            sheet.append(EXCEL_HEADER)
        else:
            workbook = openpyxl.load_workbook(EXCEL_LOG_FILE)
            sheet = workbook.active
        sheet.append(data_row)
        workbook.save(EXCEL_LOG_FILE)
        logger.debug(f"엑셀 로그 저장 완료: {timestamp}")
    except Exception:
        logger.error(f"❌ 엑셀 로그 저장 실패. 데이터: {data_row}", exc_info=True)


def load_patterns_from_gsheet():
    logger.info("구글시트에서 패턴 단어 불러오는 중...")
    try:
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]

        # --- [수정된 부분] ---
        # GSHEET_JSON 대신 resource_path(GSHEET_JSON)을 사용합니다.
        # .exe 안에 포함된 .json 파일의 실제 경로를 찾아옵니다.
        json_keyfile_path = resource_path(GSHEET_JSON)
        logger.debug(f"JSON 키 파일 경로: {json_keyfile_path}")
        creds = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile_path, scope)
        # --- [수정 완료] ---

        client = gspread.authorize(creds)
        sheet = client.open(GSHEET_NAME).worksheet(SHEET_NAME)

        patterns = sheet.col_values(PATTERN_COL_NUM)[1:]
        patterns = [p.strip() for p in patterns if p.strip()]
        logger.info(f"✅ 패턴 단어 {len(patterns)}개 불러옴 (예: {patterns[:3]}...)")
        return patterns
    except Exception:
        logger.error(f"❌ 구글시트 불러오기 실패. {GSHEET_JSON} 파일 경로 확인.", exc_info=True)
        return None


# (extract_domain_name, check_href_match, process_page 함수는 수정할 필요 없음)
# ... (이전과 동일한 process_page 함수 내용) ...
def extract_domain_name(href: str) -> str:
    try:
        parsed = urlparse(href)
        netloc = parsed.netloc
        if ".tistory.com" in netloc:
            domain_part = netloc.split(".tistory.com")[0]
            logger.debug(f"Tistory 도메인 파싱: {netloc} -> {domain_part}")
            return domain_part
        logger.debug(f"일반 도메인 파싱: {netloc}")
        return netloc
    except Exception:
        logger.warning(f"⚠ 도메인 파싱 실패 ({href})", exc_info=True)
        return href


def check_href_match(href, patterns):
    domain_name = extract_domain_name(href)
    logger.info(f"🔍 비교 대상 도메인: {domain_name}")
    for word in patterns:
        regex = rf"^{re.escape(word)}\d{{4,}}$"
        logger.debug(f"    [패턴 검사] {domain_name} vs {regex}")
        if re.match(regex, domain_name):
            logger.info(f"✅ 정규식 일치: {regex} ← {domain_name}")
            return True, word
    logger.info("❌ 정규식 불일치.")
    return False, None


def process_page(driver, patterns):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    href_result = "N/A"
    match_result = "N/A"
    action_taken = "N/A"
    try:
        logger.info("👉 [1/4] href 추출 시도 중...")
        try:
            href_elems = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "span.h5 a"))
            )
            logger.debug(f"발견된 href 요소 개수: {len(href_elems)}")
            href_elem = href_elems[-1]
            href = href_elem.get_attribute("href")
            href_result = href
            logger.info(f"🔗 href 추출 성공: {href}")
        except TimeoutException:
            logger.error("❌ [1/4 실패] href 요소를 10초 내에 찾지 못했습니다.")
            match_result = "href 추출 실패"
            action_taken = "오류"
            return href_result, match_result, action_taken
        except Exception:
            logger.error("❌ [1/4 실패] href 추출 중 알 수 없는 오류", exc_info=True)
            match_result = "href 추출 오류"
            action_taken = "오류"
            return href_result, match_result, action_taken

        logger.info("👉 [2/4] 패턴 일치 여부 확인 중...")
        try:
            is_match, matched_word = check_href_match(href_result, patterns)
        except Exception:
            logger.error(f"❌ [2/4 실패] 패턴 검사 중 오류", exc_info=True)
            match_result = "패턴 검사 오류"
            action_taken = "오류"
            return href_result, match_result, action_taken

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
    except Exception:
        logger.error(f"❌ [기타 예외] 페이지 처리 중 알 수 없는 오류", exc_info=True)
        action_taken = "알 수 없는 오류"
    finally:
        data_to_log = [timestamp, href_result, match_result, action_taken]
        logger.info(f"📋 엑셀 로그 기록 시도: {data_to_log}")
        log_to_excel(timestamp, href_result, match_result, action_taken)
        logger.debug(f"3초 대기 후 현재 창을 닫습니다...")
        time.sleep(3)
        return href_result, match_result, action_taken