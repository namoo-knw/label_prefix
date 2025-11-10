import time
import re
import logging
import gspread
import openpyxl  # 엑셀 처리를 위해 추가
import os  # 파일 존재 여부 확인을 위해 추가
from datetime import datetime  # 타임스탬프를 위해 추가
from openpyxl.utils.exceptions import InvalidFileException
from urllib.parse import urlparse
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

# --- 로거 설정 ---
# main.py에서 설정한 로거를 가져옵니다.
logger = logging.getLogger("main_logger")

# --- 구글시트 설정 ---
GSHEET_JSON = "indexcell-e71d69f270ca.json"  # 서비스 계정 JSON 파일
GSHEET_NAME = "[RPA] 테스트용"  # 구글시트 문서 이름
SHEET_NAME = "패턴단어"  # 시트 이름
PATTERN_COL_NUM = 3  # 단어가 들어있는 컬럼 (A=1, B=2, C=3)

# --- 엑셀 로그 설정 ---
EXCEL_LOG_FILE = "automation_log.xlsx"
EXCEL_HEADERS = ["작업시간", "추출된 href", "패턴 일치 결과", "작업 방법"]


def log_to_excel(log_data: list):
    """
    작업 내역을 엑셀 파일에 한 줄 추가합니다.
    log_data: 엑셀 헤더 순서에 맞는 데이터 리스트
    """
    try:
        # 1. 파일이 존재하는지 확인
        if not os.path.exists(EXCEL_LOG_FILE):
            # 1a. 파일이 없으면: 새 워크북 생성 및 헤더 추가
            logger.info(f"새 엑셀 로그 파일 생성: {EXCEL_LOG_FILE}")
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Automation Log"
            sheet.append(EXCEL_HEADERS)
        else:
            # 1b. 파일이 있으면: 기존 워크북 로드
            workbook = openpyxl.load_workbook(EXCEL_LOG_FILE)
            sheet = workbook.active

        # 2. 데이터 행 추가
        sheet.append(log_data)

        # 3. 파일 저장
        workbook.save(EXCEL_LOG_FILE)
        logger.debug(f"엑셀 로그 저장 완료: {log_data[0]}")

    except InvalidFileException:
        logger.error(f"❌ 엑셀 파일({EXCEL_LOG_FILE})이 손상되었거나 엑셀 파일이 아닙니다. 로그를 기록할 수 없습니다.")
    except PermissionError:
        logger.warning(f"⚠ 엑셀 파일({EXCEL_LOG_FILE})이 다른 프로그램에서 열려있어 저장할 수 없습니다. (파일을 닫아주세요)")
    except Exception:
        logger.error("❌ 엑셀 로그 저장 중 알 수 없는 오류 발생", exc_info=True)


def load_patterns_from_gsheet():
    """구글시트에서 패턴 단어 불러오기"""
    try:
        logger.info("구글시트에서 패턴 단어 불러오는 중...")
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_name(GSHEET_JSON, scope)
        client = gspread.authorize(creds)
        sheet = client.open(GSHEET_NAME).worksheet(SHEET_NAME)

        # C열 전체 읽기 (첫 번째 행 제외)
        patterns = sheet.col_values(PATTERN_COL_NUM)[1:]
        # 빈 문자열 제거
        patterns = [p.strip() for p in patterns if p.strip()]

        logger.info(f"✅ 패턴 단어 {len(patterns)}개 불러옴 (예: {patterns[:3]}...)")
        return patterns
    except Exception:
        logger.error(f"❌ 구글시트 불러오기 실패", exc_info=True)
        return []


def extract_domain_name(href: str) -> str:
    """URL에서 순수 도메인 이름만 추출 (예: https://salvla1234.tistory.com → salvla1234)"""
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
    """도메인 부분이 '패턴단어 + 숫자 4자리 이상' 형식에 부합하는지 확인"""
    if not href:
        logger.warning("⚠ href가 비어있어 패턴 검사를 건너뜁니다.")
        return False, None

    domain_name = extract_domain_name(href)
    logger.info(f"🔍 비교 대상 도메인: {domain_name}")

    for word in patterns:
        # 정규식 생성: (단어)(숫자 4개 이상)
        regex = rf"^{re.escape(word)}\d{{4,}}$"
        logger.debug(f"   [패턴 검사] {domain_name} vs {regex}")

        if re.match(regex, domain_name):
            logger.info(f"✅ 정규식 일치: {regex} ← {domain_name}")
            return True, word

    logger.info("❌ 정규식 불일치.")
    return False, None


def process_page(driver, patterns):
    """현재 페이지에서 href 검사 후 동작 실행 및 엑셀 로그 기록"""

    # --- 로깅용 변수 초기화 ---
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    href_result = "N/A"
    match_result = "N/A"
    action_taken = "알 수 없는 오류"  # 기본값을 오류로 설정

    try:
        # 1️⃣ href 추출 단계
        logger.info("👉 [1/4] href 추출 시도 중...")
        href = None  # href 변수 초기화
        try:
            # CSS 선택자 span.h5 a 가 여러 개일 경우를 대비해 마지막 요소를 찾습니다.
            href_elems = WebDriverWait(driver, 15).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "span.h5 a"))
            )
            logger.debug(f"발견된 href 요소 개수: {len(href_elems)}")

            href_elem = href_elems[-1]  # 가장 마지막 요소 선택
            href = href_elem.get_attribute("href")

            if not href:
                logger.warning("⚠ href 요소를 찾았으나 href 속성이 비어있습니다.")
                href_result = "속성 없음"  # 엑셀 기록용
            else:
                logger.info(f"🔗 href 추출 성공: {href}")
                href_result = href  # 엑셀 기록용

        except Exception:
            # 요소를 못찾으면 TimeoutException 등이 발생합니다.
            logger.warning(f"⚠ [1단계 경고] href 요소(span.h5 a)를 찾는 데 실패했습니다. '작업 미루기'로 진행합니다.", exc_info=True)
            href_result = "요소 없음"  # 엑셀 기록용

        # 2️⃣ 패턴 일치 여부 확인
        logger.info("👉 [2/4] 패턴 일치 여부 확인 중...")
        is_match = False
        matched_word = None
        try:
            # href가 None이 아닐 경우에만 패턴 검사
            if href:
                is_match, matched_word = check_href_match(href, patterns)
                match_result = f"일치 ({matched_word})" if is_match else "불일치"  # 엑셀 기록용
            else:
                logger.info("href가 없어 패턴 검사를 건너뜁니다.")
                match_result = "검사 안함 (href 없음)"  # 엑셀 기록용

        except Exception:
            logger.error(f"❌ [2단계 실패] 패턴 검사 중 예외 발생", exc_info=True)
            match_result = "패턴 검사 오류"
            # 여기서 return하지 않고, '작업 미루기'로 진행되도록 합니다.

        # 3️⃣ 패턴 일치 시 키보드 입력
        if is_match:
            try:
                logger.info(f"👉 [3/4] 패턴 '{matched_word}' 일치 → 키보드 'E' 입력 시도")
                actions = ActionChains(driver)
                actions.send_keys('e').perform()
                logger.info("⌨️ 'E' 키 입력 완료")
                action_taken = "E (패턴 일치)"  # 엑셀 기록용
            except Exception:
                logger.error(f"❌ [3단계 실패] 키보드 입력 중 오류", exc_info=True)
                action_taken = "E (입력 실패)"  # 엑셀 기록용

        # 4️⃣ 패턴 불일치 시 '작업 미루기' 버튼 클릭 (is_match가 False이거나, 패턴 검사 오류 시)
        else:
            try:
                logger.info("👉 [3/4] 패턴 불일치 → '작업 미루기' 버튼 클릭 시도 중...")
                postpone_btn = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[text()='작업 미루기']"))
                )
                postpone_btn.click()
                logger.info("✅ '작업 미루기' 버튼 클릭 완료")

                logger.info("👉 [4/4] '아무에게나 미루기' 버튼 클릭 시도 중...")
                # '작업 미루기' 클릭 후 나타나는 팝업 메뉴(모달) 대기
                assign_any_btn = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[text()='아무에게나 미루기']"))
                )
                assign_any_btn.click()
                logger.info("✅ '아무에게나 미루기' 버튼 클릭 완료")
                action_taken = "작업 미루기"  # 엑셀 기록용
            except Exception:
                logger.error(f"❌ [4단계 실패] '작업 미루기' 또는 '아무에게나 미루기' 버튼 클릭 중 오류", exc_info=True)
                action_taken = "작업 미루기 (클릭 실패)"  # 엑셀 기록용

        # 5️⃣ 대기 (다음 작업 전 안정성을 위해)
        logger.debug("3초 대기 후 현재 창을 닫습니다...")
        time.sleep(3)

    except Exception:
        logger.error(f"❌ [기타 예외] 페이지 처리 중 알 수 없는 오류", exc_info=True)
        action_taken = "페이지 처리 중 심각한 오류"  # 엑셀 기록용

    finally:
        # --- [최종 로깅] ---
        # try가 성공하든, except로 빠지든 항상 실행되어 엑셀 로그를 남깁니다.
        log_data = [timestamp, href_result, match_result, action_taken]
        logger.info(f"📋 엑셀 로그 기록 시도: {log_data}")
        log_to_excel(log_data)