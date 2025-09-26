import gspread
from oauth2client.service_account import ServiceAccountCredentials
import logging



# --- [로거 설정] ---
logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(asctime)s - %(message)s')
logger = logging.getLogger("prefix_logger")

# --- [구글 시트 연결 설정] ---
def load_patterns_from_gsheet():
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        "rpa-study-458606-4b6b76d1acf4.json", scope
    )
    client = gspread.authorize(creds)

    # 문서 열기
    sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1FrJNlluakEjYkoqpvaftM056YLbR1gw9u3XvtRlSLXY/edit")
    worksheet = sheet.worksheet("패턴단어")

    # C열 데이터 불러오기 (헤더 제외)
    patterns = worksheet.col_values(3)[1:]
    return patterns


# --- [href 매칭 확인] ---
def check_href_match(href, patterns):
    for word in patterns:
        if word.strip() and word in href:
            return True, word  # 매칭된 단어 반환
    return False, None


# --- [실행 예시] ---
if __name__ == "__main__":
    href = "https://write88721.tistory.com/"
    patterns = load_patterns_from_gsheet()

    is_match, matched_word = check_href_match(href, patterns)

    if is_match:
        print(f"✅ 매칭됨: '{matched_word}' 가 href에 포함됨")
    else:
        print("❌ 매칭되는 단어 없음")
