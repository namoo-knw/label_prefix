import sys
import logging
import os
import time
from datetime import datetime
from PySide6.QtWidgets import QApplication, QMessageBox
from PySide6.QtCore import QThread, Signal, Slot
from PySide6.QtUiTools import QUiLoader

# --- [수정] 빠뜨렸던 Selenium Import 구문 추가 ---
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
# --- [수정 끝] ---

# 로컬 모듈 임포트
from label_admin import label_login, close_chrome, HOME_URL
from prefix_util import load_patterns_from_gsheet, process_page

# --- [로거 설정] ---
LOG_LEVEL = logging.DEBUG
LOG_FORMAT = '[%(levelname)s] (%(name)s) %(asctime)s - %(message)s'
DATE_FORMAT = '%Y-%m-%d %H:%M:%S'
SAVE_FOLDER = "save"
os.makedirs(SAVE_FOLDER, exist_ok=True)
timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")  # 년도 추가됨
log_file_name = f"automation_{timestamp_str}.log"
LOG_FILENAME = os.path.join(SAVE_FOLDER, log_file_name)

logger = logging.getLogger("main_logger")
logger.setLevel(LOG_LEVEL)
formatter = logging.Formatter(fmt=LOG_FORMAT, datefmt=DATE_FORMAT)

if not logger.hasHandlers():
    # 1. 콘솔 핸들러 (StreamHandler)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(LOG_LEVEL)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # 2. 파일 핸들러 (FileHandler)
    file_handler = logging.FileHandler(LOG_FILENAME, mode='a', encoding='utf-8')
    file_handler.setLevel(LOG_LEVEL)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

logger.info(f"UI 모드 자동화 작업 시작. 로그 파일: {LOG_FILENAME}")


# --- [백그라운드 Selenium 작업을 위한 QThread] ---

class Worker(QThread):
    """
    Selenium 백그라운드 작업을 처리할 스레드.
    UI가 멈추는 것을 방지합니다.
    """
    # UI로 보낼 신호(Signal) 정의
    status_updated = Signal(str)
    work_finished_one = Signal(int, str)
    automation_finished = Signal(str)
    login_result = Signal(bool, str)

    def __init__(self, user_id, user_pw, headless, parent=None):
        super().__init__(parent)
        self.user_id = user_id
        self.user_pw = user_pw
        self.headless = headless
        self.driver = None
        self.patterns = []
        self.total_count = 0
        self._is_running = True  # 스레드 중지 플래그

    def run(self):
        """
        스레드가 .start()될 때 실행되는 메인 함수
        """
        try:
            self.status_updated.emit("구글 시트에서 패턴 로드 중...")
            self.patterns = load_patterns_from_gsheet()
            if not self.patterns:
                self.login_result.emit(False, "❌ 구글시트에서 패턴을 불러오지 못했습니다.")
                return
            self.status_updated.emit("패턴 로드 완료. 로그인 시도 중...")

            # 1. 로그인
            self.driver = label_login(self.user_id, self.user_pw, self.headless)

            if not self.driver:
                self.login_result.emit(False, "❌ 로그인 실패. 아이디/비밀번호를 확인하세요.")
                return

            self.login_result.emit(True, "✅ 로그인 성공!")

            # 2. 메인 작업 루프 (main.py의 main_task_loop 로직)
            self.main_task_loop()

        except Exception as e:
            logger.error(f"Worker 스레드 실행 중 오류: {e}", exc_info=True)
            self.automation_finished.emit(f"❌ 작업 중 심각한 오류 발생: {e}")
        finally:
            if self.driver:
                close_chrome(self.driver)
            logger.info("Worker 스레드 종료.")

    def stop(self):
        """
        '작업 중지' 버튼이 호출할 함수
        """
        self.status_updated.emit("🛑 작업 중지 요청됨... 현재 작업 완료 후 종료합니다.")
        self._is_running = False

    def main_task_loop(self):
        """
        (이전 main.py의 main_task_loop 로직)
        UI 스레드에 맞게 수정됨
        """
        original_window = self.driver.current_window_handle
        logger.debug(f"메인 윈도우 핸들 저장: {original_window}")

        while self._is_running:
            self.status_updated.emit("🚀 새 작업 가져오는 중... (메인 페이지 이동)")

            try:
                self.driver.get(HOME_URL)

                # --- [수정] WebDriverWait을(를) 여기서 사용 ---
                logger.info("'작업 시작' 버튼(#reviewStart)을 찾아 클릭합니다...")
                WebDriverWait(self.driver, 15).until(
                    EC.element_to_be_clickable((By.ID, "reviewStart"))
                ).click()
                self.status_updated.emit("✅ '작업 시작' 클릭. 새 창 대기 중...")

                # --- [수정] WebDriverWait을(를) 여기서 사용 ---
                WebDriverWait(self.driver, 15).until(EC.number_of_windows_to_be(2))

                all_windows = self.driver.window_handles
                new_window = next((w for w in all_windows if w != original_window), None)

                if not new_window:
                    logger.warning("⚠ 새 창을 찾지 못했습니다. 1초 대기 후 재시도.")
                    time.sleep(1)
                    continue

                self.driver.switch_to.window(new_window)
                self.status_updated.emit(f"✅ 새 작업창으로 전환. href 검사 중...")

                # --- 작업 처리 ---
                href, match, action = process_page(self.driver, self.patterns)

                if action == "E":
                    self.status_updated.emit(f"✅ '{match}' 패턴 일치. 'E' 입력 완료.")
                elif action == "미루기":
                    self.status_updated.emit(f"❌ 패턴 불일치. '작업 미루기' 완료.")
                else:
                    self.status_updated.emit(f"⚠ 알 수 없는 작업 수행. (href: {href})")

                self.total_count += 1
                self.work_finished_one.emit(self.total_count, action)

                # --- 창 닫고 복귀 ---
                self.driver.close()
                self.driver.switch_to.window(original_window)

                if not self._is_running:
                    break

                time.sleep(2)

            except Exception as e:
                if not self._is_running:
                    break

                logger.error(f"❌ 작업 루프 중 오류: {e}", exc_info=True)
                self.status_updated.emit(f"❌ 작업 루프 오류 발생. 5초 후 재시도...")

                try:
                    all_windows = self.driver.window_handles
                    for w in all_windows:
                        if w != original_window:
                            self.driver.switch_to.window(w)
                            self.driver.close()
                    self.driver.switch_to.window(original_window)
                except Exception:
                    logger.error("오류 복구 실패. 스레드 종료.")
                    self.automation_finished.emit("❌ 오류 복구 실패. 작업 중단.")
                    self._is_running = False

                time.sleep(5)

        self.automation_finished.emit("✅ 작업이 안전하게 중지되었습니다.")


# --- [PySide6 UI 메인 윈도우 (RuntimeError 수정된 버전)] ---

class MainWindow:
    def __init__(self):
        loader = QUiLoader()
        self.ui = loader.load("main_window.ui", None)
        if not self.ui:
            logger.error("FATAL: main_window.ui 파일을 로드할 수 없습니다.")
            logger.error("ui_main.py와 같은 폴더에 main_window.ui 파일이 있는지 확인하세요.")
            return

        self.worker = None

        # UI 위젯에 함수 연결
        self.ui.btn_Start.clicked.connect(self.start_automation)
        self.ui.btn_Stop.clicked.connect(self.stop_automation)

        self.ui.btn_Stop.setEnabled(False)

    @Slot()
    def start_automation(self):
        user_id = self.ui.lineEdit_ID.text().strip()
        user_pw = self.ui.lineEdit_PW.text().strip()
        headless = self.ui.checkBox_Headless.isChecked()

        if not user_id or not user_pw:
            QMessageBox.warning(self.ui, "입력 오류", "아이디와 비밀번호를 모두 입력해야 합니다.")
            return

        self.ui.btn_Start.setEnabled(False)
        self.ui.btn_Stop.setEnabled(True)
        self.ui.groupBox_Login.setEnabled(False)
        self.ui.label_StartTime.setText(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.ui.label_TotalCount.setText("0")
        self.ui.textBrowser_Status.clear()
        self.append_status("작업 스레드 초기화 중...")

        self.worker = Worker(user_id, user_pw, headless)

        self.worker.status_updated.connect(self.append_status)
        self.worker.work_finished_one.connect(self.update_count)
        self.worker.automation_finished.connect(self.on_automation_finished)
        self.worker.login_result.connect(self.on_login_result)

        self.worker.start()

    @Slot()
    def stop_automation(self):
        if self.worker:
            self.worker.stop()

        self.ui.btn_Stop.setEnabled(False)
        self.append_status("...작업 중지를 요청했습니다. 현재 작업 완료 대기 중...")

    @Slot(str)
    def append_status(self, message):
        logger.info(f"[UI] {message}")
        current_time = datetime.now().strftime("%H:%M:%S")
        self.ui.textBrowser_Status.append(f"[{current_time}] {message}")

    @Slot(int, str)
    def update_count(self, total_count, action_taken):
        self.ui.label_TotalCount.setText(str(total_count))
        self.ui.statusbar.showMessage(f"마지막 작업: {action_taken} (총 {total_count}건)", 3000)

    @Slot(bool, str)
    def on_login_result(self, success, message):
        self.append_status(message)
        if not success:
            self.ui.btn_Start.setEnabled(True)
            self.ui.btn_Stop.setEnabled(False)
            self.ui.groupBox_Login.setEnabled(True)
            QMessageBox.critical(self.ui, "로그인 실패", message)

    @Slot(str)
    def on_automation_finished(self, message):
        self.append_status(message)
        self.ui.btn_Start.setEnabled(True)
        self.ui.btn_Stop.setEnabled(False)
        self.ui.groupBox_Login.setEnabled(True)

        if "오류" in message:
            QMessageBox.critical(self.ui, "작업 오류", message)
        else:
            self.ui.statusbar.showMessage("작업 완료.", 5000)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    main_window = MainWindow()

    if main_window.ui:
        main_window.ui.show()
        sys.exit(app.exec())
    else:
        sys.exit(-1)