import knw_license
import sys
import logging
import os
import time
from datetime import datetime
from PySide6.QtWidgets import QApplication, QMessageBox, QMainWindow
from PySide6.QtCore import QThread, Signal, Slot
from PySide6.QtUiTools import QUiLoader

# Selenium 임포트
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

# 로컬 모듈 임포트
from label_admin import label_login, close_chrome, HOME_URL
from prefix_util import load_patterns_from_gsheet, process_page
# [수정] resource_path만 임포트 (logger는 setup_logger가 반환)
from label_log import setup_logger, resource_path

# --- [로거 설정] ---
logger, LOG_FILENAME = setup_logger()
logger.info(f"UI 모드 자동화 작업 시작. 로그 파일: {LOG_FILENAME}")


# --- [로거 설정 끝] ---


# --- [백그라운드 Selenium 작업을 위한 QThread] ---
# (Worker 클래스 내부는 수정할 필요 없음)
class Worker(QThread):
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
        self._is_running = True

    def run(self):
        try:
            self.status_updated.emit("구글 시트에서 패턴 로드 중...")
            self.patterns = load_patterns_from_gsheet()
            if not self.patterns:
                self.login_result.emit(False, "❌ 구글시트에서 패턴을 불러오지 못했습니다.")
                return
            self.status_updated.emit("패턴 로드 완료. 로그인 시도 중...")

            self.driver = label_login(self.user_id, self.user_pw, self.headless)

            if not self.driver:
                self.login_result.emit(False, "❌ 로그인 실패. 아이디/비밀번호를 확인하세요.")
                return

            self.login_result.emit(True, "✅ 로그인 성공!")

            self.main_task_loop_scenario_2()

        # noinspection PyBroadException
        except Exception as e:
            logger.error(f"Worker 스레드 실행 중 오류: {e}", exc_info=True)
            self.automation_finished.emit(f"❌ 작업 중 심각한 오류 발생: {e}")
        finally:
            if self.driver:
                close_chrome(self.driver)
            logger.info("Worker 스레드 종료.")

    def stop(self):
        self.status_updated.emit("🛑 작업 중지 요청됨... 현재 작업 완료 후 종료합니다.")
        self._is_running = False

    def main_task_loop_scenario_2(self):
        work_window = None
        try:
            self.status_updated.emit("🚀 작업 페이지로 이동 중... (1회)")
            self.driver.get(HOME_URL)
            original_window = self.driver.current_window_handle

            logger.info("'작업 시작' 버튼(#reviewStart)을 찾아 클릭합니다...")
            WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable((By.ID, "reviewStart"))
            ).click()
            self.status_updated.emit("✅ '작업 시작' 클릭. 새 창 대기 중...")

            WebDriverWait(self.driver, 15).until(EC.number_of_windows_to_be(2))
            all_windows = self.driver.window_handles
            work_window = next((w for w in all_windows if w != original_window), None)

            if not work_window:
                logger.warning("⚠ 새 창을 찾지 못했습니다. 작업 중단.")
                self.automation_finished.emit("❌ 새 작업창을 열지 못했습니다.")
                return

            self.driver.switch_to.window(work_window)
            self.status_updated.emit(f"✅ 새 작업창으로 전환 완료. 이 창에서 반복 작업을 시작합니다.")
            logger.info(f"작업창으로 전환 완료 (Handle: {work_window}). 무한 루프 시작...")

            while self._is_running:
                self.status_updated.emit("👉 다음 작업 처리 중... (href 대기)")

                href, match, action = process_page(self.driver, self.patterns)

                if action == "E (패턴 일치)":
                    self.status_updated.emit(f"✅ '{match}' 패턴 일치. 'E' 입력 완료.")
                elif action == "작업 미루기":
                    self.status_updated.emit(f"❌ 패턴 불일치. '작업 미루기' 완료.")
                else:
                    self.status_updated.emit(f"⚠ {action} 수행. (href: {href})")

                self.total_count += 1
                self.work_finished_one.emit(self.total_count, action)

                if not self._is_running:
                    break

        # noinspection PyBroadException
        except Exception as e:
            if not self._is_running:
                logger.info("작업 중지 요청으로 인해 루프를 종료합니다.")
            else:
                logger.error(f"❌ 작업 루프 중 오류: {e}", exc_info=True)
                self.status_updated.emit(f"❌ 작업 루프 오류 발생. 5초 후 재시도...")

                try:
                    if work_window not in self.driver.window_handles:
                        logger.error("❌ 작업창이 닫힌 것을 감지. 스레드를 종료합니다.")
                        self.automation_finished.emit("❌ 작업창이 닫혔습니다. 작업 중단.")
                        self._is_running = False
                    else:
                        time.sleep(5)
                        # noinspection PyBroadException
                except Exception:
                    logger.error("오류 복구 중 치명적 오류. 스레드 종료.")
                    self.automation_finished.emit("❌ 드라이버 오류. 작업 중단.")
                    self._is_running = False

        self.automation_finished.emit("✅ 작업이 안전하게 중지되었습니다.")


# --- [PySide6 UI 메인 윈도우] ---

class MainWindow:
    def __init__(self):

        loader = QUiLoader()

        # --- [수정된 부분] ---
        # .ui 파일 경로를 resource_path()로 감쌉니다.
        ui_file_path = resource_path("main_window.ui")
        logger.debug(f"UI 파일 경로: {ui_file_path}")
        self.ui: QMainWindow = loader.load(ui_file_path, None)
        # --- [수정 완료] ---

        if not self.ui:
            logger.error("=" * 50)
            logger.error("FATAL: main_window.ui 파일을 로드할 수 없습니다.")
            logger.error(f"경로: {ui_file_path}")
            logger.error("=" * 50)
            return

        self.worker = None

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