import logging
import sys
import threading
import time
from typing import Optional, Tuple

import pkg_resources
import psutil
import pycaw.pycaw as audio
import pystray
import win32api
import win32con
import win32gui
import win32process
from PIL import Image


def get_window_handle_from_pid(pid: int) -> Optional[int]:
    def callback(cur_hwnd, cur_hwnd_list):
        _, found_pid = win32process.GetWindowThreadProcessId(cur_hwnd)
        if found_pid == pid:
            cur_hwnd_list.append(cur_hwnd)

    hwnd_list = []
    win32gui.EnumWindows(callback, hwnd_list)

    if len(hwnd_list) > 0:
        return hwnd_list[0]
    else:
        return None


class AudioManager:

    def __init__(self, pid: Optional[int]):
        self.pid = pid
        self.last_volume = 1
        self.is_process_running = True
        self.volume_control = None

        self._init_volume_control()

    def _init_volume_control(self):
        while self.volume_control is None:
            sessions = audio.AudioUtilities.GetAllSessions()
            for session in sessions:
                if session.Process and session.Process.pid == self.pid:
                    self.volume_control = session.SimpleAudioVolume
                    break
            logging.info("SimpleAudioVolume not found, try again...")
            time.sleep(LOOP_INTERVAL)

    def set_volume(self, volume: float):
        if self.last_volume != volume:
            self.volume_control.SetMasterVolume(volume, None)
            self.last_volume = volume

    def main_loop(self, hwnd: int):
        while True:
            if not self.is_process_running:
                exit_now(0)
            current_hwnd = win32gui.GetForegroundWindow()
            if current_hwnd == hwnd:
                self.set_volume(1)
            else:
                self.set_volume(0)
            time.sleep(LOOP_INTERVAL)


audio_manager = None


def exit_now(code):
    logging.info("Exiting program...")
    icon.stop()
    if audio_manager is not None:
        audio_manager.last_volume = 0
        audio_manager.set_volume(1)
    sys.exit(code)


def check_startup():
    current_process = psutil.Process()
    res = 0
    for process in processes:
        if process.name() == current_process.name() and process.pid != current_process.pid:
            res += 1
    if res > 1:
        sys.exit(1)


def on_quit_clicked():
    audio_manager.is_process_running = False


def get_icon():
    image = Image.open(pkg_resources.resource_filename(__name__, "static/silence.ico"))
    menu = pystray.Menu(pystray.MenuItem('退出', on_quit_clicked))
    cur_icon = pystray.Icon('星铁后台静音', image, '星铁后台静音', menu)
    return cur_icon


def get_process_info(name: str) -> Tuple[int, int]:
    pid, hwnd = None, None
    for proc in processes:
        if proc.name() == name:
            pid = proc.pid
            hwnd = get_window_handle_from_pid(pid)
    if pid is None:
        logging.error("Process %s not found", TARGET_PROCESS_NAME)
        win32api.MessageBox(0, "请先启动游戏本体", "星铁后台静音", win32con.MB_ICONWARNING)
        exit_now(1)
    while hwnd is None:
        logging.info("Window handle not found, try again...")
        hwnd = get_window_handle_from_pid(pid)
    return pid, hwnd


def check_process_running():
    while True:
        if not psutil.pid_exists(sr_pid):
            logging.info("Process %s was close", TARGET_PROCESS_NAME)
            audio_manager.is_process_running = False
            break
        time.sleep(LOOP_INTERVAL * 10)


TARGET_PROCESS_NAME = "StarRail.exe"
LOOP_INTERVAL = 0.2  # 循环时间间隔，单位：秒
if __name__ == "__main__":
    processes = list(psutil.process_iter())
    check_startup()
    logging.basicConfig(level=logging.INFO)
    icon = get_icon()
    icon.run_detached()

    sr_pid, sr_hwnd = get_process_info(TARGET_PROCESS_NAME)
    audio_manager = AudioManager(sr_pid)
    logging.info("Process %s found, PID: %s, HWND: %s", TARGET_PROCESS_NAME, sr_pid, sr_hwnd)
    threading.Thread(target=check_process_running, daemon=True).start()
    icon.notify("游戏关闭后自动退出，或从任务栏退出", "星铁后台静音 启动成功")
    audio_manager.main_loop(sr_hwnd)

# pyinstaller -Fw --add-data "static/silence.ico;static" -i "static/silence.ico" background_muter.py
