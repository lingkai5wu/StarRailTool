import atexit
import logging
import time
from typing import Optional, Tuple

import psutil
import pycaw.pycaw as audio
import win32gui
import win32process


def get_process_info(name: str) -> Tuple[Optional[int], Optional[int]]:
    for proc in psutil.process_iter():
        if proc.name() == name:
            pid = proc.pid
            hwnd = get_window_handle_from_pid(pid)
            return pid, hwnd
    return None, None


def get_window_handle_from_pid(pid: int) -> Optional[int]:
    def callback(hwnd, cur_hwnd_list):
        _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
        if found_pid == pid:
            cur_hwnd_list.append(hwnd)
        return True

    hwnd_list = []
    win32gui.EnumWindows(callback, hwnd_list)

    if len(hwnd_list) > 0:
        return hwnd_list[0]
    else:
        return None


class AudioManager:
    last_volume = 1

    def __init__(self, pid: int):
        self.pid = pid
        self.volume_control: Optional[audio.ISimpleAudioVolume] = None
        self._init_volume_control()

    def _init_volume_control(self):
        sessions = audio.AudioUtilities.GetAllSessions()
        for session in sessions:
            if session.Process and session.Process.pid == self.pid:
                self.volume_control = session.SimpleAudioVolume
                break

    def set_volume(self, volume: float):
        if self.last_volume != volume:
            self.volume_control.SetMasterVolume(volume, None)
            self.last_volume = volume


def main_loop(pid: int, hwnd: int):
    audio_manager = AudioManager(pid)

    def exit_handler():
        logging.info("Exiting program...")
        audio_manager.set_volume(1)

    atexit.register(exit_handler)

    while True:
        if not psutil.pid_exists(sr_pid):
            exit(0)
        current_hwnd = win32gui.GetForegroundWindow()
        if current_hwnd == hwnd:
            audio_manager.set_volume(1)
        else:
            audio_manager.set_volume(0)
        time.sleep(LOOP_INTERVAL)


TARGET_PROCESS_NAME = "StarRail.exe"
LOOP_INTERVAL = 0.1  # 循环时间间隔，单位：秒

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    sr_pid, sr_hwnd = get_process_info(TARGET_PROCESS_NAME)
    if sr_pid is None:
        logging.error("Process %s not found", TARGET_PROCESS_NAME)
        exit(1)

    logging.info("Process %s found, PID: %s, HWND: %s", TARGET_PROCESS_NAME, sr_pid, sr_hwnd)

    main_loop(sr_pid, sr_hwnd)
