import ttkbootstrap as ttkb  # ttkbootstrap 임포트
from ttkbootstrap.constants import *  # PLACE, LEFT, etc. 상수
from tkinter import messagebox
import cv2
import win32com.client  # WMI (pywin32)
import ctypes
from ctypes import wintypes
import psutil
import subprocess
import os
import tempfile
import qrcode
from PIL import ImageTk
import re
import logging

# ---------------------------------------------
# 로깅 기본 설정 (디버그 레벨)
# ---------------------------------------------
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")

# ---------------------------------------------
# LRESULT, LONG_PTR를 플랫폼(32/64bit)에 맞게 정의
# ---------------------------------------------
if ctypes.sizeof(ctypes.c_void_p) == 8:
    LRESULT = ctypes.c_longlong
    LONG_PTR = ctypes.c_longlong
else:
    LRESULT = ctypes.c_long
    LONG_PTR = ctypes.c_long

# ---------------------------------------------
# Windows API 상수 및 함수 서명 지정
# ---------------------------------------------
WM_NCDESTROY = 0x0082
WM_INPUT = 0x00FF
RID_INPUT = 0x10000003
GWL_WNDPROC = -4
RIDI_DEVICENAME = 0x20000007
RIM_TYPEKEYBOARD = 1
RIDEV_INPUTSINK = 0x00000100
RIDEV_NOLEGACY = 0x00000030  # legacy 메시지 차단

RI_KEY_BREAK = 0x01
RI_KEY_E0 = 0x02

WM_DEVICECHANGE = 0x0219
DBT_DEVICEARRIVAL = 0x8000
DBT_DEVICEREMOVECOMPLETE = 0x8004
DBT_DEVTYP_DEVICEINTERFACE = 0x00000005

# user32 라이브러리 로드
user32 = ctypes.windll.user32

# SetWindowLongPtrW, CallWindowProcW 함수 원형 지정
user32.SetWindowLongPtrW.restype = LONG_PTR
user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, wintypes.INT, LONG_PTR]

user32.CallWindowProcW.restype = LRESULT
user32.CallWindowProcW.argtypes = [
    LONG_PTR,
    wintypes.HWND,
    wintypes.UINT,
    wintypes.WPARAM,
    wintypes.LPARAM
]

# ---------------------------------------------
# RAWINPUTDEVICE 구조체 정의
# ---------------------------------------------
class RAWINPUTDEVICE(ctypes.Structure):
    _fields_ = [
        ("usUsagePage", ctypes.c_ushort),
        ("usUsage", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("hwndTarget", ctypes.c_void_p)
    ]

# ---------------------------------------------
# RAWINPUTHEADER 구조체 정의
# ---------------------------------------------
class RAWINPUTHEADER(ctypes.Structure):
    _fields_ = [
        ("dwType", ctypes.c_uint),
        ("dwSize", ctypes.c_uint),
        ("hDevice", ctypes.c_void_p),
        ("wParam", ctypes.c_ulong)
    ]

# ---------------------------------------------
# RAWKEYBOARD 구조체 정의
# ---------------------------------------------
class RAWKEYBOARD(ctypes.Structure):
    _fields_ = [
        ("MakeCode", ctypes.c_ushort),
        ("Flags", ctypes.c_ushort),
        ("Reserved", ctypes.c_ushort),
        ("VKey", ctypes.c_ushort),
        ("Message", ctypes.c_uint),
        ("ExtraInformation", ctypes.c_ulong)
    ]

# ---------------------------------------------
# RAWINPUT 구조체 정의
# ---------------------------------------------
class RAWINPUT(ctypes.Structure):
    class _u(ctypes.Union):
        _fields_ = [("keyboard", RAWKEYBOARD)]
    _anonymous_ = ("u",)
    _fields_ = [
        ("header", RAWINPUTHEADER),
        ("u", _u)
    ]

# ---------------------------------------------
# 장치 이름 얻는 함수
# ---------------------------------------------
def get_device_name(hDevice):
    """주어진 hDevice 핸들을 통해 장치 이름을 얻어옵니다."""
    size = ctypes.c_uint(0)
    if user32.GetRawInputDeviceInfoW(hDevice, RIDI_DEVICENAME, None, ctypes.byref(size)) == 0:
        buffer = ctypes.create_unicode_buffer(size.value)
        if user32.GetRawInputDeviceInfoW(hDevice, RIDI_DEVICENAME, buffer, ctypes.byref(size)) > 0:
            return buffer.value
    return None

# ---------------------------------------------
# Raw Input 등록 함수
# ---------------------------------------------
def register_raw_input(hwnd):
    """지정된 윈도우 핸들에 대해 Raw Input을 등록합니다."""
    rid = RAWINPUTDEVICE()
    rid.usUsagePage = 0x01   # Generic Desktop Controls
    rid.usUsage = 0x06       # Keyboard
    rid.dwFlags = RIDEV_INPUTSINK | RIDEV_NOLEGACY
    rid.hwndTarget = hwnd
    if not user32.RegisterRawInputDevices(ctypes.byref(rid), 1, ctypes.sizeof(rid)):
        raise ctypes.WinError()

# ---------------------------------------------
# 가상 키 코드 -> 문자열 매핑
# ---------------------------------------------
VK_MAPPING = {
    0x30: "0",   0x31: "1",   0x32: "2",   0x33: "3",   0x34: "4",
    0x35: "5",   0x36: "6",   0x37: "7",   0x38: "8",   0x39: "9",
    0x41: "A",   0x42: "B",   0x43: "C",   0x44: "D",   0x45: "E",
    0x46: "F",   0x47: "G",   0x48: "H",   0x49: "I",   0x4A: "J",
    0x4B: "K",   0x4C: "L",   0x4D: "M",   0x4E: "N",   0x4F: "O",
    0x50: "P",   0x51: "Q",   0x52: "R",   0x53: "S",   0x54: "T",
    0x55: "U",   0x56: "V",   0x57: "W",   0x58: "X",   0x59: "Y",
    0x5A: "Z",
    0x20: "SPACE",
    0x0D: "ENTER",
    0x1B: "ESC",
    0x09: "TAB",
    0x08: "BACK",
    0x70: "F1",  0x71: "F2",  0x72: "F3",  0x73: "F4",
    0x74: "F5",  0x75: "F6",  0x76: "F7",  0x77: "F8",
    0x78: "F9",  0x79: "F10", 0x7A: "F11", 0x7B: "F12",
    0x2D: "INS",
    0x2E: "DEL",
    0x25: "LEFT",
    0x26: "UP",
    0x27: "RIGHT",
    0x28: "DOWN",
    0x14: "CAPS",
    0x90: "NUMLOCK",
    0x60: "NUM0",
    0x61: "NUM1",
    0x62: "NUM2",
    0x63: "NUM3",
    0x64: "NUM4",
    0x65: "NUM5",
    0x66: "NUM6",
    0x67: "NUM7",
    0x68: "NUM8",
    0x69: "NUM9",
    0x6A: "NUM *",
    0x6B: "NUM +",
    0x6C: "NUMENTER",
    0x6D: "NUM -",
    0x6E: "NUM .",
    0x6F: "NUM /",
    0x2C: "PRT",
    0xBB: "=",
    0xBD: "-",
    0xC0: "`",
    0xDB: "[",
    0xDD: "]",
    0xDC: "\\",
    0xBA: ";",
    0xDE: "'",
    0xBC: ",",
    0xBE: ".",
    0xBF: "/",
    0xA0: "LSHIFT",
    0xA1: "RSHIFT",
    0x11: "CTRL",
    0x5B: "WINDOW",
    0x12: "ALT",
    0x15: "한/영",
    0x19: "한자",
}

# WNDPROC 타입 선언 (윈도우 프로시저 콜백)
WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)


# =============================================
# LaptopTestApp 클래스 (ttkbootstrap.Window 기반)
# =============================================
class LaptopTestApp(ttkb.Window):
    def __init__(self, themename="flatly"):
        """
        노트북 기능 테스트 프로그램 초기화 및 UI 구성
        ttkbootstrap.Window를 상속받아 테마를 적용합니다.
        """
        super().__init__(themename=themename)  # 부모 클래스 초기화 (테마 적용)
        self.title("노트북 테스트 프로그램 (ttkbootstrap 적용)")
        self.geometry("700x750")  # 기본 크기 약간 확대

        # 내부 키보드의 Raw Input device 문자열 (화이트리스트)
        self.INTERNAL_HWIDS = [
            "\\ACPI#MSF0001"
        ]

        # 테스트 완료 여부 저장 딕셔너리
        self.test_done = {
            "키보드": False,
            "카메라": False,
            "USB": False,
            "충전기": False,
        }

        # 배터리 리포트 파일 경로 (초기 None)
        self.report_path = None

        self.failed_keys = []  # 누르지 못한 키 목록
        self.disabled_hwids = []

        # 메인 프레임 생성
        main_frame = ttkb.Frame(self, padding=20)
        main_frame.pack(fill=BOTH, expand=True)

        # 타이틀 라벨
        title_label = ttkb.Label(main_frame, text="노트북 기능 테스트", font=("맑은 고딕", 18, "bold"))
        title_label.pack(pady=10)

        self.progress_label = ttkb.Label(main_frame, text=self.get_progress_text(), font=("맑은 고딕", 12))
        self.progress_label.pack(pady=5)

        self.test_status_labels = {}

        # ----------------- 키보드 테스트 -----------------
        kb_frame = ttkb.Frame(main_frame)
        kb_frame.pack(fill=X, pady=5)

        kb_button = ttkb.Button(kb_frame, text="키보드 테스트", bootstyle=SUCCESS, command=self.open_keyboard_test)
        kb_button.pack(side=LEFT)

        kb_status = ttkb.Label(kb_frame, text="테스트 전", bootstyle="danger")  # 빨간색 계열
        kb_status.pack(side=LEFT, padx=10)
        self.test_status_labels["키보드"] = kb_status

        self.failed_keys_button = ttkb.Button(kb_frame, text="누르지 못한 키 보기",
                                              state="disabled",
                                              bootstyle=WARNING,
                                              command=self.show_failed_keys)
        self.failed_keys_button.pack(side=LEFT, padx=5)

        # ----------------- USB 테스트 -----------------
        self.usb_ports = {
            "port1": False,
            "port2": False,
            "port3": False,
        }
        self.usb_test_complete = False

        usb_frame = ttkb.Frame(main_frame)
        usb_frame.pack(fill=X, pady=5)

        self.usb_button = ttkb.Button(usb_frame, text="USB 연결 확인", bootstyle=SUCCESS, command=self.start_usb_check)
        self.usb_button.pack(side=LEFT)

        self.usb_refresh_button = ttkb.Button(usb_frame, text="새로고침",
                                              bootstyle=SECONDARY,
                                              command=self.refresh_usb_check,
                                              state="disabled")
        self.usb_refresh_button.pack(side=LEFT, padx=5)

        usb_status = ttkb.Label(usb_frame, text="테스트 전", bootstyle="danger")
        usb_status.pack(side=LEFT, padx=5)
        self.test_status_labels["USB"] = usb_status

        # self.usb_port_labels = {}
        # self.create_usb_port_labels(main_frame)

        # ----------------- 카메라 테스트 -----------------
        cam_frame = ttkb.Frame(main_frame)
        cam_frame.pack(fill=X, pady=5)

        cam_button = ttkb.Button(cam_frame, text="카메라(웹캠) 테스트", bootstyle=SUCCESS, command=self.open_camera_test)
        cam_button.pack(side=LEFT)

        cam_status = ttkb.Label(cam_frame, text="테스트 전", bootstyle="danger")
        cam_status.pack(side=LEFT, padx=10)
        self.test_status_labels["카메라"] = cam_status

        # ----------------- 충전 테스트 -----------------
        self.c_type_ports = {"충전기": False}
        self.c_type_test_complete = False

        c_type_frame = ttkb.Frame(main_frame)
        c_type_frame.pack(fill=X, pady=5)

        self.c_type_button = ttkb.Button(c_type_frame, text="충전 테스트 시작", bootstyle=SUCCESS,
                                         command=self.start_c_type_check)
        self.c_type_button.pack(side=LEFT)

        c_type_status = ttkb.Label(c_type_frame, text="테스트 전", bootstyle="danger")
        c_type_status.pack(side=LEFT, padx=10)
        self.test_status_labels["충전기"] = c_type_status

        self.c_type_port_labels = {}
        self.create_c_type_port_labels(main_frame)

        # ----------------- 배터리 리포트 및 QR 코드 -----------------
        battery_report_frame = ttkb.Frame(main_frame)
        battery_report_frame.pack(fill=X, pady=5)

        self.battery_report_button = ttkb.Button(
            battery_report_frame,
            text="배터리 리포트 생성",
            bootstyle=PRIMARY,
            command=self.generate_battery_report
        )
        self.battery_report_button.pack(side=LEFT)

        self.view_report_button = ttkb.Button(
            battery_report_frame,
            text="리포트 확인",
            bootstyle=INFO,
            command=self.view_battery_report
        )
        self.view_report_button.pack(side=LEFT, padx=5)

        self.qr_code_button = ttkb.Button(
            battery_report_frame,
            text="QR 코드 생성",
            bootstyle=PRIMARY,
            command=self.generate_qr_code
        )
        self.qr_code_button.pack(side=LEFT, padx=5)


    # ----------------- 진행 상황 관련 메서드 -----------------
    def get_progress_text(self):
        """전체 테스트 진행 상황을 텍스트로 반환합니다."""
        done_count = sum(self.test_done.values())
        total = len(self.test_done)
        return f"{done_count}/{total} 완료"

    def mark_test_complete(self, test_name):
        """특정 테스트 완료 후 상태 및 UI 갱신"""
        if test_name in self.test_done and not self.test_done[test_name]:
            self.test_done[test_name] = True
            self.progress_label.config(text=self.get_progress_text())
            status_label = self.test_status_labels[test_name]
            # 테스트 완료 시 파란색 계열("info")로 표시
            status_label.config(text="테스트 완료", bootstyle="info")
            if all(self.test_done.values()):
                messagebox.showinfo("모든 테스트 완료", "모든 테스트를 완료했습니다.\n수고하셨습니다!")

    # ----------------- 키보드 테스트 -----------------
    def open_keyboard_test(self):
        """키보드 테스트 창을 열고 Raw Input을 등록합니다."""
        kb_window = ttkb.Toplevel(self)  # ttkb.Toplevel 사용
        kb_window.title("키보드 테스트")
        kb_window.geometry("1200x500")

        info_label = ttkb.Label(kb_window, text="이 창에 포커스를 두고\n모든 키를 한 번씩 눌러보세요.\n완료 시 창이 닫힙니다.")
        info_label.pack(pady=5)

        # 키보드 레이아웃 구성
        keyboard_layout = [
            ["F1", "F2", "F3", "F4", "F5", "F6",
             "F7", "F8", "F9", "F10", "F11", "F12"],
            ["`", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "-", "=", "BACK",
             "INS", "DEL"],
            ["Q", "W", "E", "R", "T", "Y", "U", "I", "O", "P", "[", "]", "\\"],
            ["A", "S", "D", "F", "G", "H", "J", "K", "L", ";", "'", "ENTER"],
            ["Z", "X", "C", "V", "B", "N", "M", ",", ".", "/", "SPACE"],
            ["UP", "LEFT", "DOWN", "RIGHT"],
            ["CAPS", "NUMLOCK"],
            ["NUM7", "NUM8", "NUM9", "NUM /",
             "NUM4", "NUM5", "NUM6", "NUM *",
             "NUM1", "NUM2", "NUM3", "NUM -",
             "NUM0", "NUM .", "NUM +", "NUMENTER"],
            ["ESC", "TAB", "LSHIFT", "RSHIFT", "CTRL", "WINDOW", "ALT", "PRT", "한/영", "한자"]
        ]

        self.all_keys = set()
        self.key_widgets = {}
        for row_keys in keyboard_layout:
            row_frame = ttkb.Frame(kb_window)
            row_frame.pack(pady=5)
            for key in row_keys:
                key_upper = key.upper()
                self.all_keys.add(key_upper)
                btn = ttkb.Label(
                    row_frame, text=key, width=5, bootstyle="inverse-light",
                    font=("맑은 고딕", 10, "bold")
                )
                btn.pack(side=LEFT, padx=3)
                self.key_widgets[key_upper] = btn

        self.keys_not_pressed = set(self.all_keys)

        # Raw Input 등록
        hwnd = kb_window.winfo_id()
        register_raw_input(hwnd)

        # Raw Input 윈도우 프로시저 정의
        def raw_input_wnd_proc(hWnd, msg, wParam, lParam):
            if msg == WM_NCDESTROY:
                if self._kb_old_wnd_proc is not None:
                    user32.SetWindowLongPtrW(hWnd, GWL_WNDPROC, self._kb_old_wnd_proc)
                    self._kb_old_wnd_proc = None
                return 0

            if msg == WM_INPUT:
                logging.debug("raw_input_wnd_proc: WM_INPUT 메시지 처리 시작")
                size = ctypes.c_uint(0)
                if user32.GetRawInputData(lParam, RID_INPUT, None, ctypes.byref(size),
                                          ctypes.sizeof(RAWINPUTHEADER)) == 0:
                    buffer = ctypes.create_string_buffer(size.value)
                    if user32.GetRawInputData(lParam, RID_INPUT, buffer, ctypes.byref(size),
                                              ctypes.sizeof(RAWINPUTHEADER)) == size.value:
                        raw = ctypes.cast(buffer, ctypes.POINTER(RAWINPUT)).contents
                        if raw.header.dwType == RIM_TYPEKEYBOARD:
                            if (raw.u.keyboard.Flags & RI_KEY_BREAK) == 0:  # Key Down 이벤트
                                vkey = raw.u.keyboard.VKey
                                logging.debug(f"raw_input_wnd_proc: 키 입력 감지, vkey={vkey}")
                                # 키 심볼 결정
                                if vkey == 0x0D:
                                    key_sym = "NUMENTER" if (raw.u.keyboard.Flags & RI_KEY_E0) else "ENTER"
                                elif vkey == 0x10:
                                    if raw.u.keyboard.MakeCode == 0x2A:
                                        key_sym = "LSHIFT"
                                    elif raw.u.keyboard.MakeCode == 0x36:
                                        key_sym = "RSHIFT"
                                    else:
                                        key_sym = "SHIFT"
                                elif vkey == 0x2D:
                                    key_sym = "INS" if (raw.u.keyboard.Flags & RI_KEY_E0) else "NUMINS"
                                elif vkey == 0x2E:
                                    key_sym = "DEL" if (raw.u.keyboard.Flags & RI_KEY_E0) else "NUMDEL"
                                elif vkey == 0x26:
                                    key_sym = "UP" if (raw.u.keyboard.Flags & RI_KEY_E0) else "NUMUP"
                                elif vkey == 0x25:
                                    key_sym = "LEFT" if (raw.u.keyboard.Flags & RI_KEY_E0) else "NUMLEFT"
                                elif vkey == 0x28:
                                    key_sym = "DOWN" if (raw.u.keyboard.Flags & RI_KEY_E0) else "NUMDOWN"
                                elif vkey == 0x27:
                                    key_sym = "RIGHT" if (raw.u.keyboard.Flags & RI_KEY_E0) else "NUMRIGHT"
                                elif vkey in VK_MAPPING:
                                    key_sym = VK_MAPPING[vkey]
                                else:
                                    key_sym = None

                                if key_sym:
                                    device_name = get_device_name(raw.header.hDevice)
                                    if device_name:
                                        device_name_lower = device_name.lower().replace("\\", "#")
                                        is_internal = any(
                                            internal_id.lower().replace("\\", "#") in device_name_lower
                                            for internal_id in self.INTERNAL_HWIDS
                                        )
                                    else:
                                        is_internal = False
                                        logging.debug(f"키: {key_sym} is_internal: {is_internal}")

                                    if is_internal:
                                        self.on_raw_key(key_sym)
                return 0

            if not user32.IsWindow(hWnd):
                return 0

            if self._kb_old_wnd_proc:
                return user32.CallWindowProcW(self._kb_old_wnd_proc, hWnd, msg, wParam, lParam)
            else:
                return user32.DefWindowProcW(hWnd, msg, wParam, lParam)

        def on_close_keyboard_window():
            """키보드 창 종료 시 누르지 않은 키가 있으면 기록합니다."""
            if self.keys_not_pressed:
                self.failed_keys = list(self.keys_not_pressed)
                self.test_status_labels["키보드"].config(text="오류 발생", bootstyle="danger")
                self.failed_keys_button.config(state="normal")
            self.close_keyboard_window()

        kb_window.protocol("WM_DELETE_WINDOW", on_close_keyboard_window)
        self._raw_input_wnd_proc = WNDPROC(raw_input_wnd_proc)
        cb_func_ptr = ctypes.cast(self._raw_input_wnd_proc, ctypes.c_void_p).value
        cb_func_ptr = LONG_PTR(cb_func_ptr)
        old_proc = user32.SetWindowLongPtrW(hwnd, GWL_WNDPROC, cb_func_ptr)
        self._kb_old_wnd_proc = old_proc
        self._kb_hwnd = hwnd
        self.kb_window_ref = kb_window

    def close_keyboard_window(self):
        """키보드 테스트 창 종료 시 Raw Input 프로시저 복원 및 창 닫기"""
        if hasattr(self, '_kb_hwnd') and self._kb_hwnd and self._kb_old_wnd_proc is not None:
            user32.SetWindowLongPtrW(self._kb_hwnd, GWL_WNDPROC, self._kb_old_wnd_proc)
            self._kb_old_wnd_proc = None
        if hasattr(self, 'kb_window_ref'):
            self.kb_window_ref.destroy()

    def on_raw_key(self, key):
        """키 입력 시 해당 키를 표시하고, 모든 키 입력이 완료되면 테스트 완료 처리"""
        if key in self.keys_not_pressed:
            self.keys_not_pressed.remove(key)
            widget = self.key_widgets.get(key)
            if widget:
                widget.config(bootstyle="inverse-dark")  # 키를 누른 후 색을 바꿔줌
            if not self.keys_not_pressed:
                self.failed_keys_button.config(state="disabled")
                self.close_keyboard_window()
                self.mark_test_complete("키보드")

    def show_failed_keys(self):
        """누르지 못한 키 목록을 별도 창에 표시"""
        if self.failed_keys:
            failed_win = ttkb.Toplevel(self)
            failed_win.title("미처 누르지 못한 키 목록")
            failed_win.geometry("300x200")

            info_label = ttkb.Label(failed_win, text="누르지 못한 키:")
            info_label.pack(padx=10, pady=10)

            failed_keys_str = ", ".join(sorted(self.failed_keys))
            keys_label = ttkb.Label(failed_win, text=failed_keys_str, font=("맑은 고딕", 12))
            keys_label.pack(padx=10, pady=10)
        else:
            messagebox.showinfo("확인", "누르지 못한 키가 없습니다.")

    # ----------------- USB 테스트 관련 메서드 -----------------
    # def create_usb_port_labels(self, main_frame):
    #     """USB 포트 상태를 표시할 라벨 생성"""
    #     usb_port_frame = ttkb.Frame(main_frame)
    #     usb_port_frame.pack(fill=X, pady=3)
    #     for port_name in self.usb_ports:
    #         label = ttkb.Label(usb_port_frame, text=f"{port_name}: 연결 안됨",
    #                            width=16, bootstyle="inverse-light")
    #         label.pack(side=LEFT, padx=5)
    #         self.usb_port_labels[port_name] = label

    def start_usb_check(self):
        """USB 테스트 초기화 및 시작"""
        self.usb_ports = {"port1": False, "port2": False, "port3": False}
        self.usb_test_complete = False
        # for port_name, label in self.usb_port_labels.items():
        #     label.config(text=f"{port_name}: 연결 안됨", bootstyle="inverse-light")
        self.usb_button.config(state="disabled")
        self.usb_refresh_button.config(state="normal")
        self.test_status_labels["USB"].config(text="테스트 중", bootstyle="warning")
        self.refresh_usb_check()

    def refresh_usb_check(self):
        """USB 연결 상태를 확인하여 UI 갱신 및 테스트 완료 처리"""
        try:
            wmi_obj = win32com.client.GetObject("winmgmts:")
            pnp_entities = wmi_obj.InstancesOf("Win32_PnPEntity")
            for entity in pnp_entities:
                if hasattr(entity, 'PNPDeviceID') and entity.PNPDeviceID:
                    device_path = entity.PNPDeviceID.upper()
                    if not device_path.startswith("USB\\"):
                        continue
                    match = re.search(r'&0&(\d)$', device_path)
                    if match:
                        port_number = match.group(1)
                        if port_number in ['1', '2', '3']:
                            key = f"port{port_number}"
                            self.usb_ports[key] = True
                            # self.usb_port_labels[key].config(text=f"{key}: 연결됨", bootstyle="inverse-success")
            if all(self.usb_ports.values()):
                self.usb_test_complete = True
                self.usb_refresh_button.config(state="disabled")
                self.mark_test_complete("USB")
                messagebox.showinfo("USB Test", "모든 USB 포트 테스트 완료!")
        except Exception as e:
            messagebox.showerror("USB Error", f"USB 포트 확인 중 오류 발생:\n{e}")

    # ----------------- 카메라 테스트 -----------------
    def open_camera_test(self):
        """카메라(웹캠) 테스트 창을 열고 프레임을 표시합니다."""
        cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
        if not cap.isOpened():
            messagebox.showerror("카메라 오류", "카메라를 열 수 없습니다. 장치를 확인해주세요.")
            self.test_status_labels["카메라"].config(text="오류 발생", bootstyle="danger")
            return

        window_name = "Camera Test - X to exit"
        try:
            while True:
                ret, frame = cap.read()
                if not ret:
                    messagebox.showerror("카메라 오류", "카메라 프레임을 읽을 수 없습니다.")
                    break
                cv2.imshow(window_name, frame)
                key = cv2.waitKey(1) & 0xFF
                if key == 27:  # ESC 키
                    break
                if cv2.getWindowProperty(window_name, cv2.WND_PROP_VISIBLE) < 1:
                    break
        finally:
            cap.release()
            cv2.destroyAllWindows()
            self.mark_test_complete("카메라")

    # ----------------- 충전 테스트 -----------------
    def create_c_type_port_labels(self, main_frame):
        """충전 포트 상태를 표시할 라벨 생성"""
        c_type_port_frame = ttkb.Frame(main_frame)
        c_type_port_frame.pack(fill=X, pady=3)
        for port_name in self.c_type_ports:
            label = ttkb.Label(c_type_port_frame, text=f"{port_name}: 연결 안됨",
                               width=20, bootstyle="inverse-light")
            label.pack(side=LEFT, padx=5)
            self.c_type_port_labels[port_name] = label

    def start_c_type_check(self):
        """충전 테스트를 시작하고 포트 상태를 확인합니다."""
        self.c_type_ports = {"충전기": False}
        for port, lbl in self.c_type_port_labels.items():
            lbl.config(text=f"{port}: 연결 안됨", bootstyle="inverse-light")
        self.c_type_test_complete = False
        self.test_status_labels["충전기"].config(text="테스트 중", bootstyle="warning")
        self.check_c_type_port()

    def check_c_type_port(self):
        """현재 배터리 충전 상태를 확인하여 포트 상태를 갱신합니다."""
        battery = psutil.sensors_battery()
        if battery is None:
            messagebox.showerror("충전기 Error", "배터리 정보를 가져올 수 없습니다.")
            return

        if not battery.power_plugged:
            messagebox.showinfo(
                "충전기 Test",
                "충전기가 연결되지 않았습니다.\n해당 포트에 충전기를 연결 후 다시 확인하세요."
            )
            return

        if not self.c_type_ports["충전기"]:
            self.c_type_ports["충전기"] = True
            self.c_type_port_labels["충전기"].config(
                text="전원 연결됨 (충전 중)",
                bootstyle="inverse-success"
            )
        else:
            messagebox.showinfo("충전 Test", "충전 확인되었습니다.")

        if all(self.c_type_ports.values()):
            self.c_type_test_complete = True
            self.test_status_labels["충전기"].config(text="테스트 완료", bootstyle="info")
            self.mark_test_complete("충전기")

    # ----------------- 배터리 리포트 생성 -----------------
    def generate_battery_report(self):
        """powercfg 명령어를 이용하여 배터리 리포트를 생성합니다."""
        try:
            temp_dir = tempfile.mkdtemp()
            self.report_path = os.path.join(temp_dir, "battery_report.html")
            subprocess.run(
                ["powercfg", "/batteryreport", "/output", self.report_path],
                check=True,
                capture_output=True,
                text=True
            )
            messagebox.showinfo("배터리 리포트",
                                f"배터리 리포트가 생성되었습니다.\n파일 경로:\n{self.report_path}")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("배터리 리포트 오류",
                                 f"명령 실행 중 오류 발생:\n{e.stderr}")
        except Exception as e:
            messagebox.showerror("배터리 리포트 오류",
                                 f"오류 발생:\n{e}")

    def view_battery_report(self):
        """생성된 배터리 리포트 파일을 엽니다."""
        if self.report_path and os.path.exists(self.report_path):
            try:
                os.startfile(self.report_path)
            except Exception as e:
                messagebox.showerror("리포트 확인 오류", f"리포트를 열 수 없습니다:\n{e}")
        else:
            messagebox.showwarning("리포트 없음", "아직 배터리 리포트가 생성되지 않았습니다.\n먼저 '배터리 리포트 생성' 버튼을 눌러주세요.")

    def generate_qr_code(self):
        """테스트 결과를 JSON 형식으로 구성 후 QR 코드를 생성하여 표시합니다."""
        import json
        results = {
            "keyboard": {
                "status": "pass" if self.test_done.get("키보드") else "fail",
                "failed_keys": sorted(self.failed_keys) if not self.test_done.get("키보드") else []
            },
            "usb": {
                "status": "pass" if self.test_done.get("USB") else "fail",
                "failed_ports": [port for port, connected in self.usb_ports.items() if not connected]
            },
            "camera": {
                "status": "pass" if self.test_done.get("카메라") else "fail"
            },
            "charger": {
                "status": "pass" if self.test_done.get("충전기") else "fail"
            },
            "battery_report": "생성됨" if self.report_path and os.path.exists(self.report_path) else "생성되지 않음"
        }
        qr_data = json.dumps(results, ensure_ascii=False, indent=2)
        try:
            qr = qrcode.QRCode(
                version=None,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=4,
                border=4,
            )
            qr.add_data(qr_data)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            qr_img = ImageTk.PhotoImage(img)
            qr_window = ttkb.Toplevel(self)
            qr_window.title("상세 테스트 결과 QR 코드")
            qr_label = ttkb.Label(qr_window, image=qr_img)
            qr_label.image = qr_img  # 이미지 참조 유지
            qr_label.pack(padx=10, pady=10)
        except Exception as e:
            messagebox.showerror("QR 코드 생성 오류", f"QR 코드 생성 중 오류 발생:\n{e}")


# =============================================
# 메인 실행부
# =============================================
if __name__ == "__main__":
    # 원하는 테마 이름: "flatly", "litera", "journal", "cosmo", "cyborg" 등
    app = LaptopTestApp(themename="flatly")
    app.mainloop()
