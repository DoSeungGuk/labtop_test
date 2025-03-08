import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import cv2
import win32com.client  # WMI (pywin32)
import ctypes
from ctypes import wintypes  # ctypes.wintypes 임포트
import psutil
import subprocess
import os
import tempfile
import webbrowser
import qrcode
from PIL import Image, ImageTk
import re 

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

# user32 핸들
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

# RAWINPUTDEVICE 구조체
class RAWINPUTDEVICE(ctypes.Structure):
    _fields_ = [
        ("usUsagePage", ctypes.c_ushort),
        ("usUsage", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("hwndTarget", ctypes.c_void_p)
    ]

# RAWINPUTHEADER 구조체
class RAWINPUTHEADER(ctypes.Structure):
    _fields_ = [
        ("dwType", ctypes.c_uint),
        ("dwSize", ctypes.c_uint),
        ("hDevice", ctypes.c_void_p),
        ("wParam", ctypes.c_ulong)
    ]

# RAWKEYBOARD 구조체
class RAWKEYBOARD(ctypes.Structure):
    _fields_ = [
        ("MakeCode", ctypes.c_ushort),
        ("Flags", ctypes.c_ushort),
        ("Reserved", ctypes.c_ushort),
        ("VKey", ctypes.c_ushort),
        ("Message", ctypes.c_uint),
        ("ExtraInformation", ctypes.c_ulong)
    ]

# RAWINPUT 구조체
class RAWINPUT(ctypes.Structure):
    class _u(ctypes.Union):
        _fields_ = [("keyboard", RAWKEYBOARD)]
    _anonymous_ = ("u",)
    _fields_ = [
        ("header", RAWINPUTHEADER),
        ("u", _u)
    ]

# 장치 이름 얻는 함수
def get_device_name(hDevice):
    size = ctypes.c_uint(0)
    if user32.GetRawInputDeviceInfoW(hDevice, RIDI_DEVICENAME, None, ctypes.byref(size)) == 0:
        buffer = ctypes.create_unicode_buffer(size.value)
        if user32.GetRawInputDeviceInfoW(hDevice, RIDI_DEVICENAME, buffer, ctypes.byref(size)) > 0:
            return buffer.value
    return None

# Raw Input 등록 함수
def register_raw_input(hwnd):
    rid = RAWINPUTDEVICE()
    rid.usUsagePage = 0x01   # Generic Desktop Controls
    rid.usUsage = 0x06       # Keyboard
    rid.dwFlags = RIDEV_INPUTSINK | RIDEV_NOLEGACY
    rid.hwndTarget = hwnd
    if not user32.RegisterRawInputDevices(ctypes.byref(rid), 1, ctypes.sizeof(rid)):
        raise ctypes.WinError()

# 가상 키 코드 -> 문자열
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

WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

class LaptopTestApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("노트북 테스트 프로그램")
        self.geometry("460x600")

        # 내부 키보드의 Raw Input device 문자열(화이트리스트)
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

        # 배터리 리포트 파일 경로를 저장할 변수 (초기에는 None)
        self.report_path = None

        self.failed_keys = []  # 누르지 못한 키
        self.disabled_hwids = []

         # 메인 프레임 생성
        main_frame = ttk.Frame(self)
        main_frame.pack(padx=20, pady=20, fill="both", expand=True)

        # 타이틀 라벨
        title_label = ttk.Label(main_frame, text="노트북 기능 테스트", font=("Arial", 16))
        title_label.pack(pady=10)

        self.progress_label = ttk.Label(main_frame, text=self.get_progress_text(), font=("Arial", 12))
        self.progress_label.pack(pady=5)

        self.test_status_labels = {}

        # 키보드 테스트
        kb_frame = ttk.Frame(main_frame)
        kb_frame.pack(fill="x", pady=3)
        kb_button = ttk.Button(kb_frame, text="키보드 테스트", command=self.open_keyboard_test)
        kb_button.pack(side="left")
        kb_status = ttk.Label(kb_frame, text="테스트 전", foreground="red")
        kb_status.pack(side="left", padx=10)
        self.test_status_labels["키보드"] = kb_status

        self.failed_keys_button = ttk.Button(kb_frame, text="누르지 못한 키 보기", command=self.show_failed_keys, state="disabled")
        self.failed_keys_button.pack(side="left", padx=5)

        # USB 테스트 (port 수를 조절합니다.)
        self.usb_ports = {
            "port1": False,
            "port2": False,
            "port3": False,
        }
        self.usb_test_complete = False

        usb_frame = ttk.Frame(main_frame)
        usb_frame.pack(fill="x", pady=3)

        self.usb_button = ttk.Button(usb_frame, text="USB 연결 확인", command=self.start_usb_check)
        self.usb_button.pack(side="left")

        self.usb_refresh_button = ttk.Button(usb_frame, text="새로고침", command=self.refresh_usb_check, state="disabled")
        self.usb_refresh_button.pack(side="left", padx=5)

        usb_status = ttk.Label(usb_frame, text="테스트 전", foreground="red")
        usb_status.pack(side="left", padx=5)
        self.test_status_labels["USB"] = usb_status

        self.usb_port_labels = {}
        self.create_usb_port_labels(main_frame)

        # 카메라(웹캠) 테스트
        cam_frame = ttk.Frame(main_frame)
        cam_frame.pack(fill="x", pady=3)

        cam_button = ttk.Button(cam_frame, text="카메라(웹캠) 테스트", command=self.open_camera_test)
        cam_button.pack(side="left")

        cam_status = ttk.Label(cam_frame, text="테스트 전", foreground="red")
        cam_status.pack(side="left", padx=10)
        self.test_status_labels["카메라"] = cam_status

        # ---- 충전기 테스트 부분 ----
        # 충전기 포트 상태
        self.c_type_ports = {
            "충전기": False,
        }
        self.c_type_test_complete = False

        c_type_frame = ttk.Frame(main_frame)
        c_type_frame.pack(fill="x", pady=3)

        # 질문 코드에서 start_btn → lab_test.py에서는 self.c_type_button
        self.c_type_button = ttk.Button(c_type_frame, text="충전 테스트 시작", command=self.start_c_type_check)
        self.c_type_button.pack(side="left")

        # 질문 코드에서 refresh_btn → lab_test.py에서는 self.c_type_refresh_button
        self.c_type_refresh_button = ttk.Button(c_type_frame, text="포트 확인", command=self.refresh_c_type_check, state="disabled")
        self.c_type_refresh_button.pack(side="left", padx=5)

        c_type_status = ttk.Label(c_type_frame, text="테스트 전", foreground="red")
        c_type_status.pack(side="left", padx=10)
        self.test_status_labels["충전기"] = c_type_status

        self.c_type_port_labels = {}
        self.create_c_type_port_labels(main_frame)

        # ---- 배터리 리포트 버튼 추가 ----
        battery_report_frame = ttk.Frame(main_frame)
        battery_report_frame.pack(fill="x", pady=3)

        # 배터리 리포트 생성 버튼
        self.battery_report_button = ttk.Button(
            battery_report_frame, 
            text="배터리 리포트 생성", 
            command=self.generate_battery_report  # 버튼 클릭 시 generate_battery_report 메서드 호출
        )
        self.battery_report_button.pack(side="left")

        # 배터리 리포트 확인 버튼
        self.view_report_button = ttk.Button(
            battery_report_frame,
            text="리포트 확인",
            command=self.view_battery_report  # 버튼 클릭 시 view_battery_report 메서드 호출
        )
        self.view_report_button.pack(side="left", padx=5)

        # 테스트 결과와 배터리 리포트를 모아서 QR 코드 생성 버튼
        self.qr_code_button = ttk.Button(
            battery_report_frame,
            text="QR 코드 생성",
            command=self.generate_qr_code
        )
        self.qr_code_button.pack(side="left", padx=5)


    # 진행 상황 텍스트
    def get_progress_text(self):
        done_count = sum(self.test_done.values())
        total = len(self.test_done)
        return f"{done_count}/{total} 완료"

    def mark_test_complete(self, test_name):
        """특정 테스트 완료 처리 후 전체 진행 상태 업데이트"""
        if test_name in self.test_done and not self.test_done[test_name]:
            self.test_done[test_name] = True
            self.progress_label.config(text=self.get_progress_text())

            # 상태 라벨 업데이트
            status_label = self.test_status_labels[test_name]
            status_label.config(text="테스트 완료", foreground="blue")

            # 모든 테스트가 완료되었는지 확인
            if all(self.test_done.values()):
                messagebox.showinfo("모든 테스트 완료", "모든 테스트를 완료했습니다.\n수고하셨습니다!")

    # ------------------ 키보드 테스트 ------------------
    def open_keyboard_test(self):
        kb_window = tk.Toplevel(self)
        kb_window.title("키보드 테스트")
        kb_window.geometry("1200x500")
        info_label = ttk.Label(kb_window, text="이 창에 포커스를 두고\n모든 키를 한 번씩 눌러보세요.\n완료 시 창이 닫힙니다.")
        info_label.pack(pady=5)

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
            row_frame = ttk.Frame(kb_window)
            row_frame.pack(pady=5)
            for key in row_keys:
                key_upper = key.upper()
                self.all_keys.add(key_upper)
                btn = tk.Label(
                    row_frame, text=key, width=5, borderwidth=1, relief="solid",
                    font=("Arial", 12), background="lightgray"
                )
                btn.pack(side="left", padx=3)
                self.key_widgets[key_upper] = btn

        self.keys_not_pressed = set(self.all_keys)

        # Raw Input 등록
        hwnd = kb_window.winfo_id()
        register_raw_input(hwnd)

        def raw_input_wnd_proc(hWnd, msg, wParam, lParam):
            if msg == WM_NCDESTROY:
                if self._kb_old_wnd_proc is not None:
                    user32.SetWindowLongPtrW(hWnd, GWL_WNDPROC, self._kb_old_wnd_proc)
                    self._kb_old_wnd_proc = None
                return 0

            if msg == WM_INPUT:
                size = ctypes.c_uint(0)
                if user32.GetRawInputData(lParam, RID_INPUT, None, ctypes.byref(size),
                                          ctypes.sizeof(RAWINPUTHEADER)) == 0:
                    buffer = ctypes.create_string_buffer(size.value)
                    if user32.GetRawInputData(lParam, RID_INPUT, buffer, ctypes.byref(size),
                                              ctypes.sizeof(RAWINPUTHEADER)) == size.value:
                        raw = ctypes.cast(buffer, ctypes.POINTER(RAWINPUT)).contents
                        if raw.header.dwType == RIM_TYPEKEYBOARD:
                            if (raw.u.keyboard.Flags & RI_KEY_BREAK) == 0:  # Key Down
                                vkey = raw.u.keyboard.VKey
                                if vkey == 0x0D:
                                    if raw.u.keyboard.Flags & RI_KEY_E0:
                                        key_sym = "NUMENTER"
                                    else:
                                        key_sym = "ENTER"
                                elif vkey == 0x10:
                                    if raw.u.keyboard.MakeCode == 0x2A:
                                        key_sym = "LSHIFT"
                                    elif raw.u.keyboard.MakeCode == 0x36:
                                        key_sym = "RSHIFT"
                                    else:
                                        key_sym = "SHIFT"
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
            if self.keys_not_pressed:
                self.failed_keys = list(self.keys_not_pressed)
                self.test_status_labels["키보드"].config(text="오류 발생", foreground="red")
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
        if hasattr(self, '_kb_hwnd') and self._kb_hwnd and self._kb_old_wnd_proc is not None:
            user32.SetWindowLongPtrW(self._kb_hwnd, GWL_WNDPROC, self._kb_old_wnd_proc)
            self._kb_old_wnd_proc = None
        if hasattr(self, 'kb_window_ref'):
            self.kb_window_ref.destroy()

    def on_raw_key(self, key):
        if key in self.keys_not_pressed:
            self.keys_not_pressed.remove(key)
            widget = self.key_widgets.get(key)
            if widget:
                widget.config(background="black", foreground="white")
            if not self.keys_not_pressed:
                messagebox.showinfo("키보드 테스트", "모든 키를 눌렀습니다! 테스트 통과!")
                self.failed_keys_button.config(state="disabled")
                self.close_keyboard_window()
                self.mark_test_complete("키보드")

    def show_failed_keys(self):
        if self.failed_keys:
            failed_win = tk.Toplevel(self)
            failed_win.title("미처 누르지 못한 키 목록")
            failed_win.geometry("300x200")
            info_label = ttk.Label(failed_win, text="누르지 못한 키:")
            info_label.pack(padx=10, pady=10)
            failed_keys_str = ", ".join(sorted(self.failed_keys))
            keys_label = ttk.Label(failed_win, text=failed_keys_str, font=("Arial", 12))
            keys_label.pack(padx=10, pady=10)
        else:
            messagebox.showinfo("확인", "누르지 못한 키가 없습니다.")

    # ------------------ USB 테스트 ------------------
    def create_usb_port_labels(self, main_frame):
        usb_port_frame = ttk.Frame(main_frame)
        usb_port_frame.pack(fill="x", pady=3)
        for port_name in self.usb_ports:
            label = ttk.Label(usb_port_frame, text=f"{port_name}: 연결 안됨",
                              width=16, borderwidth=1, relief="solid")
            label.pack(side="left", padx=5)
            self.usb_port_labels[port_name] = label

    def start_usb_check(self):
        self.usb_ports = {
            "port1": False,
            "port2": False,
            "port3": False,
        }
        self.usb_test_complete = False
        for port_name, label in self.usb_port_labels.items():
            label.config(text=f"{port_name}: 연결 안됨")

        self.usb_button.config(state="disabled")
        self.usb_refresh_button.config(state="normal")
        self.test_status_labels["USB"].config(text="테스트 중", foreground="orange")

        messagebox.showinfo("USB Test", "USB 포트에 USB 장치를 연결하고 새로고침을 누르세요.\n포트당 하나의 장치만 연결해야 합니다.")

    def refresh_usb_check(self):
        try:
            wmi_obj = win32com.client.GetObject("winmgmts:")
            pnp_entities = wmi_obj.InstancesOf("Win32_PnPEntity")

            # usb_ports와 usb_port_labels의 키를 "port1", "port2", "port3"으로 통일합니다.
            # 예를 들어, self.usb_ports = {"port1": False, "port2": False, "port3": False}
            for entity in pnp_entities:
                if hasattr(entity, 'PNPDeviceID') and entity.PNPDeviceID:
                    # 대소문자 구분 없이 비교하기 위해 대문자로 변환
                    device_path = entity.PNPDeviceID.upper()    
                    
                    # device_path가 "USB\"로 시작하는지 확인
                    if not device_path.startswith("USB\\"):
                        continue

                    # device_path의 끝이 "&0&<숫자>"로 끝나는지 정규표현식으로 추출
                    match = re.search(r'&0&(\d)$', device_path)
                    if match:
                        port_number = match.group(1)  # 추출된 숫자 (문자열)
                        # 해당 숫자가 1, 2, 3 중 하나인 경우에만 처리
                        if port_number in ['1', '2', '3']:
                            # self.usb_ports의 키는 "port1", "port2", "port3"로 구성되어 있다고 가정
                            key = f"port{port_number}"
                            self.usb_ports[key] = True
                            self.usb_port_labels[key].config(text=f"{key}: 연결됨", background="lightgreen")
                            
            if all(self.usb_ports.values()):
                self.usb_test_complete = True
                self.usb_refresh_button.config(state="disabled")
                self.mark_test_complete("USB")
                messagebox.showinfo("USB Test", "모든 USB 포트 테스트 완료!")
            else:
                messagebox.showinfo("USB Test",
                                    "USB 포트에 USB 장치를 연결하고 새로고침을 누르세요.\n포트당 하나의 장치만 연결해야 합니다.")

        except Exception as e:
            messagebox.showerror("USB Error", f"Error checking USB ports:\n{e}")
    # ------------------ 카메라 테스트 ------------------
    def open_camera_test(self):
        cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
        if not cap.isOpened():
            messagebox.showerror("카메라 오류", "카메라를 열 수 없습니다. 장치를 확인해주세요.")
            self.test_status_labels["카메라"].config(text="오류 발생", foreground="red")
            return

        messagebox.showinfo("카메라 테스트", "카메라 창이 뜨면 영상이 보이는지 확인하세요.\n상단 X 버튼을 마우스로 누르면 종료됩니다.")
        window_name = "Camera Test - X to exit"
        while True:
            ret, frame = cap.read()
            if not ret:
                messagebox.showerror("카메라 오류", "카메라 프레임을 읽을 수 없습니다.")
                break
            cv2.imshow(window_name, frame)
            key = cv2.waitKey(1) & 0xFF
            if key == 27:  # ESC
                break
            if cv2.getWindowProperty(window_name, cv2.WND_PROP_VISIBLE) < 1:
                break

        cap.release()
        cv2.destroyAllWindows()
        self.mark_test_complete("카메라")

    # ------------------ 충전 테스트 ------------------
    def create_c_type_port_labels(self, main_frame):
        c_type_port_frame = ttk.Frame(main_frame)
        c_type_port_frame.pack(fill="x", pady=3)
        for port_name in self.c_type_ports:
            label = ttk.Label(c_type_port_frame, text=f"{port_name}: 연결 안됨",
                              width=20, borderwidth=1, relief="solid")
            label.pack(side="left", padx=5)
            self.c_type_port_labels[port_name] = label

    def start_c_type_check(self):
        """
        충전 테스트 초기화 후 안내 메시지.
        '포트 확인' 버튼을 통해 개별 포트 연결 여부를 체크한다.
        """
        # 포트 상태 초기화
        self.c_type_ports = {"충전기": False}
        for port, lbl in self.c_type_port_labels.items():
            lbl.config(text=f"{port}: 연결 안됨", background="SystemButtonFace")

        self.c_type_test_complete = False
        self.test_status_labels["충전기"].config(text="테스트 중", foreground="orange")

        messagebox.showinfo(
            "충전 Test",
            "충전기를 꽂고 포트 확인 버튼을 눌러주세요."
        )
        self.c_type_refresh_button.config(state="normal")

    def refresh_c_type_check(self):
        """
        '포트 확인' 버튼을 누르면 실제 포트 상태(check_c_type_port)를 확인하고
        테스트 상태를 업데이트한다.
        """
        self.check_c_type_port()
        self.update_c_type_status()

    def check_c_type_port(self):
        """
        현재 배터리 상태(충전 여부)를 확인한 뒤,
        아직 확인되지 않은 포트를 순서대로 처리.
        """
        battery = psutil.sensors_battery()
        if battery is None:
            messagebox.showerror("충전기 Error", "배터리 정보를 가져올 수 없습니다.")
            return

        if not battery.power_plugged:
            # 충전기가 전혀 연결 안 된 경우
            messagebox.showinfo(
                "충전기 Test",
                "충전기가 연결되지 않았습니다.\n해당 포트에 충전기를 연결 후 다시 확인하세요."
            )
            return

        if not self.c_type_ports["충전기"]:
            self.c_type_ports["충전기"] = True
            self.c_type_port_labels["충전기"].config(
                text="전원 연결됨 (충전 중)",
                background="lightgreen"
            )

        else:
            # 두 포트 모두 이미 확인됨
            messagebox.showinfo("충전 Test", "충전 확인되었습니다.")

        # 모두 연결되었다면 테스트 완료
        if all(self.c_type_ports.values()):
            self.c_type_test_complete = True
            self.c_type_refresh_button.config(state="disabled")
            self.test_status_labels["충전기"].config(text="테스트 완료", foreground="blue")
            # lab_test.py에서는 mark_test_complete() 호출 필요
            self.mark_test_complete("충전기")
            messagebox.showinfo("충전 Test", "충전 테스트 완료!")

    def update_c_type_status(self):
        """
        포트 상태에 따라 '테스트 중' 또는 '테스트 완료' 라벨을 갱신한다.
        """
        if all(self.c_type_ports.values()):
            self.test_status_labels["충전기"].config(text="테스트 완료", foreground="blue")
        else:
            self.test_status_labels["충전기"].config(text="테스트 중", foreground="orange")


    # ==================== 배터리 리포트 생성 메서드 =============
    def generate_battery_report(self):
        """
        powercfg /batteryreport 명령어를 실행하여 배터리 리포트를 생성하고,
        생성된 리포트 파일의 경로를 사용자에게 알립니다.
        """
        try:
            # 임시 디렉토리 생성 (리포트 파일을 저장할 위치)
            temp_dir = tempfile.mkdtemp()
            # 리포트 파일 경로 설정 (HTML 파일)
            self.report_path = os.path.join(temp_dir, "battery_report.html")
            
            # powercfg /batteryreport 명령어 실행 ("/html" 매개변수 제거)
            subprocess.run(
                ["powercfg", "/batteryreport", "/output", self.report_path],
                check=True,           # 오류 발생 시 예외 발생
                capture_output=True,  # stdout, stderr 캡처
                text=True             # 출력 결과를 문자열로 반환
            )
            
            # 리포트 생성 성공 시 사용자에게 파일 경로 안내
            messagebox.showinfo("배터리 리포트", 
                                f"배터리 리포트가 생성되었습니다.\n파일 경로:\n{self.report_path}")
        except subprocess.CalledProcessError as e:
            # 명령어 실행 중 오류가 발생한 경우
            messagebox.showerror("배터리 리포트 오류", 
                                 f"명령 실행 중 오류 발생:\n{e.stderr}")
        except Exception as e:
            # 기타 예외 발생 시
            messagebox.showerror("배터리 리포트 오류", 
                                 f"오류 발생:\n{e}")
            
    
    # ################# 배터리 리포트 확인 메서드 ######################
    def view_battery_report(self):
        """
        생성된 배터리 리포트 파일을 확인합니다.
        만약 리포트가 생성되지 않았다면 에러 메시지를 표시합니다.
        """
        if self.report_path and os.path.exists(self.report_path):
            try:
                # Windows의 경우 os.startfile()을 사용하여 파일 열기
                os.startfile(self.report_path)
            except Exception as e:
                # 파일 열기 실패 시 에러 메시지 출력
                messagebox.showerror("리포트 확인 오류", f"리포트를 열 수 없습니다:\n{e}")
        else:
            messagebox.showwarning("리포트 없음", "아직 배터리 리포트가 생성되지 않았습니다.\n먼저 '배터리 리포트 생성' 버튼을 눌러주세요.")

    def generate_qr_code(self):
        """
        테스트 결과와 배터리 리포트 정보를 모아서 QR 코드를 생성합니다.
        생성된 QR 코드는 새 창에서 이미지로 표시됩니다.
        """
        # 테스트 결과 문자열 생성
        results_text = "테스트 결과:\n"
        for test, done in self.test_done.items():
            status = "완료" if done else "미완료"
            results_text += f"{test}: {status}\n"

        # 배터리 리포트 내용 읽기 (리포트 파일이 존재하면)
        report_content = ""
        if self.report_path and os.path.exists(self.report_path):
            try:
                with open(self.report_path, "r", encoding="utf-8") as f:
                    report_content = f.read()
            except Exception as e:
                report_content = f"리포트 읽기 오류: {e}"
        else:
            report_content = "생성된 배터리 리포트가 없습니다."

        # QR 코드에 포함할 전체 데이터 (테스트 결과 + 배터리 리포트 내용)
        qr_data = results_text + "\n배터리 리포트 내용:\n" + report_content

        # (주의) QR 코드 용량 제한으로 인해 데이터가 너무 길면 스캔이 어려울 수 있음.
        try:
            # QR 코드 객체 생성 (버전은 fit=True로 설정하여 데이터에 맞게 자동 확장)
            qr = qrcode.QRCode(
                version=None,  # 데이터에 따라 버전을 자동 결정
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=4,
            )
            qr.add_data(qr_data)
            qr.make(fit=True)

            # QR 코드 이미지 생성 (PIL 이미지 객체)
            img = qr.make_image(fill_color="black", back_color="white")
            
            # Tkinter에서 사용 가능한 이미지 객체로 변환
            qr_img = ImageTk.PhotoImage(img)
            
            # QR 코드를 표시할 새 창 생성
            qr_window = tk.Toplevel(self)
            qr_window.title("테스트 결과 및 배터리 리포트 QR 코드")
            
            # QR 코드 이미지를 표시할 라벨 생성
            qr_label = tk.Label(qr_window, image=qr_img)
            qr_label.image = qr_img  # 이미지 객체 참조 유지 (가비지 컬렉션 방지)
            qr_label.pack(padx=10, pady=10)
        except Exception as e:
            messagebox.showerror("QR 코드 생성 오류", f"QR 코드 생성 중 오류 발생:\n{e}")



if __name__ == "__main__":
    app = LaptopTestApp()
    app.mainloop()
