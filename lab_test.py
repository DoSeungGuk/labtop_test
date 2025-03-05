# lab_test.py
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import subprocess
import cv2
import win32com.client  # WMI (pywin32)
import win32api
import pyaudio
import wave
import time
import ctypes
from ctypes import wintypes  # ctypes.wintypes 임포트

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

# user32 핸들
user32 = ctypes.windll.user32

# SetWindowLongPtrW, CallWindowProcW 함수 원형 지정
# restype과 argtypes를 명시해주어야 호출 규약이 정확히 일치함
user32.SetWindowLongPtrW.restype = LONG_PTR
user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, wintypes.INT, LONG_PTR]

user32.CallWindowProcW.restype = LRESULT
user32.CallWindowProcW.argtypes = [
    LONG_PTR,           # WNDPROC 포인터(정수 주소)
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
    # 주키(문자열)
    0x30: "0",   0x31: "1",   0x32: "2",   0x33: "3",   0x34: "4",
    0x35: "5",   0x36: "6",   0x37: "7",   0x38: "8",   0x39: "9",
    0x41: "A",   0x42: "B",   0x43: "C",   0x44: "D",   0x45: "E",
    0x46: "F",   0x47: "G",   0x48: "H",   0x49: "I",   0x4A: "J",
    0x4B: "K",   0x4C: "L",   0x4D: "M",   0x4E: "N",   0x4F: "O",
    0x50: "P",   0x51: "Q",   0x52: "R",   0x53: "S",   0x54: "T",
    0x55: "U",   0x56: "V",   0x57: "W",   0x58: "X",   0x59: "Y",
    0x5A: "Z",

    # 특수키 (Space, Enter, ESC, Tab 등)
    0x20: "SPACE",
    0x0D: "ENTER",
    0x1B: "ESC",
    0x09: "TAB",

    # 기능키(F1~F12)
    0x70: "F1",  0x71: "F2",  0x72: "F3",  0x73: "F4",
    0x74: "F5",  0x75: "F6",  0x76: "F7",  0x77: "F8",
    0x78: "F9",  0x79: "F10", 0x7A: "F11", 0x7B: "F12",

    # 편집/제어키 (Insert, Delete, Home, End, PgUp, PgDn)
    0x2D: "INS",
    0x2E: "DEL",
    0x24: "HOME",
    0x23: "END",
    0x21: "PGUP",
    0x22: "PGDN",

    # 방향키
    0x25: "LEFT",
    0x26: "UP",
    0x27: "RIGHT",
    0x28: "DOWN",

    # Lock키 (CapsLock, NumLock, ScrollLock)
    0x14: "CAPSLOCK",
    0x90: "NUMLOCK",
    0x91: "SCROLLLOCK",

    # 숫자패드(Numpad)
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
    0x6D: "NUM -",
    0x6E: "NUM .",
    0x6F: "NUM /",

    # 기타 (PrintScreen, Pause/Break 등)
    0x2C: "PRTSCR",
    0x13: "PAUSE",

    # 수정/조합키(Shift, Ctrl, Alt 등) - 일반적으로 누르면 VKey 이벤트가 뜨긴 하나,
    # RawInput에서 활용시 눌렀다는 것만 인지 가능(문자 입력 X)
    0x10: "SHIFT",
    0x11: "CTRL",
    0x12: "ALT",
}

# 새로운 WndProc 시그니처
WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

class LaptopTestApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("노트북 테스트 프로그램")
        self.geometry("460x600")

        # 내부 키보드의 Raw Input device 문자열(화이트리스트)
        self.INTERNAL_HWIDS = [
            "\\ACPI#MSFT0001"
        ]

        # 기타 테스트 관련 변수들
        self.disabled_hwids = []
        self.test_done = {
            "키보드": False
        }
        self.test_status_labels = {}

        main_frame = ttk.Frame(self)
        main_frame.pack(padx=20, pady=20, fill="both", expand=True)
        title_label = ttk.Label(main_frame, text="노트북 기능 테스트", font=("Arial", 16))
        title_label.pack(pady=10)
        self.progress_label = ttk.Label(main_frame, text=self.get_progress_text(), font=("Arial", 12))
        self.progress_label.pack(pady=5)

        kb_frame = ttk.Frame(main_frame)
        kb_frame.pack(fill="x", pady=3)
        kb_button = ttk.Button(kb_frame, text="키보드 테스트", command=self.open_keyboard_test)
        kb_button.pack(side="left")
        kb_status = ttk.Label(kb_frame, text="테스트 전", foreground="red")
        kb_status.pack(side="left", padx=10)
        self.test_status_labels["키보드"] = kb_status

        # (추가 테스트 버튼 생략)

    def get_progress_text(self):
        done_count = sum(self.test_done.values())
        total = len(self.test_done)
        return f"{done_count}/{total} 완료"

    def mark_test_complete(self, test_name):
        if test_name in self.test_done and not self.test_done[test_name]:
            self.test_done[test_name] = True
            self.progress_label.config(text=self.get_progress_text())
            status_label = self.test_status_labels[test_name]
            status_label.config(text="테스트 완료", foreground="blue")
            if all(self.test_done.values()):
                messagebox.showinfo("모든 테스트 완료", "모든 테스트를 완료했습니다.\n수고하셨습니다!")

    def open_keyboard_test(self):
        kb_window = tk.Toplevel(self)
        kb_window.title("키보드 테스트")
        kb_window.geometry("1280x720")
        info_label = ttk.Label(kb_window, text="이 창에 포커스를 두고\n모든 키를 한 번씩 눌러보세요.\n완료 시 창이 닫힙니다.")
        info_label.pack(pady=8)

        keyboard_layout = [
            # 함수키 라인 (F1 ~ F12)
            ["F1", "F2", "F3", "F4", "F5", "F6",
            "F7", "F8", "F9", "F10", "F11", "F12"],

            # 숫자열 & 편집 키 (INS, DEL, HOME, END, PGUP, PGDN)
            ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0",
            "INS", "DEL", "HOME", "END", "PGUP", "PGDN"],

            # QWERTY 1
            ["Q", "W", "E", "R", "T", "Y", "U", "I", "O", "P"],
            # QWERTY 2
            ["A", "S", "D", "F", "G", "H", "J", "K", "L", "ENTER"],
            # QWERTY 3
            ["Z", "X", "C", "V", "B", "N", "M", "SPACE"],

            # 방향키
            ["UP", "LEFT", "DOWN", "RIGHT"],

            # Lock 키들
            ["CAPSLOCK", "NUMLOCK", "SCROLLLOCK"],

            # 숫자패드
            ["NUM7", "NUM8", "NUM9", "NUM /",
            "NUM4", "NUM5", "NUM6", "NUM *",
            "NUM1", "NUM2", "NUM3", "NUM -",
            "NUM0", "NUM .", "NUM +"],

            # 기타
            ["ESC", "TAB", "SHIFT", "CTRL", "ALT", "PRTSCR", "PAUSE"]
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
                    row_frame, text=key, width=10, borderwidth=1, relief="solid",
                    font=("Arial", 12), background="lightgray"
                )
                btn.pack(side="left", padx=3)
                self.key_widgets[key_upper] = btn

        self.keys_not_pressed = set(self.all_keys)

        # Raw Input 등록
        hwnd = kb_window.winfo_id()
        register_raw_input(hwnd)

        # 새 콜백 함수 정의 (WndProc)
        def raw_input_wnd_proc(hWnd, msg, wParam, lParam):
            if msg == WM_NCDESTROY:
                # 창 파괴 시 원래 WndProc 복원
                if self._kb_old_wnd_proc is not None:
                    user32.SetWindowLongPtrW(hWnd, GWL_WNDPROC, self._kb_old_wnd_proc)
                    self._kb_old_wnd_proc = None
                return 0

            if msg == WM_INPUT:
                # WM_INPUT 메시지 처리
                size = ctypes.c_uint(0)
                if user32.GetRawInputData(lParam, RID_INPUT, None, ctypes.byref(size), ctypes.sizeof(RAWINPUTHEADER)) == 0:
                    buffer = ctypes.create_string_buffer(size.value)
                    if user32.GetRawInputData(
                        lParam, RID_INPUT, buffer, ctypes.byref(size),
                        ctypes.sizeof(RAWINPUTHEADER)
                    ) == size.value:
                        raw = ctypes.cast(buffer, ctypes.POINTER(RAWINPUT)).contents
                        if raw.header.dwType == RIM_TYPEKEYBOARD:
                            # 키 다운 이벤트인지 체크
                            if (raw.u.keyboard.Flags & RI_KEY_BREAK) == 0:
                                vkey = raw.u.keyboard.VKey
                                if vkey in VK_MAPPING:
                                    key_sym = VK_MAPPING[vkey]
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

            # 창이 이미 파괴된 상태인지 체크 (옵션)
            if not user32.IsWindow(hWnd):
                return 0

            # 그 외 메시지를 원래의 WndProc에 전달
            if self._kb_old_wnd_proc:
                return user32.CallWindowProcW(self._kb_old_wnd_proc, hWnd, msg, wParam, lParam)
            else:
                return user32.DefWindowProcW(hWnd, msg, wParam, lParam)

        # Python이 콜백을 GC하지 않도록 클래스 멤버 변수로 참조
        self._raw_input_wnd_proc = WNDPROC(raw_input_wnd_proc)

        # SetWindowLongPtrW를 통해 새 WndProc 설정, Old WndProc 정수 주소 반환
        cb_func_ptr = ctypes.cast(self._raw_input_wnd_proc, ctypes.c_void_p).value
        cb_func_ptr = LONG_PTR(cb_func_ptr)  # 64/32bit에 맞게 long/long long 변환
        old_proc = user32.SetWindowLongPtrW(hwnd, GWL_WNDPROC, cb_func_ptr)
        
        # old_proc은 정수 주소. 이후 CallWindowProcW 호출 시 사용
        self._kb_old_wnd_proc = old_proc
        self._kb_hwnd = hwnd
        self.kb_window_ref = kb_window

    # 창 종료 시점에 WndProc 복원 및 창 닫기
    def close_keyboard_window(self):
        if hasattr(self, '_kb_hwnd') and self._kb_hwnd and self._kb_old_wnd_proc is not None:
            user32.SetWindowLongPtrW(self._kb_hwnd, GWL_WNDPROC, self._kb_old_wnd_proc)
            self._kb_old_wnd_proc = None
        if hasattr(self, 'kb_window_ref'):
            self.kb_window_ref.destroy()

    # 키를 누르면 라벨 색상 변경 후 모두 눌리면 테스트 완료
    def on_raw_key(self, key):
        print("Raw key processed:", key)
        if key in self.keys_not_pressed:
            self.keys_not_pressed.remove(key)
            widget = self.key_widgets.get(key)
            if widget:
                widget.config(background="black", foreground="white")
            if not self.keys_not_pressed:
                messagebox.showinfo("키보드 테스트", "모든 키를 눌렀습니다! 테스트 통과!")
                # 종료 전 원복
                self.close_keyboard_window()
                self.mark_test_complete("키보드")

    # 디스플레이 정보 예시
    def show_display_info(self):
        try:
            wmi_obj = win32com.client.GetObject("winmgmts:")
            monitors = wmi_obj.InstancesOf("Win32_VideoController")
            disp_text = ""
            for m in monitors:
                name = m.Name
                width = getattr(m, "CurrentHorizontalResolution", "Unknown")
                height = getattr(m, "CurrentVerticalResolution", "Unknown")
                disp_text += f"그래픽 장치: {name}\n해상도: {width} x {height}\n\n"
            messagebox.showinfo("디스플레이 정보", disp_text.strip())
        except Exception as e:
            messagebox.showerror("에러", f"디스플레이 정보를 가져오는 중 오류 발생:\n{e}")
        finally:
            self.mark_test_complete("디스플레이")

    # 나머지 테스트(카메라, 마우스, Wi-Fi, USB, 배터리, 사운드)는 동일하게 구성

if __name__ == "__main__":
    app = LaptopTestApp()
    app.mainloop()
