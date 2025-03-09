# lab_test.py
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import subprocess
import cv2
import win32com.client  # WMI (pywin32)
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
RI_KEY_E0 = 0x02

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
    0x08: "BACK",

    # 기능키(F1~F12)
    0x70: "F1",  0x71: "F2",  0x72: "F3",  0x73: "F4",
    0x74: "F5",  0x75: "F6",  0x76: "F7",  0x77: "F8",
    0x78: "F9",  0x79: "F10", 0x7A: "F11", 0x7B: "F12",

    # 편집/제어키 (Insert, Delete, Home, End, PgUp, PgDn)
    0x2D: "INS",
    0x2E: "DEL",
    # 0x24: "HOME",
    # 0x23: "END",
    # 0x21: "PGUP",
    # 0x22: "PGDN",

    # 방향키
    0x25: "LEFT",
    0x26: "UP",
    0x27: "RIGHT",
    0x28: "DOWN",

    # Lock키 (CapsLock, NumLock, ScrollLock)
    0x14: "CAPS",
    0x90: "NUMLOCK",

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
    0x6C: "NUMENTER",
    0x6D: "NUM -",
    0x6E: "NUM .",
    0x6F: "NUM /",


    # 기타 (PrintScreen, Pause/Break 등)
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
    # 0x13: "PAUSE",

    # 수정/조합키(Shift, Ctrl, Alt 등) - 일반적으로 누르면 VKey 이벤트가 뜨긴 하나,
    # RawInput에서 활용시 눌렀다는 것만 인지 가능(문자 입력 X)
    0xA0: "LSHIFT",
    0xA1: "RSHIFT",
    0x11: "CTRL",
    0x5B: "WINDOW",
    0x12: "ALT",
    0x15: "한/영",
    0x19: "한자",
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
            "\\ACPI#MSF0001"
        ]

        # 기타 테스트 관련 변수들
        self.disabled_hwids = []
        self.test_done = {
            "키보드": False,
            "카메라": False,
            "USB": False,
        }

        # 각 테스트의 상태를 저장할 딕셔너리
        self.test_status_labels = {}
        self.failed_keys = []  # 누르지 못한 키를 저장할 변수

        main_frame = ttk.Frame(self)
        main_frame.pack(padx=20, pady=20, fill="both", expand=True)
        title_label = ttk.Label(main_frame, text="노트북 기능 테스트", font=("Arial", 16))
        title_label.pack(pady=10)
        self.progress_label = ttk.Label(main_frame, text=self.get_progress_text(), font=("Arial", 12))
        self.progress_label.pack(pady=5)

        # 1) 키보드
        kb_frame = ttk.Frame(main_frame)
        kb_frame.pack(fill="x", pady=3)
        kb_button = ttk.Button(kb_frame, text="키보드 테스트", command=self.open_keyboard_test)
        kb_button.pack(side="left")
        kb_status = ttk.Label(kb_frame, text="테스트 전", foreground="red")
        kb_status.pack(side="left", padx=10)
        self.test_status_labels["키보드"] = kb_status

        # 누르지 못한 키를 확인할 버튼(기본 비활성화)
        self.failed_keys_button = ttk.Button(kb_frame, text="누르지 못한 키 보기", command=self.show_failed_keys, state="disabled")
        self.failed_keys_button.pack(side="left", padx=5)

        # USB test variables
        self.usb_ports = {
            "Port_#0001.Hub_#0001": False,
            "Port_#0002.Hub_#0001": False,
            "Port_#0003.Hub_#0001": False,
        }
        self.usb_test_complete = False

        # 2) USB
        usb_frame = ttk.Frame(main_frame)
        usb_frame.pack(fill="x", pady=3)

        self.usb_button = ttk.Button(usb_frame, text="USB 연결 확인", command=self.start_usb_check)
        self.usb_button.pack(side="left")

        self.usb_refresh_button = ttk.Button(usb_frame, text="새로고침", command=self.refresh_usb_check, state="disabled")
        self.usb_refresh_button.pack(side="left", padx=5)

        usb_status = ttk.Label(usb_frame, text="테스트 전", foreground="red")
        usb_status.pack(side="left", padx=10)
        self.test_status_labels["USB"] = usb_status

        self.usb_port_labels = {}
        self.create_usb_port_labels(main_frame)


        # 3) 카메라
        cam_frame = ttk.Frame(main_frame)
        cam_frame.pack(fill="x", pady=3)

        cam_button = ttk.Button(cam_frame, text="카메라(웹캠) 테스트", command=self.open_camera_test)
        cam_button.pack(side="left")

        cam_status = ttk.Label(cam_frame, text="테스트 전", foreground="red")
        cam_status.pack(side="left", padx=10)
        self.test_status_labels["카메라"] = cam_status


    # ============ 진행 상황 텍스트 ============
    def get_progress_text(self):
        done_count = sum(self.test_done.values())
        total = len(self.test_done)
        return f"{done_count}/{total} 완료"

    def mark_test_complete(self, test_name):
        """
        특정 테스트를 완료 처리하고,
        진행 상황 갱신 후, 모든 테스트가 끝났으면 팝업창 표시
        - 완료 시 test_status_labels[test_name]를 "테스트 완료"(파란색)로 변경
        """
        if test_name in self.test_done and not self.test_done[test_name]:
            self.test_done[test_name] = True
            # 진행 상황 라벨 갱신
            self.progress_label.config(text=self.get_progress_text())
            
            # 상태 라벨 업데이트트 
            status_label = self.test_status_labels[test_name]
            status_label.config(text="테스트 완료", foreground="blue")

            # 모든 테스트가 완료되었는지 확인인
            if all(self.test_done.values()):
                # 완료 확인 표시창(팝업업)
                messagebox.showinfo("모든 테스트 완료", "모든 테스트를 완료했습니다.\n수고하셨습니다!")

    def show_failed_keys(self):
        # 누르지 못한 키가 있을 때 이를 표시하는 창 생성
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

    # ============ 키보드 테스트 ============
    def open_keyboard_test(self):
        kb_window = tk.Toplevel(self)
        kb_window.title("키보드 테스트")
        kb_window.geometry("1200x500")
        info_label = ttk.Label(kb_window, text="이 창에 포커스를 두고\n모든 키를 한 번씩 눌러보세요.\n완료 시 창이 닫힙니다.")
        info_label.pack(pady=5)

        keyboard_layout = [
            # 함수키 라인 (F1 ~ F12)
            ["F1", "F2", "F3", "F4", "F5", "F6",
             "F7", "F8", "F9", "F10", "F11", "F12"],

            # 숫자열 & 편집 키 (INS, DEL 등)
            ["`", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "-", "=", "BACK",
             "INS", "DEL"],
                
            # QWERTY 1
            ["Q", "W", "E", "R", "T", "Y", "U", "I", "O", "P", "[", "]", "\\"],
            # QWERTY 2
            ["A", "S", "D", "F", "G", "H", "J", "K", "L", ";", "'", "ENTER"],
            # QWERTY 3
            ["Z", "X", "C", "V", "B", "N", "M", ",", ".", "/", "SPACE"],

            # 방향키
            ["UP", "LEFT", "DOWN", "RIGHT"],

            # Lock 키들
            ["CAPS", "NUMLOCK"],

            # 숫자패드
            ["NUM7", "NUM8", "NUM9", "NUM /",
             "NUM4", "NUM5", "NUM6", "NUM *",
             "NUM1", "NUM2", "NUM3", "NUM -",
             "NUM0", "NUM .", "NUM +", "NUMENTER"],

            # 기타
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

        # 새 콜백 함수 정의 (WndProc)
        def raw_input_wnd_proc(hWnd, msg, wParam, lParam):
            if msg == WM_NCDESTROY:
                if self._kb_old_wnd_proc is not None:
                    user32.SetWindowLongPtrW(hWnd, GWL_WNDPROC, self._kb_old_wnd_proc)
                    self._kb_old_wnd_proc = None
                return 0

            if msg == WM_INPUT:
                size = ctypes.c_uint(0)
                if user32.GetRawInputData(lParam, RID_INPUT, None, ctypes.byref(size), ctypes.sizeof(RAWINPUTHEADER)) == 0:
                    buffer = ctypes.create_string_buffer(size.value)
                    if user32.GetRawInputData(lParam, RID_INPUT, buffer, ctypes.byref(size), ctypes.sizeof(RAWINPUTHEADER)) == size.value:
                        raw = ctypes.cast(buffer, ctypes.POINTER(RAWINPUT)).contents
                        if raw.header.dwType == RIM_TYPEKEYBOARD:
                            if (raw.u.keyboard.Flags & RI_KEY_BREAK) == 0:
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
                                elif vkey == 0x2D:
                                    if raw.u.keyboard.Flags & RI_KEY_E0:
                                        key_sym = "INS"
                                    else:
                                        key_sym = "NUMINS"
                                elif vkey == 0x2E:
                                    if raw.u.keyboard.Flags & RI_KEY_E0:
                                        key_sym = "DEL"
                                    else:
                                        key_sym = "NUMDEL"
                                elif vkey == 0x26:
                                    if raw.u.keyboard.Flags & RI_KEY_E0:
                                        key_sym = "UP"
                                    else:
                                        key_sym = "NUMUP"
                                elif vkey == 0x25:
                                    if raw.u.keyboard.Flags & RI_KEY_E0:
                                        key_sym = "LEFT"
                                    else:
                                        key_sym = "NUMLEFT"
                                elif vkey == 0x28:
                                    if raw.u.keyboard.Flags & RI_KEY_E0:
                                        key_sym = "DOWN"
                                    else:
                                        key_sym = "NUMDOWN"
                                elif vkey == 0x27:
                                    if raw.u.keyboard.Flags & RI_KEY_E0:
                                        key_sym = "RIGHT"
                                    else:
                                        key_sym = "NUMRIGHT"
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
                                    print("키:", key_sym, "is_internal:", is_internal)
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
                # 버튼 활성화: 누르지 못한 키가 있을 경우
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
        print("Raw key processed:", key)
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

    # ============ 카메라 테스트 ============
    def open_camera_test(self):
        cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
        if not cap.isOpened():
            messagebox.showerror("카메라 오류", "카메라를 열 수 없습니다. 장치를 확인해주세요.")
            self.test_status_labels["카메라"].config(text="오류 발생", foreground="red")
            return

        messagebox.showinfo("카메라 테스트", "카메라 창이 뜨면 영상이 보이는지 확인하세요.상단 X 버튼을 마우스로 누르면 종료됩니다.")
        window_name = "Camera Test - X to exit"
        while True:
            ret, frame = cap.read()
            if not ret:
                messagebox.showerror("카메라 오류", "카메라 프레임을 읽을 수 없습니다.")
                break
            cv2.imshow(window_name, frame)
            key = cv2.waitKey(1) & 0xFF
            # ESC 키를 누르면 종료
            if key == 27:
                break
            # X 버튼(창 닫기)로 창이 닫혔는지 체크
            if cv2.getWindowProperty(window_name, cv2.WND_PROP_VISIBLE) < 1:
                break

        cap.release()
        cv2.destroyAllWindows()
        self.mark_test_complete("카메라")


    # ============ USB 체크 ============
    def create_usb_port_labels(self, main_frame):
        usb_port_frame = ttk.Frame(main_frame)
        usb_port_frame.pack(fill="x", pady=3)
        for port_name in self.usb_ports:
            label = ttk.Label(usb_port_frame, text=f"{port_name}: 연결 안됨", width=30, borderwidth=1, relief="solid")
            label.pack(side="left", padx=5)
            self.usb_port_labels[port_name] = label


    def start_usb_check(self):
        """
        Initializes the USB check.
        """
        # Reset USB ports
        self.usb_ports = {
            "Port_#0001.Hub_#0001": False,
            "Port_#0002.Hub_#0001": False,
            "Port_#0003.Hub_#0001": False,
        }
        self.usb_test_complete = False


        # Update port labels
        for port_name, label in self.usb_port_labels.items():
            label.config(text=f"{port_name}: 연결 안됨")

        self.usb_button.config(state="disabled")
        self.usb_refresh_button.config(state="normal")
        self.test_status_labels["USB"].config(text="테스트 중", foreground="orange")

        messagebox.showinfo("USB Test", "USB 포트에 USB 장치를 연결하고 새로고침을 누르세요. 포트당 하나의 장치만 연결해야 합니다.")

    def refresh_usb_check(self):
        """
        Checks the USB connections using WMI and updates the port states.
        """
        try:
            wmi_obj = win32com.client.GetObject("winmgmts:")
            pnp_entities = wmi_obj.InstancesOf("Win32_PnPEntity")

            # 포트별 예상되는 장치 인스턴스 경로 패턴
            expected_device_paths = {
                "Port_#0001.Hub_#0001": "USB\\VID_25A7&PID_2410\\5&218DD721&0&1",
                "Port_#0002.Hub_#0001": "USB\\VID_25A7&PID_2410\\5&218DD721&0&2",
                "Port_#0003.Hub_#0001": "USB\\VID_25A7&PID_2410\\5&218DD721&0&3",
            }

            # # 각 포트 상태 초기화
            # for port in self.usb_ports:
            #     self.usb_ports[port] = False
            #     self.usb_port_labels[port].config(text=f"{port}: 연결 안됨", background="SystemButtonFace") # 초기화된 상태로 설정

            # 모든 PnP 장치 정보를 순회하면서, 장치 인스턴스 경로를 기반으로 포트 연결 확인
            for entity in pnp_entities:
                if entity.PNPDeviceID:
                    device_path = entity.PNPDeviceID.upper()
                    for port, expected_path_pattern in expected_device_paths.items():
                        print("expected_device_paths: ", expected_device_paths)
                        print("expected_path_pattern: ", expected_path_pattern)
                        print("port: ", port)
                        expected_path_prefix = expected_path_pattern.split("\\*")[0].upper()
                        print("expected_path_prefix: ", expected_path_prefix)
                        if expected_path_prefix in device_path:
                            self.usb_ports[port] = True
                            self.usb_port_labels[port].config(text=f"{port}: 연결됨", background="lightgreen")
                            break

            # check if all port is connected
            if all(self.usb_ports.values()):
                self.usb_test_complete = True
                self.usb_refresh_button.config(state="disabled")
                self.mark_test_complete("USB")
                messagebox.showinfo("USB Test", "모든 USB 포트 테스트 완료!")
            else:
                messagebox.showinfo("USB Test", "USB 포트에 USB 장치를 연결하고 새로고침을 누르세요. 포트당 하나의 장치만 연결해야 합니다.")

        except Exception as e:
            messagebox.showerror("USB Error", f"Error checking USB ports:\n{e}")




if __name__ == "__main__":
    app = LaptopTestApp()
    app.mainloop()
