import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from tkinter import messagebox
from PIL import Image, ImageTk, ImageFont, ImageDraw  # 이미지 처리 라이브러리
import os
import random  # USB 연결 상태를 랜덤하게 테스트하기 위해 추가
import cv2
import win32com.client  # WMI (pywin32)
import ctypes
from ctypes import wintypes
import psutil
import subprocess
import tempfile
import qrcode
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
# Raw Input 해제 플래그 (RIDEV_REMOVE)
RIDEV_REMOVE = 0x00000001

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

    # 해당 장치의 legacy 메시지(WM_KEYDOWN 등)를 생성하지 않게 하여, 키보드 테스트 창에서는 Raw Input 방식으로만 키 이벤트를 받음음
    rid.dwFlags = RIDEV_INPUTSINK | RIDEV_NOLEGACY
    rid.hwndTarget = hwnd
    if not user32.RegisterRawInputDevices(ctypes.byref(rid), 1, ctypes.sizeof(rid)):
        raise ctypes.WinError()

# ---------------------------------------------
# Raw Input 해제 함수
# ---------------------------------------------
def unregister_raw_input():
    """지정된 윈도우 핸들에 대해 등록된 Raw Input을 해제합니다."""
    rid = RAWINPUTDEVICE()
    rid.usUsagePage = 0x01
    rid.usUsage = 0x06
    rid.dwFlags = RIDEV_REMOVE
    rid.hwndTarget = 0
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


#######################################################
#######################################################

class TestApp(ttkb.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("KkomDae diagnostics")
        self.geometry("875x700") 
        self.resizable(False, False)
        self._style = ttkb.Style()  # 스타일 객체 생성

        # 내부 키보드의 Raw Input device 문자열 (화이트리스트)
        self.INTERNAL_HWIDS = [
            "\\ACPI#MSF0001"
        ]

        # 배터리 리포트 파일 경로 (초기 None)
        self.report_path = None

        self.failed_keys = []  # 누르지 못한 키 목록
        self.disabled_hwids = []

        # 열려있는 테스트 창 관리
        self.active_test_windows = {}

        # 🔹 폰트 파일 직접 로드
        self.samsung_bold_path = "SamsungSharpSans-Bold.ttf"  
        self.samsung_regular_path = "SamsungOne-400.ttf"
        self.notosans_path = "NotoSansKR-VariableFont_wght.ttf"

        # 🔹 Frame 스타일 설정
        self._style.configure("Blue.TFrame", background="#0078D7")   # 타이틀 배경 파란색
        self._style.configure("White.TFrame", background="white")   # 테스트 영역 배경 흰색
        
        # 테스트 완료 여부를 저장할 딕셔너리 초기화
        self.test_done = {
            "키보드": False,
            "카메라": False,
            "USB": False,
            "충전": False,
            "배터리": False,
            "QR코드": False
        }
        # 🔹 테스트 아이콘 및 설명 데이터
        self.test_icons = {
            "키보드": "keyboard.png",
            "카메라": "camera.png",
            "USB": "usb.png",
            "충전": "charging.png",
            "배터리": "battery.png",
            "QR코드": "qrcode.png"
        }

        self.test_descriptions = {
            "키보드": "키 입력이 정상적으로 작동하는지 확인합니다.",
            "카메라": "카메라(웹캠)가 정상적으로 작동하는지 확인합니다.",
            "USB": "모든 USB 포트가 정상적으로 인식되는지 확인합니다.",
            "충전": "노트북이 정상적으로 충전되는지 확인합니다.",
            "배터리": "배터리 리포트를 생성하여 성능을 확인합니다.",
            "QR코드": "테스트 결과를 QR 코드로 생성합니다."
        }

        # 테스트 전/중/완료 상태 문자열 설정
        self.test_status = {
            "키보드": "테스트 전",
            "카메라": "테스트 전",
            "충전": "테스트 전",
            "배터리": "생성 전",
            "QR코드": "생성 전"
        }

        self.test_status_ing = {
            "키보드": "테스트 중",
            "카메라": "테스트 중",
            "충전": "테스트 중",
            "배터리": "생성 중",
            "QR코드": "생성 중"
        }

        self.test_status_done = {
            "키보드": "테스트 완료",
            "카메라": "테스트 완료",
            # USB 완료 상태는 check_usb_ports 함수에서 동적으로 구성됨
            "USB": "",
            "충전": "테스트 완료",
            "배터리": "생성 완료",
            "QR코드": "생성 완료"
        }

        # 테스트 완료 여부를 저장할 딕셔너리 초기화
        self.test_done = {
            "키보드": False,
            "카메라": False,
            "USB": False,
            "충전": False,
            "배터리": False,
            "QR코드": False
        }

        # 🔹 USB 포트 상태 (처음엔 모두 비연결 상태)
        self.usb_ports = {
            "port1": False,
            "port2": False,
            "port3": False,
        }
        self.usb_test_complete = False

        self.test_status_labels = {}

        # 타이틀 영역 생성
        self.create_title_section()
        
        # 테스트 항목 UI 구성
        self.create_test_items()
        
        self.c_type_port_labels = {}

    def open_test_window(self, test_name, create_window_func):
        # 이미 해당 테스트의 창이 열려 있다면 경고 후 반환
        if test_name in self.active_test_windows:
            messagebox.showwarning("경고", f"{test_name} 테스트 창이 이미 열려 있습니다.")
            return
        # 새 창 생성
        window = create_window_func()

        # 창이 닫힐 때 딕셔너리에서 제거
        self.active_test_windows[test_name] = window
        return window

    def on_test_window_close(self, test_name):
        if test_name in self.active_test_windows:
            del self.active_test_windows[test_name]

    def create_title_section(self):
        title_frame = ttkb.Frame(self, style="Blue.TFrame")
        title_frame.place(relx=0, rely=0, relwidth=1, relheight=0.35)

        # SSAFY 로고 이미지 삽입
        img_path = "ssafy_logo.png"
        image = Image.open(img_path)
        image = image.resize((80, 60), Image.LANCZOS)
        self.ssafy_logo = ImageTk.PhotoImage(image)
        img_label = ttkb.Label(title_frame, image=self.ssafy_logo, background="#0078D7", anchor="w")
        img_label.grid(row=0, column=0, padx=30, pady=(30, 10), sticky="w")  # 하단 여백을 조절

        # 컨테이너 프레임 생성 (타이틀과 서브타이틀)
        text_container = ttkb.Frame(title_frame, style="Blue.TFrame")
        text_container.grid(row=1, column=0, padx=20, sticky="w")

        self.title_img = self.create_text_image(
            "KkomDae diagnostics", (800, 45), self.samsung_bold_path, 28, (255, 255, 255), align_left=True
        )
        title_label = ttkb.Label(text_container, image=self.title_img, background="#0078D7", anchor="w")
        title_label.grid(row=0, column=0, sticky="w", pady=(0, 0))

        # 첫 번째 서브타이틀 라인
        self.subtitle_img1 = self.create_text_image(
            "KkomDae diagnostics로 노트북을 빠르고 꼼꼼하게 검사해보세요.",
            (800,27),  # 높이 조정
            self.notosans_path, 14, (255, 255, 255, 255), align_left=True
        )
        subtitle_label1 = ttkb.Label(text_container, image=self.subtitle_img1, background="#0078D7", anchor="w")
        subtitle_label1.grid(row=1, column=0, sticky="w", pady=(0, 0))

        # 두 번째 서브타이틀 라인
        self.subtitle_img2 = self.create_text_image(
            "로고를 클릭하면 테스트 or 생성을 시작할 수 있습니다.",
            (800, 27),  # 높이 조정
            self.notosans_path, 14, (255, 255, 255, 255), align_left=True
        )  
        subtitle_label2 = ttkb.Label(text_container, image=self.subtitle_img2, background="#0078D7", anchor="w")
        subtitle_label2.grid(row=2, column=0, sticky="w", pady=(0, 0))

    def create_text_image(self, text, size, font_path, font_size, color, align_left=False):
        """ 텍스트를 이미지로 변환 (왼쪽 정렬 옵션 추가) """
        img = Image.new("RGBA", size, (0, 0, 0, 0))  # 투명한 배경
        draw = ImageDraw.Draw(img)

        # 폰트 로드 (경로 기반)
        try:
            font = ImageFont.truetype(font_path, font_size)
        except IOError:
            print(f"⚠️ 폰트 '{font_path}'을 찾을 수 없습니다. 기본 폰트 사용")
            font = ImageFont.load_default()

        # 텍스트 위치 설정
        text_x = 10 if align_left else (size[0] - draw.textbbox((0, 0), text, font=font)[2]) // 2
        text_y = (size[1] - font_size) // 2
        draw.text((text_x, text_y), text, font=font, fill=color, spacing=2)

        return ImageTk.PhotoImage(img)

    def create_test_items(self):
        """ 테스트 항목 UI 생성 """
        test_frame = ttkb.Frame(self, style="White.TFrame")  # ✅ 흰색 배경 적용
        test_frame.place(relx=0.1, rely=0.35, relwidth=0.8, relheight=0.6)

        self.tests = ["키보드", "카메라", "USB", "충전", "배터리", "QR코드"]

        for idx, test_name in enumerate(self.tests):
            self.create_test_item(test_frame, test_name, row=idx//3, col=idx%3)

    def create_test_item(self, parent, name, row, col):
        """ 개별 테스트 항목 생성 (각 테스트마다 아이콘과 설명 다르게 설정) """
        frame = ttkb.Frame(parent, padding=10)  # ✅ 부모 배경이 흰색이므로 그대로 둠
        frame.grid(row=row, column=col, padx=20, pady=10)

        # 아이콘 이미지 불러오기
        icon_path = self.test_icons.get(name, "default.png")  # 기본값 설정
        icon_img = Image.open(icon_path).resize((50, 50), Image.LANCZOS)
        icon_photo = ImageTk.PhotoImage(icon_img)

        icon_label = ttkb.Label(frame, image=icon_photo)
        icon_label.image = icon_photo  # 참조 유지
        icon_label.pack()

        name_label = ttkb.Label(frame, text=name, font=("맑은 고딕", 14, "bold"))
        name_label.pack()

        desc_label = ttkb.Label(frame, text=self.test_descriptions.get(name, ""), font=("맑은 고딕", 10), wraplength=180, justify="center")
        desc_label.pack()

        status_label = ttkb.Label(frame, text=self.test_status.get(name, ""), bootstyle="danger", font=("맑은 고딕", 12))
        status_label.pack()
        self.test_status_labels[name] = status_label

        # ----------------- 키보드 테스트 -----------------
        if name == "키보드":
            self.failed_keys_button = ttkb.Button(frame, text="누르지 못한 키 보기",
                                                state="disabled",
                                                bootstyle=WARNING,
                                                command=self.show_failed_keys)
            self.failed_keys_button.pack(side=LEFT, padx=5)

        # ----------------- USB 테스트 -----------------
        # USB 항목의 경우 상태 레이블은 보이지 않도록 처리
        if name == "USB":
            # 상태 레이블은 생성은 하지만, 화면에서 숨김
            status_label.pack_forget()
            self.usb_status_label = status_label  # USB 상태 라벨 저장
            self.usb_port = []
            port_frame = ttkb.Frame(frame)
            port_frame.pack(pady=0)

            for port in range(1, 4):
                # 초기 상태: 미연결
                port_label = ttkb.Label(
                    port_frame,
                    text=f"port{port}",
                    font=("맑은 고딕", 12),
                    bootstyle="danger",
                    width=7  # 여백 조절용
                )

                port_label.pack(side="left", padx=2, pady=0)
                self.usb_port.append(port_label)

            self.usb_refresh_button = ttkb.Button(frame, text="새로고침",
                                                  bootstyle = SECONDARY,
                                                  command=self.refresh_usb_check,
                                                  state="disabled")
            
            self.usb_refresh_button.pack(side=TOP, padx=5)

        else:
            status_label.pack()


        if name == "배터리":
            self.battery_report_button = ttkb.Button(frame, text="리포트 확인하기",
                                                     bootstyle=SECONDARY,
                                                     command=self.view_battery_report
                                                     )
            self.battery_report_button.pack(side=TOP)

        frame.bind("<Button-1>", lambda e: self.start_test(name))
        icon_label.bind("<Button-1>", lambda e: self.start_test(name))

    # 기존 start_test 메서드 수정 
    def start_test(self, name):
        """카드 클릭 시 해당 테스트의 별도 GUI를 실행합니다."""
        status_label = self.test_status_labels.get(name)
        status_label.config(text=self.test_status_ing.get(name, ""), bootstyle="warning")
        if name == "키보드":
            self.open_keyboard_test()  # 기존 키보드 테스트 창
        elif name == "카메라":
            self.open_camera_test()  # 아래에서 새롭게 구현할 카메라 테스트
        elif name == "USB":
            self.start_usb_check()     # USB 테스트
        elif name == "충전":
            self.start_c_type_check()  # 충전 테스트 
        elif name == "배터리":
            self.generate_battery_report()  # 배터리 리포트 
        elif name == "QR코드":
            self.generate_qr_code()         # QR 코드 생성 
    

    # ----------------- 진행 상황 관련 메서드 -----------------

    def mark_test_complete(self, test_name):
        """특정 테스트 완료 후 상태 및 UI 갱신"""
        if test_name in self.test_done and not self.test_done[test_name]:
            self.test_done[test_name] = True
            status_label = self.test_status_labels[test_name]
            # 테스트 완료 시 파란색 계열("info")로 표시
            if test_name in ["배터리", "QR코드"]:
                 status_label.config(text="생성 완료", bootstyle="info")
            else:
                status_label.config(text="테스트 완료", bootstyle="info")

            if all(self.test_done.values()):
                messagebox.showinfo("모든 테스트 완료", "모든 테스트를 완료했습니다.\n수고하셨습니다!")

    # ----------------- 키보드 테스트 관련 메서드 -----------------
    def open_keyboard_test(self):
        def create_window():
            kb_window = ttkb.Toplevel(self)
            kb_window.title("키보드 테스트")
            kb_window.geometry("1200x500")
            # ... (키보드 테스트 창 구성 코드)
            return kb_window
        
        kb_window = self.open_test_window("키보드", create_window)
        
        if kb_window is None:
            return  # 이미 열려있으면 실행하지 않음
        
        # 반드시 kb_win을 부모로 하여 위젯을 생성
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

        # 모든 키와 각 키에 대응하는 위젯 저장
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

        # 아직 누르지 않은 키 목록
        self.keys_not_pressed = set(self.all_keys)

        # Raw Input 등록 (현재 창의 핸들 가져오기)
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

                # 첫 호출 시 데이터 크기 측정
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
                unregister_raw_input()
                self.failed_keys = list(self.keys_not_pressed)
                self.test_status_labels["키보드"].config(text="오류 발생", bootstyle="danger")
                self.failed_keys_button.config(state="normal")
            self.close_keyboard_window()

        # 윈도우 프로시저 콜백 타입 선언
        kb_window.protocol("WM_DELETE_WINDOW", on_close_keyboard_window)
        self._raw_input_wnd_proc = WNDPROC(raw_input_wnd_proc)

        # 기존 윈도우 프로시저 저장 후 새 프로시저 설정
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

        self.on_test_window_close("키보드")
    
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
    # def create_usb_port_labels(self, frame):
    #     """USB 포트 상태를 표시할 라벨 생성"""
    #     usb_port_frame = ttkb.Frame(frame)
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
        self.usb_refresh_button.config(state="normal", bootstyle="info")
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
                            self.usb_port[int(port_number)-1].config(text=f"{key}", bootstyle="info")
                            print(self.usb_ports)
                            print(self.usb_port)
                            print()
            if all(self.usb_ports.values()):
                self.usb_test_complete = True
                self.usb_refresh_button.config(state="disabled")
                self.mark_test_complete("USB")
                messagebox.showinfo("USB Test", "모든 USB 포트 테스트 완료!")
                self.test_status_labels["USB"].config(text="테스트 완료", bootstyle="info")

        except Exception as e:
            messagebox.showerror("USB Error", f"USB 포트 확인 중 오류 발생:\n{e}")

    # ----------------- 카메라 테스트 관련 메서드 ----------------
    def open_camera_test(self):
        """카메라(웹캠) 테스트 창을 열고 프레임을 표시합니다."""
        # 카메라 테스트가 이미 실행 중이면 추가 호출하지 않도록 합니다.
        if getattr(self, "camera_test_running", False):
            messagebox.showinfo("정보", "카메라 테스트가 이미 실행 중입니다.")
            return

        # 테스트 실행 상태 플래그 설정 및 버튼 비활성화 (중복 호출 방지)
        self.camera_test_running = True

        # 카메라 열기 (인덱스 0은 기본 카메라)
        self.cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
        if not self.cap.isOpened():
            messagebox.showerror("카메라 오류", "카메라를 열 수 없습니다. 장치를 확인해주세요.")
            self.camera_test_running = False
            return

        # 윈도우 이름 설정 및 명시적으로 윈도우 생성
        self.window_name = "Camera Test - X to exit"
        cv2.namedWindow(self.window_name)

        # Tkinter의 after() 메서드를 사용하여 주기적으로 프레임 업데이트를 시작합니다.
        self.update_camera_frame()

    def update_camera_frame(self):
        """Tkinter의 after()를 사용해 주기적으로 카메라 프레임을 업데이트합니다."""
        if not self.camera_test_running:
            return

        ret, frame = self.cap.read()
        if not ret:
            messagebox.showerror("카메라 오류", "카메라 프레임을 읽을 수 없습니다.")
            self.close_camera_test()
            return

        # 프레임을 윈도우에 표시합니다.
        cv2.imshow(self.window_name, frame)
        key = cv2.waitKey(1) & 0xFF
        # ESC 키를 누르거나 윈도우가 닫힌 경우 테스트 종료
        if key == 27 or cv2.getWindowProperty(self.window_name, cv2.WND_PROP_VISIBLE) < 1:
            self.close_camera_test()
            return
        
        # # 10밀리초 후에 다시 update_camera_frame() 호출 (GUI 이벤트 루프에 등록)
        self.after(10, self.update_camera_frame)

    def close_camera_test(self):
        """카메라 테스트 종료 후 자원 해제 및 상태 복원"""
        self.cap.release()
        cv2.destroyAllWindows()
        self.mark_test_complete("카메라")
        self.camera_test_running = False
        self.test_status_labels["카메라"].config(text="테스트 완료", bootstyle="info")

    # ----------------- 충전 테스트 관련 메서드 -----------------
    def create_c_type_port_labels(self, frame):
        """충전 포트 상태를 표시할 라벨 생성"""
        c_type_port_frame = ttkb.Frame(frame)
        c_type_port_frame.pack(fill=X, pady=3)
        for port_name in self.c_type_ports:
            label = ttkb.Label(c_type_port_frame, text=f"{port_name}: 연결 안됨",
                               width=20, bootstyle="inverse-light")
            label.pack(side=LEFT, padx=5)
            # self.c_type_port_labels[port_name] = label

    def start_c_type_check(self):
        """충전 테스트를 시작하고 포트 상태를 확인합니다."""
        self.c_type_ports = {"충전": False}
        # for port, lbl in self.c_type_port_labels.items():
        #     lbl.config(text=f"{port}: 연결 안됨", bootstyle="inverse-light")
        self.c_type_test_complete = False
        self.test_status_labels["충전"].config(text="테스트 중", bootstyle="warning")
        self.check_c_type_port()

    def check_c_type_port(self):
        """현재 배터리 충전 상태를 확인하여 포트 상태를 갱신합니다."""
        battery = psutil.sensors_battery()
        if battery is None:
            messagebox.showerror("충전 Error", "배터리 정보를 가져올 수 없습니다.")
            return

        if not battery.power_plugged:
            messagebox.showinfo(
                "충전 Test",
                "충전기가 연결되지 않았습니다.\n해당 포트에 충전기를 연결 후 다시 확인하세요."
            )
            return

        if not self.c_type_ports["충전"]:
            self.c_type_ports["충전"] = True
            # self.c_type_port_labels["충전"].config(
            #     text="전원 연결됨 (충전 중)",
            #     bootstyle="inverse-success"
            # )
        else:
            messagebox.showinfo("충전 Test", "충전 확인되었습니다.")

        if all(self.c_type_ports.values()):
            self.c_type_test_complete = True
            self.test_status_labels["충전"].config(text="테스트 완료", bootstyle="info")
            self.mark_test_complete("충전")

    # ----------------- 배터리 리포트 생성 관련 메서드 -----------------
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
            self.battery_report_button.config(bootstyle="info")
            self.mark_test_complete("배터리")
            self.test_status_labels["배터리"].config(text="생성 완료", bootstyle="info")

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

    # ----------------- qt코드드 생성 관련 메서드 -----------------
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
                "status": "pass" if self.test_done.get("충전") else "fail"
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
            self.mark_test_complete("QR코드")
            self.test_status_labels["QR코드"].config(text="생성 완료", bootstyle="info")

        except Exception as e:
            messagebox.showerror("QR 코드 생성 오류", f"QR 코드 생성 중 오류 발생:\n{e}")

if __name__ == "__main__":
    app = TestApp()
    app.mainloop()
