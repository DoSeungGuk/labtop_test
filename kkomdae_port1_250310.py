# ===============================
# 표준 라이브러리 및 외부 라이브러리 임포트
# ===============================
import sys
import os
import re
import subprocess
import logging
import ctypes
from ctypes import wintypes
from tkinter import messagebox

# 외부 라이브러리
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from PIL import Image, ImageTk, ImageFont, ImageDraw, ImageEnhance
import cv2
import win32com.client
import psutil
import qrcode

# ===============================
# Windows API 상수 및 구조체 정의
# ===============================
# 플랫폼에 따라 LRESULT, LONG_PTR 타입 결정
if ctypes.sizeof(ctypes.c_void_p) == 8:
    LRESULT = ctypes.c_longlong
    LONG_PTR = ctypes.c_longlong
else:
    LRESULT = ctypes.c_long
    LONG_PTR = ctypes.c_long

# Windows 메시지 상수
WM_NCDESTROY = 0x0082
WM_INPUT = 0x00FF
RID_INPUT = 0x10000003
GWL_WNDPROC = -4
RIDI_DEVICENAME = 0x20000007
RIM_TYPEKEYBOARD = 1
RIDEV_INPUTSINK = 0x00000100
RIDEV_NOLEGACY = 0x00000030  # legacy 메시지 차단
RIDEV_REMOVE = 0x00000001   # Raw Input 해제 플래그

RI_KEY_BREAK = 0x01
RI_KEY_E0 = 0x02

WM_DEVICECHANGE = 0x0219
DBT_DEVICEARRIVAL = 0x8000
DBT_DEVICEREMOVECOMPLETE = 0x8004
DBT_DEVTYP_DEVICEINTERFACE = 0x00000005

# user32 라이브러리 로드 및 함수 서명 지정
user32 = ctypes.windll.user32
user32.SetWindowLongPtrW.restype = LONG_PTR
user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, wintypes.INT, LONG_PTR]
user32.CallWindowProcW.restype = LRESULT
user32.CallWindowProcW.argtypes = [LONG_PTR, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]

# Raw Input 관련 구조체 정의
class RAWINPUTDEVICE(ctypes.Structure):
    _fields_ = [
        ("usUsagePage", ctypes.c_ushort),
        ("usUsage", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("hwndTarget", ctypes.c_void_p)
    ]

class RAWINPUTHEADER(ctypes.Structure):
    _fields_ = [
        ("dwType", ctypes.c_uint),
        ("dwSize", ctypes.c_uint),
        ("hDevice", ctypes.c_void_p),
        ("wParam", ctypes.c_ulong)
    ]

class RAWKEYBOARD(ctypes.Structure):
    _fields_ = [
        ("MakeCode", ctypes.c_ushort),
        ("Flags", ctypes.c_ushort),
        ("Reserved", ctypes.c_ushort),
        ("VKey", ctypes.c_ushort),
        ("Message", ctypes.c_uint),
        ("ExtraInformation", ctypes.c_ulong)
    ]

class RAWINPUT(ctypes.Structure):
    class _u(ctypes.Union):
        _fields_ = [("keyboard", RAWKEYBOARD)]
    _anonymous_ = ("u",)
    _fields_ = [
        ("header", RAWINPUTHEADER),
        ("u", _u)
    ]

# WNDPROC 타입 선언 (윈도우 프로시저 콜백)
WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

# 가상 키 코드 -> 문자열 매핑 딕셔너리
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
    0x60: "N 0",
    0x61: "N 1",
    0x62: "N 2",
    0x63: "N 3",
    0x64: "N 4",
    0x65: "N 5",
    0x66: "N 6",
    0x67: "N 7",
    0x68: "N 8",
    0x69: "N 9",
    0x6A: "N *",
    0x6B: "N +",
    0x6C: "N ENTER",
    0x6D: "N -",
    0x6E: "N .",
    0x6F: "N /",
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
    0x5B: "WIN",
    0x12: "ALT",
    0x15: "한/영",
    0x19: "한자",
}

# exe 빌드 시 파일 경를 찾기 위한 함수
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# ===============================
# Raw Input 관련 유틸리티 함수
# ===============================
def get_device_name(hDevice: int) -> str:
    """
    주어진 hDevice 핸들을 통해 장치 이름을 반환합니다.
    """
    size = ctypes.c_uint(0)
    if user32.GetRawInputDeviceInfoW(hDevice, RIDI_DEVICENAME, None, ctypes.byref(size)) == 0:
        buffer = ctypes.create_unicode_buffer(size.value)
        if user32.GetRawInputDeviceInfoW(hDevice, RIDI_DEVICENAME, buffer, ctypes.byref(size)) > 0:
            return buffer.value
    return None

def register_raw_input(hwnd: int) -> None:
    """
    지정된 윈도우 핸들에 대해 Raw Input을 등록합니다.
    legacy 메시지(WM_KEYDOWN 등)를 생성하지 않도록 설정합니다.
    """
    rid = RAWINPUTDEVICE()
    rid.usUsagePage = 0x01   # Generic Desktop Controls
    rid.usUsage = 0x06       # Keyboard
    rid.dwFlags = RIDEV_INPUTSINK | RIDEV_NOLEGACY
    rid.hwndTarget = hwnd
    if not user32.RegisterRawInputDevices(ctypes.byref(rid), 1, ctypes.sizeof(rid)):
        raise ctypes.WinError()

def unregister_raw_input() -> None:
    """
    등록된 Raw Input을 해제합니다.
    """
    rid = RAWINPUTDEVICE()
    rid.usUsagePage = 0x01
    rid.usUsage = 0x06
    rid.dwFlags = RIDEV_REMOVE
    rid.hwndTarget = 0
    if not user32.RegisterRawInputDevices(ctypes.byref(rid), 1, ctypes.sizeof(rid)):
        raise ctypes.WinError()

# ===============================
# TestApp 클래스 정의 (메인 GUI 애플리케이션)
# ===============================
class TestApp(ttkb.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("KkomDae Diagnostics")
        self.geometry("1200x950")
        self.resizable(False, False)
        self._style = ttkb.Style()

        # 변수 및 상태 초기화
        self._init_variables()

        # UI 구성
        self.create_title_section()
        self.create_test_items()

    def _init_variables(self) -> None:
        """
        내부 변수와 상태를 초기화합니다.
        """
        # 내부 키보드의 Raw Input device 화이트리스트
        self.INTERNAL_HWIDS = ["\\ACPI#MSF0001"]

        # 테스트 완료 여부 딕셔너리
        self.test_done = {
            "키보드": False,
            "카메라": False,
            "USB": False,
            "충전": False,
            "배터리": False,
            "QR코드": False
        }

        # 테스트 상태 문자열 설정
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
            "USB": "",
            "충전": "테스트 완료",
            "배터리": "생성 완료",
            "QR코드": "생성 완료"
        }
        # 테스트 상태 라벨 저장 딕셔너리
        self.test_status_labels = {}

        # 열려있는 테스트 창 관리 딕셔너리
        self.active_test_windows = {}

        # 폰트 경로 설정
        self.samsung_bold_path = resource_path("SamsungSharpSans-Bold.ttf")
        self.samsung_regular_path = resource_path("SamsungOne-400.ttf")
        self.samsung_700_path = resource_path("SamsungOne-700.ttf")
        self.notosans_path = resource_path("NotoSansKR-VariableFont_wght.ttf")

        # resource_path 함수를 이용해 이미지 파일의 경로를 동적으로 설정
        self.test_icons = {
            "키보드": resource_path("keyboard.png"),
            "카메라": resource_path("camera.png"),
            "USB": resource_path("usb.png"),
            "충전": resource_path("charging.png"),
            "배터리": resource_path("battery.png"),
            "QR코드": resource_path("qrcode.png")
        }

        self.test_descriptions = {
            "키보드": "키 입력이 정상적으로 작동하는지 확인합니다.",
            "카메라": "카메라(웹캠)가 정상적으로 작동하는지 확인합니다.",
            "USB": "모든 USB 포트가 정상적으로 인식되는지 확인합니다.",
            "충전": "노트북이 정상적으로 충전되는지 확인합니다.",
            "배터리": "배터리 리포트를 생성하여 성능을 확인합니다.",
            "QR코드": "테스트 결과를 QR 코드로 생성합니다."
        }

        # USB 관련 변수 초기화
        self.usb_ports = {"port1": False}
        self.usb_test_complete = False

        # 배터리 리포트 파일 경로 초기화
        self.report_path = None

        # 키보드 테스트 관련 변수
        self.failed_keys = []
        self.keys_not_pressed = set()
        self.all_keys = set()
        self.key_widgets = {}

    # -------------------------------
    # UI 구성 메서드들
    # -------------------------------
        # 🔹 Frame 스타일 설정
        self._style.configure("Blue.TFrame", background="#0078D7")   # 타이틀 배경 파란색
        self._style.configure("White.TFrame", background="white")   # 테스트 영역 배경 흰색

    def create_title_section(self) -> None:
        """
        상단 타이틀 영역을 생성합니다.
        """
        title_frame = ttkb.Frame(self, style="Blue.TFrame")
        title_frame.place(relx=0, rely=0, relwidth=1, relheight=0.27)

        # SSAFY 로고 이미지 삽입
        img_path = resource_path("ssafy_logo.png")
        image = Image.open(img_path).resize((80, 60), Image.LANCZOS)
        self.ssafy_logo = ImageTk.PhotoImage(image)
        img_label = ttkb.Label(title_frame, image=self.ssafy_logo, background="#0078D7", anchor="w")
        img_label.grid(row=0, column=0, padx=30, pady=(30, 10), sticky="w")

        # 타이틀 및 서브타이틀 텍스트 이미지 생성
        text_container = ttkb.Frame(title_frame, style="Blue.TFrame")
        text_container.grid(row=1, column=0, padx=20, sticky="w")

        self.title_img = self.create_text_image(
            "KkomDae Diagnostics", (800, 45), self.samsung_regular_path, 35, (255, 255, 255), align_left=True
        )
        title_label = ttkb.Label(text_container, image=self.title_img, background="#0078D7", anchor="w")
        title_label.grid(row=0, column=0, sticky="w")

        self.subtitle_img1 = self.create_text_image(
            "KkomDae Diagnostics로 노트북을 빠르고 꼼꼼하게 검사해보세요.",
            (800, 30), self.notosans_path, 17, (255, 255, 255, 255), align_left=True
        )
        subtitle_label1 = ttkb.Label(text_container, image=self.subtitle_img1, background="#0078D7", anchor="w")
        subtitle_label1.grid(row=1, column=0, sticky="w")

        self.subtitle_img2 = self.create_text_image(
            "로고를 클릭하면 테스트 or 생성을 시작할 수 있습니다.",
            (800, 30), self.notosans_path, 17, (255, 255, 255, 255), align_left=True
        )
        subtitle_label2 = ttkb.Label(text_container, image=self.subtitle_img2, background="#0078D7", anchor="w")
        subtitle_label2.grid(row=2, column=0, sticky="w")

    def create_text_image(self, text: str, size: tuple, font_path: str, font_size: int, color: tuple, align_left: bool = False) -> ImageTk.PhotoImage:
        """
        텍스트를 이미지로 변환하여 반환합니다.
        """
        img = Image.new("RGBA", size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        try:
            font = ImageFont.truetype(font_path, font_size)
        except IOError:
            print(f"⚠️ 폰트 '{font_path}'을 찾을 수 없습니다. 기본 폰트 사용")
            font = ImageFont.load_default()

        # 텍스트 위치 계산
        text_bbox = draw.textbbox((0, 0), text, font=font)
        text_x = 10 if align_left else (size[0] - text_bbox[2]) // 2
        text_y = (size[1] - font_size) // 2
        draw.text((text_x, text_y), text, font=font, fill=color, spacing=2, stroke_width=0.2)
        return ImageTk.PhotoImage(img)

    def create_test_items(self) -> None:
        """
        각 테스트 항목(키보드, 카메라, USB, 충전, 배터리, QR코드)의 UI를 생성합니다.
        2행 3열의 격자 배치로 구성합니다.
        """
        test_frame = ttkb.Frame(self, style="White.TFrame")
        test_frame.place(relx=0.1, rely=0.35, relwidth=0.8, relheight=0.6)
        self.tests = ["키보드", "카메라", "USB", "충전", "배터리", "QR코드"]

        # 2행으로 균등하게 분배 (각 행의 최소 높이 200)
        for row in range(2):
            test_frame.grid_rowconfigure(row, weight=1, minsize=200)
        # 3열로 균등하게 분배 (각 열의 최소 폭 250)
        for col in range(3):
            test_frame.grid_columnconfigure(col, weight=1, minsize=250) # minsize를 250으로 늘려줌

        # 각 테스트 항목을 2행 3열의 격자에 배치합니다.
        for idx, name in enumerate(self.tests):
            row = idx // 3  # 0,1,2 -> 0 / 3,4,5 -> 1
            col = idx % 3   # 0,3 -> 0 / 1,4 -> 1 / 2,5 -> 2
            self._create_test_item(test_frame, name, row, col)

    def _create_test_item(self, parent, name: str, row: int, col: int) -> None:
        """
        각 테스트 항목의 UI를 생성하고, 격자에 배치합니다.
        """
        # 컨테이너 프레임을 고정 크기로 생성 (크기는 원하는 대로 조정)
        frame = ttkb.Frame(parent, padding=10, width=250, height=200) # width를 250으로 수정
        frame.grid(row=row, column=col, padx=10, pady=10, sticky="nsew") # sticky 옵션 추가로 전체 격자 채우기
        frame.grid_propagate(False)  # 내부 위젯 크기에 의해 자동 조정되지 않도록 함

        # [Row 0] 아이콘 전용 프레임 (고정 크기, 최상단에 배치)
        icon_frame = ttkb.Frame(frame, width=55, height=55)
        icon_frame.grid(row=0, column=0,sticky= "n", pady=(0, 5), padx=10)
        icon_frame.grid_propagate(False)
        # 아이콘 이미지 로드 및 명암(채도) 낮추기
        icon_path = self.test_icons.get(name, "default.png")
        icon_img = Image.open(icon_path).resize((50, 50), Image.LANCZOS)
        enhancer = ImageEnhance.Color(icon_img)
        icon_img = enhancer.enhance(0)  # 채도를 0으로 낮춰 흑백 효과
        icon_photo = ImageTk.PhotoImage(icon_img)
        icon_label = ttkb.Label(icon_frame, image=icon_photo,justify='center')
        icon_label.image = icon_photo  # 이미지 참조 유지
        icon_label.pack(expand=True, fill="both") # grid 에서 pack으로 수정해줍니다.

        if name == "QR코드":
            icon_label.pack(expand=True, fill="both", padx=67) # qr 코드만 따로 padx를 적용해줍니다.
        else:
            icon_label.pack(expand=True, fill="both") # grid 에서 pack으로 수정해줍니다.
        if name == "배터리":
            icon_label.pack(expand=True, fill="both", padx=55) # qr 코드만 따로 padx를 적용해줍니다.
        else:
            icon_label.pack(expand=True, fill="both") # grid 에서 pack으로 수정해줍니다.


        # [Row 1] 테스트 이름 레이블
        name_label = ttkb.Label(frame, text=name, font=("맑은 고딕", 14, "bold"), foreground="#666666")
        name_label.grid(row=1, column=0, sticky="ew", pady=(5, 0)) # sticky="ew" 추가

        # [Row 2] 테스트 설명 레이블
        desc_label = ttkb.Label(
            frame,
            text=self.test_descriptions.get(name, ""),
            font=("맑은 고딕", 10),
            wraplength=180,
            # justify="center",
            foreground="#666666"
        )
        desc_label.grid(row=2, column=0, sticky="ew", pady=(5, 0)) # sticky="ew" 추가

        # [Row 3] 테스트 상태 레이블
        status_label = ttkb.Label(frame, text=self.test_status.get(name, ""), bootstyle="danger",
                                font=("맑은 고딕", 12))
        status_label.grid(row=3, column=0, sticky="ew", pady=(5, 0)) # sticky="ew" 추가
        self.test_status_labels[name] = status_label

        # [Row 4 및 Row 5] 추가 버튼 및 관련 UI 구성 (항목별로 다르게 처리)
        if name == "키보드":
            # 기존 변수명 유지: failed_keys_button
            self.failed_keys_button = ttkb.Button(
                frame,
                text="누르지 못한 키 보기",
                state="disabled",
                bootstyle=WARNING,
                command=self.show_failed_keys
            )
            self.failed_keys_button.grid(row=4, column=0, sticky="ew", pady=(5, 0)) # sticky="ew" 추가
        elif name == "USB":
            # USB의 경우 상태 레이블은 숨기고, 포트 상태와 새로고침 버튼을 별도의 행에 배치
            status_label.grid_forget()
            self.usb_status_label = status_label
            # USB 포트 상태 레이블들을 담을 프레임
            usb_ports_frame = ttkb.Frame(frame)
            usb_ports_frame.grid(row=3, column=0, sticky="ew") # sticky="ew" 추가
            usb_ports_frame.grid_columnconfigure(0, weight=1)
            # usb_ports_frame.grid_columnconfigure(1, weight=1)
            # usb_ports_frame.grid_columnconfigure(2, weight=1)
            self.usb_port = []
            for port in range(1, 2):
                port_frame = ttkb.Frame(usb_ports_frame)
                port_frame.grid(row=0, column=port-1, padx=2, sticky='ew')
                port_label = ttkb.Label(
                    port_frame,
                    text=f"port{port}",
                    font=("맑은 고딕", 12),
                    bootstyle="danger",
                    width=7
                )
                port_label.pack(expand=True, fill='x')
                self.usb_port.append(port_label)
            # 새로고침 버튼 (기존 변수명 유지: usb_refresh_button)
            self.usb_refresh_button = ttkb.Button(
                frame,
                text="새로고침",
                bootstyle=SECONDARY,
                command=self.refresh_usb_check,
                state="disabled"
            )
            self.usb_refresh_button.grid(row=4, column=0, sticky="ew", pady=(5, 0)) # sticky="ew" 추가
        elif name == "배터리":
            # 기존 변수명 유지: battery_report_button
            self.battery_report_button = ttkb.Button(
                frame,
                text="리포트 확인하기",
                bootstyle=SECONDARY,
                command=self.view_battery_report
            )
            self.battery_report_button.grid(row=4, column=0, sticky="ew", pady=(5, 0)) # sticky="ew" 추가
        # 항목 전체를 클릭하면 해당 테스트 시작 (아이콘 레이블 등에도 이벤트 바인딩)
        frame.bind("<Button-1>", lambda e: self.start_test(name))
        icon_label.bind("<Button-1>", lambda e: self.start_test(name))

    # -------------------------------
    # 테스트 시작 및 완료 처리 메서드
    # -------------------------------
    def start_test(self, name: str) -> None:
        """
        테스트 카드 클릭 시 해당 테스트 실행.
        """
        status_label = self.test_status_labels.get(name)
        status_label.config(text=self.test_status_ing.get(name, ""), bootstyle="warning")
        if name == "키보드":
            self.open_keyboard_test()
        elif name == "카메라":
            self.open_camera_test()
        elif name == "USB":
            self.start_usb_check()
        elif name == "충전":
            self.start_c_type_check()
        elif name == "배터리":
            self.generate_battery_report()
        elif name == "QR코드":
            self.generate_qr_code()

    def mark_test_complete(self, test_name: str) -> None:
        """
        특정 테스트 완료 후 상태 업데이트 및 모든 테스트 완료시 메시지 출력.
        """
        if test_name in self.test_done and not self.test_done[test_name]:
            self.test_done[test_name] = True
            status_label = self.test_status_labels[test_name]
            if test_name in ["배터리", "QR코드"]:
                status_label.config(text="생성 완료", bootstyle="info")
            else:
                status_label.config(text="테스트 완료", bootstyle="info")
            if all(self.test_done.values()):
                messagebox.showinfo("모든 테스트 완료", "모든 테스트를 완료했습니다.\n수고하셨습니다!")

    def open_test_window(self, test_name: str, create_window_func) -> ttkb.Toplevel:
        """
        이미 열려있는 테스트 창이 있는지 확인 후, 새 창을 생성합니다.
        """
        if test_name in self.active_test_windows:
            messagebox.showwarning("경고", f"{test_name} 테스트 창이 이미 열려 있습니다.")
            return
        window = create_window_func()
        self.active_test_windows[test_name] = window
        return window

    def on_test_window_close(self, test_name: str) -> None:
        """
        테스트 창 종료 시 관리 딕셔너리에서 제거합니다.
        """
        if test_name in self.active_test_windows:
            del self.active_test_windows[test_name]

    # -------------------------------
    # 키보드 테스트 관련 메서드
    # -------------------------------
    def open_keyboard_test(self) -> None:
        """
        키보드 테스트 창을 열어 Raw Input 이벤트를 처리합니다.
        """
        def create_window() -> ttkb.Toplevel:
            kb_window = ttkb.Toplevel(self)
            kb_window.title("키보드 테스트")
            kb_window.geometry("1200x500")
            # 키보드 테스트 창 구성
            info_label = ttkb.Label(kb_window, text="모든 키를 한 번씩 눌러보세요.\n완료 시 창이 닫힙니다.")
            info_label.pack(pady=5)
            return kb_window

        kb_window = self.open_test_window("키보드", create_window)
        if kb_window is None:
            return

        # 키보드 레이아웃 구성 (실제 키보드 레이아웃 반영)

        keyboard_layout = [
            # 첫 번째 행: ESC, F1 ~ F12, PRT, INS, DEL, N /, N *
            [("ESC", 5), ("F1", 5), ("F2", 5), ("F3", 5), ("F4", 5), ("F5", 5),
            ("F6", 5), ("F7", 5), ("F8", 5), ("F9", 5), ("F10", 5), ("F11", 5),
            ("F12", 5), ("PRT", 5), ("INS", 5), ("DEL", 4), ("N /", 4), ("N *", 4)],
            # 두 번째 행: `, 1 ~ 0, -, =, BACK, N -, N +, NUMLOCK  (총합 88)
            [("`", 5), ("1", 5), ("2", 5), ("3", 5), ("4", 5), ("5", 5),
            ("6", 5), ("7", 5), ("8", 5), ("9", 5), ("0", 5), ("-", 5),
            ("=", 5), ("BACK", 8), ("N -", 5), ("N +", 5), ("NUMLOCK", 5)],
            # 세 번째 행: TAB, Q ~ P, [, ], \, N 7, N 8, N 9 (총합 88)
            [("TAB", 8), ("Q", 5), ("W", 5), ("E", 5), ("R", 5), ("T", 5),
            ("Y", 5), ("U", 5), ("I", 5), ("O", 5), ("P", 5), ("[", 5),
            ("]", 5), ("\\", 5), ("N 7", 5), ("N 8", 5), ("N 9", 5)],
            # 네 번째 행: CAPS, A, S ~ L, ;, ', ENTER, N 4, N 5, N 6

            [("CAPS", 8), ("A", 7), ("S", 5), ("D", 5), ("F", 5), ("G", 5),
            ("H", 5), ("J", 5), ("K", 5), ("L", 5), (";", 5), ("'", 5),
            ("ENTER", 9), ("N 4", 5), ("N 5", 5), ("N 6", 5)],
            # 다섯 번째 행: LSHIFT, Z, X, C, V, B, N, M, ,, ., /, RSHIFT, N 1, N 2, N 3
            [("LSHIFT", 12), ("Z", 5), ("X", 5), ("C", 5), ("V", 5), ("B", 5),
            ("N", 5), ("M", 5), (",", 5), (".", 5), ("/", 6), ("RSHIFT", 12),
            ("N 1", 5), ("N 2", 5), ("N 3", 5)],
            # 여섯 번째 행: CTRL, (빈 키), WIN, ALT, SPACE, 한/영, 한자, LEFT, DOWN, UP, RIGHT, N 0, N ., N ENTER
            [("CTRL", 5), ("", 5), ("WIN", 5), ("ALT", 5), ("SPACE", 27), ("한/영", 5),
            ("한자", 5), ("LEFT", 5), ("DOWN", 5), ("UP", 5), ("RIGHT", 5), 
            ("N 0", 5), ("N .", 5), ("N ENTER", 5)]
        ]

        # 키보드 레이아웃 구성 (실제 키보드 배열과 유사)
        self.all_keys = set()
        self.key_widgets = {}
        for row_keys in keyboard_layout:
            row_frame = ttkb.Frame(kb_window)
            row_frame.pack(pady=5)
            for key, width in row_keys:
                if key == "":  # 빈 문자열이면 키 입력 대상에서 제외
                    spacer = ttkb.Label(row_frame, text="", width=width)
                    spacer.pack(side=LEFT, padx=3)
                    continue
                key_upper = key.upper()
                self.all_keys.add(key_upper)
                # 각 키의 너비를 튜플의 두 번째 요소로 지정
                btn = ttkb.Label(row_frame, text=key, width=width, bootstyle="inverse-light",
                                font=("맑은 고딕", 10, "bold"))
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
                                    key_sym = "N ENTER" if (raw.u.keyboard.Flags & RI_KEY_E0) else "ENTER"
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

        kb_window.protocol("WM_DELETE_WINDOW", on_close_keyboard_window)
        self._raw_input_wnd_proc = WNDPROC(raw_input_wnd_proc)
        cb_func_ptr = ctypes.cast(self._raw_input_wnd_proc, ctypes.c_void_p).value
        cb_func_ptr = LONG_PTR(cb_func_ptr)
        old_proc = user32.SetWindowLongPtrW(hwnd, GWL_WNDPROC, cb_func_ptr)
        self._kb_old_wnd_proc = old_proc
        self._kb_hwnd = hwnd
        self.kb_window_ref = kb_window


    def close_keyboard_window(self) -> None:
        """
        키보드 테스트 종료 시 Raw Input 프로시저 복원 및 창 닫기
        """
        if hasattr(self, '_kb_hwnd') and self._kb_hwnd and self._kb_old_wnd_proc is not None:
            user32.SetWindowLongPtrW(self._kb_hwnd, GWL_WNDPROC, self._kb_old_wnd_proc)
            self._kb_old_wnd_proc = None
        if hasattr(self, 'kb_window_ref'):
            self.kb_window_ref.destroy()
        self.on_test_window_close("키보드")

    def on_raw_key(self, key: str) -> None:
        """
        키 입력 이벤트 처리. 해당 키가 눌리면 상태 업데이트 후 모든 키 입력 시 테스트 완료.
        """
        if key in self.keys_not_pressed:
            self.keys_not_pressed.remove(key)
            widget = self.key_widgets.get(key)
            if widget:
                widget.config(bootstyle="inverse-dark")
            if not self.keys_not_pressed:
                unregister_raw_input()
                self.failed_keys_button.config(state="disabled")
                self.close_keyboard_window()
                self.mark_test_complete("키보드")

    def show_failed_keys(self) -> None:
        """
        누르지 않은 키 목록을 별도 창에 표시합니다.
        """
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

    # -------------------------------
    # USB 테스트 관련 메서드
    # -------------------------------
    def start_usb_check(self) -> None:
        """
        USB 테스트 초기화 후 상태 갱신 및 새로고침 버튼 활성화
        """
        self.usb_test_complete = False
        self.usb_refresh_button.config(state="normal", bootstyle="info")
        self.test_status_labels["USB"].config(text="테스트 중", bootstyle="warning")
        self.refresh_usb_check()

    def refresh_usb_check(self) -> None:
        """
        USB 연결 상태를 확인하여 UI 업데이트 후 모든 포트 연결시 테스트 완료 처리
        """
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
                            self.usb_port[int(port_number)-1].config(text=key, bootstyle="info")
            if all(self.usb_ports.values()):
                self.usb_test_complete = True
                self.usb_refresh_button.config(state="disabled")
                self.mark_test_complete("USB")
                messagebox.showinfo("USB Test", "모든 USB 포트 테스트 완료!")
                self.test_status_labels["USB"].config(text="테스트 완료", bootstyle="info")
        except Exception as e:
            messagebox.showerror("USB Error", f"USB 포트 확인 중 오류 발생:\n{e}")

    # -------------------------------
    # 카메라 테스트 관련 메서드
    # -------------------------------
    def open_camera_test(self) -> None:
        """
        카메라(웹캠) 테스트 창을 열어 프레임을 표시합니다.
        """
        if getattr(self, "camera_test_running", False):
            messagebox.showinfo("정보", "카메라 테스트가 이미 실행 중입니다.")
            return
        self.camera_test_running = True
        self.cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
        if not self.cap.isOpened():
            messagebox.showerror("카메라 오류", "카메라를 열 수 없습니다. 장치를 확인해주세요.")
            self.camera_test_running = False
            return
        self.window_name = "Camera Test - X to exit"
        cv2.namedWindow(self.window_name)
        self.update_camera_frame()

    def update_camera_frame(self) -> None:
        """
        Tkinter after()를 이용하여 주기적으로 카메라 프레임을 업데이트합니다.
        """
        if not self.camera_test_running:
            return
        ret, frame = self.cap.read()
        if not ret:
            messagebox.showerror("카메라 오류", "카메라 프레임을 읽을 수 없습니다.")
            self.close_camera_test()
            return
        cv2.imshow(self.window_name, frame)
        key = cv2.waitKey(1) & 0xFF
        if key == 27 or cv2.getWindowProperty(self.window_name, cv2.WND_PROP_VISIBLE) < 1:
            self.close_camera_test()
            return
        self.after(10, self.update_camera_frame)

    def close_camera_test(self) -> None:
        """
        카메라 테스트 종료 후 자원 해제 및 상태 복원.
        """
        self.cap.release()
        cv2.destroyAllWindows()
        self.mark_test_complete("카메라")
        self.camera_test_running = False
        self.test_status_labels["카메라"].config(text="테스트 완료", bootstyle="info")

    # -------------------------------
    # 충전 테스트 관련 메서드
    # -------------------------------
    def start_c_type_check(self) -> None:
        """
        충전 테스트를 시작하고 충전 포트 상태를 확인합니다.
        """
        self.c_type_ports = {"충전": False}
        self.c_type_test_complete = False
        self.test_status_labels["충전"].config(text="테스트 중", bootstyle="warning")
        self.check_c_type_port()

    def check_c_type_port(self) -> None:
        """
        배터리 충전 상태를 확인하여 포트 상태를 갱신합니다.
        """
        battery = psutil.sensors_battery()
        if battery is None:
            messagebox.showerror("충전 Error", "배터리 정보를 가져올 수 없습니다.")
            return
        if not battery.power_plugged:
            messagebox.showinfo("충전 Test", "충전기가 연결되지 않았습니다.\n해당 포트에 충전기를 연결 후 다시 확인하세요.")
            return
        if not self.c_type_ports["충전"]:
            self.c_type_ports["충전"] = True
        else:
            messagebox.showinfo("충전 Test", "충전 확인되었습니다.")
        if all(self.c_type_ports.values()):
            self.c_type_test_complete = True
            self.test_status_labels["충전"].config(text="테스트 완료", bootstyle="info")
            self.mark_test_complete("충전")

    # -------------------------------
    # 배터리 리포트 관련 메서드
    # -------------------------------
    def generate_battery_report(self) -> None:
        """
        powercfg 명령어를 통해 배터리 리포트를 생성합니다.
        """
        try:
            # 다운로드 폴더 경로를 가져옵니다.
            downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
            
            # 다운로드 폴더가 없는 경우, 생성합니다.
            if not os.path.exists(downloads_path):
                os.makedirs(downloads_path)

            self.report_path = os.path.join(downloads_path, "battery_report.html")
            subprocess.run(["powercfg", "/batteryreport", "/output", self.report_path],
                           check=True, capture_output=True, text=True)
            messagebox.showinfo("배터리 리포트", f"배터리 리포트가 생성되었습니다.\n파일 경로:\n{self.report_path}")
            self.battery_report_button.config(bootstyle="info")
            self.mark_test_complete("배터리")
            self.test_status_labels["배터리"].config(text="생성 완료", bootstyle="info")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("배터리 리포트 오류", f"명령 실행 중 오류 발생:\n{e.stderr}")
        except Exception as e:
            messagebox.showerror("배터리 리포트 오류", f"오류 발생:\n{e}")

    def view_battery_report(self) -> None:
        """
        생성된 배터리 리포트 파일을 엽니다.
        """
        if self.report_path and os.path.exists(self.report_path):
            try:
                os.startfile(self.report_path)
            except Exception as e:
                messagebox.showerror("리포트 확인 오류", f"리포트를 열 수 없습니다:\n{e}")
        else:
            messagebox.showwarning("리포트 없음", "아직 배터리 리포트가 생성되지 않았습니다.\n먼저 '배터리 리포트 생성' 버튼을 눌러주세요.")

    # -------------------------------
    # QR 코드 생성 관련 메서드
    # -------------------------------
    def generate_qr_code(self) -> None:
        """
        테스트 결과를 JSON 형식으로 구성 후 QR 코드를 생성하여 표시합니다.
        """
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

# ===============================
# 애플리케이션 실행
# ===============================
if __name__ == "__main__":
    app = TestApp()
    app.mainloop()
