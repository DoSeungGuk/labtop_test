import ctypes
import ctypes.wintypes as wintypes
import sys

# 64비트 환경에 맞춰 WPARAM과 LPARAM 타입 정의
if sys.maxsize > 2**32:
    LPARAM_TYPE = ctypes.c_longlong  # 64비트
    WPARAM_TYPE = ctypes.c_ulonglong   # 64비트
else:
    LPARAM_TYPE = ctypes.c_long       # 32비트
    WPARAM_TYPE = ctypes.c_ulong      # 32비트

# wintypes에 ULONG_PTR가 없는 경우 직접 정의
if hasattr(wintypes, 'ULONG_PTR'):
    ULONG_PTR = wintypes.ULONG_PTR
else:
    if ctypes.sizeof(ctypes.c_void_p) == ctypes.sizeof(ctypes.c_ulong):
        ULONG_PTR = ctypes.c_ulong
    else:
        ULONG_PTR = ctypes.c_ulonglong

# Windows API 상수 정의
WH_KEYBOARD_LL = 13      # 저수준 키보드 후크 ID
WM_KEYDOWN = 0x0100      # 키 눌림 메시지
WM_KEYUP   = 0x0101      # 키 떼짐 메시지

# KBDLLHOOKSTRUCT 구조체 정의 (키보드 이벤트 정보를 담는 구조체)
class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("vkCode", wintypes.DWORD),      # 가상 키 코드
        ("scanCode", wintypes.DWORD),    # 스캔 코드
        ("flags", wintypes.DWORD),       # 이벤트 플래그
        ("time", wintypes.DWORD),        # 이벤트 발생 시간
        ("dwExtraInfo", ULONG_PTR)       # 추가 정보
    ]

# 후크 콜백 함수 타입 정의 (LowLevelKeyboardProc)
LowLevelKeyboardProc = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_int, WPARAM_TYPE, LPARAM_TYPE)

def hook_proc(nCode, wParam, lParam):
    """
    후크 콜백 함수: 키 이벤트 발생 시 호출되어 가상 키 코드를 출력
    """
    if nCode == 0:
        # lParam을 KBDLLHOOKSTRUCT 포인터로 캐스팅하여 이벤트 정보 추출
        kb_struct = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
        if wParam == WM_KEYDOWN:
            print("키 눌림 - 가상 키 코드(VK Code):", kb_struct.vkCode)
        elif wParam == WM_KEYUP:
            print("키 떼짐 - 가상 키 코드(VK Code):", kb_struct.vkCode)
    # CallNextHookEx 호출 시, 올바른 타입으로 인자 캐스팅
    return user32.CallNextHookEx(None, nCode, WPARAM_TYPE(wParam), LPARAM_TYPE(lParam))

# 콜백 함수 포인터 생성
hook_proc_ptr = LowLevelKeyboardProc(hook_proc)

# Windows API 라이브러리 로드
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# CallNextHookEx의 인자 타입 설정 (64비트에 맞게)
user32.CallNextHookEx.argtypes = [ctypes.c_void_p, ctypes.c_int, WPARAM_TYPE, LPARAM_TYPE]
user32.CallNextHookEx.restype = ctypes.c_long

# 전역 키보드 후크 설치 (hInstance 대신 None 사용)
hook_id = user32.SetWindowsHookExW(WH_KEYBOARD_LL, hook_proc_ptr, None, 0)
if not hook_id:
    error_code = kernel32.GetLastError()
    print("후크 설치에 실패했습니다. 오류 코드:", error_code)
    exit(1)

print("키보드 후크가 설치되었습니다. 키를 누르면 가상 키 코드가 출력됩니다.")

# 메시지 루프를 돌면서 후크가 계속 동작하도록 함
msg = wintypes.MSG()
while user32.GetMessageW( ctypes.byref(msg), None, 0, 0) != 0:
    user32.TranslateMessage(ctypes.byref(msg))
    user32.DispatchMessageW(ctypes.byref(msg))


