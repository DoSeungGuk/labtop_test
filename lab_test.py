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

class LaptopTestApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("노트북 테스트 프로그램")
        self.geometry("460x600")

        # 미리 조사한 '내장' 장치(HWID) 화이트리스트 (예시)
        self.INTERNAL_HWIDS = [
            # 내장 키보드
            "ACPI\\VEN_MSFT&DEV_0001",
            "ACPI\\MSFT0001",
            "*MSFT0001",
            # 내장 터치패드
            "HID\\VEN_MSFT&DEV_0001&SUBSYS_SYNA0001&Col01",
            "HID\\MSFT0001&Col01",
            "HID\\*MSFT0001&Col01",
            "HID\\VID_06CB&UP:0001_U:0002",
        ]

        # 외부 HID를 일시 비활성화했다가, 복원할 때 사용하기 위해
        self.disabled_hwids = [] 
        
        # 1) 테스트 완료 여부 추적: 각 테스트 이름 -> True/False
        #    이 부분은 동일
        self.test_done = {
            "키보드": False,
            "디스플레이": False,
            "카메라": False,
            "마우스": False,
            "Wi-Fi": False,
            "USB": False,
            "배터리": False,
            "사운드": False
        }

        # 각 테스트의 상태 라벨을 저장할 딕셔너리
        # 예: self.test_status_labels["키보드"] = 해당 라벨 위젯
        self.test_status_labels = {}

        # 2) 메인 프레임 구성
        main_frame = ttk.Frame(self)
        main_frame.pack(padx=20, pady=20, fill="both", expand=True)

        title_label = ttk.Label(main_frame, text="노트북 기능 테스트", font=("Arial", 16))
        title_label.pack(pady=10)

        # 진행 상황 라벨 (예: "0/8 완료")
        self.progress_label = ttk.Label(main_frame, text=self.get_progress_text(), font=("Arial", 12))
        self.progress_label.pack(pady=5)

        # ============ 버튼 + 상태 라벨들 ============
        # 각 테스트마다 버튼과 "테스트 전/완료" 라벨을 한 줄에 배치
        # 1) 키보드
        kb_frame = ttk.Frame(main_frame)
        kb_frame.pack(fill="x", pady=3)

        kb_button = ttk.Button(kb_frame, text="키보드 테스트", command=self.open_keyboard_test)
        kb_button.pack(side="left")

        kb_status = ttk.Label(kb_frame, text="테스트 전", foreground="red")
        kb_status.pack(side="left", padx=10)
        self.test_status_labels["키보드"] = kb_status

        # 2) 디스플레이
        disp_frame = ttk.Frame(main_frame)
        disp_frame.pack(fill="x", pady=3)

        disp_button = ttk.Button(disp_frame, text="디스플레이 정보 확인", command=self.show_display_info)
        disp_button.pack(side="left")

        disp_status = ttk.Label(disp_frame, text="테스트 전", foreground="red")
        disp_status.pack(side="left", padx=10)
        self.test_status_labels["디스플레이"] = disp_status

        # 3) 카메라
        cam_frame = ttk.Frame(main_frame)
        cam_frame.pack(fill="x", pady=3)

        cam_button = ttk.Button(cam_frame, text="카메라(웹캠) 테스트", command=self.open_camera_test)
        cam_button.pack(side="left")

        cam_status = ttk.Label(cam_frame, text="테스트 전", foreground="red")
        cam_status.pack(side="left", padx=10)
        self.test_status_labels["카메라"] = cam_status

        # 4) 마우스
        mouse_frame = ttk.Frame(main_frame)
        mouse_frame.pack(fill="x", pady=3)

        mouse_button = ttk.Button(mouse_frame, text="마우스 패드 테스트", command=self.open_mouse_test)
        mouse_button.pack(side="left")

        mouse_status = ttk.Label(mouse_frame, text="테스트 전", foreground="red")
        mouse_status.pack(side="left", padx=10)
        self.test_status_labels["마우스"] = mouse_status

        # 5) Wi-Fi
        wifi_frame = ttk.Frame(main_frame)
        wifi_frame.pack(fill="x", pady=3)

        wifi_button = ttk.Button(wifi_frame, text="Wi-Fi 상태 확인", command=self.check_wifi_status)
        wifi_button.pack(side="left")

        wifi_status = ttk.Label(wifi_frame, text="테스트 전", foreground="red")
        wifi_status.pack(side="left", padx=10)
        self.test_status_labels["Wi-Fi"] = wifi_status

        # 6) USB
        usb_frame = ttk.Frame(main_frame)
        usb_frame.pack(fill="x", pady=3)

        usb_button = ttk.Button(usb_frame, text="USB 연결 확인", command=self.check_usb_devices)
        usb_button.pack(side="left")

        usb_status = ttk.Label(usb_frame, text="테스트 전", foreground="red")
        usb_status.pack(side="left", padx=10)
        self.test_status_labels["USB"] = usb_status

        # 7) 배터리
        battery_frame = ttk.Frame(main_frame)
        battery_frame.pack(fill="x", pady=3)

        battery_button = ttk.Button(battery_frame, text="충전기/배터리 상태 확인", command=self.check_battery_status)
        battery_button.pack(side="left")

        battery_status = ttk.Label(battery_frame, text="테스트 전", foreground="red")
        battery_status.pack(side="left", padx=10)
        self.test_status_labels["배터리"] = battery_status

        # 8) 사운드
        sound_frame = ttk.Frame(main_frame)
        sound_frame.pack(fill="x", pady=3)

        sound_button = ttk.Button(sound_frame, text="사운드 입출력(마이크,스피커) 테스트", command=self.sound_test)
        sound_button.pack(side="left")

        sound_status = ttk.Label(sound_frame, text="테스트 전", foreground="red")
        sound_status.pack(side="left", padx=10)
        self.test_status_labels["사운드"] = sound_status

    # ============ 진행 상황 텍스트 ============

    def get_progress_text(self):
        """
        ex) "3/8 완료" 형태의 문자열을 반환
        """
        done_count = sum(self.test_done.values())  # True인 것들 세기
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

            # 상태 라벨 업데이트
            status_label = self.test_status_labels[test_name]
            status_label.config(text="테스트 완료", foreground="blue")

            # 모든 테스트가 완료되었는지 확인
            if all(self.test_done.values()):
                # 완료 확인 표시창(팝업)
                messagebox.showinfo("모든 테스트 완료", "모든 테스트를 완료했습니다.\n수고하셨습니다!")

    # ============ 키보드 테스트 (예시) ============
    def open_keyboard_test(self):
        kb_window = tk.Toplevel(self)
        kb_window.title("키보드 테스트")
        kb_window.geometry("600x300")

        info_label = ttk.Label(kb_window, text="이 창에 포커스를 두고\n모든 키를 한 번씩 눌러보세요.\n완료 시 창이 닫힙니다.")
        info_label.pack(pady=5)

        keyboard_layout = [
            ["1","2","3","4","5","6","7","8","9","0"],
            ["Q","W","E","R","T","Y","U","I","O","P"],
            ["A","S","D","F","G","H","J","K","L","ENTER"],
            ["Z","X","C","V","B","N","M","SPACE"]
        ]

        self.all_keys = set()
        self.key_widgets = {}
        self.keys_not_pressed = set()

        frame = ttk.Frame(kb_window)
        frame.pack(pady=10)

        for row_keys in keyboard_layout:
            row_frame = ttk.Frame(frame)
            row_frame.pack(pady=5)
            for key in row_keys:
                key_upper = key.upper()
                self.all_keys.add(key_upper)
                btn = tk.Label(row_frame, text=key, width=5, borderwidth=1, relief="solid",
                               font=("Arial", 12), background="lightgray")
                btn.pack(side="left", padx=3)
                self.key_widgets[key_upper] = btn

        self.keys_not_pressed = set(self.all_keys)
        kb_window.bind("<KeyPress>", self.on_key_press_visual)
        self.kb_window_ref = kb_window  # 창 핸들 보관

    def on_key_press_visual(self, event):
        pressed = event.keysym.upper()
        if pressed == "RETURN":
            pressed = "ENTER"
        elif pressed == "SPACE":
            pressed = "SPACE"

        if pressed in self.keys_not_pressed:
            self.keys_not_pressed.remove(pressed)
            widget = self.key_widgets.get(pressed)
            if widget:
                widget.config(background="black", foreground="white")

            if not self.keys_not_pressed:
                messagebox.showinfo("키보드 테스트", "모든 키를 눌렀습니다! 테스트 통과!")
                if hasattr(self, 'kb_window_ref'):
                    self.kb_window_ref.destroy()
                self.mark_test_complete("키보드")
                event.widget.unbind("<KeyPress>")

    # ============ 디스플레이 정보 테스트 (예시) ============
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

    # ============ 카메라 테스트 (예시) ============
    def open_camera_test(self):
        cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
        if not cap.isOpened():
            messagebox.showerror("카메라 오류", "카메라를 열 수 없습니다. 장치를 확인해주세요.")
            self.mark_test_complete("카메라")
            return

        messagebox.showinfo("카메라 테스트", "카메라 창이 뜨면 영상이 보이는지 확인하세요.\nESC 키를 누르면 종료됩니다.")
        while True:
            ret, frame = cap.read()
            if not ret:
                messagebox.showerror("카메라 오류", "카메라 프레임을 읽을 수 없습니다.")
                break
            cv2.imshow("Camera Test - ESC to exit", frame)
            key = cv2.waitKey(1) & 0xFF
            if key == 27:  # ESC
                break

        cap.release()
        cv2.destroyAllWindows()
        self.mark_test_complete("카메라")

    # ============ 마우스 테스트 (개선 예시) ============
    def open_mouse_test(self):
        self.mouse_window = tk.Toplevel(self)
        self.mouse_window.title("마우스 패드 테스트")
        self.mouse_window.geometry("640x360")
        self.mouse_window.resizable(False, False)

        info = ttk.Label(self.mouse_window, text=(
            "화면이 16x9 영역으로 나누어집니다.\n"
            "모든 구역을 커서로 방문하세요.\n"
            "움직임을 멈추면 커서가 중앙으로 돌아옵니다."
        ))
        info.pack()

        self.canvas = tk.Canvas(self.mouse_window, width=640, height=360, bg="white")
        self.canvas.pack()

        self.cols = 16
        self.rows = 9
        self.cell_w = 640 // self.cols
        self.cell_h = 360 // self.rows

        self.visited_cells = set()
        self.total_cells = self.cols * self.rows

        self.last_motion_time = time.time()
        self.mouse_stop_job = None

        self.canvas.bind("<Motion>", self.on_mouse_move)
        self.mouse_window.protocol("WM_DELETE_WINDOW", self.on_mouse_close)

        self.center_mouse()  # 시작 시 중앙으로

    def center_mouse(self):
        rootx = self.mouse_window.winfo_rootx()
        rooty = self.mouse_window.winfo_rooty()
        center_x = rootx + 640 // 2
        center_y = rooty + 360 // 2
        win32api.SetCursorPos((center_x, center_y))

    def on_mouse_move(self, event):
        cell_x = event.x // self.cell_w
        cell_y = event.y // self.cell_h
        if 0 <= cell_x < self.cols and 0 <= cell_y < self.rows:
            self.visited_cells.add((cell_x, cell_y))

        if len(self.visited_cells) == self.total_cells:
            messagebox.showinfo("마우스 패드 테스트", "모든 구역을 방문했습니다! 테스트 통과!")
            self.mouse_window.destroy()
            self.mark_test_complete("마우스")
            return

        self.last_motion_time = time.time()
        if self.mouse_stop_job is not None:
            self.mouse_window.after_cancel(self.mouse_stop_job)
        self.mouse_stop_job = self.mouse_window.after(500, self.check_mouse_stop)

    def check_mouse_stop(self):
        now = time.time()
        if now - self.last_motion_time >= 0.5:
            self.center_mouse()

    def on_mouse_close(self):
        if self.mouse_stop_job:
            self.mouse_window.after_cancel(self.mouse_stop_job)
        self.mouse_window.destroy()
        # 사용자가 중도에 창 닫아도, 일단 테스트 "완료" 처리할지, 미완료로 둘지는 선택
        # 여기서는 "일단 완료"로 처리 가정
        self.mark_test_complete("마우스")

    # ============ Wi-Fi 체크 (예시) ============
    def check_wifi_status(self):
        try:
            result = subprocess.check_output(["netsh", "wlan", "show", "interfaces"], encoding="cp866")
            messagebox.showinfo("Wi-Fi 상태", f"인터페이스 정보:\n\n{result}")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Wi-Fi 오류", f"Wi-Fi 정보를 가져오는 중 오류 발생:\n{e}")
        except Exception as e:
            messagebox.showerror("Wi-Fi 오류", f"Wi-Fi 정보를 가져오는 중 알 수 없는 오류:\n{e}")
        finally:
            self.mark_test_complete("Wi-Fi")

    # ============ USB 체크 (예시) ============
    def check_usb_devices(self):
        try:
            wmi_obj = win32com.client.GetObject("winmgmts:")
            pnp_entities = wmi_obj.InstancesOf("Win32_PnPEntity")
            usb_list = []
            for entity in pnp_entities:
                if "USB" in (entity.Name or "") or "USB" in (entity.DeviceID or ""):
                    usb_list.append(entity.Name)
            if usb_list:
                usb_info = "\n".join(usb_list)
                messagebox.showinfo("USB 장치 목록", f"다음 USB 장치가 감지되었습니다:\n\n{usb_info}")
            else:
                messagebox.showinfo("USB 장치 목록", "USB 장치를 찾을 수 없습니다.\n(연결 안 되었거나 인식 못 했을 수 있음)")
        except Exception as e:
            messagebox.showerror("USB 오류", f"USB 정보를 가져오는 중 오류 발생:\n{e}")
        finally:
            self.mark_test_complete("USB")

    # ============ 배터리 테스트 (예시) ============
    def check_battery_status(self):
        try:
            status = win32api.GetSystemPowerStatus()
            ac_line = status.ACLineStatus
            battery_flag = status.BatteryFlag
            battery_percent = status.BatteryLifePercent

            ac_text = "연결됨" if ac_line == 1 else "연결 안 됨"
            if battery_percent == 255:
                battery_text = "알 수 없음"
            else:
                battery_text = f"{battery_percent}%"

            if battery_flag & 8:
                charge_state_text = "충전 중"
            elif ac_line == 1:
                charge_state_text = "AC 연결(충전 혹은 충전 필요 없음)"
            else:
                charge_state_text = "배터리 사용 중(방전)"

            info_text = f"AC 어댑터 상태: {ac_text}\n배터리 잔량: {battery_text}\n충전 상태: {charge_state_text}"
            messagebox.showinfo("배터리/충전기 상태", info_text)
        except Exception as e:
            messagebox.showerror("배터리 오류", f"정보 가져오는 중 오류:\n{e}")
        finally:
            self.mark_test_complete("배터리")

    # ============ 사운드 테스트 (예시) ============
    def sound_test(self):
        try:
            RATE = 44100
            CHUNK = 1024
            CHANNELS = 2
            RECORD_SECONDS = 3

            p = pyaudio.PyAudio()

            stream_in = p.open(format=pyaudio.paInt16,
                               channels=CHANNELS,
                               rate=RATE,
                               input=True,
                               frames_per_buffer=CHUNK)
            messagebox.showinfo("사운드 테스트", "마이크로 3초간 녹음합니다. 말해보세요!")
            frames = []
            for _ in range(int(RATE / CHUNK * RECORD_SECONDS)):
                data = stream_in.read(CHUNK)
                frames.append(data)

            stream_in.stop_stream()
            stream_in.close()

            messagebox.showinfo("사운드 테스트", "녹음이 완료되었습니다. 이제 재생합니다.")

            stream_out = p.open(format=pyaudio.paInt16,
                                channels=CHANNELS,
                                rate=RATE,
                                output=True)
            for frame in frames:
                stream_out.write(frame)

            stream_out.stop_stream()
            stream_out.close()
            p.terminate()

            messagebox.showinfo("사운드 테스트", "재생이 완료되었습니다. 정상적으로 들렸다면 테스트 성공입니다.")
        except Exception as e:
            messagebox.showerror("사운드 테스트 오류", f"오류 발생:\n{e}")
        finally:
            self.mark_test_complete("사운드")


if __name__ == "__main__":
    app = LaptopTestApp()
    app.mainloop()
