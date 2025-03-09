import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
import tkinter as tk
from PIL import Image, ImageTk, ImageFont, ImageDraw  # 이미지 처리 라이브러리
import os
import random  # USB 연결 상태를 랜덤하게 테스트하기 위해 추가

class TestApp(ttkb.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("KkomDae diagnostics")
        self.geometry("925x675")  # 16:9 비율
        self.resizable(False, False)
        self._style = ttkb.Style()  # 스타일 객체 생성

        # 🔹 폰트 파일 직접 로드 (설치 불필요)
        self.samsung_bold_path = "SamsungSharpSans-Bold.ttf"  
        self.samsung_regular_path = "SamsungOne-400.ttf"
        self.notosans_path = "NotoSansKR-VariableFont_wght.ttf"

        # 🔹 Frame 스타일 설정
        self._style.configure("Blue.TFrame", background="#0078D7")   # 타이틀 배경 파란색
        self._style.configure("White.TFrame", background="white")   # 테스트 영역 배경 흰색

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
            "카메라": "카메라(웹캠)이 정상적으로 작동하는지 확인합니다.",
            "USB": "모든 USB 포트가 정상적으로 인식되는지 확인합니다.",
            "충전": "노트북의 충전이 정상적으로 되는지 확인합니다.",
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

        # 🔹 USB 포트 상태 (처음엔 모두 비연결 상태)
        self.usb_ports = {
            1: False,
            2: False,
            3: False
        }

        # 타이틀 영역 생성
        self.create_title_section()
        
        # 테스트 항목 UI 구성
        self.create_test_items()

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
            (800, 33),  # 높이 조정
            self.notosans_path, 18, (255, 255, 255, 255), align_left=True
        )
        subtitle_label1 = ttkb.Label(text_container, image=self.subtitle_img1, background="#0078D7", anchor="w")
        subtitle_label1.grid(row=1, column=0, sticky="w", pady=(0, 0))

        # 두 번째 서브타이틀 라인
        self.subtitle_img2 = self.create_text_image(
            "로고를 클릭하면 테스트 or 생성을 시작할 수 있습니다.",
            (800, 33),  # 높이 조정
            self.notosans_path, 18, (255, 255, 255, 255), align_left=True
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
        frame.grid(row=row, column=col, padx=20, pady=20)

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

        # USB 항목의 경우 상태 레이블은 보이지 않도록 처리
        if name == "USB":
            # 상태 레이블은 생성은 하지만, 화면에서 숨김
            status_label.pack_forget()
        else:
            status_label.pack()

        if name == "USB":
            self.usb_status_label = status_label  # USB 상태 라벨 저장
            self.usb_port_labels = []
            port_frame = ttkb.Frame(frame)
            port_frame.pack(pady=0)

            for port in range(1, 4):
                # 초기 상태: 미연결
                port_label = ttkb.Label(
                    port_frame,
                    text=f"포트 {port}",
                    font=("맑은 고딕", 12),
                    bootstyle="danger",
                    width=7,  # 여백 조절용
    
                )
                port_label.pack(side="left", padx=2, pady=0)
                self.usb_port_labels.append(port_label)
    
            print(self.usb_port_labels)
        frame.bind("<Button-1>", lambda e: self.start_test(name, status_label))
        icon_label.bind("<Button-1>", lambda e: self.start_test(name, status_label))


    def start_test(self, name, label):
        """ 테스트 실행 시 상태 변경 """
        if name == "USB":
            # USB 테스트의 경우, 상태 레이블에 테스트 중 상태 multiline 문자열 적용
            label.config(text=self.test_status_ing.get(name, ""), bootstyle="warning")
            # 2초 후에 USB 포트 상태 업데이트 함수 호출
            self.after(2000, self.check_usb_ports)
        else:
            label.config(text=self.test_status_ing.get(name, ""), bootstyle="warning")
            self.after(2000, lambda: label.config(text=self.test_status_done.get(name, ""), bootstyle="info"))

    # check_usb_ports 함수 수정
    def check_usb_ports(self):
        """ USB 포트 테스트 (랜덤하게 USB 연결을 시뮬레이션 후 각 포트 레이블 업데이트) """
        # 3개의 포트 상태를 랜덤으로 결정
        new_ports = random.choices([True, False], k=3)
        print(new_ports)  # 디버깅용 출력
        # 각 포트에 대해 업데이트
        for i in range(3):
            if new_ports[i]:
                self.usb_ports[i+1] = True
                self.usb_port_labels[i].config(text=f"포트 {i+1}", bootstyle="success")
            else:
                self.usb_ports[i+1] = False
                self.usb_port_labels[i].config(text=f"포트 {i+1}", bootstyle="danger")


if __name__ == "__main__":
    app = TestApp()
    app.mainloop()
