import psutil

battery = psutil.sensors_battery()

if battery.power_plugged:
    print("충전기가 연결됨 (충전 중)")
else:
    print("충전기 연결되지 않음 (배터리 방전 중)")

print(f"배터리 잔량: {battery.percent}%")

# import win32com.client
# import pythoncom

# def power_event_listener():
#     pythoncom.CoInitialize()
#     c = win32com.client.DispatchWithEvents(
#         "WbemScripting.SWbemLocator",
#         WMIEventHandler
#     )
#     wmi_service = c.ConnectServer(".", "root\\CIMV2")
#     # Win32_Battery 클래스의 상태 변경 이벤트를 구독
#     query = "SELECT * FROM __InstanceModificationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_Battery'"
#     event_watcher = wmi_service.ExecNotificationQuery(query)
#     print("배터리 상태 변경 추적 시작...")

#     while True:
#         battery_event = event_watcher.NextEvent()
#         new_state = battery_event.TargetInstance.BatteryStatus
#         # BatteryStatus: 1(방전), 2(AC 연결, 충전 중), 3(완충)
#         if new_state == 2:
#             print("충전기 연결됨 (충전 중)")
#         elif new_state == 1:
#             print("충전기 분리됨 (방전 중)")
#         elif new_state == 3:
#             print("충전 완료됨 (완충)")

# class WMIEventHandler:
#     pass

# if __name__ == "__main__":
#     power_event_listener()

# import win32con
# import win32gui
# import win32api

# class PowerBroadcastWindow:
#     def __init__(self):
#         message_map = {
#             win32con.WM_POWERBROADCAST: self.on_power_broadcast,
#             win32con.WM_DESTROY: self.on_destroy,
#         }

#         wc = win32gui.WNDCLASS()
#         wc.lpfnWndProc = message_map
#         wc.lpszClassName = "PowerBroadcastMonitor"
#         class_atom = win32gui.RegisterClass(wc)
#         self.hwnd = win32gui.CreateWindow(
#             class_atom, "Power Broadcast",
#             0, 0, 0, 0, 0, 0, 0, win32api.GetModuleHandle(None), None
#         )
#         win32gui.PumpMessages()

#     def on_power_broadcast(self, hwnd, msg, wparam, lparam):
#         if wparam == win32con.PBT_APMPOWERSTATUSCHANGE:
#             battery = psutil.sensors_battery()
#             if battery.power_plugged:
#                 print("충전기 연결됨")
#             else:
#                 print("충전기 분리됨")
#         return True

#     def on_destroy(self, hwnd, msg, wparam, lparam):
#         win32gui.PostQuitMessage(0)

# if __name__ == '__main__':
#     PowerBroadcastWindow()

# import time
# import win32api

# def check_power_status():
#     """
#     주기적으로 시스템 전원 상태를 확인하고, 충전기 연결 여부를 출력합니다.
#     """
#     while True:
#         try:
#             status = win32api.GetSystemPowerStatus()
#             ac_line = status.ACLineStatus
            
#             if ac_line == 1:
#                 print("충전기 연결됨")
#             else:
#                 print("충전기 연결되지 않음")

#         except Exception as e:
#             print(f"전원 상태 확인 중 오류 발생: {e}")

#         time.sleep(2)  # 2초마다 확인

# if __name__ == "__main__":
#     check_power_status()

