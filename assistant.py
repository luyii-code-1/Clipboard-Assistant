import win32clipboard
import win32con
import time
import subprocess
import threading
import ctypes
from pynput.keyboard import Controller, Key
import psutil
import win32gui
import win32process
import win32api
from winrt.windows.ui.notifications import ToastNotificationManager, ToastNotification
from winrt.windows.data.xml.dom import XmlDocument
import traceback
from win10toast_click import ToastNotifier


keyboard = Controller()
localsend_path = r'F:\LocalSend_1.16.0_Green\localsend_app.exe'


def is_localsend_running():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and 'localsend' in proc.info['name'].lower():
            return True
    return False


def launch_localsend():
    #if not is_localsend_running():
        
    subprocess.Popen([localsend_path])


def get_localsend_hwnd():
    def callback(hwnd, hwnds):
        try:
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if "LocalSend" in title:
                    hwnds.append(hwnd)
        except Exception as e:
            print("EnumWindows 回调异常:", e)
        return True  # 必须返回True
    hwnds = []
    win32gui.EnumWindows(callback, hwnds)
    return hwnds[0] if hwnds else None

def bring_to_front():
    hwnd = get_localsend_hwnd()
    if hwnd:
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            # 尝试用SetForegroundWindow
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception as e:
                print(f"SetForegroundWindow失败: {e}")
                # 尝试用ctypes强制切前台
                try:
                    foreground_thread_id = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())[0]
                    target_thread_id = win32process.GetWindowThreadProcessId(hwnd)[0]
                    win32api.AttachThreadInput(target_thread_id, foreground_thread_id, True)
                    win32gui.SetForegroundWindow(hwnd)
                    win32api.AttachThreadInput(target_thread_id, foreground_thread_id, False)
                except Exception as e2:
                    print(f"AttachThreadInput方式也失败: {e2}")
        except Exception as e:
            print(f"激活窗口失败: {e}")
    else:
        print("未找到 LocalSend 窗口句柄")


def simulate_paste():
    time.sleep(2)  # 等待窗口激活
    keyboard.press(Key.ctrl)
    keyboard.press('v')
    keyboard.release('v')
    keyboard.release(Key.ctrl)


toaster = ToastNotifier()

def on_notification_click():
    print("通知被点击，启动 LocalSend 并激活窗口")
    launch_localsend()
    # 等待窗口出现，最多等3秒
    for _ in range(30):
        hwnd = get_localsend_hwnd()
        if hwnd:
            break
        time.sleep(0.1)
    bring_to_front()
    simulate_paste()  # 激活窗口后立即粘贴

def show_notification():
    toaster.show_toast(
        "LocalSend助手",
        "检测到剪贴板图片，点击此处发送到 LocalSend",
        duration=5,
        threaded=True,
        callback_on_click=on_notification_click  # 添加回调
    )

last_clipboard = None  # 提升为全局变量

def clipboard_listener():
    global last_clipboard
    while True:
        try:
            win32clipboard.OpenClipboard()
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_DIB):
                if last_clipboard != 'image':
                    last_clipboard = 'image'
                    print("检测到图片，准备通知")
                    show_notification()
            else:
                last_clipboard = None
        except Exception as e:
            print("剪贴板监听异常:", e)
            traceback.print_exc()
        finally:
            try:
                win32clipboard.CloseClipboard()
            except:
                pass
        time.sleep(0.5)

if __name__ == "__main__":
    print("剪贴板监听已启动...")
    t = threading.Thread(target=clipboard_listener, daemon=True)
    t.start()
    while True:
        time.sleep(10)