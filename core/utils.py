import time, pyautogui, os
from PIL import Image

def wait_for_element(image, timeout=10, confidence=0.8):
    """等待图像出现在屏幕上"""
    start = time.time()
    while time.time() - start < timeout:
        try:
            if isinstance(image, Image.Image):
                if pyautogui.locateOnScreen(image, confidence=confidence):
                    return True
            elif isinstance(image, str) and os.path.exists(image):
                if pyautogui.locateOnScreen(image, confidence=confidence):
                    return True
        except:
            pass
        time.sleep(0.5)
    return False

def nowstr():
    from datetime import datetime
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
