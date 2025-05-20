import easyocr
import pyautogui
import time
from PIL import ImageGrab

def capture_and_ocr():
    print("5秒以内に画面を範囲選択してください（Windows + Shift + S）...")
    pyautogui.hotkey('winleft', 'shift', 's')
    time.sleep(5)

    img = ImageGrab.grabclipboard()
    if img is None:
        print("画像が取得できませんでした。")
        return

    img.save("screenshot.png")

    reader = easyocr.Reader(['ja'])
    result = reader.readtext("screenshot.png", detail=0)

    lines = [line.strip() for line in result if line.strip()]
    flow = '→'.join(lines)
    print("検出されたフロー:")
    print(flow)

if __name__ == "__main__":
    capture_and_ocr()
