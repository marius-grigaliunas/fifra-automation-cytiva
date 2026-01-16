import pyautogui
import time

print("Move mouse to print button and press Ctrl+C")
time.sleep(3)
try:
    while True:
        x, y = pyautogui.position()
        print(f"\rX: {x} Y: {y}", end="")
        time.sleep(0.1)
except KeyboardInterrupt:
    print(f"\n\nPrint button at: X={x}, Y={y}")