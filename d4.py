import cv2
import numpy as np
import pyautogui
import time
import win32api
import win32con
import win32gui
import random
import os
from PIL import ImageGrab

def find_image_on_screen(image_path, window_rect, threshold=0.8):
    left, top, right, bottom = window_rect

    # Capture the specified window area using Pillow
    img = ImageGrab.grab(bbox=(left, top, right, bottom), include_layered_windows=False, all_screens=True)
    img = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)

    # Save the captured image to the current directory
    screenshot_path = os.path.join(os.getcwd(), 'screenshot.png')
    cv2.imwrite(screenshot_path, img)
    print(f"Screenshot saved to {screenshot_path}")

    # Read the target image
    target_image = cv2.imread(image_path, cv2.IMREAD_UNCHANGED)

    if target_image.shape[2] == 4:  # If the target image has an alpha channel
        target_image = cv2.cvtColor(target_image, cv2.COLOR_BGRA2BGR)

    # Match the image
    result = cv2.matchTemplate(img, target_image, cv2.TM_CCOEFF_NORMED)
    
    # Get the location of the matched image
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

    # If the match score is above the threshold, return the center point of the target image
    if max_val >= threshold:
        target_height, target_width = target_image.shape[:2]
        center_point = (max_loc[0] + target_width // 2, max_loc[1] + target_height // 2)
        center_point_global = (center_point[0] + left, center_point[1] + top)
        return center_point_global
    else:
        return None

def click_at_position(x, y):
    # Add a random offset of 0-10 pixels
    offset_x = random.randint(0, 10)
    offset_y = random.randint(0, 10)
    x += offset_x
    y += offset_y

    # Get the current mouse position
    current_pos = win32api.GetCursorPos()

    # Move the mouse to the target position
    win32api.SetCursorPos((x, y))

    # Simulate mouse left button down and up
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)

    # Return the mouse to its original position
    win32api.SetCursorPos(current_pos)

def bring_window_to_front(window_title):
    hwnd = win32gui.FindWindow(None, window_title)
    if hwnd:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        print(f"Window '{window_title}' position: left={left}, top={top}, right={right}, bottom={bottom}")
        return (left, top, right, bottom)
    else:
        print(f"Window with title '{window_title}' not found.")
        return None

if __name__ == "__main__":
    window_title = "Diablo IV"
    image_path = "E:\\rip\\test.png"
    interval_seconds = 5  # Check every N seconds
    threshold = 0.8  # Set the match score threshold

    while True:
        try:
            hwnd = win32gui.FindWindow(None, window_title)
            if hwnd:
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                window_rect = (left, top, right, bottom)
                target_position = find_image_on_screen(image_path, window_rect, threshold)
                
                if target_position:
                    print(f"Found target image at {target_position}, bringing window to front and clicking...")
                    bring_window_to_front(window_title)  # Bring window to the front
                    time.sleep(1)  # Wait for window switch to complete
                    click_at_position(*target_position)
                else:
                    print("Target image not found.")
            else:
                print("Unable to find the window.")
        
            time.sleep(interval_seconds)
        except Exception as e:
            print(f"An error occurred: {e}. Restarting...")