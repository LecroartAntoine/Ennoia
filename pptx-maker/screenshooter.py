import win32gui, win32con, pyautogui, time

def take_screenshot():
    a, winlist = [], []
    def enum_cb(hwnd, _):
        winlist.append((hwnd, win32gui.GetWindowText(hwnd)))
    win32gui.EnumWindows(enum_cb, a)
    window = [(hwnd, title) for hwnd, title in winlist if '3d slicer' in title.lower()]

    hwnd = window[0][0]

    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    win32gui.BringWindowToTop(hwnd)
    win32gui.SetForegroundWindow(hwnd)
    pos = win32gui.GetWindowRect(hwnd)
    bbox = (pos[0]+900, pos[1]+300, pos[2]-1410, pos[3]-450)

    pyautogui.press('r') 
    pyautogui.press('num1') 
    
    p = pyautogui.screenshot(region=bbox)
    p.save('Fichiers/Images/front.png')

    pyautogui.press('num2')
    pyautogui.press('num2')
    pyautogui.press('num2')

    p = pyautogui.screenshot(region=bbox)
    p.save('Fichiers/Images/bottom45.png')

    pyautogui.press('num2')
    pyautogui.press('num2')
    pyautogui.press('num2')
    
    p = pyautogui.screenshot(region=bbox)
    p.save('Fichiers/Images/bottom.png')

    pyautogui.press('num3')

    p = pyautogui.screenshot(region=bbox)
    p.save('Fichiers/Images/left.png')

    pyautogui.press('num4')
    pyautogui.press('num4')
    pyautogui.press('num4')

    p = pyautogui.screenshot(region=bbox)
    p.save('Fichiers/Images/left45.png')

    pyautogui.press('num1')
    pyautogui.press('num4')
    pyautogui.press('num4')
    pyautogui.press('num4')

    p = pyautogui.screenshot(region=bbox)
    p.save('Fichiers/Images/right45.png')

    pyautogui.press('num4')
    pyautogui.press('num4')
    pyautogui.press('num4')

    p = pyautogui.screenshot(region=bbox)
    p.save('Fichiers/Images/right.png')

    pyautogui.press('num1')
    pyautogui.press('num8')
    pyautogui.press('num8')
    pyautogui.press('num8')

    p = pyautogui.screenshot(region=bbox)
    p.save('Fichiers/Images/front45.png')
