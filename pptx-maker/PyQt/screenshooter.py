import win32gui, win32con, pyautogui, os

def take_screenshot():
    try:
        a, winlist = [], []
        def enum_cb(hwnd, _):
            winlist.append((hwnd, win32gui.GetWindowText(hwnd)))
        win32gui.EnumWindows(enum_cb, a)
        window = [(hwnd, title) for hwnd, title in winlist if 'materialise' in title.lower()]

        hwnd = window[0][0]

        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        win32gui.BringWindowToTop(hwnd)
        win32gui.SetForegroundWindow(hwnd)
        pos = win32gui.GetWindowRect(hwnd)
        bbox = (pos[0]+200, pos[1]+185, pos[2]-975, pos[3]-400)

        pyautogui.press('num7')
        pyautogui.keyDown('alt')
        pyautogui.press('z')
        pyautogui.keyUp('alt')
        
        p = pyautogui.screenshot(region=bbox)
        p.save(f'{os.path.join(os.path.dirname(os.path.abspath(__file__)), "Temp/front.png")}')

        pyautogui.press('num2')
        pyautogui.keyDown('alt')
        pyautogui.press('z')
        pyautogui.keyUp('alt')

        p = pyautogui.screenshot(region=bbox)
        p.rotate(180).save(f'{os.path.join(os.path.dirname(os.path.abspath(__file__)), "Temp/bottom.png")}')

        bbox = (pos[0]+325, pos[1]+185, pos[2]-1100, pos[3]-400)

        pyautogui.press('num6')
        pyautogui.keyDown('alt')
        pyautogui.press('z')
        pyautogui.keyUp('alt') 

        p = pyautogui.screenshot(region=bbox)
        p.save(f'{os.path.join(os.path.dirname(os.path.abspath(__file__)), "Temp/left.png")}')

        pyautogui.press('num4')
        pyautogui.keyDown('alt')
        pyautogui.press('z')
        pyautogui.keyUp('alt')

        bbox = (pos[0]+200, pos[1]+185, pos[2]-1100, pos[3]-400)

        p = pyautogui.screenshot(region=bbox)
        p.save(f'{os.path.join(os.path.dirname(os.path.abspath(__file__)), "Temp/right.png")}')

        return(True)
    
    except:
        return(False)

