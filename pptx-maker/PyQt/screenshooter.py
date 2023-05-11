import win32gui, win32con, pyautogui

def take_screenshot():
    try:
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
        bbox = (pos[0]+900, pos[1]+290, pos[2]-1410, pos[3]-460)

        pyautogui.press('num1') 
        pyautogui.press('r') 
        
        p = pyautogui.screenshot(region=bbox)
        p.save('Temp/front.png')

        pyautogui.press('num2')
        pyautogui.press('num2')
        pyautogui.press('num2')
        pyautogui.press('r') 

        p = pyautogui.screenshot(region=bbox)
        p.save('Temp/bottom45.png')

        pyautogui.press('num2')
        pyautogui.press('num2')
        pyautogui.press('num2')
        pyautogui.press('r') 
        
        p = pyautogui.screenshot(region=bbox)
        p.save('Temp/bottom.png')

        pyautogui.press('num3')
        pyautogui.press('r') 

        p = pyautogui.screenshot(region=bbox)
        p.save('Temp/left.png')

        pyautogui.press('num4')
        pyautogui.press('num4')
        pyautogui.press('num4')
        pyautogui.press('r') 

        p = pyautogui.screenshot(region=bbox)
        p.save('Temp/left45.png')

        pyautogui.press('num1')
        pyautogui.press('num4')
        pyautogui.press('num4')
        pyautogui.press('num4')
        pyautogui.press('r') 

        p = pyautogui.screenshot(region=bbox)
        p.save('Temp/right45.png')

        pyautogui.press('num4')
        pyautogui.press('num4')
        pyautogui.press('num4')
        pyautogui.press('r') 

        p = pyautogui.screenshot(region=bbox)
        p.save('Temp/right.png')

        pyautogui.press('num1')
        pyautogui.press('num8')
        pyautogui.press('num8')
        pyautogui.press('num8')
        pyautogui.press('r') 

        p = pyautogui.screenshot(region=bbox)
        p.save('Temp/front45.png')
        pyautogui.press('r') 

        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        return(True)
    
    except:
        return(False)
