import win32gui

def enumHandler(hwnd, lParam):
    if win32gui.IsWindowVisible(hwnd):
        print 'Here is the window text:', win32gui.GetWindowText(hwnd)
        if 'Skype for Business' in win32gui.GetWindowText(hwnd):
            win32gui.SetForegroundWindow(hwnd)

win32gui.EnumWindows(enumHandler, None)
