import re
import math

def scaling(window_handle):
    #To detect high DPI displays and avoid need to set Windows compatibility flags
    import os
    if os.name == "nt":
        from ctypes import windll, pointer, wintypes
        try:
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass  # this will fail on Windows Server and maybe early Windows
        DPI100pc = 96  # DPI 96 is 100% scaling
        DPI_type = 0  # MDT_EFFECTIVE_DPI = 0, MDT_ANGULAR_DPI = 1, MDT_RAW_DPI = 2
        winH = wintypes.HWND(window_handle)
        monitorhandle = windll.user32.MonitorFromWindow(winH, wintypes.DWORD(2))  # MONITOR_DEFAULTTONEAREST = 2
        X = wintypes.UINT()
        Y = wintypes.UINT()
        try:
            windll.shcore.GetDpiForMonitor(monitorhandle, DPI_type, pointer(X), pointer(Y))
            #return X.value, Y.value, (X.value + Y.value) / (2 * DPI100pc)
            s=math.sqrt(pow(X.value/96,2) + pow(Y.value/96,2))
            return str(s)
        except Exception:
            return  1  # Assume standard Windows DPI & scaling
    else:
        return 1  # What to do for other OSs?
