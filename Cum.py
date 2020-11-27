import ctypes
from ctypes import wintypes
import time

user32 = ctypes.WinDLL('user32', use_last_error=True)

INPUT_MOUSE    = 0
INPUT_KEYBOARD = 1
INPUT_HARDWARE = 2

KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP       = 0x0002
KEYEVENTF_UNICODE     = 0x0004
KEYEVENTF_SCANCODE    = 0x0008

MAPVK_VK_TO_VSC = 0

# msdn.microsoft.com/en-us/library/dd375731
VK_TAB  = 0x09
VK_MENU = 0x12
windowskey = 0x5B


# C struct definitions

wintypes.ULONG_PTR = wintypes.WPARAM

class MOUSEINPUT(ctypes.Structure):
    _fields_ = (("dx",          wintypes.LONG),
                ("dy",          wintypes.LONG),
                ("mouseData",   wintypes.DWORD),
                ("dwFlags",     wintypes.DWORD),
                ("time",        wintypes.DWORD),
                ("dwExtraInfo", wintypes.ULONG_PTR))

class KEYBDINPUT(ctypes.Structure):
    _fields_ = (("wVk",         wintypes.WORD),
                ("wScan",       wintypes.WORD),
                ("dwFlags",     wintypes.DWORD),
                ("time",        wintypes.DWORD),
                ("dwExtraInfo", wintypes.ULONG_PTR))

    def __init__(self, *args, **kwds):
        super(KEYBDINPUT, self).__init__(*args, **kwds)
        # some programs use the scan code even if KEYEVENTF_SCANCODE
        # isn't set in dwFflags, so attempt to map the correct code.
        if not self.dwFlags & KEYEVENTF_UNICODE:
            self.wScan = user32.MapVirtualKeyExW(self.wVk,
                                                 MAPVK_VK_TO_VSC, 0)

class HARDWAREINPUT(ctypes.Structure):
    _fields_ = (("uMsg",    wintypes.DWORD),
                ("wParamL", wintypes.WORD),
                ("wParamH", wintypes.WORD))

class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = (("ki", KEYBDINPUT),
                    ("mi", MOUSEINPUT),
                    ("hi", HARDWAREINPUT))
    _anonymous_ = ("_input",)
    _fields_ = (("type",   wintypes.DWORD),
                ("_input", _INPUT))

LPINPUT = ctypes.POINTER(INPUT)

def _check_count(result, func, args):
    if result == 0:
        raise ctypes.WinError(ctypes.get_last_error())
    return args

user32.SendInput.errcheck = _check_count
user32.SendInput.argtypes = (wintypes.UINT, # nInputs
                             LPINPUT,       # pInputs
                             ctypes.c_int)  # cbSize

# Functions

def PressKey(hexKeyCode):
    x = INPUT(type=INPUT_KEYBOARD,
              ki=KEYBDINPUT(wVk=hexKeyCode))
    user32.SendInput(1, ctypes.byref(x), ctypes.sizeof(x))

def ReleaseKey(hexKeyCode):
    x = INPUT(type=INPUT_KEYBOARD,
              ki=KEYBDINPUT(wVk=hexKeyCode,
                            dwFlags=KEYEVENTF_KEYUP))
    user32.SendInput(1, ctypes.byref(x), ctypes.sizeof(x))

def AltTab():
    """Press Alt+Tab and hold Alt key for 2 seconds
    in order to see the overlay.
    """
    PressKey(VK_MENU)   # Alt
    PressKey(VK_TAB)    # Tab
    ReleaseKey(VK_TAB)  # Tab~
    time.sleep(2)
    ReleaseKey(VK_MENU) # Alt~

def cum():
    Spacebar = 0x20
    WKey = 0x57
    OKey = 0x4f
    RKey = 0x52
    DKey = 0x44
    PKey = 0x50
    AKey = 0x41
    Enter = 0x0D

    GreaterThan = 0xBE

    CKey = 0x43
    UKey = 0x55
    MKey = 0x4D
    
    IKey = 0x49

    SKey = 0x53
    NKey = 0x4E


    Control = 0x11
    Shift = 0x10


    PressKey(windowskey)
    ReleaseKey(windowskey)
    time.sleep(1)
    
    PressKey(WKey)
    ReleaseKey(WKey)
    
    PressKey(OKey)
    ReleaseKey(OKey)
    
    PressKey(RKey)
    ReleaseKey(RKey)
    
    PressKey(DKey)
    ReleaseKey(DKey)

    PressKey(PKey)
    ReleaseKey(PKey)

    PressKey(AKey)
    ReleaseKey(AKey)

    PressKey(DKey)
    ReleaseKey(DKey)

    time.sleep(1)
    PressKey(Enter)
    ReleaseKey(Enter)

    time.sleep(3)
    for _i in range(12):
        PressKey(Control)
        PressKey(Shift)
        PressKey(GreaterThan)
        ReleaseKey(Control)
        ReleaseKey(Shift)
        ReleaseKey(GreaterThan)

    PressKey(CKey)
    ReleaseKey(CKey)

    PressKey(UKey)
    ReleaseKey(UKey)

    PressKey(MKey)
    ReleaseKey(MKey)

    PressKey(Spacebar)
    ReleaseKey(Spacebar)

    PressKey(SKey)
    ReleaseKey(SKey)

    PressKey(AKey)
    ReleaseKey(AKey)

    PressKey(NKey)
    ReleaseKey(NKey)

    PressKey(SKey)
    ReleaseKey(SKey)

    time.sleep(2)

    PressKey(Control)
    PressKey(PKey)
    ReleaseKey(Control)
    ReleaseKey(PKey)

    time.sleep(1)

    PressKey(VK_MENU)
    PressKey(PKey)
    ReleaseKey(VK_MENU)
    ReleaseKey(PKey)

if __name__ == "__main__":
    cum()
