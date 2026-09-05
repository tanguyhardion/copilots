"""
Utilities for robustly connecting to running Microsoft Excel instances and active workbooks via pywin32.
"""

import os
import re
import ctypes
from ctypes import wintypes
from typing import Tuple, Optional, Any, List

import pythoncom
import win32com.client


class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", wintypes.DWORD),
        ("Data2", wintypes.WORD),
        ("Data3", wintypes.WORD),
        ("Data4", wintypes.BYTE * 8),
    ]


IID_IDispatch = GUID(
    0x00020400,
    0x0000,
    0x0000,
    (wintypes.BYTE * 8)(0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46),
)
OBJID_NATIVEOM = 0xFFFFFFF0  # -16


def col_index_to_letter(col_idx: int) -> str:
    """Convert 1-based column index to Excel column letter (e.g. 1 -> A, 27 -> AA)."""
    result = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def col_letter_to_index(col_letter: str) -> int:
    """Convert Excel column letter to 1-based index (e.g. A -> 1, AA -> 27)."""
    idx = 0
    for char in col_letter.upper():
        if "A" <= char <= "Z":
            idx = idx * 26 + (ord(char) - ord("A") + 1)
    return idx


def hex_to_bgr_int(hex_color: str) -> int:
    """Convert hex color string (e.g. '1F4E78' or '#1F4E78') to BGR integer for Excel COM."""
    clean = hex_color.lstrip("#").strip()
    if len(clean) == 3:
        clean = "".join([c * 2 for c in clean])
    if len(clean) != 6:
        clean = "000000"
    r = int(clean[0:2], 16)
    g = int(clean[2:4], 16)
    b = int(clean[4:6], 16)
    return (b << 16) | (g << 8) | r


def _get_excel_from_accessible_window() -> Tuple[Optional[Any], Optional[Any]]:
    """Query visible Excel workbook panes (EXCEL7 window inside XLMAIN) via IAccessible."""
    user32 = ctypes.windll.user32
    oleacc = ctypes.windll.oleacc

    AccessibleObjectFromWindow = oleacc.AccessibleObjectFromWindow
    AccessibleObjectFromWindow.argtypes = [
        wintypes.HWND,
        wintypes.DWORD,
        ctypes.POINTER(GUID),
        ctypes.POINTER(ctypes.c_void_p),
    ]
    AccessibleObjectFromWindow.restype = ctypes.c_long

    excel7_hwnds: List[int] = []

    def enum_child(hwnd, _):
        cls = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, cls, 256)
        if cls.value == "EXCEL7":
            excel7_hwnds.append(hwnd)
        return True

    WNDENUMCHILD = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)

    def enum_top(hwnd, _):
        cls = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, cls, 256)
        if cls.value == "XLMAIN":
            user32.EnumChildWindows(hwnd, WNDENUMCHILD(enum_child), 0)
        return True

    WNDENUMTOP = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    user32.EnumWindows(WNDENUMTOP(enum_top), 0)

    for hwnd in excel7_hwnds:
        ptr = ctypes.c_void_p()
        res = AccessibleObjectFromWindow(
            hwnd, OBJID_NATIVEOM, ctypes.byref(IID_IDispatch), ctypes.byref(ptr)
        )
        if res == 0 and ptr.value:
            try:
                disp = win32com.client.Dispatch(ptr.value)
                # disp is a Window object in Excel
                app = disp.Application
                wb = app.ActiveWorkbook
                return app, wb
            except Exception:
                pass
    return None, None


def _get_excel_from_processes() -> Tuple[Optional[Any], Optional[Any]]:
    """Inspect running EXCEL.EXE processes to find open workbook files via GetObject."""
    try:
        wmi = win32com.client.GetObject("winmgmts:")
        for p in wmi.InstancesOf("Win32_Process"):
            if not p.Name or "excel" not in p.Name.lower():
                continue
            cmd = p.CommandLine or ""
            if "-Embedding" in cmd:
                continue
            matches = re.findall(r'["\']?([^"\']+\.xls[xmb]?)["\']?', cmd, re.IGNORECASE)
            for path in matches:
                path = path.strip('"\'')
                if os.path.exists(path):
                    try:
                        wb = win32com.client.GetObject(path)
                        if wb:
                            return wb.Application, wb
                    except Exception:
                        pass
    except Exception:
        pass
    return None, None


def get_active_excel_and_wb() -> Tuple[Any, Any]:
    """
    Robustly acquires the active Microsoft Excel application and active workbook.

    Handles:
    1. Direct GetActiveObject("Excel.Application")
    2. Accessible window inspection (EXCEL7 pane inside XLMAIN)
    3. Process command-line GetObject inspection

    Returns:
        Tuple of (excel_app, active_workbook)
    Raises:
        Exception if Excel is not running or no workbook is currently open.
    """
    pythoncom.CoInitialize()

    # 1. Try GetActiveObject
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        if excel.Workbooks.Count > 0:
            try:
                return excel, excel.ActiveWorkbook
            except Exception:
                return excel, excel.Workbooks(1)
    except Exception:
        pass

    # 2. Try window accessibility (EXCEL7 pane inside XLMAIN)
    excel, wb = _get_excel_from_accessible_window()
    if excel and wb:
        return excel, wb

    # 3. Try process command line GetObject
    excel, wb = _get_excel_from_processes()
    if excel and wb:
        return excel, wb

    raise Exception(
        "No active workbook found in Microsoft Excel. Please ensure Excel is open with a workbook."
    )


def get_active_excel_app() -> Any:
    """Acquires active Excel.Application instance, or starts one if not running."""
    pythoncom.CoInitialize()
    try:
        excel, _ = get_active_excel_and_wb()
        return excel
    except Exception:
        pass

    try:
        return win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        return excel
