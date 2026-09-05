"""
Utilities for robustly connecting to running Microsoft Word instances and documents.
"""

import os
import re
import subprocess
import ctypes
from ctypes import wintypes
from typing import Tuple, Optional, Any

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


def _get_word_from_accessible_window() -> Tuple[Optional[Any], Optional[Any]]:
    """Query visible Word document panes (_WwG in OpusApp) via IAccessible."""
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

    wwg_hwnds = []

    def enum_child(hwnd, _):
        cls = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, cls, 256)
        if cls.value == "_WwG":
            wwg_hwnds.append(hwnd)
        return True

    WNDENUMCHILD = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)

    def enum_top(hwnd, _):
        cls = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, cls, 256)
        if cls.value == "OpusApp":
            user32.EnumChildWindows(hwnd, WNDENUMCHILD(enum_child), 0)
        return True

    WNDENUMTOP = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    user32.EnumWindows(WNDENUMTOP(enum_top), 0)

    for hwnd in wwg_hwnds:
        ptr = ctypes.c_void_p()
        res = AccessibleObjectFromWindow(
            hwnd, OBJID_NATIVEOM, ctypes.byref(IID_IDispatch), ctypes.byref(ptr)
        )
        if res == 0 and ptr.value:
            try:
                disp = win32com.client.Dispatch(ptr.value)
                return disp.Application, disp.Document
            except Exception:
                pass
    return None, None


def _get_word_from_processes() -> Tuple[Optional[Any], Optional[Any]]:
    """Inspect running WINWORD.EXE command lines to find open document paths and bind via GetObject."""
    try:
        wmi = win32com.client.GetObject("winmgmts:")
        for p in wmi.InstancesOf("Win32_Process"):
            if not p.Name or "winword" not in p.Name.lower():
                continue
            cmd = p.CommandLine or ""
            if "-Embedding" in cmd:
                continue
            matches = re.findall(r'["\']?([^"\']+\.docx?m?)["\']?', cmd, re.IGNORECASE)
            for path in matches:
                path = path.strip('"\'')
                if os.path.exists(path):
                    try:
                        doc = win32com.client.GetObject(path)
                        if doc:
                            return doc.Application, doc
                    except Exception:
                        pass
    except Exception:
        pass
    return None, None


def get_active_word_and_doc() -> Tuple[Any, Any]:
    """
    Robustly acquires the active Microsoft Word application and active document.

    Handles:
    1. Direct GetActiveObject("Word.Application")
    2. Lingering ghost -Embedding instances (invisible with 0 docs)
    3. Document windows via IAccessible (AccessibleObjectFromWindow)
    4. Running WINWORD.EXE processes with open document files (via GetObject)
    5. ProtectedViewWindows if a document is opened in Protected View

    Returns: (word_app, active_doc)
    Raises: Exception if Word is not running or no document is open.
    """
    pythoncom.CoInitialize()

    # 1. Try GetActiveObject
    ghost_word = None
    try:
        word = win32com.client.GetActiveObject("Word.Application")
        if word.Documents.Count > 0:
            try:
                return word, word.ActiveDocument
            except Exception:
                return word, word.Documents(1)
        if (
            hasattr(word, "ProtectedViewWindows")
            and word.ProtectedViewWindows.Count > 0
        ):
            return word, word.ProtectedViewWindows(1).Document
        if not word.Visible and word.Documents.Count == 0:
            ghost_word = word
    except Exception:
        pass

    # 2. Try window accessibility (_WwG document window in OpusApp)
    word, doc = _get_word_from_accessible_window()
    if word and doc:
        if ghost_word:
            try:
                ghost_word.Quit()
            except Exception:
                pass
        return word, doc

    # 3. Try process command line GetObject (e.g. Word launched with /n <file>)
    word, doc = _get_word_from_processes()
    if word and doc:
        if ghost_word:
            try:
                ghost_word.Quit()
            except Exception:
                pass
        return word, doc

    if ghost_word:
        try:
            ghost_word.Quit()
        except Exception:
            pass

    raise Exception(
        "No active document found in Word. Please ensure Word is open with a document."
    )


def get_active_word_app() -> Any:
    """Acquires Word.Application instance, or starts one if not running."""
    pythoncom.CoInitialize()
    try:
        word, _ = get_active_word_and_doc()
        return word
    except Exception:
        pass

    try:
        return win32com.client.GetActiveObject("Word.Application")
    except Exception:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True
        return word
