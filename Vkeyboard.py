import win32api
import win32con
from PyQt5.QtCore import QMimeData, QUrl
from PyQt5.QtWidgets import QApplication


def CtrlAltZ():
    win32api.keybd_event(0x11, 0, 0, 0)
    win32api.keybd_event(0x12, 0, 0, 0)
    win32api.keybd_event(0x5a, 0, 0, 0)
    win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(0x12, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(0x5a, 0, win32con.KEYEVENTF_KEYUP, 0)


def CtrlAltW():
    win32api.keybd_event(0x11, 0, 0, 0)
    win32api.keybd_event(0x12, 0, 0, 0)
    win32api.keybd_event(0x57, 0, 0, 0)
    win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(0x12, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(0x57, 0, win32con.KEYEVENTF_KEYUP, 0)


def CtrlV():
    win32api.keybd_event(0x11, 0, 0, 0)
    win32api.keybd_event(0x56, 0, 0, 0)
    win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(0x56, 0, win32con.KEYEVENTF_KEYUP, 0)


def Enter():
    win32api.keybd_event(0x0D, 0, 0, 0)
    win32api.keybd_event(0x0D, 0, win32con.KEYEVENTF_KEYUP, 0)


def copyFile(file_path):
    data = QMimeData()
    url = QUrl.fromLocalFile(file_path)
    data.setUrls([url])
    QApplication.clipboard().setMimeData(data)


def setClipText(text):
    clip = QApplication.clipboard()
    clip.setText(text)
