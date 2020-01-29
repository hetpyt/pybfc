#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import win32gui
import win32con
from time import sleep
import platform
import glob
from os import path

# Config
SRC_DIR = "G:\\RecordFile"
DST_DIR = "F:\\RecordFile\\test"
MAX_TASKS = 3
##

WIN_VERSION = platform.win32_ver()[0]

FILE_EXTENTION = "264"

TARGET_MAIN_WND_TITLE = "H264 конвертировать AVI"
TARGET_DLG_WND_TITLE = "Новые задачи преобразования"
TARGET_NEWTASK_BTTN_TITLE = "Новая задача"
TARGET_ADD_BTTN_TITLE = "добавлять"
TARGET_BTTN_CLASS = "Button"
STD_DLG_OPEN_TITLE = "Открыть"
if WIN_VERSION == '7':
    STD_DLG_SAVEAS_TITLE = "Сохранить как"
else: # 10
    STD_DLG_SAVEAS_TITLE = "Сохранение"
STD_DLG_OPEN_BTN_TITLE = "&Открыть"
STD_DLG_SAVEAS_BTN_TITLE = "Со&хранить"

LVM_GETITEMCOUNT = 4100

_hSelectFile0 = 0
_hSelectFile1 = 0

def push_button(h):
    win32gui.SetForegroundWindow(h);
    win32gui.SendMessage(h, win32con.WM_ACTIVATE, 1, 0);
    win32gui.SendMessage(h, win32con.WM_ENABLE, 1, 0);
    win32gui.SendMessage(h, win32con.WM_SETFOCUS, 1, 0);
    win32gui.PostMessage(h, win32con.BM_CLICK, 0, 0);

def wait_for_wnd(title, timeout = 10):
    duration = 0
    step = 0.5
    while True:
        h = win32gui.FindWindow(None, title)
        if h != 0:
            return h
        sleep(step)
        duration += step
        if  duration >= timeout:
            raise Exception("waiting for window '{}' timeout".format(title))

def wait_for_wnd_closed(title, timeout = 10):
    duration = 0
    step = 0.5
    while True:
        h = win32gui.FindWindow(None, title)
        if h == 0:
            return
        sleep(step)
        duration += step
        if  duration >= timeout:
            raise Exception("waiting for window '{}' closed timeout".format(title))
            
def wait_for_wnd_closed_h(h, timeout = 10):
    duration = 0
    step = 0.5
    while win32gui.IsWindow(h):
        sleep(step)
        duration += step
        if  duration >= timeout:
            raise Exception("waiting for window handle '{}' closed timeout".format(h))
    
def EnumWndProc(h, param):
    global _hSelectFile0, _hSelectFile1
    sClassName= win32gui.GetClassName(h)
    if sClassName == TARGET_BTTN_CLASS:
        sTitle = win32gui.GetWindowText(h)
        if sTitle == "..":
            #print("found control handle {} of '{}' with title '{}'".format(h, sClassName, sTitle))
            if _hSelectFile0 == 0:
                _hSelectFile0 = h
            else:
                _hSelectFile1 = h
                return False
    return True

def process_std_dialog(hBtnToOpen, DlgTitle, ConfirmBtnTitle, FileName):
    push_button(hBtnToOpen)
    hDlg = wait_for_wnd(DlgTitle)
    if hDlg == 0:
        raise Exception("standard dialog window '{}' not found". format(DlgTitle))
    hEdit = win32gui.FindWindowEx(hDlg, 0, 'Edit', '')
    if hEdit == 0:
        raise Exception("standard dialog file name field not found")
    #win32gui.SetWindowText(hEdit, FileName)
    win32gui.SendMessage(hEdit, win32con.WM_SETTEXT, len(FileName), FileName)
    
    hConfirmBtn = win32gui.FindWindowEx(hDlg, 0, TARGET_BTTN_CLASS, ConfirmBtnTitle)
    if hConfirmBtn == 0:
        raise Exception("standard dialog confirm button not found")

    push_button(hConfirmBtn)
    try:
        wait_for_wnd_closed_h(hDlg)
    except Exception as e:
        raise Exception("standard dialog '{}' not closed for the time provided".format(DlgTitle))

def add_task(SrcFileName, DestFileName):
    global _hSelectFile0, _hSelectFile1
    # get main window
    hMainWnd = win32gui.FindWindow(None, TARGET_MAIN_WND_TITLE)
    if hMainWnd == 0:
        raise Exception("main window not found")
    # get button "New task"
    hNewTaskBttn = win32gui.FindWindowEx(hMainWnd, 0, TARGET_BTTN_CLASS, TARGET_NEWTASK_BTTN_TITLE)
    if hNewTaskBttn == 0:
        raise Exception("new task button not found")
    # push button New task
    push_button(hNewTaskBttn)
    # waiting for new task dialog
    hNewTaskDlg = wait_for_wnd(TARGET_DLG_WND_TITLE)
    if hNewTaskDlg == 0:
        raise Exception("new task dialog window not found")
    # enumerate controls on the window
    _hSelectFile0 = 0
    _hSelectFile1 = 0
    win32gui.EnumChildWindows(hNewTaskDlg, EnumWndProc, 0)
    if _hSelectFile0 == 0 and _hSelectFile1 == 0:
        raise Exception("file selection buttons not found")

    try:
        process_std_dialog(_hSelectFile0, STD_DLG_OPEN_TITLE, STD_DLG_OPEN_BTN_TITLE, SrcFileName)
        process_std_dialog(_hSelectFile1, STD_DLG_SAVEAS_TITLE, STD_DLG_SAVEAS_BTN_TITLE, DestFileName)
    except Exception as e:
        raise Exception("can't add tast because {}".format(e))
        
    hAddTaskBttn = win32gui.FindWindowEx(hNewTaskDlg, 0, TARGET_BTTN_CLASS, TARGET_ADD_BTTN_TITLE)
    push_button(hAddTaskBttn)
    try:
        wait_for_wnd_closed_h(hNewTaskDlg)
    except Exception as e:
        raise Exception("new task dialog not closed for the time provided")
        

def get_tasks_count():
    # get main window
    hMainWnd = win32gui.FindWindow(None, TARGET_MAIN_WND_TITLE)
    if hMainWnd == 0:
        raise Exception("main window not found")
    # get list
    hList = win32gui.FindWindowEx(hMainWnd, 0, "SysListView32", None)
    if hList == 0:
        raise Exception("list window not found")
    Count = win32gui.SendMessage(hList, LVM_GETITEMCOUNT, 0, 0);
    return Count

for SrcFilename in glob.glob(path.join(SRC_DIR ,"*." + FILE_EXTENTION)):
    DstFileName = path.join(DST_DIR, path.basename(SrcFilename) + ".avi")
    if path.exists(DstFileName):
        print('file "{}" already exists'.format(DstFileName))
        continue
    print('add task "{}" -> "{}"'.format(SrcFilename, DstFileName))
    add_task(SrcFilename, DstFileName)
    while get_tasks_count() >= MAX_TASKS:
        sleep(0.5)

