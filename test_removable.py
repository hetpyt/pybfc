#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import win32con
from win32api import GetLogicalDriveStrings, GetVolumeInformation
from win32file import GetDriveType

#define TEST='removable drives'

def get_drives_list(drive_types=(win32con.DRIVE_REMOVABLE,)):
    drives_str = GetLogicalDriveStrings()
    drives = [item for item in drives_str.split("\x00") if item]
    return [item[:2] for item in drives if drive_types is None or GetDriveType(item) in drive_types]

if __name__ == "__main__":
    drives = get_drives_list(drive_types = (win32con.DRIVE_REMOVABLE,))
    print(TEST)
    for drive in drives:
        info = GetVolumeInformation(drive)
        print('{} {}'.format(drive, info[0]))