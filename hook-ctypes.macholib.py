# -*- coding: utf-8-sig -*-
"""
File Name ：    hook-ctypes.macholib.py
Author :        Eric
Create date ：  2020/11/13
"""
from PyInstaller.utils.hooks import copy_metadata
datas = copy_metadata('cffi') + copy_metadata('greenlet')