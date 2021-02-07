#!/usr/bin/python
# -*- coding: utf-8 -*-

import sys
import os

class Logger(object):
    def __init__(self, filename="Default.log"):
        self.terminal = sys.stdout
        path = filename[0:filename.rfind("\\")]
        if not os.path.isdir(path):  # 无文件夹时创建
            os.makedirs(path)
        if not os.path.isfile(filename):  # 无文件时创建
            fd = open(filename, mode="w", encoding="utf-8")
        else:
            fd=open(filename, "a")
        self.log = fd

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)

    def flush(self):
        pass