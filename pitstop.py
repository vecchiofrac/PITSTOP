# -*- coding: utf-8 -*-

import wx
from PITSTOP import *

app = wx.App(False)

frame = main.MainWindow(None, 'PITSTOP')
app.MainLoop()
