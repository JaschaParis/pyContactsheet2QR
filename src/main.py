import wx
import logging
from view.mainframe import MainFrame

if __name__ == "__main__":
    logging.info("Starting program.")
    app = wx.App()
    frame = MainFrame()
    app.MainLoop()