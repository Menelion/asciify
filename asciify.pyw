#!/usr/bin/env python
# -*- coding: utf-8 -*-

""" Asciify, version 0.1
A GUI app designed to recursively transform unicode file names to ASCII.
This software is free and MIT-licensed. It's written in hope that it will be useful and implies no warranties.
Inspired by and initially developed for Radio Music Train."""

import logging
logging.basicConfig(filename="asciify.log", filemode="w", level=logging.DEBUG)
import wx, os, thread
import win32com.client as com
from unidecode import unidecode

__author__ = u"Nostië & Menelion Elensúlë"
__copyright__ = u"Copyright © %s, 2015" % __author__
__credits__ = [u"Nostië Elensúlë", u"Menelion Elensúlë"]
__license__ = u"MIT"
__version__ = u"0.1"
__maintainer__ = u"Menelion Elensúlë"
__email__ = u"info@oire.me"
__status__ = u"Development"


class MainWindow(wx.Frame):
	""" Main app frame."""

	def __init__(self, parent, title=u"Asciify"):
		wx.Frame.__init__(self, parent, -1, title)

		# The menus
		mainMenu = wx.MenuBar()
		fileMenu = wx.Menu()
		fileMenu.Append(wx.ID_EXIT, u"E&xit", u"Exit Asciify")
		self.Bind(wx.EVT_MENU, self.onCloseApp, id=wx.ID_EXIT)
		mainMenu.Append(fileMenu, u"&File")
		helpMenu = wx.Menu()
		helpMenu.Append(wx.ID_ABOUT, u"&About Asciify...\tF1", u"Displays info about the program, version and copyright")
		self.Bind(wx.EVT_MENU, self.onAbout, id=wx.ID_ABOUT)
		mainMenu.Append(helpMenu, u"&Help")
		self.SetMenuBar(mainMenu)
		self.CreateStatusBar()

		# Main interface panel
		panel = wx.Panel(self)
		sfLabel = wx.StaticText(panel, -1, u"Please select the folder to be processed:")
		self.folderEdit = wx.TextCtrl(panel, -1, style=wx.TE_READONLY)
		browseBtn = wx.Button(panel, -1, u"&Browse...")
		self.startBtn = wx.Button(panel, -1, u"Start &Processing")
		self.startBtn.Enable(False)
		self.Bind(wx.EVT_BUTTON, self.onBrowse, browseBtn)
		self.Bind(wx.EVT_BUTTON, self.onStartProcessing, self.startBtn)
		self.folderEdit.SetFocus()

		folderSizer = wx.BoxSizer(wx.HORIZONTAL)
		folderSizer.Add(self.folderEdit)
		folderSizer.Add(browseBtn)
		sizer = wx.BoxSizer(wx.VERTICAL)
		sizer.Add(sfLabel)
		sizer.Add(folderSizer)
		sizer.Add(self.startBtn)
		panel.SetSizerAndFit(sizer)
		panel.Layout()
		sizer = wx.BoxSizer()
		sizer.Add(panel, 1, wx.EXPAND)
		self.SetSizerAndFit(sizer)


	# Event handlers
	def onCloseApp(self, event):
		self.Close()

	def onAbout(self, event):
		aboutDlg = wx.AboutDialogInfo()
		aboutDlg.SetName(u"Asciify")
		aboutDlg.SetVersion(__version__)
		aboutDlg.SetCopyright(__copyright__)
		aboutDlg.SetDescription(u"An application designed to recursively transform unicode file names to ASCII.")
		wx.AboutBox(aboutDlg)

	def onBrowse(self, event):
		dlg = wx.DirDialog(None, u"Choose a folder", style=wx.DD_DEFAULT_STYLE|wx.DD_DIR_MUST_EXIST)
		if dlg.ShowModal() == wx.ID_OK:
			self.folder = dlg.GetPath().replace(u"\\", u"\\\\")
			self.folderEdit.SetValue(dlg.GetPath())
			self.startBtn.Enable(True)
		dlg.Destroy()
		self.folderEdit.SetFocus()

	def onStartProcessing(self, event):
		# Checking once more
		if not os.path.isdir(self.folder):
			errorDlg = wx.MessageDialog(None, u"The folder you have selected does not exist. Please choose another folder.", u"Error!", wx.OK|wx.ICON_ERROR)
			errorDlg.ShowModal()
			errorDlg.Destroy()
		else: # the path is valid
			self.startBtn.Enable(False)
			self.count = 0
			self.processDlg = None
			self.timer = wx.Timer(self)
			self.Bind(wx.EVT_TIMER, self.processFiles, self.timer)
			self.timer.Start(1000)

	# Everything related to files processing itself
	def getFolderSize(self, root=u"."):
		"""Gets the size of a folder including all subfolders in a very fast and efficient manner. Returns the value in bytes."""
		fso = com.Dispatch("Scripting.FileSystemObject")
		folder = fso.GetFolder(root)
		return folder.Size

	def processFiles(self, event):
		""" Recursively processes all folders and files transforming their names into ASCII."""
		if not self.processDlg:
			self.processDlg = wx.ProgressDialog(u"Asciify - Processing", u"Processing files...", 100, style=wx.PD_CAN_ABORT|wx.PD_REMAINING_TIME)
			keepGoing = True
			folderSize = float(self.getFolderSize(self.folder))
			percent = 0
			for root, dirs, files in os.walk(self.folder):
				for file in files:
					srcFile = os.path.join(root, file)
					destFile = os.path.join(root, unidecode(file))
					try:
						logging.info(u"Renaming to %s" % destFile)
						os.rename(srcFile, destFile)
					except Exception as e:
						logging.exception(e)
					try:
						currentPercent = int(round(os.path.getsize(destFile)/folderSize*100, 0))
						percent += currentPercent
						(keepGoing, skip) = self.processDlg.Update(percent, srcFile)
					except Exception as exc:
						logging.exception(exc)
			if not keepGoing or percent == 100:
				self.processDlg.Destroy()
				self.timer.Stop()

# the WX Application itself
class AsciifyApp(wx.App):
	def OnInit(self):
		frame = MainWindow(None, u"Asciify")
		self.SetTopWindow(frame)
		frame.Show(True)
		return True

app = AsciifyApp()
app.MainLoop()