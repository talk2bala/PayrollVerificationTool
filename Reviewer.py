import time
#import serial
import sys
import os
import xlrd
import wx
import filecmp
import zipfile
import wx.animate

#*****************************************************************************************************
#
#Class to communicate to ADB and CMD
#
#*****************************************************************************************************
class AndroidDebugBridge(object):

    def call_adb(self, command):
        command_result = ''
        command_text = 'adb %s' % command
        results = os.popen(command_text, 'r')
        while 1:
            line = results.readline()
            if not line: break
            command_result += line
        return command_result

        
    def attached_devices(self):
        """ Return a list of attached devices."""
        result = self.call_adb("devices")
        devices = result.partition('\n')[2].replace('\n', '').split('\tdevice')
        return [device for device in devices if len(device) > 2]
       
    def push(self, local, remote):
        result = self.call_adb("push %s %s" % (local, remote))
        return result
        
    def pull(self, remote, local):
        result = self.call_adb("pull %s %s" % (remote, local))
        return result

    def execute_pythoncmd(self, command):
        command_result = ''
        command_text = 'python %s' % command
        results = os.popen(command_text, "r")
        while 1:
            line = results.readline()
            if not line: break
            command_result += line
        return command_result

#*****************************************************************************************************
#
#           Class to Open and Edit File
#
#*****************************************************************************************************

class EditWindow(wx.Frame):
    def __init__(self):
        super(EditWindow, self).__init__(None)
        self.filename = ''
        self.dirname = ''
        self.CreateInteriorWindowComponents()
        self.CreateExteriorWindowComponents()
        self.displayfile(self.dirname+self.filename)
        self.Show(True)

    def CreateInteriorWindowComponents(self):
        ''' Create "interior" window components. In this case it is just a
            simple multiline text control. '''
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        grid0 = wx.GridBagSizer(hgap=5, vgap=5)
        grid1 = wx.GridBagSizer(hgap=5, vgap=5)

        self.control = wx.TextCtrl(self, size=(350,200),style=wx.TE_MULTILINE)
        grid0.Add(self.control, pos=(1,0))

        self.save = wx.Button(self, label="Save")
        self.Bind(wx.EVT_BUTTON, self.OnSave,self.save)        
        grid1.Add(self.save, pos=(1,0))

        self.cancel = wx.Button(self, label="Cancel")
        self.Bind(wx.EVT_BUTTON, self.OnEditCancel,self.cancel)
        grid1.Add(self.cancel, pos=(1,1))
    
        mainSizer.Add(grid0, 0, wx.CENTER)
        mainSizer.Add(grid1, 0, wx.CENTER)
        self.SetSizerAndFit(mainSizer)

    def CreateExteriorWindowComponents(self):
        ''' Create "exterior" window components, such as menu and status
            bar. '''
        self.SetTitle()

    def SetTitle(self):
        # MainWindow.SetTitle overrides wx.Frame.SetTitle, so we have to
        # call it using super:
        super(EditWindow, self).SetTitle('Edit %s'%self.filename)

    def displayfile(self,filename):
        "test"

        dlg = wx.FileDialog(self, "Choose a file", self.dirname, "", "*.*", wx.OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.filename = dlg.GetFilename()
            self.dirname = dlg.GetDirectory()
            f = open(os.path.join(self.dirname, self.filename), 'r')
            self.control.SetValue(f.read())
            f.close()
        dlg.Destroy()

        if dlg.ShowModal() != wx.ID_OK:
            self.Close()

    def OnEditCancel(self, event):
        self.Close()  # Close the window.

    def OnSave(self, event):
        textfile = open(os.path.join(self.dirname, self.filename), 'w')
        textfile.write(self.control.GetValue())
        textfile.close()
        self.Close()

#*****************************************************************************************************
#
#           Main Class to control Window
#
#*****************************************************************************************************

class Reviewer(wx.Frame):
    "This is a Main Class to display the window"

    def __init__(self):
        wx.Frame.__init__(self, None, wx.ID_ANY, "Payroll Verification Tool", size=(800,600))
        #Creating Objects
        self.MakeCRC_Panel = MakeCRC_Panel(self)
        self.SoftReviewer_Panel = SoftReviewer_Panel(self)
        self.About_Panel = About_Panel(self)

        self.MakeCRC_Panel.Hide()
        self.SoftReviewer_Panel.Hide()
        self.About_Panel.Show()

        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.sizer.Add(self.MakeCRC_Panel, 1, wx.EXPAND)
        self.sizer.Add(self.SoftReviewer_Panel, 1, wx.EXPAND)
        self.SetSizer(self.sizer)

        self.CreateStatusBar()
        toolMenuBar = wx.MenuBar()
        toolMenu = wx.Menu()
        MakeCRC = toolMenu.Append(wx.ID_ANY,"Format XL&S","Tool for Handling Microsoft Excel 2003 Format")
        self.Bind(wx.EVT_MENU, self.OnMakeCRC, MakeCRC)
        toolMenu.AppendSeparator()
        SoftReviewer = toolMenu.Append(wx.ID_ANY,"&Format XLS&X","Tool for Handling Microsoft Excel 2007 Format")
        self.Bind(wx.EVT_MENU, self.OnSoftReviewer, SoftReviewer)
        toolMenu.AppendSeparator()
        About = toolMenu.Append(wx.ID_ANY,"&About","Info Regarding Tool and Developer")
        self.Bind(wx.EVT_MENU, self.OnAbout, About)
        toolMenu.AppendSeparator()
        Exit = toolMenu.Append(wx.ID_EXIT,"&Exit","Exit the Program")
        self.Bind(wx.EVT_MENU, self.OnExit, Exit)

        toolMenuBar.Append(toolMenu, "&Menu")
        self.SetMenuBar(toolMenuBar)

    def OnMakeCRC(self, event):
        "To launch Make CRC tool"
        self.About_Panel.Hide()
        self.SoftReviewer_Panel.Hide()
        self.MakeCRC_Panel.Show()
        self.Layout()

    def OnSoftReviewer(self, event):
        "To launch  Soft Reviewer tool"
        self.About_Panel.Hide()
        self.MakeCRC_Panel.Hide()
        self.SoftReviewer_Panel.Show()
        self.Layout()

    def OnAbout(self, event):
        "To launch  About"
        self.MakeCRC_Panel.Hide()
        self.SoftReviewer_Panel.Hide()
        self.About_Panel.Show()
        self.Layout()

    def OnExit(self, event):
        "Exit"
        self.Close(True)

#*****************************************************************************************************
#
#           MakeCRC Class
#
#*****************************************************************************************************
class About_Panel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent=parent)

        img_real = wx.Image('./res_common/Intro.JPG', wx.BITMAP_TYPE_ANY)
        self.imageCtrl = wx.StaticBitmap(self, wx.ID_ANY, wx.BitmapFromImage(img_real))
        self.imageCtrl.SetBitmap(wx.BitmapFromImage(img_real))

        self.mainSizer = wx.BoxSizer(wx.VERTICAL)
        self.mainSizer.Add(self.imageCtrl, 0, wx.CENTER)
        self.SetSizerAndFit(self.mainSizer)
    

#*****************************************************************************************************
#
#           MakeCRC Class
#
#*****************************************************************************************************
class MakeCRC_Panel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent=parent)

        # create some sizers
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        grid0 = wx.GridBagSizer(hgap=5, vgap=5)
        grid = wx.GridBagSizer(hgap=10, vgap=10)
        grid2 = wx.GridBagSizer(hgap=5, vgap=5)

        self.heading = wx.StaticText(self, label="XLS Tool")
        font = wx.Font(22,wx.DEFAULT, wx.NORMAL, wx.NORMAL, True)
        self.heading.SetFont(font)
        grid0.Add(self.heading, pos=(1,0))        

        #Adding Combo Box for Port Selection
        self.quote = wx.StaticText(self, label="Select your Previous Month Sheet ")
        grid.Add(self.quote, pos=(2,0))
        self.Prev_Month_Sheet = wx.TextCtrl(self, size=(300,20))
        grid.Add(self.Prev_Month_Sheet, pos=(2,1))

        self.button =wx.Button(self, label="Browse")
        self.Bind(wx.EVT_BUTTON, self.OnBrowse1,self.button)
        grid.Add(self.button, pos=(2,2))


        # A button to Edit Local Config file
        self.Current_Month = wx.StaticText(self, label="Select Your Current Month Sheet ")
        grid.Add(self.Current_Month, pos=(3,0))
        self.Curr_Month_Sheet = wx.TextCtrl(self, size=(300,20))
        grid.Add(self.Curr_Month_Sheet, pos=(3,1))

        self.button =wx.Button(self, label="Browse")
        self.Bind(wx.EVT_BUTTON, self.OnBrowse2,self.button)
        grid.Add(self.button, pos=(3,2))
        

        # Buttons to Select all/MakeCRC/Exit
        self.scanthrough =wx.Button(self, label="Scan through...")
        self.Bind(wx.EVT_BUTTON, self.OnScanthrough,self.scanthrough)
        grid.Add(self.scanthrough, pos=(5,0))

        self.MAKECRC =wx.Button(self, label="Generate Payroll Sheet")
        self.Bind(wx.EVT_BUTTON, self.OnMakeCRC,self.MAKECRC)
        grid.Add(self.MAKECRC, pos=(5,2))


        gif_fname = './res_common/Progress.gif'
        self.gifleft = wx.animate.GIFAnimationCtrl(self, -1, gif_fname)
        self.gifleft.GetPlayer().UseBackgroundColour(True)
        grid2.Add(self.gifleft, pos=(2,0))
        self.gifleft.Hide()

        # A multiline TextCtrl - This is here to show how the events work in this program, don't pay too much attention to it
        self.logger = wx.TextCtrl(self, size=(600,150), style=wx.TE_MULTILINE | wx.TE_READONLY)
        grid2.Add(self.logger, pos=(2,1))

        self.gifright = wx.animate.GIFAnimationCtrl(self, -1, gif_fname)
        self.gifright.GetPlayer().UseBackgroundColour(True)
        grid2.Add(self.gifright, pos=(2,2))
        self.gifright.Hide()

        mainSizer.Add(grid0, 0, wx.CENTER)
        mainSizer.Add(grid, 0, wx.CENTER)
        mainSizer.Add(grid2, 0, wx.CENTER)
        self.SetSizerAndFit(mainSizer)

    def EvtComboBox(self, event):
        "test"

    def OnBrowse1(self, event):
        "Edit Local_Config.txt file"
        
        dlg = wx.FileDialog(self, "Choose a file", "c:\\", "", "*.*", wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.path1 = dlg.GetPath();
            self.Prev_Month_Sheet.WriteText(self.path1);
        dlg.Destroy()
            
    def OnBrowse2(self, event):
        "Edit Local_Config.txt file"
        
        dlg = wx.FileDialog(self, "Choose a file", "c:\\", "", "*.*", wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.path2 = dlg.GetPath();
            self.Curr_Month_Sheet.WriteText(self.path2);
        dlg.Destroy()

    def OnScanthrough(self,event):
        "Core function to identify the probelms"
        File1_Presence = True;#is_SystemFilesPresent(self.path1, "") Needs to be implemented
        File2_Presence = True;#is_SystemFilesPresent(self.path2, "") Needs to be implemented

        if( File1_Presence and File2_Presence):
            self.logger.AppendText("File is Present \n")
            Prev_wb = xlrd.open_workbook(self.path1);
            Prev_sh = wb.sheet_by_index(0);
            Curr_wb = xlrd.open_workbook(self.path2);
            Curr_sh = wb.sheet_by_index(0);

        else:
            self.logger.AppendText("Seems to be one of the input file is missing!!!! Please Check \n")

    def OnSelectAll(self,event):
        "Function to select all check boxes in single click"
        self.dbcrc.SetValue(True)
        self.filecrc.SetValue(True)
        self.fpricrc.SetValue(True)
        self.efscrc.SetValue(True)
        self.fpritest.SetValue(True)
        self.swreqdoc.SetValue(True)

    def OnMakeCRC(self,event):
        "This is a function which does everything in single click"
        SL = '' #Global variable to store SIM LOCK Info
        NTCODE = '' #Global variable to store NT Code Info
        #Starts Animation
        self.gifleft.Show()
        self.gifright.Show()
        self.gifleft.Play()
        self.gifright.Play()
        self.Layout()
        #Input Validation
        if validate_Input(self,'makecrc')==True:
            # configure the serial connections (the parameters differs on the device you are connecting to)
            ser = serial.Serial(
                port=self.edithear.GetValue(),
                baudrate=9600,
                parity=serial.PARITY_ODD,
                stopbits=serial.STOPBITS_TWO,
                bytesize=serial.SEVENBITS
            )
            if ser.isOpen():
                #system file check
                if is_SystemFilesPresent("./res_makecrc/", "Local_Config.txt")==True:
                    LocalConfig = open('./res_makecrc/Local_Config.txt','r')
                    Lines = LocalConfig.readlines()
                    LocalConfig.close()
                    PRXLS_Fname = Lines[0].split('$')[1]
                    PRXLS_Filename = PRXLS_Fname[0:len(PRXLS_Fname)-1]
                    #initial Setup
                    if display_Warning(self)==True:

                        #Filesystem Clean up
                        do_SystemCleanup(self)
                        #PR File check
                        if is_SystemFilesPresent("./res_makecrc/", PRXLS_Filename)==True:
                            SL = get_ValuefromXLS("./res_makecrc/"+PRXLS_Filename, trim_value(Lines[24].split('$')[1]))
                            NTCODE = get_ValuefromXLS("./res_makecrc/"+PRXLS_Filename, trim_value(Lines[25].split('$')[1]))
                        else:
                            self.logger.AppendText("PR File is missing, Proceeding with User Input\n")
                            dlg = wx.MessageDialog(self,"Is it SIM LOCK Version?", "Simlock", wx.YES|wx.NO|wx.ICON_QUESTION)
                            res = dlg.ShowModal()
                            dlg.Destroy()
                            if res == wx.ID_YES:
                                SL = "Yes"
                            else:
                                SL = "No"

                            NTcodedefault = NTCODE = "\"0\",\"FFF,FFF,FFFFFFFF,FFFFFFFF,FF\""

                            dlg = wx.TextEntryDialog(self,"Enter the NTCODE to be written", "NTCODE", NTcodedefault ,wx.OK|wx.CENTRE)
                            res = dlg.ShowModal()
                            if res == wx.ID_OK:
                                NTCODE = dlg.GetValue()

                        if write_IDDE(self, ser) == True:
                            if write_NTCODE(self, ser, NTCODE) == True:
                                if do_SIMLOCK(self, ser, SL) == True:
                                    #Execute Core logic
                                    result = Execute_Core_Logic(self, ser)
                                    if result == True :
                                        dlg = wx.MessageDialog(self, "Done", "Result", wx.OK | wx.ICON_INFORMATION )
                                        dlg.ShowModal()
                                        dlg.Destroy()
                                    else:
                                        Message = "Stopped due to problems. See logs for more details."
                                        dlg = wx.MessageDialog(self, "Stopped due to problems. See logs for more details.", "Result", wx.OK | wx.ICON_ERROR )
                                        dlg.ShowModal()
                                        dlg.Destroy()
                                    ser.close()
                                else:
                                    self.logger.AppendText("Failed in doing SIMLOCK\n")
                            else:
                                self.logger.AppendText("Failed in writing NTCODE\n")
                        else:
                            self.logger.AppendText("Failed in doing IDDE. Check your HW\n")
                    else:
                        self.logger.AppendText("User Cancelled Execution due to lack of initial setup in Hardware\n")
                else:
                    self.logger.AppendText("Local_Config.txt file is missing\n")
                ser.close()
            else:
                self.logger.AppendText("Please Check your Port Settings\n")
            ser.close()
        else:
            self.logger.AppendText("Input Validation Failed\n")

        self.gifleft.Stop()
        self.gifright.Stop()    
        self.gifleft.Hide()
        self.gifright.Hide()
        self.Layout()

    def OnClearLogs(self,event):
        self.logger.Clear()


#*****************************************************************************************************
#
#           Soft Reviewer Class
#
#*****************************************************************************************************

class SoftReviewer_Panel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent=parent)

        # create some sizers
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        grid0 = wx.GridBagSizer(hgap=5, vgap=5)
        grid1 = wx.GridBagSizer(hgap=5, vgap=5)
        grid2 = wx.GridBagSizer(hgap=15, vgap=15)
        grid3 = wx.GridBagSizer(hgap=5, vgap=5)
        grid4 = wx.GridBagSizer(hgap=5, vgap=5)

        self.heading = wx.StaticText(self, label="Soft Reviewer")
        font = wx.Font(22,wx.DEFAULT, wx.NORMAL, wx.NORMAL, True)
        self.heading.SetFont(font)
        grid0.Add(self.heading, pos=(1,0))        

        #Adding Combo Box for Port Selection
        self.quote = wx.StaticText(self, label="Select your COM PORT ")
        grid1.Add(self.quote, pos=(1,0))
        self.portList = populate_COMPORT(self)
        self.edithear = wx.ComboBox(self, size=(100, -1), choices=self.portList, style=wx.CB_DROPDOWN)
        grid1.Add(self.edithear, pos=(1,2))

        # A button to Edit Local Config file
        self.BrowseLable = wx.StaticText(self, label="Edit your Local Config ")
        grid1.Add(self.BrowseLable, pos=(3,0))
        self.button =wx.Button(self, label="Edit")
        self.Bind(wx.EVT_BUTTON, self.OnBrowse,self.button)
        grid1.Add(self.button, pos=(3,2))

        # Checkbox 1
        self.dbcrc = wx.CheckBox(self, label="DBCRC")
        grid2.Add(self.dbcrc, pos=(1,0), flag=wx.BOTTOM, border=5)
        self.filecrc = wx.CheckBox(self, label="FILECRC")
        grid2.Add(self.filecrc, pos=(2,0), flag=wx.BOTTOM, border=5)
        self.fpricrc = wx.CheckBox(self, label="FPRICRC")
        grid2.Add(self.fpricrc, pos=(3,0), flag=wx.BOTTOM, border=5)
        self.efscrc = wx.CheckBox(self, label="EFS CRC")
        grid2.Add(self.efscrc, pos=(4,0), flag=wx.BOTTOM, border=5)
        self.fpritest = wx.CheckBox(self, label="FPRI Test Cases")
        grid2.Add(self.fpritest, pos=(1,1), flag=wx.BOTTOM, border=5)
        self.swreqdoc = wx.CheckBox(self, label="SW Req Doc")
        grid2.Add(self.swreqdoc, pos=(2,1), flag=wx.BOTTOM, border=5)
        self.SanityRep = wx.CheckBox(self, label="Sanity Report")
        grid2.Add(self.SanityRep, pos=(3,1), flag=wx.BOTTOM, border=5)
        self.bin = wx.CheckBox(self, label="BIN")
        grid2.Add(self.bin, pos=(4,1), flag=wx.BOTTOM, border=5)
        self.rom = wx.CheckBox(self, label="ROM")
        grid2.Add(self.rom, pos=(1,2), flag=wx.BOTTOM, border=5)
        self.dz = wx.CheckBox(self, label="DZ")
        grid2.Add(self.dz, pos=(2,2), flag=wx.BOTTOM, border=5)
        self.bin_dz = wx.CheckBox(self, label="BIN-DZ")
        grid2.Add(self.bin_dz, pos=(3,2), flag=wx.BOTTOM, border=5)
        self.kdz = wx.CheckBox(self, label="KDZ")
        grid2.Add(self.kdz, pos=(4,2), flag=wx.BOTTOM, border=5)
        self.fota = wx.CheckBox(self, label="FOTA")
        grid2.Add(self.fota, pos=(1,3), flag=wx.BOTTOM, border=5)
        self.dll = wx.CheckBox(self, label="DLL")
        grid2.Add(self.dll, pos=(2,3), flag=wx.BOTTOM, border=5)
        self.mmc = wx.CheckBox(self, label="MMC")
        grid2.Add(self.mmc, pos=(3,3), flag=wx.BOTTOM, border=5)
        self.pri = wx.CheckBox(self, label="PRI")
        grid2.Add(self.pri, pos=(4,3), flag=wx.BOTTOM, border=5)
        self.swrelnote = wx.CheckBox(self, label="SW Rel Note")
        grid2.Add(self.swrelnote, pos=(1,4), flag=wx.BOTTOM, border=5)
        self.tagcompare = wx.CheckBox(self, label="Tag Comparision")
        grid2.Add(self.tagcompare, pos=(2,4), flag=wx.BOTTOM, border=5)

        # Buttons to Select all/MakeCRC/Exit
        self.update =wx.Button(self, label="Update from Model Config")
        self.Bind(wx.EVT_BUTTON, self.updatemodelconfig,self.update)
        grid3.Add(self.update, pos=(1,0))

        self.selectall =wx.Button(self, label="Select All")
        self.Bind(wx.EVT_BUTTON, self.OnSelectAll,self.selectall)
        grid3.Add(self.selectall, pos=(1,1))

        self.softreview =wx.Button(self, label="Review")
        self.Bind(wx.EVT_BUTTON, self.OnReview,self.softreview)
        grid3.Add(self.softreview, pos=(1,2))

        self.Logclear =wx.Button(self, label="Clear Logs")
        self.Bind(wx.EVT_BUTTON, self.OnClearLogs,self.Logclear)
        grid3.Add(self.Logclear, pos=(1,3))
        
        #logger Heading
        self.loggerHeading = wx.StaticText(self, label="Logger")
        grid4.Add(self.loggerHeading, pos=(1,1))

        gif_fname = './res_common/Progress.gif'
        self.gifleft = wx.animate.GIFAnimationCtrl(self, -1, gif_fname)
        self.gifleft.GetPlayer().UseBackgroundColour(True)
        grid4.Add(self.gifleft, pos=(2,0))
        self.gifleft.Hide()

        # A multiline TextCtrl - This is here to show how the events work in this program, don't pay too much attention to it
        self.logger = wx.TextCtrl(self, size=(600,300), style=wx.TE_MULTILINE | wx.TE_READONLY)
        grid4.Add(self.logger, pos=(2,1))

        self.gifright = wx.animate.GIFAnimationCtrl(self, -1, gif_fname)
        self.gifright.GetPlayer().UseBackgroundColour(True)
        grid4.Add(self.gifright, pos=(2,2))
        self.gifright.Hide()

        mainSizer.Add(grid0, 0, wx.CENTER)
        mainSizer.Add(grid1, 0, wx.CENTER)
        mainSizer.Add(grid2, 0, wx.CENTER)
        mainSizer.Add(grid3, 0, wx.CENTER)
        mainSizer.Add(grid4, 0, wx.CENTER)
        self.SetSizerAndFit(mainSizer)

    def OnBrowse(self, event):
        "Edit Local_Config.txt file"
        dirname = "./res_softreviewer/"
        filename = "Local_Config.txt"
        if is_SystemFilesPresent(dirname, filename)==True:
            sub_Frame = EditWindow()
        else:
            self.logger.AppendText('Local Config file is missing\n')
            dlg = wx.MessageDialog(self, "Local Config File is missing", "Warning", wx.OK | wx.ICON_INFORMATION )
            dlg.ShowModal()
            dlg.Destroy()  

    def OnSelectAll(self, event):
        "A funtion to select all check boxes in single click"
        self.dbcrc.SetValue(True)
        self.filecrc.SetValue(True)
        self.fpricrc.SetValue(True)
        self.efscrc.SetValue(True)
        self.fpritest.SetValue(True)
        self.swreqdoc.SetValue(True)
        self.SanityRep.SetValue(True)
        self.bin.SetValue(True)
        self.rom.SetValue(True)
        self.dz.SetValue(True)
        self.bin_dz.SetValue(True)
        self.kdz.SetValue(True)
        self.fota.SetValue(True)
        self.dll.SetValue(True)
        self.mmc.SetValue(True)
        self.pri.SetValue(True)
        self.swrelnote.SetValue(True)

    def OnClearLogs(self, event):
        "To clear the logger"
        self.logger.Clear()

    def updatemodelconfig(self, event):
        "Updating check boxes from Model Config"
        Lines = ''
        if is_SystemFilesPresent('./res_softreviewer/', 'Model_Config.txt')==True:
            ModelConfig = open('./res_softreviewer/Model_Config.txt','r')
            Lines = ModelConfig.readlines()
            ModelConfig.close()
        else:
            self.logger.AppendText('Model Config file is missing\n')
            dlg = wx.MessageDialog(self, "Model Config File is missing", "Warning", wx.OK | wx.ICON_INFORMATION )
            dlg.ShowModal()
            dlg.Destroy()  
        
        if Lines[0].split(':')[1] == 'Y\n': self.dbcrc.SetValue(True)
        else: self.dbcrc.SetValue(False)

        if Lines[1].split(':')[1] == 'Y\n': self.filecrc.SetValue(True)
        else: self.filecrc.SetValue(False)

        if Lines[2].split(':')[1] == 'Y\n': self.fpricrc.SetValue(True)
        else: self.fpricrc.SetValue(False)

        if Lines[3].split(':')[1] == 'Y\n': self.efscrc.SetValue(True)
        else: self.efscrc.SetValue(False)

        if Lines[4].split(':')[1] == 'Y\n': self.fpritest.SetValue(True)
        else: self.fpritest.SetValue(False)

        if Lines[5].split(':')[1] == 'Y\n': self.bin.SetValue(True)
        else: self.bin.SetValue(False)

        if Lines[6].split(':')[1] == 'Y\n': self.rom.SetValue(True)
        else: self.rom.SetValue(False)

        if Lines[7].split(':')[1] == 'Y\n': self.dz.SetValue(True)
        else: self.dz.SetValue(False)

        if Lines[8].split(':')[1] == 'Y\n': self.bin_dz.SetValue(True)
        else: self.bin_dz.SetValue(False)

        if Lines[9].split(':')[1] == 'Y\n': self.kdz.SetValue(True)
        else: self.kdz.SetValue(False)

        if Lines[10].split(':')[1] == 'Y\n': self.fota.SetValue(True)
        else: self.fota.SetValue(False)

        if Lines[11].split(':')[1] == 'Y\n': self.dll.SetValue(True)
        else: self.dll.SetValue(False)

        if Lines[12].split(':')[1] == 'Y\n': self.mmc.SetValue(True)
        else: self.mmc.SetValue(False)

        if Lines[13].split(':')[1] == 'Y\n': self.pri.SetValue(True)
        else: self.pri.SetValue(False)

        if Lines[14].split(':')[1] == 'Y\n': self.SanityRep.SetValue(True)
        else: self.SanityRep.SetValue(False)

        if Lines[15].split(':')[1] == 'Y\n': self.swrelnote.SetValue(True)
        else: self.swrelnote.SetValue(False)

        if Lines[16].split(':')[1] == 'Y\n': self.swreqdoc.SetValue(True)
        else: self.swreqdoc.SetValue(False)

        if Lines[17].split(':')[1] == 'Y\n': self.tagcompare.SetValue(True)
        else: self.tagcompare.SetValue(False)

    def OnReview(self, event):
        "Test"

        SL = '' #Global variable to store SIM LOCK Info
        NTCODE = '' #Global variable to store NT Code Info
        EFS_Created = False
        FPRI_Created = False
        #Starts Animation
        self.gifleft.Show()
        self.gifright.Show()
        self.gifleft.Play()
        self.gifright.Play()
        self.Layout()

        #Input Validation
        if validate_Input(self,'softreviewer')==True:
            # configure the serial connections (the parameters differs on the device you are connecting to)
            ser = serial.Serial(
                port=self.edithear.GetValue(),
                baudrate=9600,
                parity=serial.PARITY_ODD,
                stopbits=serial.STOPBITS_TWO,
                bytesize=serial.SEVENBITS
            )
            if ser.isOpen():
                if is_SystemFilesPresent('./res_softreviewer/','Local_Config.txt')== True:
                    File_Obj_Local = open('./res_softreviewer/Local_Config.txt','r')
                    LocalLines = File_Obj_Local.readlines()
                    File_Obj_Local.close()
                    #initializing all required Variables from Local Config.
                    X = LocalLines[0].split('$')[1]
                    User_BIN_Path = X[0:len(X)-1]
                    Y = LocalLines[1].split('$')[1]
                    User_DOC_Path = Y[0:len(Y)-1]
                    Z = LocalLines[2].split('$')[1]
                    User_PR_Path  = Z[0:len(Z)-1]
                    User_PR_File = LocalLines[3].split('$')[1]
                    User_PR_Filename = User_PR_File[0:len(User_PR_File)-1]
                    Src_Ver       = get_SRC_VER_info(LocalLines[4].split('$')[1]);
                    EFSCRC_File_Path = LocalLines[8].split('$')[1]
                    EFSCRC_File = LocalLines[9].split('$')[1]
                    FPRI_File_Path = LocalLines[10].split('$')[1]
                    FPRI_File = LocalLines[11].split('$')[1]
                    Remote_EFSCRC_Path = LocalLines[12].split('$')[1]
                    Remote_FPRITest_Path = LocalLines[13].split('$')[1]
                    FingerPrint = ''
                    W = LocalLines[14].split('$')[1]
                    SRN_Path = W[0:len(W)-1]
                    SW_Req_Doc = LocalLines[15].split('$')[0]
                    SW_Req_Doc = SW_Req_Doc[0:len(SW_Req_Doc)-1]
                    QM_info = LocalLines[15].split('$')[1].split('/')[0]
                    QM_info = QM_info[0:len(QM_info)-1]
                    Tester_info = LocalLines[15].split('$')[1].split('/')[1]
                    Tester_info = Tester_info[0:len(Tester_info)-1]
                    User_SW_Req_Doc_Path = LocalLines[16].split('$')[1]
                    User_SW_Req_Doc_Path = User_SW_Req_Doc_Path[0:len(User_SW_Req_Doc_Path)-1]
                    User_SW_Req_Doc_File = LocalLines[17].split('$')[1]
                    User_SW_Req_Doc_File = User_SW_Req_Doc_File[0:len(User_SW_Req_Doc_File)-1]
                    LocalPath_EFS = os.getcwd()+"\\EFS_CRC.txt"
                    LocalPath_FPRI = os.getcwd()+"\\FPRI_TestCase.txt"
                    DOC_PRI_Path = LocalLines[18].split('$')[1]
                    DOC_PRI_Path = DOC_PRI_Path[0:len(DOC_PRI_Path)-1]
                    DOC_TagCompare_Path = LocalLines[19].split('$')[1]
                    DOC_TagCompare_Path = DOC_TagCompare_Path[0:len(DOC_TagCompare_Path)-1]
                    DOC_SanityRep_Path = LocalLines[20].split('$')[1]
                    DOC_SanityRep_Path = DOC_SanityRep_Path[0:len(DOC_SanityRep_Path)-1]

                    if is_SystemFilesPresent(User_PR_Path,User_PR_Filename) == True: 
                        SL = get_ValuefromXLS(User_PR_Path+User_PR_Filename, trim_value(LocalLines[34].split('$')[1]))
                        NTCODE = get_ValuefromXLS(User_PR_Path+User_PR_Filename, trim_value(LocalLines[35].split('$')[1]))
                        if display_Warning(self)==True:                    
                            #Filesystem Clean up
                            do_SystemCleanup(self)
                            if write_IDDE(self, ser) == True:
                                if write_NTCODE(self, ser, NTCODE) == True:
                                    if do_SIMLOCK(self, ser, SL) == True:
                                        if do_SecondarySetup(self, ser)== True:
                                            # List of files to be reviewed from user
                                            list_of_files = os.listdir(User_BIN_Path)

                                            SWV = get_SWV(ser);
                                            SWOV = get_SWOV(ser);
                                            Ver_Name = get_KDZName(SWV);
                                            Model = get_Modelinfo(get_ValuefromXLS(User_PR_Path+User_PR_Filename, trim_value(LocalLines[28].split('$')[1])));
                                            Suffix = get_ValuefromXLS(User_PR_Path+User_PR_Filename, trim_value(LocalLines[29].split('$')[1])) 
                                            Curr_Ver = get_CURR_VER_info(SWV);
                                            Country = get_ValuefromXLS(User_PR_Path+User_PR_Filename, trim_value(LocalLines[30].split('$')[1]))

                                            #File Format requirements
                                            BIN_Name = "BIN_"+ SWV +".zip"
                                            ROM_Name = "ROM_"+ SWV +".zip"
                                            DZ_Name = "DZ_"+ SWV +".zip"
                                            SanReport_Name = "Sanity_Rep_"+ SWV + ".xls"
                                            PRI_Name = "PRI_"+ SWV +".xls"
                                            KDZ_Name1 = Ver_Name+"_00.kdz"
                                            KDZ_Name2 = Ver_Name+"_00.txt"
                                            FOTA_Name = Model+"_"+Suffix+"_"+Src_Ver+"-"+Curr_Ver+".up"
                                            SRN_File = "SRN_"+SWV+".xls"
                                            DLL_Name = ".dll"


                                            #Reading Config file - to identify what all files to be checked. 
                                            File_Obj = open('./res_softreviewer/Model_Config.txt','r')
                                            Lines = File_Obj.readlines()
                                            File_Obj.close()
                                            
                                            Output = open("./Review_Result.txt",'w')

                                            Output.write("***************************************************************\n")
                                            Output.write("                     E R R O R  R E P O R T                    \n")
                                            Output.write("***************************************************************\n\n")


                                            #DB_CRC
                                            if self.dbcrc.IsChecked() == True:
                                                Output.write("DBCRC Test\n")
                                                Output.write("----------\n")
                                                if list_of_files.count('DBCRC') >= 1: #Folder present or not
                                                    User_list1 = os.listdir(User_BIN_Path+'DBCRC/')
                                                    Sys_list1 = os.listdir('./DBCRC/')
                                                    if len(User_list1) != 0: # Folder is not emplty
                                                        if cmp(User_list1,Sys_list1) == 0: # File name are same?
                                                            if filecmp.cmp(User_BIN_Path+'DBCRC/'+User_list1[0],'./DBCRC/'+Sys_list1[0]) == True: # file content comparision
                                                                Output.write("\tDB CRC Check --> Success")#pass;
                                                            else:
                                                                Output.write("\tMismatch found in DBCRC file!!!")
                                                        else:
                                                            Output.write("\tThere is difference in file names - DBCRC!!!")
                                                    else:
                                                        Output.write("\tDBCRC Folder is EMPTY!!!")
                                                else:
                                                    Output.write("\tDBCRC Folder is missing!!!")
                                            #FILE_CRC
                                            if self.filecrc.IsChecked() == True:
                                                Output.write("\n\n")
                                                Output.write("FILECRC Test\n")
                                                Output.write("------------\n")
                                                if list_of_files.count('FILECRC') >= 1: #Folder present or not
                                                    User_list2 = os.listdir(User_BIN_Path+'FILECRC/')
                                                    Sys_list2 = os.listdir('./FILECRC/')
                                                    if len(User_list2) != 0: # Folder is not emplty
                                                        if cmp(User_list2,Sys_list2) == 0: # File name are same?
                                                            if filecmp.cmp(User_BIN_Path+'FILECRC/'+User_list2[0],'./FILECRC/'+Sys_list2[0]) == True: # file content comparision
                                                                Output.write("\tFILE CRC Check --> Success")#pass;
                                                            else:
                                                                Output.write("\tMismatch found in FILECRC file!!!")
                                                        else:
                                                            Output.write("\tThere is difference in file names - FILECRC!!!")
                                                    else:
                                                        Output.write("\tFILECRC Folder is EMPTY!!!")
                                                else:
                                                    Output.write("\tFILECRC Folder is missing!!!")
                                            #FPRI_CRC
                                            if self.fpricrc.IsChecked() == True:
                                                Output.write("\n\n")
                                                Output.write("FPRICRC Test\n")
                                                Output.write("------------\n")
                                                if list_of_files.count('FPRICRC') >= 1: #Folder present or not
                                                    User_list3 = os.listdir(User_BIN_Path+'FPRICRC/')
                                                    Sys_list3 = os.listdir('./FPRICRC/')
                                                    if len(User_list3) != 0: # Folder is not emplty
                                                        if cmp(User_list3,Sys_list3) == 0: # File name are same?
                                                            if filecmp.cmp(User_BIN_Path+'FPRICRC/'+User_list3[0],'./FPRICRC/'+Sys_list3[0]) == True: # file content comparision
                                                                Output.write("\tFPRI CRC Check --> Success")#pass;
                                                            else:
                                                                Output.write("\tMismatch found in FPRICRC file!!!")
                                                        else:
                                                            Output.write("\tThere is difference in file names - FPRICRC!!!")
                                                    else:
                                                        Output.write("\tFPRICRC Folder is EMPTY!!!")
                                                else:
                                                    Output.write("\tFPRICRC Folder is missing!!!")

                                            #EFS_CRC
                                            if self.efscrc.IsChecked() == True:
                                                Output.write("\n\n")
                                                Output.write("EFSCRC Test\n")
                                                Output.write("-----------\n")
                                                dlg = wx.MessageDialog(self,"Create EFSCRC through Hidden Menu then Proceed\nFile Created?", "Confirmation", wx.YES|wx.NO|wx.ICON_QUESTION)
                                                res = dlg.ShowModal()
                                                dlg.Destroy()
                                                if res == wx.ID_YES:
                                                    RemotePath_EFS = Remote_EFSCRC_Path[0:len(Remote_EFSCRC_Path)-1]
                                                    x = AndroidDebugBridge();
                                                    x.pull(RemotePath_EFS,LocalPath_EFS)
                                                    list_t = os.listdir('./')
                                                    if list_t.count("EFS_CRC.txt") >= 1 : #pulled file is present properly
                                                        if validate_EFS_FPRI_Check(EFSCRC_File_Path,EFSCRC_File,"EFS",True,Output)== True:
                                                            Output.write("\tEFSCRC Test --> Success")
                                                        else:
                                                            Output.write("\tEFSCRC Test --> Failed!!!")
                                                    else:
                                                        self.logger.AppendText("Seems to be EFS CRC File Not Copied/Generated Properly by tool. Check Local Config path and Phone\n")
                                                        Output.write("Note: Only file presence has been checked. This file is not created by tool and checked.\n")
                                                        if validate_EFS_FPRI_Check(EFSCRC_File_Path,EFSCRC_File,"EFS",False,Output)== True:
                                                            Output.write("\tEFSCRC Test --> Success")
                                                        else:
                                                            Output.write("\tEFSCRC Test --> Failed!!!")
                                                else:
                                                    self.logger.AppendText("User Skipped generation/Creation of EFSCRC\n")
                                                    Output.write("Note: Only file presence has been checked. This file is not created by tool and checked.\n")
                                                    if validate_EFS_FPRI_Check(EFSCRC_File_Path,EFSCRC_File,"EFS",False,Output)== True:
                                                        Output.write("\tEFSCRC Test --> Success")
                                                    else:
                                                        Output.write("\tEFSCRC Test --> Failed!!!")

                                            #FPRI_Test cases
                                            if self.fpritest.IsChecked() == True:
                                                Output.write("\n\n")
                                                Output.write("FPRI Test Cases\n")
                                                Output.write("---------------\n")
                                                dlg = wx.MessageDialog(self,"Create FPRITestCases through Hidden Menu then Proceed\nFile Created?", "Confirmation", wx.YES|wx.NO|wx.ICON_QUESTION)
                                                res = dlg.ShowModal()
                                                dlg.Destroy()
                                                if res == wx.ID_YES:
                                                    RemotePath_FPRI = Remote_FPRITest_Path[0:len(Remote_FPRITest_Path)-1]
                                                    x = AndroidDebugBridge();
                                                    x.pull(RemotePath_FPRI,LocalPath_FPRI)
                                                    list_t = os.listdir('./')
                                                    if list_t.count("FPRI_TestCase.txt") >= 1 : #pulled file is present properly
                                                        if validate_EFS_FPRI_Check(FPRI_File_Path,FPRI_File,"FPRI",True,Output)== True:
                                                            Output.write("\tFPRI Test Cases --> Success")
                                                        else:
                                                            Output.write("\tFPRI Test Cases --> Failed!!!")
                                                    else:
                                                        self.logger.AppendText("Seems to be FPRI TestCase File Not Copied/Generated Properly by tool. Check Local Config path and Phone\n")
                                                        Output.write("Note: Only file presence and failed cases have been checked. This file is not created by tool and checked.\n")
                                                        if validate_EFS_FPRI_Check(FPRI_File_Path,FPRI_File,"FPRI",False,Output)== True:
                                                            Output.write("\tFPRI Test Cases --> Success")
                                                        else:
                                                            Output.write("\tFPRI Test Cases --> Failed!!!")
                                                else:
                                                    self.logger.AppendText("User Skipped generation/Creation of FPRI Testcase\n")
                                                    Output.write("Note: Only file presence and failed cases have been checked. This file is not created by tool and checked.\n")
                                                    if validate_EFS_FPRI_Check(FPRI_File_Path,FPRI_File,"FPRI",False,Output)== True:
                                                        Output.write("\tFPRI Test Cases --> Success")
                                                    else:
                                                        Output.write("\tFPRI Test Cases --> Failed!!!")

                                            #BIN
                                            if self.bin.IsChecked() == True:
                                                Output.write("\n\n")
                                                Output.write("BIN File Test\n")
                                                Output.write("-------------\n")
                                                if list_of_files.count(BIN_Name) >= 1:
                                                    if verify_ZIPFile_Content(User_BIN_Path+BIN_Name,"BIN",SWV,Output) == True:
                                                        Output.write("\tBIN File Check --> Success")#pass;# here further we need to include to check the logic of ZIP file
                                                    else:
                                                        Output.write("\tBIN File Check --> Failed!!!")
                                                else:
                                                    Output.write("\t"+ BIN_Name + " File is MISSING or Check the file naming convention!!!")
                                            #ROM
                                            if self.rom.IsChecked() == True:
                                                Output.write("\n\n")
                                                Output.write("ROM File Test\n")
                                                Output.write("-------------\n")
                                                if list_of_files.count(ROM_Name) >= 1:
                                                    if verify_ZIPFile_Content(User_BIN_Path+ROM_Name,"ROM",SWV,Output)==True:
                                                        Output.write("\tROM File Check --> Success")#pass;# here further we need to include to check the logic of ZIP file
                                                    else:
                                                        Output.write("\tROM File Check --> Failed!!!")
                                                else:
                                                    Output.write("\t>>"+ ROM_Name + " File is MISSING or Check the file naming convention!!!")
                                            #DZ
                                            if self.dz.IsChecked() == True:
                                                Output.write("\n\n")
                                                Output.write("DZ File Test\n")
                                                Output.write("------------\n")
                                                if list_of_files.count(DZ_Name) >= 1:
                                                    if verify_ZIPFile_Content(User_BIN_Path+DZ_Name,"DZ",'DZ_'+SWV+'.dz',Output)==True:
                                                        Output.write("\tDZ File Check --> Success")#pass;# here further we need to include to check the logic of ZIP file
                                                    else:
                                                        Output.write("\tDZ File Check --> Failed!!!")
                                                else:
                                                    Output.write("\t>>"+ DZ_Name + " File is MISSING or Check the file naming convention!!!")
                                            #BIN_DZ Check
                                            if self.bin_dz.IsChecked() == True:
                                                Output.write("\n\n")
                                                Output.write("BIN-DZ File Test\n")
                                                Output.write("----------------\n")
                                                if list_of_files.count(BIN_Name) >= 1:
                                                    if verify_ZIPFile_Content(User_BIN_Path+BIN_Name,"DZ", SWV+'.dz',Output) == True:
                                                        Output.write("\tBIN-DZ File Check --> Success")#pass;# here further we need to include to check the logic of ZIP file
                                                    else:
                                                        Output.write("\tBIN-DZ File Check --> Failed!!!")
                                                else:
                                                    Output.write("\t>>"+ BIN_Name + " File is MISSING or Check the file naming convention!!!")

                                            #MMC
                                            if self.mmc.IsChecked() == True:
                                                MMC = get_ValuefromXLS(User_PR_Path+User_PR_Filename, trim_value(LocalLines[31].split('$')[1]))
                                                MMC_TYPE = get_ValuefromXLS(User_PR_Path+User_PR_Filename, trim_value(LocalLines[32].split('$')[1]))
                                                MMC_SIZE = get_ValuefromXLS(User_PR_Path+User_PR_Filename, trim_value(LocalLines[33].split('$')[1]))
                                                MOD = get_ValuefromXLS(User_PR_Path+User_PR_Filename, trim_value(LocalLines[28].split('$')[1]))

                                                Output.write("\n\n")
                                                Output.write("MMC File Test\n")
                                                Output.write("-------------\n")
                                                if MMC !='' and MMC_TYPE != '':
                                                    MMC_Name = MOD+"_"+Suffix+"_"+remove_BlankSpace(MMC_TYPE)+"_"+remove_BlankSpace(str(int(MMC_SIZE)))+".zip"
                                                    if list_of_files.count(MMC_Name) >= 1:
                                                        if zipfile.is_zipfile(User_BIN_Path+MMC_Name)== True:
                                                            Output.write("\tMMC File Check --> Success")#pass;# here further we need to include to check the logic of ZIP file
                                                        else:
                                                            Output.write("\t>>ZIP file is seems to be not a valid ZIP file. Please Check!!!\n")
                                                            Output.write("\tMMC File Check --> Failed!!!")
                                                    else:
                                                        Output.write("\t>>"+ MMC_Name + " File is MISSING or Check the file naming convention!!!")
                                                else:
                                                    Output.write("\tPR does not have any info regarding MMC Data. Please Check your Model Config file that whether you are really having MMC data for this model.")
                                            #PRI
                                            if self.pri.IsChecked() == True:
                                                DRM = get_ValuefromXLS(User_PR_Path+User_PR_Filename, trim_value(LocalLines[37].split('$')[1]))
                                                NWL = get_ValuefromXLS(User_PR_Path+User_PR_Filename, trim_value(LocalLines[36].split('$')[1]))

                                                Output.write("\n\n")
                                                Output.write("PRI Test\n")
                                                Output.write("---------\n")
                                                if list_of_files.count(PRI_Name) >= 1: # File name is present or not....
                                                    Output.write("Note: Only BIN Folder Factory PRI Content has been verified. Make sure that you are coping the same in to DOC also.\n")
                                                    if is_SystemFilesPresent(DOC_PRI_Path,PRI_Name) == True:
                                                        PRI_Content_Check_Module(User_BIN_Path+PRI_Name,User_PR_Path+User_PR_Filename, Output, SWV, SWOV, Suffix, DRM, NTCODE, SL, NWL);
                                                    else:
                                                        Output.write("\t>>Factory PRI is missing in DOC or Check the Naming Convention!!!\n")
                                                        Output.write("\tPRI File Check --> Failed!!!")
                                                else:
                                                    Output.write("\t"+ PRI_Name + " File is MISSING in BIN!!! or Check the file naming convention!!!\n")
                                                    Output.write("\tPRI File Check --> Failed!!!")
                                            #Sanity Report
                                            if self.SanityRep.IsChecked() == True:
                                                Output.write("\n\n")
                                                Output.write("RnD Report Test\n")
                                                Output.write("---------------\n")
                                                if list_of_files.count(SanReport_Name) >= 1:
                                                    if is_SystemFilesPresent(DOC_SanityRep_Path,SanReport_Name) == True:
                                                        Output.write("\tRnD Report File Check --> Success")#pass;
                                                    else:
                                                        Output.write("\t>>Sanity Report is missing in DOC or Check the Naming Convention!!!\n")
                                                        Output.write("\tRnD Report File Check --> Failed!!!")
                                                else:
                                                    Output.write("\t"+ SanReport_Name + " File is MISSING in BIN or Check the file naming convention!!!\n")
                                                    Output.write("\tRnD Report File Check --> Failed!!!")
                                            #KDZ
                                            if self.kdz.IsChecked() == True:
                                                Output.write("\n\n")
                                                Output.write("KDZ Test\n")
                                                Output.write("--------\n")
                                                if list_of_files.count(KDZ_Name1) >= 1 :
                                                    if list_of_files.count(KDZ_Name2) >= 1 :
                                                        Output.write("\tKDZ File Check --> Success")#pass;
                                                    else:
                                                        Output.write("\t"+ KDZ_Name2 + " File is MISSING or Check the file naming convention!!!")
                                                else:
                                                    Output.write("\t>>"+ KDZ_Name1 + " File is MISSING or Check the file naming convention!!!")
                                            #FOTA
                                            if self.fota.IsChecked() == True:
                                                Output.write("\n\n")
                                                Output.write("FOTA File Test\n")
                                                Output.write("--------------\n")
                                                if list_of_files.count(FOTA_Name) >= 1:
                                                    Output.write("\tFOTA File Check --> Success")#pass;
                                                else:
                                                    Output.write("\t"+ FOTA_Name + " File is MISSING or Check the file naming convention!!!")
                                            #DLL
                                            if self.dll.IsChecked() == True:
                                                Output.write("\n\n")
                                                Output.write("WEB DLL Test\n")
                                                Output.write("------------\n")
                                                if is_DLLfile_Present(list_of_files) == False:
                                                    Output.write("\t>>WEB DLL file is MISSING or Check the file naming convention!!!")
                                                else:
                                                    Output.write("\tDLL File Check --> Success")
                                            #SRN
                                            if self.swrelnote.IsChecked() == True:
                                                Output.write("\n\n")
                                                Output.write("SW Relese Note Test\n")
                                                Output.write("-------------------\n")
                                                if is_SystemFilesPresent(SRN_Path,SRN_File) == False:
                                                    Output.write("\t>>SRN file is MISSING or Check the file naming convention!!!")
                                                else:
                                                    Output.write("Note: However, manual verification is required for Devlog/CodeDifflog/etc. This is just a file presence check.\n")
                                                    Output.write("\tSRN File Check --> Success")

                                            #SW Req Doc
                                            if self.swreqdoc.IsChecked() == True:
                                                Output.write("\n\n")
                                                Output.write("SW Request Doc\n")
                                                Output.write("--------------\n")
                                                if make_SW_Req_Doc(self, User_PR_Path, User_PR_Filename, SWV, QM_info, Tester_info, User_DOC_Path,ser,'softreviewer')== True:
#                                                    Output.write("\tSW Request Doc Created --> Success\n")
                                                    if is_SystemFilesPresent(User_SW_Req_Doc_Path,User_SW_Req_Doc_File) == True:
                                                        if User_SW_Req_Doc_File == "SW_Req_"+SWV+".txt":
                                                            pass;
                                                        else:
                                                            Output.write("\t>>SW_Req_"+SWV+".txt file naming Convention is not Correct!!!\n")
                                                    else:
                                                        Output.write("\t>>SW Request Doc Missing!!! Please Check\n")

                                                    if filecmp.cmp("./SW_Req_"+SWV+".txt",User_SW_Req_Doc_Path+User_SW_Req_Doc_File) == True:
                                                        Output.write("\tSW Request Doc Comparision --> Success")
                                                    else:
                                                        Output.write("\tFound Mismatch in content--> Fail")
                                                else:
                                                    Output.write("\tSW Request Doc Cration Failed. --> Fail")

                                            #Tag Comparision
                                            if self.tagcompare.IsChecked() == True:
                                                Output.write("\n\n")
                                                Output.write("Tag Comparision\n")
                                                Output.write("--------------\n")
                                                if is_SystemFilesPresent(DOC_TagCompare_Path,"Tag_Compare"+SWV+".xls") == True:
                                                    Output.write("\t>>Tag Comaprision File Check --> Success")
                                                else:
                                                    Output.write("\tTag Comaprision File is Missing!!! Please Check\n")
                                                    Output.write("\t>>Tag Comaprision File Check --> Fail")

                                            Output.close()

                                            dlg = wx.MessageDialog(self, "Done", "Result", wx.OK | wx.ICON_INFORMATION )
                                            dlg.ShowModal()
                                            dlg.Destroy()

                                            ser.close()
                                        else:
                                            self.logger.AppendText("Failed in Secondary Setup [while taking DBCRC/FILECRC/FPRICRC]\n")
                                    else:
                                        self.logger.AppendText("Failed in doing SIMLOCK\n")
                                else:
                                    self.logger.AppendText("Failed in writing NTCODE\n")
                            else:
                                self.logger.AppendText("Failed in doing IDDE. Check your HW\n")
                        else:
                            self.logger.AppendText("User Cancelled the execution due to non-compilance of initial setup.\n")
                    else:
                        self.logger.AppendText("PR file is missing.\n")
                else:
                    self.logger.AppendText("Local_Config.txt file is missing\n")
                ser.close()            
            else:
                self.logger.AppendText("Please Check your Port Settings\n")
            ser.close()                
        else:
            self.logger.AppendText("Input Validation Failed\n")

        self.gifleft.Stop()
        self.gifright.Stop()    
        self.gifleft.Hide()
        self.gifright.Hide()
        self.Layout()

#*********************************************************************************************************************************

def write_IDDE(self, ser):
    "Function to write IDDE"
    ser.write('at%idde\r\n')
    out = ''
    time.sleep(2)
    while ser.inWaiting() > 0:
        out += ser.read(2)
    #Process Output
    out = out.split('\n')
    if (out[1].find(' OK') != -1):
        self.logger.AppendText("IDDE Success\n")
        return True;
    else:
        return False;

def write_NTCODE(self, ser,NTCODE):
    "Function to write NTCODE"
    if NTCODE == "\"0\",\"FFF,FFF,FFFFFFFF,FFFFFFFF,FF\"":
        self.logger.AppendText("NT CODE Writing Skipped due to OPEN version\n")
        return True;
    else:
        ser.write('at%ntcode='+NTCODE+'\r\n')
        out = ''
        time.sleep(2)
        while ser.inWaiting() > 0:
            out += ser.read(2)
        #Process Output
        out = out.split('\n')
        if (out[1].find(' OK') != -1):
            self.logger.AppendText("NTCODE Write Success\n")
            return True;
        else:
            return False;

def do_SIMLOCK(self, ser,SL):
    "Function to write SIMLOCK"
    if (SL == 'Y' or SL == 'Yes' or SL == 'yes' or SL == 'YES' or SL == '(YES)' or SL == '(Yes)'):
        ser.write('at%sltype=1\r\n')
        out = ''
        time.sleep(2)
        while ser.inWaiting() > 0:
            out += ser.read(2)
        #Process Output
        out = out.split('\n')
        if (out[1].find(' OK') != -1):
            self.logger.AppendText("SIM LOCK Done\n")
            return True;
        else:
            return False;
    else:
        self.logger.AppendText("This is Non-SIMLOCK Version\n")
        return True;

def populate_COMPORT(self):
    "This function gives a list of available COM ports"
    cmd = AndroidDebugBridge();
    ret = cmd.execute_pythoncmd("-m serial.tools.list_ports")
    List = ret.split('\n')
    List = List[0:len(List)-2]
    for i in range (len(List)):
        List[i]=List[i].rstrip()
    return List;

def validate_Input(self,callfrom):
    "Validating input. Returns True if there is no problem in input"
    Message = ''
    status = True
    if self.edithear.GetValue() =='' or self.edithear.GetValue() ==' ' or self.edithear.GetValue() =='  ':
        Message = "Select valid COM PORT"
        status = False;
    elif callfrom == 'makecrc':
        if self.dbcrc.GetValue() == False and self.filecrc.GetValue() == False and self.fpricrc.GetValue() == False and self.efscrc.GetValue() == False and self.fpritest.GetValue() == False and self.swreqdoc.GetValue() == False:
            Message = "Select atleast one CheckBox" 
            status = False;
    elif callfrom == 'softreviewer':
        if self.dbcrc.GetValue() or self.filecrc.GetValue() or self.fpricrc.GetValue() or self.efscrc.GetValue() or self.fpritest.GetValue() or self.swreqdoc.GetValue() or  self.SanityRep.GetValue() or self.bin.GetValue() or self.rom.GetValue() or self.dz.GetValue() or self.bin_dz.GetValue() or self.kdz.GetValue() or self.fota.GetValue() or self.dll.GetValue() or self.mmc.GetValue() or self.pri.GetValue() or self.swrelnote.GetValue():
            pass;
        else:
            Message = "Select atleast one CheckBox" 
            status = False;
    
    if status == False:
        dlg = wx.MessageDialog(self, Message, "Warning", wx.OK)
        dlg.ShowModal()
        dlg.Destroy()

    return status;

def is_SystemFilesPresent(path, file):
    "File presence check"
    list_s = os.listdir(path)
    if list_s.count(file) >= 1:
        return True;
    else:
        return False;

def is_DLLfile_Present(list_of_files):
    "This function navigate through user folder and try to find DLL file"
    for index in range(len(list_of_files)):
        if list_of_files[index].endswith(".dll")== True:
            return True;
    return False;

def display_Warning(self):
    "To display warning"
    status = False
    dlg = wx.MessageDialog(self,"1. Make sure that your HW is Flashed with intented Binary\n2. Make sure that you have done Factory Reset without SIM before connecting to this tool\n3. Make sure that you have updated PR with latest info\n4. Make sure that you have updated Local Config as per your Local settings\n\n Proceed?", "Warning", wx.YES|wx.NO|wx.ICON_QUESTION)
    result = dlg.ShowModal()
    dlg.Destroy()
    if result == wx.ID_YES:
        self.logger.AppendText("User Pressed YES\n")
        status = True;
    if result == wx.ID_NO:
        self.logger.AppendText("User Pressed NO\n")    
        status = False;

    return status;

def do_SystemCleanup(self):
    "System clean up"
    file = "./res_makecrc/Local_Config.txt"

    File_Obj_Local = open(file,'r')
    Lines = File_Obj_Local.readlines()
    File_Obj_Local.close()

    list_root = os.listdir('./')
    if list_root.count("DBCRC") >=1 :
        list_sub1 = os.listdir('./DBCRC/')
        if len(list_sub1) != 0 :
            os.remove('./DBCRC/'+list_sub1[0])
        os.rmdir("DBCRC")
    if list_root.count("FILECRC") >=1 :
        list_sub2 = os.listdir('./FILECRC/')
        if len(list_sub2) != 0 :
            os.remove('./FILECRC/'+list_sub2[0])
        os.rmdir("FILECRC")
    if list_root.count("FPRICRC") >=1 :
        list_sub3 = os.listdir('./FPRICRC/')
        if len(list_sub3) != 0 :
            os.remove('./FPRICRC/'+list_sub3[0])
        os.rmdir("FPRICRC")


    Temp_t = Lines[6].split('$')[1]
    EFS_File = Temp_t[0:len(Temp_t)-1]
    Temp_s = Lines[7].split('$')[1]
    FPRI_File = Temp_s[0:len(Temp_s)-1]
    if list_root.count(EFS_File) >= 1:
        os.remove(EFS_File)
    if list_root.count(FPRI_File) >= 1:
        os.remove(FPRI_File)

    if list_root.count("EFS_CRC.txt") >= 1:
        if is_SystemFilesPresent('./',"EFS_CRC.txt") == True:
            os.remove("EFS_CRC.txt")
    if list_root.count("FPRI_TestCase.txt") >= 1:
        if is_SystemFilesPresent('./',"FPRI_TestCase.txt") == True:
            os.remove("FPRI_TestCase.txt")
    if list_root.count("Review_Result.txt") >= 1:
        os.remove("Review_Result.txt")
        
    for index in range(len(list_root)):
        if list_root[index].startswith("SW_Req_Doc_")== True:
            os.remove(list_root[index])

    self.logger.AppendText("File System Cleaned!!!!!!\n")

def Execute_Core_Logic(self, ser):
    "This is a core function which does everything"

    result = True
    file="./res_makecrc/Local_Config.txt"
    #Local Config File OPEN
    File_Obj_Local = open(file,'r')
    LocalLines = File_Obj_Local.readlines()
    File_Obj_Local.close()

    PR_File = LocalLines[0].split('$')[1]
    PR_File = PR_File[0:len(PR_File)-1]

    DBCRC = LocalLines[1].split('$')[1]
    DBCRC_Template = DBCRC[0:len(DBCRC)-1]
    DBCRC_Fname = LocalLines[1].split('$')[2]
    DBCRC_Filename = DBCRC_Fname[0:len(DBCRC_Fname)-1]

    FILECRC = LocalLines[2].split('$')[1]
    FILECRC_Template = FILECRC[0:len(FILECRC)-1]
    FILECRC_Fname = LocalLines[2].split('$')[2]
    FILECRC_Filename = FILECRC_Fname[0:len(FILECRC_Fname)-1]

    FPRICRC = LocalLines[3].split('$')[1]
    FPRICRC_Template = FPRICRC[0:len(FPRICRC)-1]
    FPRICRC_Fname = LocalLines[3].split('$')[2]
    FPRICRC_Filename = FPRICRC_Fname[0:len(FPRICRC_Fname)-1]

    Remote_EFSCRC = LocalLines[4].split('$')[1]
    Remote_EFSCRC_Path = Remote_EFSCRC[0:len(Remote_EFSCRC)-1]
    Remote_FPRITest = LocalLines[5].split('$')[1]
    Remote_FPRITest_Path = Remote_FPRITest[0:len(Remote_FPRITest)-1]
    Temp_t = LocalLines[6].split('$')[1]
    EFS_File = Temp_t[0:len(Temp_t)-1]
    Temp_s = LocalLines[7].split('$')[1]
    FPRI_File = Temp_s[0:len(Temp_s)-1]

    SW_Req_Doc = LocalLines[8].split('$')[0]
    SW_Req_Doc = SW_Req_Doc[0:len(SW_Req_Doc)-1]
    QM_info = LocalLines[8].split('$')[1].split('/')[0]
    QM_info = QM_info[0:len(QM_info)-1]
    Tester_info = LocalLines[8].split('$')[1].split('/')[1]
    Tester_info = Tester_info[0:len(Tester_info)-1]

    User_DOC_Path = LocalLines[9].split('$')[1]
    User_DOC_Path = User_DOC_Path[0:len(User_DOC_Path)-1]

    # at%SWV
    ser.write('at%swv\r\n')
    out = temp = ''
    time.sleep(2)
    while ser.inWaiting() > 0:
        out += ser.read(2)
    #Process Output
    out = out.split()
    temp = out[1]
    SWV_Str= temp[1:len(temp)-1]

    print SWV_Str

    # at%DBCRC
    if self.dbcrc.IsChecked() == True:
        ser.write('at%dbchk\r\n')
        out = temp = ''
        time.sleep(2)
        while ser.inWaiting() > 0:
            out += ser.read(2)
        #Process Output
        out = out.split()
        for index in range(len(out)):
            if ((len(out[index])>5) and (out[index].startswith('at%')!= True)):
                temp = out[index]
                break;
        DBCRC_Str = temp[1:len(temp)-1]
        print DBCRC_Str
        DBCRC_Filename = DBCRC_Filename.replace('FILENAME',DBCRC_Str);
        
        #Writing it to FILE
        os.mkdir("DBCRC")
        CRC = open('DBCRC/'+DBCRC_Filename+'.txt','w')
        DBCRC_Template = DBCRC_Template.replace('SWV ',SWV_Str+'\n');
        DBCRC_Template = DBCRC_Template.replace('VALUE',DBCRC_Str);
        CRC.write(DBCRC_Template)
        CRC.close()


    # at%FILECRC
    if self.filecrc.IsChecked() == True:
        ser.write('at%filecrc\r\n')
        out = temp = ''
        time.sleep(2)
        while ser.inWaiting() > 0:
            out += ser.read(2)
        #Process Output
        out = out.split()
        for index in range(len(out)):
            if ((len(out[index])>5) and (out[index].startswith('at%')!= True)):
                temp = out[index]
                break;
        FILECRC_Str = temp[1:len(temp)-1]
        print FILECRC_Str
        FILECRC_Filename = FILECRC_Filename.replace('FILENAME',FILECRC_Str);

        #Writing it to FILE
        os.mkdir("FILECRC")
        CRC = open('FILECRC/'+FILECRC_Filename+'.txt','w')
        FILECRC_Template = FILECRC_Template.replace('SWV ',SWV_Str+'\n');
        FILECRC_Template = FILECRC_Template.replace('VALUE',FILECRC_Str);
        CRC.write(FILECRC_Template)
        CRC.close()

    # at%FPRICRC
    if self.fpricrc.IsChecked() == True:
        ser.write('at%fpricrc\r\n')
        out = temp = ''
        time.sleep(2)
        while ser.inWaiting() > 0:
            out += ser.read(2)
        #Process Output
        out = out.split()
        for index in range(len(out)):
            if ((len(out[index])>5) and (out[index].startswith('at%')!= True)):
                temp = out[index]
                break;
        FPRICRC_Str = temp[1:len(temp)-1]
        print FPRICRC_Str
        FPRICRC_Filename = FPRICRC_Filename.replace('FILENAME',FPRICRC_Str);

        #Writing it to FILE
        os.mkdir("FPRICRC")
        CRC = open('FPRICRC/'+FPRICRC_Filename+'.txt','w')
        FPRICRC_Template = FPRICRC_Template.replace('SWV ',SWV_Str+'\n');
        FPRICRC_Template = FPRICRC_Template.replace('VALUE',FPRICRC_Str);
        CRC.write(FPRICRC_Template)
        CRC.close()
    
    #EFSCRC
    if self.efscrc.IsChecked() == True:
        dlg = wx.MessageDialog(self,"Create EFS CRC through Hidden Menu then Proceed\nFile Created?", "Confirmation", wx.YES|wx.NO|wx.ICON_QUESTION)
        res = dlg.ShowModal()
        dlg.Destroy()
        if res == wx.ID_YES:
            x = AndroidDebugBridge();
            x.pull(Remote_EFSCRC_Path,os.getcwd()+'\\'+EFS_File)
            if is_SystemFilesPresent("./", EFS_File) == True:
                self.logger.AppendText("ECSCRC File Generated/Copied Successfully\n")
            else:
                self.logger.AppendText("Problem in file Generation/Copy. Please Check Local Config path configuration\n")
                result = False;
        else:
            self.logger.AppendText("User Pressed NO. EFSCRC skipped\n")
            result = False;

    #FPRI Test Cases
    if self.fpritest.IsChecked() == True:
        dlg = wx.MessageDialog(self,"Create FPRI TEST CASES through Hidden Menu then Proceed\nFile Created?", "Confirmation", wx.YES|wx.NO|wx.ICON_QUESTION)
        res = dlg.ShowModal()
        dlg.Destroy()
        if res == wx.ID_YES:
            x = AndroidDebugBridge();
            x.pull(Remote_FPRITest_Path,os.getcwd()+'\\'+FPRI_File)
            if is_SystemFilesPresent("./", FPRI_File) == True:
                self.logger.AppendText("FPRITestCase File Generated/Copied Successfully\n")
            else:
                self.logger.AppendText("Problem in file Generation/Copy. Please Check Local Config path configuration\n")
                result = False;
        else:
            self.logger.AppendText("User Pressed NO. FPRITestCases skipped\n")
            result = False;

    #Make SW Req Doc
    if self.swreqdoc.IsChecked() == True:
        result = make_SW_Req_Doc(self, "./res_makecrc/", PR_File, SWV_Str, QM_info, Tester_info, User_DOC_Path,ser,'makecrc');

    return result;

def make_SW_Req_Doc(self, PR_Path, PR_File, SWV, QM, ENGINEER, User_DOC_Path,ser,callfrom):
    "To make SW Req Doc"
    result = False;
    LocalLines =''
    Country = ''
    Suffix = ''
    Model = ''
    PRNO = ''
    SIMLOCK = ''
    
    if is_SystemFilesPresent(PR_Path, PR_File) == True:
        File_src = open("./res_common/SW_Req_Doc_Template.txt",'r')
        File_tgt = open("SW_Req_"+SWV+".txt",'w')

#        wb = xlrd.open_workbook(PR_Path+PR_File)
#        sh = wb.sheet_by_index(0)

        if callfrom == 'makecrc':
            File_Obj_Local = open('./res_makecrc/Local_Config.txt','r')
            LocalLines = File_Obj_Local.readlines()
            File_Obj_Local.close()       
            Country = get_ValuefromXLS(PR_Path+PR_File, trim_value(LocalLines[20].split('$')[1]))
            Suffix = get_ValuefromXLS(PR_Path+PR_File, trim_value(LocalLines[19].split('$')[1]))
            Model = get_ValuefromXLS(PR_Path+PR_File, trim_value(LocalLines[18].split('$')[1]))
            PRNO = get_ValuefromXLS(PR_Path+PR_File, trim_value(LocalLines[17].split('$')[1]))
            SIMLOCK = get_ValuefromXLS(PR_Path+PR_File, trim_value(LocalLines[24].split('$')[1]))
        else:
            File_Obj_Local = open('./res_softreviewer/Local_Config.txt','r')
            LocalLines = File_Obj_Local.readlines()
            File_Obj_Local.close()       
            Country = get_ValuefromXLS(PR_Path+PR_File, trim_value(LocalLines[30].split('$')[1]))
            Suffix = get_ValuefromXLS(PR_Path+PR_File, trim_value(LocalLines[29].split('$')[1]))
            Model = get_ValuefromXLS(PR_Path+PR_File, trim_value(LocalLines[28].split('$')[1]))
            PRNO = get_ValuefromXLS(PR_Path+PR_File, trim_value(LocalLines[27].split('$')[1]))
            SIMLOCK = get_ValuefromXLS(PR_Path+PR_File, trim_value(LocalLines[34].split('$')[1]))

        for index in range(23):
            line =''
            line = File_src.next()
            if line == '\n':
                if index == 21:
                    List_of_files = os.listdir(User_DOC_Path)
                    for x in range (len(List_of_files)):
                        temp = List_of_files[x]
                        File_tgt.write(temp+'\n')
                else:
                    File_tgt.write(line)
            else:
                line = line[0:len(line)-1]#to remove last \n char
                if index == 0:
                    line = line + Model +'/'+ Country +'/'+ Suffix + '\n'
                elif index == 2:
                    line = line + QM + '\n'
                elif index == 6:
                    line = line + SWV + '\n'
                elif index == 7:
                    line = line + get_SWOV(ser) + '\n'
                elif index == 8:
                    line = line + get_SWCV(ser) + '\n'
                elif index == 9:
                    line = line + PRNO +'\n'
                elif index == 10:
                    line = line + SIMLOCK +'\n'
                elif index == 11:
                    line = line + get_CRC("DBCRC") +'\n'
                elif index == 12:
                    line = line + get_CRC("FILECRC") +'\n'
                elif index == 13:
                    line = line + get_CRC("FPRICRC") +'\n'
                elif index == 14:
                    line = line + 'YES \n'
                elif index == 15:
                    self.logger.AppendText("Connecting to ADB...\n")
                    y = AndroidDebugBridge()
                    FingerPrint = y.call_adb("shell getprop ro.build.fingerprint")
                    line = line + FingerPrint +'\n'
                    self.logger.AppendText("Exiting from ADB...\n")
                elif index == 17:
                    line = line + ENGINEER + '\n'
                else:
                    line = line+'\n'

                File_tgt.write(line)

        File_tgt.close()
        File_src.close()  
        self.logger.AppendText("SW Req Document Created\n")
        result = True;
    else:
        self.logger.AppendText("Skipped SW Req Doc due to unavailability of PR file\n")
        result = False;

    return result;

def get_SWV(ser):
    "This is to get the SWV"
    ser.write('at%swv\r\n')
    out = temp = ''
    time.sleep(2)
    while ser.inWaiting() > 0:
        out += ser.read(2)
    #Process Output
    out = out.split()
    temp = out[1]
    SWV = temp[1:len(temp)-1]
    return SWV;

def get_SWOV(ser):
    "This is to get the SWOV"
    ser.write('at%swov\r\n')
    out = temp = ''
    time.sleep(2)
    while ser.inWaiting() > 0:
        out += ser.read(2)
    #Process Output
    out = out.split()
    temp = out[1]
    SWOV = temp[1:len(temp)-1]
    return SWOV;

def get_SWCV(ser):
    "This is to get the SWCV"
    ser.write('at%swcv\r\n')
    out = temp = ''
    time.sleep(2)
    while ser.inWaiting() > 0:
        out += ser.read(2)
    #Process Output
    if out.find('ERROR') != -1 or out.find('Error') != -1 or len(out) <= 20:
        return "NA";
    else:
        out = out.split()
        temp = out[1]
        SWCV = temp[1:len(temp)-1]
        return SWCV;

def get_CRC(str):
    "test"
    if str == "DBCRC":
        list = os.listdir('./DBCRC/')
        Temp = list[0][0:len(list[0])-4]
        return Temp
    elif str =="FILECRC":
        list = os.listdir('./FILECRC/')
        Temp = list[0][0:len(list[0])-4]
        return Temp
    elif str =="FPRICRC":
        list = os.listdir('./FPRICRC/')
        Temp = list[0][0:len(list[0])-4]
        return Temp
    else:
        return "NULL"

def get_SRC_VER_info(Prev_SWV):
    "This function gets input regarding previous binary"
    Temp = Prev_SWV.split('-');
    Prev_Ver = Temp[2]
    Prev_Ver_Day = Temp[6]
    return Prev_Ver.capitalize()+"_"+Convert_Month_To_Number(Temp[5])+Prev_Ver_Day;

def get_KDZName(SWV):
    "This is a function which extracts the name from version name for KDZ"
    Temp = SWV.split('-')
    return Temp[2].upper();

def get_Modelinfo(Model):
    "To extract model info from version info"
#    Temp1 = SWV.split('-')
#    Temp2 = Temp1[0]
    return Model[2:len(Model)];

def get_CURR_VER_info(SWV):
    "This function gets input regarding previous binary"
    Temp = SWV.split('-');
    Curr_Ver = Temp[2]
    Curr_Ver_Day = Temp[6]
    return Curr_Ver.capitalize()+"_"+Convert_Month_To_Number(Temp[5])+Curr_Ver_Day;

def Convert_Month_To_Number(str):
    "Convert Month To Number"
    if str == 'JAN':
        return '01';
    elif str == 'FEB':
        return '02';
    elif str == 'MAR':
        return '03';
    elif str == 'APR':
        return '04';
    elif str == 'MAY':
        return '05';
    elif str == 'JUN':
        return '06';
    elif str == 'JUL':
        return '07';
    elif str == 'AUG':
        return '08';
    elif str == 'SEP':
        return '09';
    elif str == 'OCT':
        return '10';
    elif str == 'NOV':
        return '11';
    elif str == 'DEC':
        return '12';
    else:
        return "NULL";

def Convert_Day_To_Number(str):
    "Convert Day To Number"
    if str == '1':
        return '01';
    elif str == '2':
        return '02';
    elif str == '3':
        return '03';
    elif str == '4':
        return '04';
    elif str == '5':
        return '05';
    elif str == '6':
        return '06';
    elif str == '7':
        return '07';
    elif str == '8':
        return '08';
    elif str == '9':
        return '09';
    else:
        return str;

def remove_BlankSpace(Str):
    "Function to remove spaces for file naming convention"
    String =''
    Temp = Str.split(' ')
    if len(Temp) >= 1:
        for index in range(len(Temp)):
            String = String+Temp[index]

    return String;

def do_SecondarySetup(self, ser):
    ""
    #Local Config File OPEN
    File_Obj_Local = open('./res_softreviewer/Local_Config.txt','r')
    LocalLines = File_Obj_Local.readlines()
    File_Obj_Local.close()

    DBCRC = LocalLines[5].split('$')[1]
    DBCRC_Template = DBCRC[0:len(DBCRC)-1]
    DBCRC_Fname = LocalLines[5].split('$')[2]
    DBCRC_Filename = DBCRC_Fname[0:len(DBCRC_Fname)-1]

    FILECRC = LocalLines[6].split('$')[1]
    FILECRC_Template = FILECRC[0:len(FILECRC)-1]
    FILECRC_Fname = LocalLines[6].split('$')[2]
    FILECRC_Filename = FILECRC_Fname[0:len(FILECRC_Fname)-1]

    FPRICRC = LocalLines[7].split('$')[1]
    FPRICRC_Template = FPRICRC[0:len(FPRICRC)-1]
    FPRICRC_Fname = LocalLines[7].split('$')[2]
    FPRICRC_Filename = FPRICRC_Fname[0:len(FPRICRC_Fname)-1]

    # at%SWV
    ser.write('at%swv\r\n')
    out = temp = ''
    time.sleep(2)
    while ser.inWaiting() > 0:
        out += ser.read(2)
    #Process Output
    out = out.split()
    temp = out[1]
    SWV_Str= temp[1:len(temp)-1]

    # at%DBCRC
    ser.write('at%dbchk\r\n')
    out = temp = ''
    time.sleep(2)
    while ser.inWaiting() > 0:
        out += ser.read(2)
    #Process Output
    out = out.split()
    for index in range(len(out)):
        if ((len(out[index])>5) and (out[index].startswith('at%')!= True)):
            temp = out[index]
            break;
    DBCRC_Str = temp[1:len(temp)-1]
    print DBCRC_Str
    DBCRC_Filename = DBCRC_Filename.replace('FILENAME',DBCRC_Str);
    
#Writing it to FILE
    os.mkdir("DBCRC")
    CRC = open('DBCRC/'+DBCRC_Filename+'.txt','w')
    DBCRC_Template = DBCRC_Template.replace('SWV ',SWV_Str+'\n');
    DBCRC_Template = DBCRC_Template.replace('VALUE',DBCRC_Str);
    CRC.write(DBCRC_Template)
    CRC.close()


    # at%FILECRC
    ser.write('at%filecrc\r\n')
    out = temp = ''
    time.sleep(2)
    while ser.inWaiting() > 0:
        out += ser.read(2)
    #Process Output
    out = out.split()
    for index in range(len(out)):
        if ((len(out[index])>5) and (out[index].startswith('at%')!= True)):
            temp = out[index]
            break;
    FILECRC_Str = temp[1:len(temp)-1]
    print FILECRC_Str
    FILECRC_Filename = FILECRC_Filename.replace('FILENAME',FILECRC_Str);

    #Writing it to FILE
    os.mkdir("FILECRC")
    CRC = open('FILECRC/'+FILECRC_Filename+'.txt','w')
    FILECRC_Template = FILECRC_Template.replace('SWV ',SWV_Str+'\n');
    FILECRC_Template = FILECRC_Template.replace('VALUE',FILECRC_Str);
    CRC.write(FILECRC_Template)
    CRC.close()

    # at%FPRICRC
    ser.write('at%fpricrc\r\n')
    out = temp = ''
    time.sleep(2)
    while ser.inWaiting() > 0:
        out += ser.read(2)
    #Process Output
    out = out.split()
    for index in range(len(out)):
        if ((len(out[index])>5) and (out[index].startswith('at%')!= True)):
            temp = out[index]
            break;
    FPRICRC_Str = temp[1:len(temp)-1]
    print FPRICRC_Str
    FPRICRC_Filename = FPRICRC_Filename.replace('FILENAME',FPRICRC_Str);

    #Writing it to FILE
    os.mkdir("FPRICRC")
    CRC = open('FPRICRC/'+FPRICRC_Filename+'.txt','w')
    FPRICRC_Template = FPRICRC_Template.replace('SWV ',SWV_Str+'\n');
    FPRICRC_Template = FPRICRC_Template.replace('VALUE',FPRICRC_Str);
    CRC.write(FPRICRC_Template)
    CRC.close()

    return True;

def validate_EFS_FPRI_Check(Path, File, Type, ContentCheck, Output):
    "This function will test EFSCRC and FPRI Test cases file"
    is_file_present = False;
    Path = Path[0:len(Path)-1]
    File = File[0:len(File)-1]

    list_t = os.listdir(Path)
    if list_t.count(File) >= 1:
        if Type == "EFS":
            if ContentCheck == True:
                if filecmp.cmp(Path+File,os.getcwd()+"\\EFS_CRC.txt") == True:
                    return True;
                else:
                    Output.write("\tMismatch found in EFS CRC file between User and System generated file!!!\n")
                    return False;
            else:
                return True;

        if Type == "FPRI":
            if ContentCheck == True:
                if filecmp.cmp(Path+File,os.getcwd()+"\\FPRI_TestCase.txt") == True:
                    pass;
                else:
                    Output.write("\tMismatch found in FPRI TestCase file between User and System generated file!!!\n")

            File_Obj = open(Path+File,"r")
            Lines = File_Obj.readlines()
            File_Obj.close()
            Start = len(Lines)-5
            for Start in range (len(Lines)):
                if Lines[Start].find("Fail") != -1 or Lines[Start].find("fail") != -1 :
                    Sub_lines = Lines[Start].split(':')
                    if len(Sub_lines) >= 2:
                        if Sub_lines[1].find(" 0") != -1:
                            return True;
                        else:
                            Output.write("\tFPRI CRC File is having some failed cases. Please Check!!!\n")
                            return False;
            Output.write("\tSorry...Tool is not able to find Failed Cases. Please Check Manualy once.\n")
            return False;
    else:
        if Type == "EFS":
            Output.write("\tEFS CRC File is missing in user given path!!! Check Path or Filename info in Local Config\n")
        if Type == "FPRI":
            Output.write("\tFPRI Test Cases File is missing in user given path!!! Check Path or Filename info in Local Config\n")
        return False;

def verify_ZIPFile_Content(Zip_file, CallFrom, Filename, Output):
    "Zip file analysis "
    
    if zipfile.is_zipfile(Zip_file)== True:
        zf = zipfile.ZipFile(Zip_file,'r')
        list_zf = zf.namelist()

        for index in range (len(list_zf)):
            if '/' in list_zf[index]:
                Output.write("\t>>Found that, %s is zipped with folder Please Check!!!\n"%(CallFrom))
                return False;
        
        if (CallFrom == 'BIN'):
            result = True
            print list_zf
            BIN_List_Expected = extract_Fileinfo(CallFrom, Filename)
            print BIN_List_Expected
            if len(list_zf) != len(BIN_List_Expected):
                result = False
                Output.write("\t>>BIN - Number of files in ZIP is different than Actual. Please Check!!!\n")

            for ind in range (len(BIN_List_Expected)):
                if list_zf.count(BIN_List_Expected[ind]) == 1:
                    pass;
                else:
                    result = False
                    Output.write("\t>>BIN - %s is MISSING!!!!!. Please Check!!!\n"%(BIN_List_Expected[ind]))
            return result;

        elif (CallFrom == 'ROM'):
            result = True
            print list_zf
            ROM_List_Expected = extract_Fileinfo(CallFrom,Filename)
            print ROM_List_Expected
            if len(list_zf) != len(ROM_List_Expected):
                result = False
                Output.write("\t>>ROM - Number of files in ZIP is different than Actual as per Model Config. Please Check!!!\n")

            for ind in range (len(ROM_List_Expected)):
                if list_zf.count(ROM_List_Expected[ind]) == 1:
                    pass;
                else:
                    result = False
                    Output.write("\t>>ROM - %s is MISSING!!!!!. Please Check!!!\n"%(ROM_List_Expected[ind]))

            return result;

        elif (CallFrom == 'DZ'):
            if len(list_zf) == 1 :
                if list_zf[0] == Filename:
                    return True;
                else:
                    Output.write("\t>>Filename inside zipfile is wrong. Please Check!!! [%s/%s]\n" %(list_zf[0],Filename))
                    return False;
            else:
                Output.write("\t>>Found that DZ ZIP file has been zipped with no files or more than one files\n")
                return False;                
        else:
            pass;

        return True;

    else:
        Output.write("\t>>ZIP file is seems to be not a valid ZIP file. Please Check!!!\n")
        return False;

def PRI_Content_Check_Module(PRI, PR, Output,SWV,SWOV, Suffix, DRM, Ntcode, SL, NWL):
    "PRI Content check with PR"
    Error = False;

    #open PR
#    wb_pr = xlrd.open_workbook(PR)
#    sh_pr = wb_pr.sheet_by_index(0)
    
    #open PRI
    wb_pri = xlrd.open_workbook(PRI)
    sh_pri = wb_pri.sheet_by_index(0)
    
    #list of items [Suffix, SWV, SWOV, NTCode, DRM Type]
    list_PR = [Suffix, SWV, SWOV, Ntcode, DRM ]
    list_PRI = [sh_pri.cell(9,5).value, sh_pri.cell(25,5).value, sh_pri.cell(26,5).value, sh_pri.cell(28,5).value, sh_pri.cell(48,5).value]
    
    for index in range (len(list_PR)):
        if list_PR[index] == list_PRI[index]:
            pass;
        else:
            Error = True;
            Output.write("\tFound Mismatch between PR and PRI in "+ list_PR[index]+"\n")
    #SIMLOCK and NETWORK LOCK info check.
    if SL == 'Yes' or SL == '(Yes)' or SL == 'Y':
        if sh_pri.cell(30,4).value == 'Y':
            pass;
        else:
            Error = True;
            Output.write("\tFound Mismatch between PR and PRI in SIM LOCK info\n")

    if SL == 'No' or SL == '(No)' or SL == 'N':
        if sh_pri.cell(30,4).value == 'N':
            pass;
        else:
            Error = True;
            Output.write("\tFound Mismatch between PR and PRI in SIM LOCK info\n")
    
    if NWL == '1.0' or NWL == '1' or NWL == 'Yes' or NWL == 'YES' or NWL == 'yes' or NWL == 'Y' or NWL == 'y':
        if sh_pri.cell(31,4).value == 'Y':
            pass;
        else:
            Error = True;
            Output.write("\tFound Mismatch between PR and PRI in NETWORK LOCK info\n")

    if NWL == '' or NWL == ' ':
        if sh_pri.cell(31,4).value == 'N':
            pass;
        else:
            Error = True;
            Output.write("\tFound Mismatch between PR and PRI in NETWORK LOCK info\n")


    # Config file items check

    if Error == False:
        Output.write("\tPRI Test --> Success")
    else:
        Output.write("\tPRI Test --> FAILED !!!")

def get_ValuefromXLS(PRFILE, STR_TO_SEARCH):
    " "
    found = False
    ret_val = ''

    wb = xlrd.open_workbook(PRFILE)
    sh = wb.sheet_by_index(0)


    for col in range (sh.ncols):
    #while col <= 10:
        if col==1 or col==3 or col == 5:
            continue;
        else:
            row = 0
            for row in range (sh.nrows):
                if sh.cell(row,col).value == STR_TO_SEARCH:
                    ret_val = sh.cell(row,col+1).value
                    found = True
                    break

        if found == True:
            break
    return ret_val;

def trim_value(str):
    ""
    return str[0:len(str)-1]

def extract_Fileinfo(callfrom,SWV):
    " "
    Model_Lines=''
    Final_List = []
    if is_SystemFilesPresent('./res_softreviewer/', 'Model_Config.txt')==True:
        ModelConfig = open('./res_softreviewer/Model_Config.txt','r')
        Model_Lines = ModelConfig.readlines()
        ModelConfig.close()

    if callfrom == "BIN":
        if Model_Lines[24].split(':')[1].find('XXX') == -1:
            Final_List.append((trim_value(Model_Lines[24].split(':')[1])).replace('%SWV%',SWV))
        if Model_Lines[25].split(':')[1].find('XXX') == -1:
            Final_List.append((trim_value(Model_Lines[25].split(':')[1])).replace('%SWV%',SWV))
        if Model_Lines[26].split(':')[1].find('XXX') == -1:
            Final_List.append((trim_value(Model_Lines[26].split(':')[1])).replace('%SWV%',SWV))
        if Model_Lines[27].split(':')[1].find('XXX') == -1:
            Final_List.append((trim_value(Model_Lines[27].split(':')[1])).replace('%SWV%',SWV))
        return Final_List;
    else: #callfrom == ROM
        if Model_Lines[28].split(':')[1].find('XXX') == -1:
            Final_List.append((trim_value(Model_Lines[28].split(':')[1])).replace('%SWV%',SWV))
        if Model_Lines[29].split(':')[1].find('XXX') == -1:
            Final_List.append((trim_value(Model_Lines[29].split(':')[1])).replace('%SWV%',SWV))
        if Model_Lines[30].split(':')[1].find('XXX') == -1:
            Final_List.append((trim_value(Model_Lines[30].split(':')[1])).replace('%SWV%',SWV))
        if Model_Lines[31].split(':')[1].find('XXX') == -1:
            Final_List.append((trim_value(Model_Lines[31].split(':')[1])).replace('%SWV%',SWV))
        return Final_List;

#*****************************************************************************************************
#
#           Main Program
#
#*****************************************************************************************************
if __name__ == "__main__":
    app = wx.App(False)
    frame = Reviewer() #wx.Frame(None, id=-1, title="Make CRC V 1.4", size=(700,850))
    #panel = MainWindow(frame)
    frame.Show()
    app.MainLoop()


