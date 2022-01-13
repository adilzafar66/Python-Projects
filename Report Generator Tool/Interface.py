#------------------------------------------------------------------------------------------#
# Author: Adil Zafar Khan
# Last Edit Date: 12/22/2021
# Description:
"""
    ClassBuilder is a class that contains all the module functions that execute the program.
    It can be viewed as the program manager and every class is finally initialized and
    called from here.
"""
#------------------------------------------------------------------------------------------#

#Import Required Modules
import os
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from pathlib import Path
import queue
from win32com import client
from ClassBuilder import ClassBuilder
from threading import Thread, Event
        
class Interface:

    def __init__(self):

        #Initialize main executer class
        self.classBuilder = ClassBuilder()

        #Create root window
        self.root = Tk()
        
        rootWidth = 600
        rootHeight = 440

        positionRight = int((self.root.winfo_screenwidth() - rootWidth) / 2)
        positionDown = int((self.root.winfo_screenheight() - rootHeight) / 2 - rootHeight/6) 

        self.root.geometry("{}x{}+{}+{}".format(rootWidth, rootHeight, positionRight, positionDown))
        self.root.resizable(height = False, width = False)

        #root window title
        self.root.title("Report Generator - Power Systems")
        self.subheading = StringVar()

        #General Tab Variables
        self.filePath = StringVar()
        self.projectName = StringVar()
        self.projectID = StringVar()
        self.location = StringVar()
        self.revision = StringVar()

        #Details Tab Variables
        self.client = StringVar()
        self.contractor = StringVar()
        self.site = StringVar()
        self.drawing = StringVar()

        #Short Circuit Tab Variables
        self.cycleReportDir = StringVar()
        self.momReportDir = StringVar()
        self.intReportDir = StringVar()
        self.switchReportDir = StringVar()

        #TCC, AF & Device Settings Tab Variables
        self.tccGraphsPath = StringVar()
        self.arcFlashDir = StringVar()
        self.deviceSettingsDir = StringVar()
        self.minCalorie = StringVar()

        #Appendix Tab Variables
        self.sysDataDir = StringVar()
        self.sldDir = StringVar()
        self.utility = StringVar()

        #Variables List
        self.varList = [self.projectName, self.projectID, self.location,
                        self.revision, self.client, self.contractor,
                        self.site, self.drawing, self.minCalorie, self.utility,
                        self.filePath, self.cycleReportDir, self.momReportDir,
                        self.intReportDir, self.switchReportDir, self.tccGraphsPath,
                        self.arcFlashDir, self.deviceSettingsDir, self.sysDataDir, self.sldDir]

        #Root Labels
        self.firstLabel = StringVar()
        self.secondLabel = StringVar()
        self.thirdLabel = StringVar()
        self.fourthLabel = StringVar()
        self.fifthLabel = StringVar()

        #Progress Bar Labels
        self.pbStatus = StringVar()
        self.pbValue = StringVar()

        #Buttons
        self.fileBrowseBtn = Button(text="Browse")
        self.dirBrowseBtn1 = Button(text="Browse")
        self.dirBrowseBtn2 = Button(text="Browse")
        self.dirBrowseBtn3 = Button(text="Browse")
        self.dirBrowseBtn4 = Button(text="Browse")
        

        #Entries
        self.entry1 = Entry(self.root)
        self.entry2 = Entry(self.root)
        self.entry3 = Entry(self.root)
        self.entry4 = Entry(self.root)
        
        #Path Variable
        self.pathLabel = Label(self.root, textvariable = self.filePath)

        #Canvas
        self.canvas = Canvas(self.root, width=700, height=300, bg = 'white', bd = 2)
        self.resetBtn = Button(text="Reset", command=lambda:self.canvas.delete("all"), width = 10)

        #Build the interface
        self.__addMenu()
        self.__createGrid()
        self.__addTitle()
        self.__generalBtnClick()
        self.__addLabels()
        self.__addCmdButtons()
        self.__chooseScreen()

        self.canvas = Canvas(self.root, width=700, height=268, bg = 'white', highlightthickness=1, highlightbackground="black")
        self.canvas.old_coords = None
        
        self.root.protocol("WM_DELETE_WINDOW", self.__exitProgram)
        self.root.mainloop()

    def __exitProgram(self):

        #Destroy main window
        self.root.destroy()
            
        #Delete classBuilder
        del self.classBuilder

    def __addMenu(self):

        #Creating Menubar
        menubar = Menu(self.root)
        
        #Adding File Menu and commands
        file = Menu(menubar, tearoff = 0)
        menubar.add_cascade(label ='File', menu = file)
        file.add_command(label ='Load Last Inputs', command = self.loadInputs)
        file.add_command(label ='Save Inputs', command = self.saveInputs)
        file.add_command(label ='Restart', command = self.restart)
        file.add_separator()
        file.add_command(label ='Exit', command = self.root.destroy)

        #Adding Edit Menu and commands
        edit = Menu(menubar, tearoff = 0)
        menubar.add_cascade(label ='Edit', menu = edit)
        edit.add_command(label ='Clear Inputs', command = self.clearInputs)
        edit.add_command(label ='View Inputs', command = self.viewInputs)
        edit.add_command(label ='None', command = None)
        edit.add_separator()
        edit.add_command(label ='None', command = None)
        edit.add_command(label ='None', command = None)
          
        #Adding Help Menu
        help_ = Menu(menubar, tearoff = 0)
        menubar.add_cascade(label ='Help', menu = help_)
        help_.add_command(label ='User Manual', command = self.openManual)
        help_.add_command(label ='Source', command = self.openSource)
        help_.add_separator()
        help_.add_command(label ='About', command = self.openAbout)
          
        #Display Menu
        self.root.config(menu = menubar)

    #Creates a 5x9 grid for precise placement of widgets
    def __createGrid(self):

        for cols in range(4):
            self.root.columnconfigure(cols, weight=1)

        for rows in range(8):
            self.root.rowconfigure(rows, minsize=40)

        #Configure sizes of the footer and header rows
        self.root.rowconfigure(0, minsize=60)
        self.root.rowconfigure(2, minsize=60)
        self.root.rowconfigure(8, minsize=90)

    
    #Adds a title and a subtitle to the display    
    def __addTitle(self):

        #Add title
        title = Label(self.root, text = "Report Generator")
        title.config(font=("Tahoma", 14))
        title.grid(row = 0, column = 1, columnspan = 2)

        #Add subtitle
        subtitle = Label(self.root, textvariable = self.subheading)
        subtitle.config(font=("Tahoma", 12))
        subtitle.grid(row = 2, column = 0, columnspan = 3, sticky = W, padx = (35,0))

        #Add a seperating top line
        lineTop = Canvas(self.root, height = 4, width = 525)
        lineTop.create_line(0, 2, 525, 2)
        #lineTop.grid(row = 2, column = 0, columnspan = 4, sticky = SW, padx = (35,0), pady = (0, 10))

        #Add a seperating bottom line
        lineBottom = Canvas(self.root, height = 4, width = 530)
        lineBottom.create_line(0, 2, 530, 2)
        lineBottom.grid(row = 8, column = 0, columnspan = 4, sticky = NW, padx = (35,0), pady = (10, 0))

    #Lets a user swich between different tabs
    def __chooseScreen(self):

        #General Button Tab
        generalBtn = Button(text="General", command=self.__generalBtnClick, width = 13)
        generalBtn.grid(row=1, column=0, columnspan = 4, sticky = W, padx = (35, 0))

        #Details Tab Button
        detailsBtn = Button(text="Details", command=self.__detailsBtnClick, width = 13)
        detailsBtn.grid(row=1, column=0, columnspan = 4, padx = (123, 0), sticky = W)

        #Short Circuit Button Tab
        shortCircuitBtn = Button(text="Short Circuit", command=self.__shortCircuitBtnClick, width = 13)
        shortCircuitBtn.grid(row=1, column=0, columnspan = 4, padx = (211, 0), sticky = W)

        #TCC Graph Button Tab
        tccGraphBtn = Button(text="TCC, AF & DS", command=self.__tccGraphBtnClick, width = 13)
        tccGraphBtn.grid(row=1, column=0, columnspan = 4, padx = (299, 0), sticky = W)

        #Appendix Button Tab
        appendixBtn = Button(text="Appendix", command=self.__appendixBtnClick, width = 13)
        appendixBtn.grid(row=1, column=0, columnspan = 4, padx = (387, 0), sticky = W)

        #Extra Button Tab
        appendixBtn = Button(text="Click Me!", command=self.__ExtraBtnClick, width = 13)
        appendixBtn.grid(row=1, column=0, columnspan = 4, padx = (475, 0), sticky = W)

        #Add a seperating top line
        lineTop = Canvas(self.root, height = 2, width = 526)
        lineTop.create_line(0, 2, 526, 2)
        lineTop.grid(row = 1, column = 0, columnspan = 4, sticky = SW, padx = (35,0))

    #Add labels for the widgets
    def __addLabels(self):

        #Add 6 labels which will change according to the tab selected
        self.l1 = Label(self.root, textvariable = self.firstLabel)
        self.l1.grid(row = 3, column = 0, columnspan = 2, sticky = W, padx = (35, 0))

        self.l2 = Label(self.root, textvariable = self.secondLabel)
        self.l2.grid(row = 4, column = 0, columnspan = 2, sticky = W, padx = (35, 0))
        
        self.l3 = Label(self.root, textvariable = self.thirdLabel)
        self.l3.grid(row = 5, column = 0, columnspan = 2, sticky = W, padx = (35, 0))

        self.l4 = Label(self.root, textvariable = self.fourthLabel)
        self.l4.grid(row = 6, column = 0, columnspan = 2, sticky = W, padx = (35, 0))

        self.l5 = Label(self.root, textvariable = self.fifthLabel)
        self.l5.grid(row = 7, column = 0, columnspan = 2, sticky = W, padx = (35, 0))


    #Add command buttons which includes a create button for execution
    def __addCmdButtons(self):

        #Initialize buttons
        self.importButton = Button(text="Import Data", command=self.executeFirstPhase)
        self.compileButton = Button(text="Compile PDF", command=self.executeSecondPhase)  
        self.closeButton = Button(text="Close", command=self.root.destroy)
        
        #Display buttons
        self.importButton.grid(row=8, column=0, sticky = E, ipadx = 10)
        self.compileButton.grid(row=8, column=1, columnspan = 2, ipadx = 10)
        self.closeButton.grid(row=8, column=3, ipadx = 10, sticky = W)

    def __generalBtnClick(self):

        #Display General Tab
        self.subheading.set("General")
        self.firstLabel.set("Report File:")
        self.secondLabel.set("Project Name:")
        self.thirdLabel.set("Project #:")
        self.fourthLabel.set("Location:")
        self.fifthLabel.set("Revision:")
        self.__packWidgets1()

    def __detailsBtnClick(self):

        #Display Details Tab
        self.subheading.set("Details")
        self.firstLabel.set("Client:")
        self.secondLabel.set("Contractor:")
        self.thirdLabel.set("Site:")
        self.fourthLabel.set("Drawing ref#:")
        self.fifthLabel.set("")
        self.__packWidgets2()

    def __shortCircuitBtnClick(self):

        #Display Short Circuit Tab
        self.subheading.set("Short Circuit")
        self.firstLabel.set("0.5 Cycle Report:")
        self.secondLabel.set("MOM Report:")
        self.thirdLabel.set("INT Report:")
        self.fourthLabel.set("Switch Report:")
        self.fifthLabel.set("")
        self.__packWidgets3()

    def __tccGraphBtnClick(self):

        #Display TCC, AF & DS Tab
        self.subheading.set("TCC, AF & DS")
        self.firstLabel.set("TCC Graphs:")
        self.secondLabel.set("Arc Flash:")
        self.thirdLabel.set("Device Settings:")
        self.fourthLabel.set("Min Calorie Level:")
        self.fifthLabel.set("")
        self.__packWidgets4()

    def __appendixBtnClick(self):

        #Display Appendix & BCIT Info Tab
        self.subheading.set("Appendix & Utility Info")
        self.firstLabel.set("System Data: ")
        self.secondLabel.set("SLD Diagrams:")
        self.thirdLabel.set("Utility Name:")
        self.fourthLabel.set("")
        self.fifthLabel.set("")
        self.__packWidgets5()
        
    def __ExtraBtnClick(self):

        #Display Extras Tab
        self.subheading.set("Draw Something!")
        self.firstLabel.set("")
        self.secondLabel.set("")
        self.thirdLabel.set("")
        self.fourthLabel.set("")
        self.fifthLabel.set("")
        self.__packWidgets6()

    def __dirBrowseClick(self, strVar):

        #Open directory browse window
        dirName = filedialog.askdirectory()
        strVar.set(dirName)

    def __fileBrowseClick(self, strVar):

        #Open file browse window
        fileName = filedialog.askopenfilename()
        strVar.set(fileName)

    def __unpackWidgets(self):

        #Remove canvas
        self.canvas.grid_forget()
        self.resetBtn.grid_forget()

        #Remove all the browse buttons
        self.fileBrowseBtn.grid_forget()
        self.dirBrowseBtn1.grid_forget()
        self.dirBrowseBtn2.grid_forget()
        self.dirBrowseBtn3.grid_forget()
        self.dirBrowseBtn4.grid_forget()

        #Remove all the entry forms
        self.entry1.grid_forget()
        self.entry2.grid_forget()
        self.entry3.grid_forget()
        self.entry4.grid_forget()

    def __packWidgets1(self):

        #Remove existing widgets
        self.__unpackWidgets()

        #Create widgets for Tab 1
        self.fileBrowseBtn.config(command=lambda:self.__fileBrowseClick(self.filePath))
        self.fileBrowseBtn.grid(row=3, column=1, columnspan = 2, sticky = W, ipadx = 10)

        #self.pathLabel.grid(row=3, column=2, columnspan = 2, sticky = W, padx=(20, 0))

        self.entry1.config(textvariable = self.projectName)
        self.entry1.grid(row=4, column=1, columnspan = 2, sticky = W, ipadx = 10)
        self.entry2.config(textvariable = self.projectID)
        self.entry2.grid(row=5, column=1, columnspan = 2, sticky = W, ipadx = 10)
        self.entry3.config(textvariable = self.location)
        self.entry3.grid(row=6, column=1, columnspan = 2, sticky = W, ipadx = 10)
        self.entry4.config(textvariable = self.revision)
        self.entry4.grid(row=7, column=1, columnspan = 2, sticky = W, ipadx = 10)

    def __packWidgets2(self):

        #Remove existing widgets
        self.__unpackWidgets()

        #Default value for site entry
        if self.site.get() == '':
            self.site.set(self.location.get())

        #Create widgets for Tab 1
        self.entry1.config(textvariable = self.client)
        self.entry1.grid(row=3, column=1, columnspan = 2, sticky = W, ipadx = 10)
        self.entry2.config(textvariable = self.contractor)
        self.entry2.grid(row=4, column=1, columnspan = 2, sticky = W, ipadx = 10)
        self.entry3.config(textvariable = self.site)
        self.entry3.grid(row=5, column=1, columnspan = 2, sticky = W, ipadx = 10)
        self.entry4.config(textvariable = self.drawing)
        self.entry4.grid(row=6, column=1, columnspan = 2, sticky = W, ipadx = 10)

    def __packWidgets3(self):

        #Remove existing widgets
        self.__unpackWidgets()

        #Create widgets for Tab 2
        self.dirBrowseBtn1.config(command=lambda:self.__dirBrowseClick(self.cycleReportDir))
        self.dirBrowseBtn1.grid(row=3, column=1, columnspan = 2, sticky = W, ipadx = 10)
        
        self.dirBrowseBtn2.config(command=lambda:self.__dirBrowseClick(self.momReportDir))
        self.dirBrowseBtn2.grid(row=4, column=1, columnspan = 2, sticky = W, ipadx = 10)

        self.dirBrowseBtn3.config(command=lambda:self.__dirBrowseClick(self.intReportDir))
        self.dirBrowseBtn3.grid(row=5, column=1, columnspan = 2, sticky = W, ipadx = 10)

        self.dirBrowseBtn4.config(command=lambda:self.__dirBrowseClick(self.switchReportDir))
        self.dirBrowseBtn4.grid(row=6, column=1, columnspan = 2, sticky = W, ipadx = 10)
        
        
    def __packWidgets4(self):

        #Remove existing widgets
        self.__unpackWidgets()

        #Create widgets for Tab 3
        self.dirBrowseBtn1.config(command=lambda:self.__fileBrowseClick(self.tccGraphsPath))
        self.dirBrowseBtn1.grid(row=3, column=1, columnspan = 2, sticky = W, ipadx = 10)
        
        self.dirBrowseBtn2.config(command=lambda:self.__dirBrowseClick(self.arcFlashDir))
        self.dirBrowseBtn2.grid(row=4, column=1, columnspan = 2, sticky = W, ipadx = 10)

        self.fileBrowseBtn.config(command=lambda:self.__fileBrowseClick(self.deviceSettingsDir))
        self.fileBrowseBtn.grid(row=5, column=1, columnspan = 2, sticky = W, ipadx = 10)

        self.entry1.config(textvariable = self.minCalorie)
        self.entry1.grid(row=6, column=1, columnspan = 2, sticky = W, ipadx = 10)

    def __packWidgets5(self):

        #Remove existing widgets
        self.__unpackWidgets()

        #Create widgets for Tab 4
        self.dirBrowseBtn1.config(command=lambda:self.__dirBrowseClick(self.sysDataDir))
        self.dirBrowseBtn1.grid(row=3, column=1, columnspan = 2, sticky = W, ipadx = 10)
        
        self.dirBrowseBtn2.config(command=lambda:self.__dirBrowseClick(self.sldDir))
        self.dirBrowseBtn2.grid(row=4, column=1, columnspan = 2, sticky = W, ipadx = 10)

        self.entry1.config(textvariable = self.utility)
        self.entry1.grid(row=5, column=1, columnspan = 2, sticky = W, ipadx = 10)
        
    def __packWidgets6(self):
        
        #Remove existing widgets
        self.__unpackWidgets()

        self.resetBtn.grid(row = 2, column = 1)
        self.canvas.grid(row = 3, column = 0, columnspan = 4, rowspan = 10)
        self.root.bind('<ButtonPress-1>', self.draw_line)
        self.root.bind('<ButtonRelease-1>', self.draw_line)
        self.root.bind('<B1-Motion>', self.draw)
        self.root.bind('<ButtonRelease-1>', self.reset_coords)

    def __errorHandler(self):
        
        self.__unpackWidgets()

    def openNewWindow(self, title, width, height, rows, columns):
     
        #Toplevel object which will be treated as a new window
        window = Toplevel(self.root)

        for cols in range(columns):
            window.columnconfigure(cols, weight=1)

        for rows in range(rows):
            window.rowconfigure(rows, weight=1)
        
        #Sets the title of the Toplevel widget
        window.title(title)
     
        #Sets the geometry of Toplevel
        winWidth = width
        winHeight = height

        positionRight = int((self.root.winfo_screenwidth() - winWidth) / 2)
        positionDown = int((self.root.winfo_screenheight() - winHeight) / 2 - winHeight/6) 

        window.geometry("{}x{}+{}+{}".format(winWidth, winHeight, positionRight, positionDown))
        window.focus_set()

        return window

    def startProgress(self):
        
        self.pbWindow = self.openNewWindow("Progress", 300, 120, 5, 3)
        self.pbWindow.protocol("WM_DELETE_WINDOW", self.haltThread)
        
        self.pb = Progressbar(self.pbWindow, orient=HORIZONTAL, length=280, mode='determinate')
        self.pb.grid(row=2, column=0, columnspan = 3)
        self.progressVal = self.pb['value']
        
        self.pbStatus.set("Initializing...")
        statusLabel = Label(self.pbWindow, textvariable = self.pbStatus)
        statusLabel.grid(row=1, column=0, columnspan = 3, sticky = W, padx = (10, 0))

        self.pbValue.set(f"{self.pb['value']:.0f}%")
        pbLabel = Label(self.pbWindow, textvariable = self.pbValue)
        pbLabel.grid(row=2, column=0, columnspan = 3)

        pbButton = Button(self.pbWindow, text="Abort", command=self.haltThread)
        pbButton.grid(row=3, column=0, columnspan = 3, ipadx = 10)

    def stopProgress(self):
        
        self.pb.grid_forget()
        self.pb.destroy()
        self.pbWindow.destroy()

    def haltThread(self):

        #Destroy Progress Window
        self.pbWindow.destroy()

    def openManual(self):

        os.startfile(Path(os.getcwd(), "User Guide.docx"))

    def openSource(self):

        os.startfile(os.getcwd())

    def openAbout(self):

        aboutWindow = self.openNewWindow("About", 400, 300, 6, 3)
        
        #Add subtitle
        title = Label(aboutWindow, text = "The Report Generator")
        title.config(font=("Tahoma", 12))
        title.grid(row = 0, column = 0, columnspan = 3, sticky = W, padx = (35,0))

        bodyText = Label(aboutWindow, wraplength=320, text = '''The Report Generator is a program that generates and compiles report for the Power Systems and Studies department at Prime Engineering Inc. based in Victoria, BC. The purpose of the program is to minimize manual pricessing and substituting it with automated''')
        bodyText.config(font=("Tahoma", 9))
        bodyText.grid(row = 1, column = 0, columnspan = 3, rowspan = 6, sticky = NW, padx = (35,0))
        
    def restart(self):

        self.root.destroy()
        self.__init__()

    def saveInputs(self):

        path = Path("Data")
        path.mkdir(parents = True, exist_ok = True)

        path = Path(path, "LastInputs.txt")
        
        file = open(path, "w")
        inputs = [item.get() for item in self.varList]
        file.writelines(["%s\n" % item  for item in inputs])
        file.close()

    def loadInputs(self):

        path = Path("Data", "LastInputs.txt")
        
        try:
            file = open(path, "r")
        except:
            return
        
        inputs = [item.strip() for item in file.readlines()]
        
        i = 0
        for item in self.varList:

            item.set(inputs[i])
            i += 1
            if i == len(inputs):
                break

    def clearInputs(self):

        for item in self.varList:
            item.set('')

    def viewInputs(self):

        nameList = ["Project Name:", "Project ID:", "Location:", "Revision:",
                    "Client:", "Contractor:", "Site:", "Drawing ref#:",
                    "Min Calorie:", "Utility Name:", "Report File:", "0.5 Cycle Report:",
                    "MOM Report:", "INT Report:", "Switch Report:", "TCC Graphs:",
                    "Arc Flash:", "Device Settings:", "Ref Info:", "SLDs:"]

        rows = len(self.varList)
        inputsViewWindow = self.openNewWindow("Input Viewer", 500, 450, rows, 3)
        scrollbar = Scrollbar(inputsViewWindow)
        scrollbar.grid(column = 3)
        varLabel = Text(inputsViewWindow, wrap = WORD)
        varLabel.grid(row=0, column=0, rowspan = rows, columnspan = 3)
            
        for i in range(len(self.varList)):
            varLabel.insert(END, nameList[i] + "\t\t\t" + self.varList[i].get())
            varLabel.insert(END, "\n")
            varLabel.insert(END, "-" * 20)
            varLabel.insert(END, "\n")

        varLabel.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=varLabel.yview)

    def draw(self, event):
        x, y = event.x, event.y
        if self.canvas.old_coords:
            x1, y1 = self.canvas.old_coords
            self.canvas.create_line(x, y, x1, y1)
        self.canvas.old_coords = x, y

    def draw_line(self, event):

        if str(event.type) == 'ButtonPress':
            self.canvas.old_coords = event.x, event.y

        elif str(event.type) == 'ButtonRelease':
            x, y = event.x, event.y
            x1, y1 = self.canvas.old_coords
            self.canvas.create_line(x, y, x1, y1)
            
    def reset_coords(self, event):
        self.canvas.old_coords = None
        
        
#-------------------------------------Main Executer-----------------------------------------#
        
    def executeFirstPhase(self):
        
        self.firstExecutionThread = Thread(target=self.classBuilder.firstPhase, args=(self,))
        self.firstExecutionThread.setDaemon(True)
        self.firstExecutionThread.start()

    def executeSecondPhase(self):

        self.secondExecutionThread = Thread(target=self.classBuilder.secondPhase, args=(self,))
        self.secondExecutionThread.setDaemon(True)
        self.secondExecutionThread.start()

    def pickError(self):
        
        del self.classBuilder
        self.classBuilder = ClassBuilder()
