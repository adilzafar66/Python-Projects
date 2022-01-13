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

#Import all the required libraries

import os
import sys
import time
import tkinter
import traceback
from pathlib import Path
from win32com import client
from tkinter import messagebox
from PdfHandler import PdfHandler
from DocHandler import DataImporter
from ExcelGenerator import IntReport, MomReport, SwitchReport, ArcFlashReport
from PdfGenerator import TCCGraphs, HalfCycleReports, RefInfo, FinalReport

class ClassBuilder:

    def __init__(self):

        #Initialize first phase return value that indicates if first phase has been executed
        self.firstPhaseReturn = False

    def initializeFirstPhase(self, interface):

         #________ASSIGNING PATHS TO LOCAL VARIABLES_______#

        #Report template path
        self.filePath = Path(interface.filePath.get())

        #MOM reports directory path
        self.momReportDir = Path(interface.momReportDir.get())

        #INT reports directory path
        self.intReportDir = Path(interface.intReportDir.get())

        #INT reports directory path
        self.switchReportDir = Path(interface.switchReportDir.get())

        #TCC graphs PDF path
        self.tccGraphsPath = Path(interface.tccGraphsPath.get())

        #Arc Flash reports directory path
        self.arcFlashDir = Path(interface.arcFlashDir.get())

        #Device settings reports template path
        self.deviceSettingsDir = Path(interface.deviceSettingsDir.get())

        #SLD diagrams directory path
        self.sldDir = Path(interface.sldDir.get())

        #Half Cycle reports directory path
        self.cycleReportDir = Path(interface.cycleReportDir.get())

        #Reference Info directory path
        self.sysDataDir = Path(interface.sysDataDir.get())

        #Convert minCalorie to float
        self.minCalorie = interface.minCalorie.get()

         #________CREATING OUTPUT DIRECTORY PATHS________#

        #Ouput root directory
        self.outputDir = Path(self.filePath.parent, "Output Files")

        #Excel outputs directory
        self.excelFolder = Path(self.outputDir, "Excel Outputs")
        self.excelFolder.mkdir(parents = True, exist_ok = True)

        #Edited report path
        self.reportDocPath = Path(self.outputDir, self.filePath.name)

        #PDF outputs directory
        self.pdfFolder = Path(self.outputDir, "PDF Outputs")
        self.pdfFolder.mkdir(parents = True, exist_ok = True)

        #________INITIALIZE CLASS OBJECTS________#

        #Initialize INT Reports object
        self.intReport = IntReport()

        #Initialize MOM Reports object
        self.momReport = MomReport()

        #Initialize MOM Reports object
        self.switchReport = SwitchReport()

        #Initialize AF Reports object
        self.arcFlashReport = ArcFlashReport()

        #Initialize DataImporter object
        self.dataImporter = DataImporter(self.filePath)

        #Initialize Half Cycle Reports object
        self.halfCycleReports = HalfCycleReports(self.cycleReportDir, self.pdfFolder)

        #Initialize Reference Info object
        self.refInfo = RefInfo(self.sysDataDir, self.pdfFolder)

        #Initialize TCC Graphs object
        self.tccGraphs = TCCGraphs(self.tccGraphsPath, Path(self.pdfFolder, self.tccGraphsPath.name))


    def initializeSecondPhase(self, interface):

         #________ASSIGNING PATHS TO LOCAL VARIABLES_______#

        #Report template path
        self.filePath = Path(interface.filePath.get())

        #TCC graphs PDF path
        self.tccGraphsPath = Path(interface.tccGraphsPath.get())

        #Reference Info directory path
        self.sysDataDir = Path(interface.sysDataDir.get())
        
        #Half Cycle reports directory path
        self.cycleReportDir = Path(interface.cycleReportDir.get())

         #________CREATING OUTPUT DIRECTORY PATHS________#

        #Ouput root directory
        self.outputDir = Path(self.filePath.parent, "Output Files")

        #PDF outputs directory
        self.pdfFolder = Path(self.outputDir, "PDF Outputs")
        self.pdfFolder.mkdir(parents = True, exist_ok = True)

        #Edited report path
        self.reportDocPath = Path(self.outputDir, self.filePath.name)

        #Pdf output path
        self.wordOutputPath = Path(self.outputDir, self.filePath.stem + '.pdf')

        #________INITIALIZE CLASS OBJECTS________#

        #If the first phase has not been executed, take the input report file
        if self.firstPhaseReturn and self.reportDocPath.is_file():

            #Initialize DataImporter object
            self.dataImporter = DataImporter(self.reportDocPath)
            
        else:

            #Use the report created by first phase in the output directory
            self.dataImporter = DataImporter(self.filePath)
        
        #Initialize Final Report object
        self.finalReport = FinalReport(self.wordOutputPath)
        

    def errorHandler(self, interface):
        
        #Check if any of the inputs have not been provided
        #Run a check on each input field
        #Return True if any of the inputs are not provided
        
        if interface.filePath.get() == '':

            #Show error
            messagebox.showerror("Input Error",
                                 "Error: no input found for Report directory")
            return True

        if interface.cycleReportDir.get() == '':

            #Show error
            messagebox.showerror("Input Error",
                                 "Error: no input found for Half Cycle Reports directory")
            return True

        if interface.momReportDir.get() == '':

            #Show Yes/No message box
            response = messagebox.askyesno("Input Warning",
            "Warning: no input found for MOM Reports directory. Do you want to proceed?")

            #If user chooses no, then return True
            if not response:
                return True

        if interface.intReportDir.get() == '':

            #Show Yes/No message box
            response = messagebox.askyesno("Input Warning",
            "Warning: no input found for INT Reports directory. Do you want to proceed?")

            #If user chooses no, then return True
            if not response:
                return True

        if interface.tccGraphsPath.get() == '':

            #Show error
            messagebox.showerror("Input Error",
                                 "Error: no input found for TCC Graphs directory")
            return True

        if interface.arcFlashDir.get() == '':

            #Show error
            messagebox.showerror("Input Error",
                                 "Error: no input found for Arc Flash Reports directory")
            return True

        if interface.deviceSettingsDir.get() == '':

            #Show Yes/No message box
            response = messagebox.askyesno("Input Warning",
            "Warning: no input found for Device Settings file. Do you want to proceed?")

            #If user chooses no, then return True
            if not response:
                return True

        if interface.sysDataDir.get() == '':

            #Show error
            messagebox.showerror("Input Error",
                                 "Error: no input found for System Data path")
            return True

        if interface.sldDir.get() == '':

            #Show error
            messagebox.showerror("Input Error",
                                 "Error: no input found for BCH Info path")
            return True

        if interface.utility.get() == '':

            #Show error
            messagebox.showerror("Input Error",
                                 "Error: no input found for Utility Name")
            return True

        if interface.minCalorie.get() == '':

            #Show error
            messagebox.showerror("Input Error",
                                 "Error: no input found for Min Calorie field")
            return True

        #Check if a numeric value for minCalorie is provided
        #If not then give an error
        else:
            try:
                float(interface.minCalorie.get())
            except:

                #Show error
                messagebox.showerror("Input Error",
                                     "Error: Min Calorie can only take numeric values")
                return True

        return False


    def secondPhaseErrorHandler(self, interface):

        #Check if any of the inputs have not been provided
        #Run a check on each input field
        #Return True if any of the inputs are not provided

        if interface.filePath.get() == '':

            #Show error
            messagebox.showerror("Input Error",
                                 "Error: no input found for Report directory")
            return True

        return False
        

    #Update the progress bar in the interface
    def updateProgress(self, interface, amount, status = "Processing"):

        #Set progress bar status
        interface.pbStatus.set(status + '...')

        #Increment progress bar value
        interface.pb['value'] += amount

        #Set progress bar value
        interface.pbValue.set(f"{interface.pb['value']:.0f}%")

    #Create a dictionary containing tags for INT and MOM reports
    def createTagDict(self, dirPath):

        #Initialize an empty dictionary
        tagDict = {}

        #Loop through each file name in dirPath
        for fileName in os.listdir(dirPath):

            #Get the stem of each file name
            tagDict[fileName] = fileName[ : fileName.find('.')]

        #Return the dictionary object
        return tagDict

    #Check if an excel table is imported to the report doc
    def checkImport(self, interface, imported, wbPath):

        if imported == -1:

            #Show warning
            messagebox.showwarning("Warning", "Could not import " + Path(wbPath).stem,
                                   parent = interface.pbWindow)


    #Executes the first phase of the program
    def firstPhase(self, interface):

        ##    Tags to be used in the document
        ##    [PROJECT_NAME] = Title or name of project
        ##    [PROJECT_NUMBER] = Prime project number
        ##    [PROJECT_REV] = Report revision
        ##    [PROJECT_CLIENT] = Prime client
        ##    [PROJECT_CONTRACTOR] = Contractor name
        ##    [PROJECT_SITE] = Site location
        ##    [PROJECT_LOC] = Location of project 
        ##    [DRAWING_REF] = Drawing reference
        ##    [PROJECT_UTILITY] = Utility Name


        #Create a list of tags and their corresponding insertion values
        insertionList = [["[PROJECT_NAME]", "[PROJECT_NUMBER]",
                          "[PROJECT_LOC]", "[PROJECT_REV]",
                          "[PROJECT_CLIENT]", "[PROJECT_CONTRACTOR]",
                          "[PROJECT_SITE]", "[DRAWING_REF]", "[PROJECT_UTILITY]"],
                         [interface.projectName.get(), interface.projectID.get(),
                          interface.location.get(), interface.revision.get(),
                          interface.client.get(), interface.contractor.get(),
                          interface.site.get(), interface.drawing.get(),
                          interface.utility.get()]]

        #If any errors caught, exit before doing anything
        if self.errorHandler(interface):
            return
        
        #----------------------INITIALIZING----------------------#

        #Start the progress bar
        interface.startProgress()

        #Total progress bar value
        pbRem = 100

        #-----------------------------GENERATING DATA----------------------------------#

        try:

            #Initialize all variables and objects
            self.initializeFirstPhase(interface)

            #Save inputs of the current execution
            interface.saveInputs()

            #Convert minCalorie to float
            self.minCalorie = float(self.minCalorie)
            
            #______________INT REPORTS_____________#
            
            #Update progress bar
            self.updateProgress(interface, 0, "Generating INT Reports")

            #Generate INT reports
            self.intReport.buildTables(self.intReportDir, self.excelFolder, "INT")

            #______________MOM REPORTS_____________#

            #Update progress bar
            self.updateProgress(interface, int(pbRem * 0.05), "Generating MOM Reports")

            #Generate MOM reports
            self.momReport.buildTables(self.momReportDir, self.excelFolder, "MOM")

            #______________SWT REPORTS_____________#

            #Update progress bar
            self.updateProgress(interface, int(pbRem * 0.03), "Generating SWT Report")

            #Generate SWT reports
            self.switchReport.buildTables(self.switchReportDir, self.excelFolder, "SWT")

            #______________AF REPORTS______________#
            
            #ARC FLASH REPORTS
            self.updateProgress(interface, int(pbRem * 0.05), "Generating Arc Flash Reports")

            #Loop through all the arc flah sheets in the arcFlashDir
            for fileName in os.listdir(self.arcFlashDir):

                #If the current file name ends with .xls or .xlsx
                if fileName.endswith('.xls') or fileName.endswith('.xlsx'):
                    currentPath = Path(self.arcFlashDir, fileName)

                    #Reinitialize arcFlashReport object
                    self.arcFlashReport.__init__()

                    #Generate arc flash table from current file
                    self.arcFlashReport.generateTable(currentPath, self.minCalorie, self.excelFolder)

                    #Update progress bar
                    self.updateProgress(interface, int(pbRem * 0.02), "Generating Arc Flash Reports")


            #___________HALF CYCLE REPORTS___________#

            #Update progress bar
            self.updateProgress(interface, int(pbRem * 0.02), "Generating Half Cycle Reports")

            #Generate Half Cycle Reports
            self.halfCycleReports.generateHalfCycle()
            
            
            #---------------------------------IMPORTING DATA---------------------------------#

            #Get tags for MOM and INT reports
            tagDict = self.createTagDict(self.excelFolder)

            #___________WORDS INSERTION___________#

            #Update progress bar
            self.updateProgress(interface, 0, "Inserting words into the Document")

            #Find and replace words in insertionList[0] with insertionList[1] in the document
            self.dataImporter.insertWords(insertionList[0], insertionList[1])
            
            #___________DEVICE SETTINGS___________#

            #If device settings path is provided
            if self.deviceSettingsDir != '':

                #Set the output path for device settings report as the 'Excel Outputs' directory
                filledDSPath = Path(self.excelFolder, self.deviceSettingsDir.name)

                #Update progress bar
                self.updateProgress(interface, 0, "Running Excel Macros")

                #Run the macro in the Device Settings template file to generate Dvice Settings tables
                self.dataImporter.runXlMacro(self.deviceSettingsDir, "!InputFunctions.importWorkbook",
                                        saveCopy = True, savePath = filledDSPath)

            #If device settings path is not provided
            else:

                #Set the output path as an empty string
                filledDSPath = Path('')

                #Try and look for any preexisting device settings files in the 'Excel Outputs' directory
                for fileName in os.listdir(self.excelFolder):
                    if fileName.endswith('.xlsm'):

                        #If found then set the output path as that file's path
                        filledDSPath = Path(self.excelFolder, fileName)

                        #Exit the loop
                        break

            #If the output path is set and the file exists
            if filledDSPath.is_file():

                #Get the indices of the filled sheets in the device settings file
                filledSheets = self.dataImporter.sheetsFilledList(filledDSPath)

                #Update the progress bar
                self.updateProgress(interface, int(pbRem * 0.01), "Importing Device Settings Reports")

                #By default, let the file remain open after table export
                closeExcel = False

                #Loop through the sheets that are filled
                for index in filledSheets:

                    #If it is the last filled sheet
                    if index == filledSheets[-1]:

                        #The file will be closed after table export
                        closeExcel = True

                    #Import the current sheet data as a table and set imported as True if import is succesful
                    imported = self.dataImporter.importExcelTable(filledDSPath, index, 2, tagDict, closeExcel)

                    #Check if last import was succesful
                    self.checkImport(interface, imported, filledDSPath)

                    #Update the progress bar
                    self.updateProgress(interface, int(pbRem * 0.02), "Importing Device Settings Reports")

            #___________MOM, INT & AF REPORTS___________#

            #Update the progress bar       
            self.updateProgress(interface, 0, "Importing MOM, INT, AF & SWT Reports")

            #Loop through all the .xlsx files in the 'Excel Outputs' directory
            for fileName in os.listdir(self.excelFolder):
                if fileName.endswith('.xlsx'):

                    #If it is a MOM or INT Report
                    if fileName.startswith(("MOM", "INT", "SWT")):

                        #Set the top two rows as header rows
                        headerRows = 2
                    else:

                        #Set the top row as the header row
                        headerRows = 1

                    #Set the path of the current file
                    path = Path(self.excelFolder, fileName)

                    #Import the current sheet data as a table and set imported as True if import is succesful
                    imported = self.dataImporter.importExcelTable(path, 1, headerRows, tagDict)

                    #Check if last import was succesful
                    self.checkImport(interface, imported, path)

                    #Update the progress bar
                    self.updateProgress(interface, int(pbRem * 0.01), "Importing MOM, INT, AF & SWT Reports")

            #___________SLD DIAGRAMS___________#

            #Update the progress bar
            self.updateProgress(interface, 0, "Importing SLD Diagrams")

            #Loop through all the .emf files in the SLD directory
            for fileName in os.listdir(self.sldDir):
                if fileName.endswith(".emf"):

                    #Set the path of the current file
                    path = Path(self.sldDir, fileName)

                    #Import the current .emf file as an image
                    self.dataImporter.importImage(path, path.stem)

                    #Update the progress bar
                    self.updateProgress(interface, int(pbRem * 0.02) , "Importing SLD Diagrams")

            #______________TCC TITLES______________#

            #Update progress bar
            self.updateProgress(interface, int(pbRem * 0), "Importing TCC Graph titles")

            #Split TCC Pdf file
            PdfHandler().splitPdf(self.tccGraphsPath)

            #Generate TCC title list
            self.dataImporter.createTCCTitleList(self.tccGraphsPath)

            #_____________TCC TABLE____________#

            #Update the progress bar      
            self.updateProgress(interface, int(pbRem * 0.1), "Creating TCC Table in the Document")

            #Create the TCC Table using the titleTable attribute of self.dataImporter
            self.dataImporter.createTCCTable(self.dataImporter.titleTable, "TCC Table")

            #______________TCC GRAPHS______________#

            #Update progress bar
            self.updateProgress(interface, int(pbRem * 0.02), "Generating TCC Graphs")

            #Generate TCC graphs from TCC title list
            self.tccGraphs.generateTCC(self.dataImporter.titleTable)

            #_____________REFERENCE INFO_____________#

            #Update progress bar
            self.updateProgress(interface, int(pbRem * 0.02), "Generating Reference Info")

            #Generate Reference Info
            self.refInfo.generateRefInfo()

            #Update the progress bar  
            self.updateProgress(interface, 0, "Waiting for the User")

            #Save a copy of the report doc in the outputDir
            self.dataImporter.saveAs(self.reportDocPath)

            #--------------------------------FINISHING---------------------------------#

            #Delete Dataimporter object to run destructor
            #When we reassign it in the second phase, the destructor should not run
            #The destructor runs after the object is reassigned
            #Therefore the destructor runs on the new assigned object
            del self.dataImporter

            #Set the return value as True
            self.firstPhaseReturn = True

            #Show success message
            messagebox.showinfo("Phase 1 Completed", "Please review the document before proceeding to the compilation phase.")

            #Minimize the progress bar window
            interface.pbWindow.wm_state('iconic')

            #Minimize the program window
            interface.root.wm_state('iconic')
            
            #Run the report document
            os.startfile(self.reportDocPath)
        
        #If pbWindow is closed, exit the function and hence the thread
        except tkinter.TclError:

            #Delete all the variable/objects
            #Safely exit during an exception

            interface.pickError()
            return

        #If it is any other exception
        except Exception:

            #Print the exception
            traceback.print_exc()
            
            #Delete all the variable/objects
            #Safely exit during an exception

            interface.pickError()
            return

    #Executes the second phase of the program
    def secondPhase(self, interface):

        #If any errors caught, exit before doing anything
        if self.secondPhaseErrorHandler(interface):
            return

        #If the first phase has not run yet
        if not self.firstPhaseReturn:

            #Ask the user if they wish to proceed
            response = messagebox.askyesno("Runtime Warning",
            "Warning: the first phase has not been executed yet. Are you sure you want to proceed?")

            #If the answer is no, exit the function
            if not response:
                return
        
        #----------------------INITIALIZING-------------------------#

        #Check if the progress bar window is still open
        try:

            #Maximize the progress bar window
            interface.pbWindow.wm_state('normal')

            #Update the progress bar
            self.updateProgress(interface, 0, "Saving the completed the report")

        #If the progress bar is not initialized
        except:

            #Start progress bar
            interface.startProgress()

        #----------------------GENERATING DATA------------------------#
            
        try:

            #Initialize all variables and objects
            self.initializeSecondPhase(interface)

            #Save inputs of the current execution
            interface.saveInputs()

            #Calculate the progress bar status remaining
            pbRem = 100 - interface.pb['value']
            self.updateProgress(interface, 0, "Saving the completed the report")

            #Convert the report doc containing tags to pdf
            self.updateProgress(interface, int(pbRem * 0.01), "Converting Document to PDF")
            self.dataImporter.convertToPdf(Path(self.wordOutputPath.parent, self.wordOutputPath.stem + "_Tags.pdf"))

##            #Create file names for the input half cycle report files
##            halfCyclePaths = os.listdir(self.cycleReportDir)
##            halfCycleFileNames = [Path(item).stem for item in halfCyclePaths]
##
##            #Create file names for the input reference info files
##            sysDataPaths = os.listdir(self.sysDataDir)
##            sysDataFileNames = [Path(item).stem for item in sysDataPaths]
##            tagList = [self.tccGraphsPath.stem] + halfCycleFileNames + sysDataFileNames
            
            #Create a tagList for the pdf files
            tagsPaths = os.listdir(self.pdfFolder)
            tagList = [Path(item).stem for item in tagsPaths]

            #Remove the tags from the report doc
            self.dataImporter.findDeleteAll(tagList)

            #Convert report doc to pdf again
            #This pdf doc will contain no tags
            self.dataImporter.convertToPdf(self.wordOutputPath)

            #Delete dataImporter object
            del self.dataImporter

            #-------------------COMPILING FINAL REPORT---------------------#
            
            #FINAL REPORT
            self.updateProgress(interface, int(pbRem * 0.2), "Looking for tags in the pdf")
            self.finalReport.compileSearchReport(self.pdfFolder, interface.pbWindow)
            self.updateProgress(interface, int(pbRem * 0.5), "Compiling final report")
            self.finalReport.compileFinalReport()
            
            #FINISIH
            self.updateProgress(interface, int(pbRem * 0.3), "Finishing")
            interface.stopProgress()

            messagebox.showinfo("Success", "The report has been compiled successfully!")

        #If pbWindow is closed, exit the function and hence the thread
        except tkinter.TclError:
            
            #Delete all the variable/objects
            #Safely exit during an exception

            interface.pickError()
            return

        #If it is any other exception
        except Exception:

            #Print the exception
            traceback.print_exc()

            #Delete all the variable/objects
            #Safely exit during an exception
            
            interface.pickError()
            return

