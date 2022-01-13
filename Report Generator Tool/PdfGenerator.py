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

#Import required libraries
import os
import re
import PyPDF2
import string
import shutil
import time
from PIL import Image
from pathlib import Path
from tkinter import messagebox
from pdf2image import convert_from_path
from PdfHandler import PdfHandler
from DocHandler import DataImporter

#No maximum pixel range since our input images are too big
Image.MAX_IMAGE_PIXELS = None

#Class that generates TCC Graphs merged into one pdf file
class TCCGraphs(PdfHandler):

    #Initialize the class
    def __init__(self, inFilePath, outFilePath):

        #Set the input directory path
        self.inFilePath = inFilePath

        #Set the output path for the merged pdf
        self.outFilePath = outFilePath

        #Set the directory path for the split pdfs
        self.dirPath = Path(self.inFilePath).parent
        
    #Reorders the files and generates a pdf for each file
    def reorderFiles(self, titleTable):

        #Counter to keep track of titleTable
        i = 0

        #Loop through all the files in the directory
        for fileName in os.listdir(self.dirPath):

            #If the file is a pdf file and not the parent file
            if fileName.endswith(".pdf") and fileName != Path(self.inFilePath).name:

                #Set file path
                filePath = Path(self.dirPath, fileName)

                #Add a '.0' at the end of the number in title if it is a two digit number
                if len(titleTable[i][0]) == 3:
                    #Get numTitle from titleTable
                    numTitle = titleTable[i][0][ : 3] + ".0" + titleTable[i][0][3 : ]
                else:
                    numTitle = titleTable[i][0]

                #Get fault type from titleTable
                fault = titleTable[i][2]

                #'phase' comes before 'ground'
                if fault.lower() == 'phase':
                    outfilename =  numTitle + '_' + '1' + fault + '.pdf'
                else:
                    outfilename =  numTitle + '_' + '2' + fault + '.pdf'

                #Rename the pdf
                filePath.rename(Path(self.dirPath, outfilename))
                
                #Increment counter
                i += 1
                

    #Generates and compiles a merged pdf with all the TCC Graphs in order
    def generateTCC(self, titleTable):

        #Reorder and merge. Delete the files after merging.
        self.reorderFiles(titleTable)
        self.mergePdfs(self.dirPath, self.outFilePath, True, [Path(self.inFilePath)])

#Class that generates Half Cycle Reports merged into one pdf file
class HalfCycleReports(PdfHandler):

    #Initialize the class
    def __init__(self, inFolderPath, outFolderPath):
        
        #Set the input directory path
        self.inFolderPath = inFolderPath

        #Set the output directory path
        self.outFolderPath = outFolderPath

    #Generates and compiles a merged pdf with all the Half Cycle Reports in order
    def generateHalfCycle(self):

        #Loop through all the pdf files in the input directory
        for fileName in os.listdir(self.inFolderPath):
            if fileName.endswith(".pdf"):

                #Set the paths for input and output files
                #Set the output filename as the input file name
                inFilePath = Path(self.inFolderPath, fileName)
                outFilePath = Path(self.outFolderPath, fileName)

                #Use insertBetweenPdf to just copy the pdf to output location
                #Serves no other purpose than just copying the input file to the output path
                self.insertBetweenPdf(inFilePath, outFilePath)

     
class RefInfo(PdfHandler):

    def __init__(self, inFolderPath, outFolderPath):

        #Set the input directory path
        self.inFolderPath = inFolderPath

        #Set the output directory path
        self.outFolderPath = outFolderPath
        
    def generateRefInfo(self):

        #Loop through all the pdf files in the input directory
        for fileName in os.listdir(self.inFolderPath):
            if fileName.endswith(".pdf"):

                #Set the paths for input and output files
                #Set the output filename as the input file name
                inFilePath = Path(self.inFolderPath, fileName)
                outFilePath = Path(self.outFolderPath, fileName)

                #Use insertBetweenPdf to just copy the pdf to output location
                #Serves no other purpose than just copying the input file to the output path
                self.insertBetweenPdf(inFilePath, outFilePath)
                
class FinalReport(PdfHandler):

    def __init__(self, inputPdfPath):

        #Create a 2D list to contain the path and insertion page information of a pdf
        self.pdfPathPage = []

        #Set the input pdf path
        self.inputPdfPath = inputPdfPath

    def createPdfPathPage(self, inputPdfPath, pdfPath, pageDelay):

        pdfPathPageElem = []
        
        #File naming should correspond to this line
        keyword = Path(pdfPath).stem
        
        #PageDelay skips first x pages
        pageNum = self.getPageFromKeyword(inputPdfPath, keyword, pageDelay)

        #Append path and insertion page of the pdf to the the list
        pdfPathPageElem.append(pdfPath)
        pdfPathPageElem.append(pageNum)

        return pdfPathPageElem

    #Compiles the final report
    def compileSearchReport(self, dirPath, pbWindow, pageDelay = 0):
        
        #Input path of the pdf with tags generated in first phase
        searchPdfPath = Path(Path(self.inputPdfPath).parent, Path(self.inputPdfPath).stem + "_Tags.pdf")
        searchPdfName = Path(searchPdfPath).stem

        #Initialize outputPdfPath outside the loop
        outputPdfPath = ''

        #Create a temporary folder for the pdf files with tags used just for searching.
        #This directory will be deleted later.
        tempFolder = Path(Path(self.inputPdfPath).parent, "Temp Files")
        tempFolder.mkdir(parents = True, exist_ok = True)

        #Set counter
        i = 0
        
        #Loop through all pdf files in the directory dirPath
        for fileName in os.listdir(dirPath):
            if fileName.endswith(".pdf"):

                #File path for each file in directory
                filePath = Path(dirPath, fileName)
                
                #Create a 2D list containing the path and page number where the pdf needs to be inserted in the final pdf
                self.pdfPathPage.append(self.createPdfPathPage(searchPdfPath, filePath, pageDelay))
                
                #If no page number is foud, then simply do not insert the pdf
                if self.pdfPathPage[i][1] == None:

                    #Show warning that the pdf could not be merged
                    messagebox.showwarning("Warning",
                                           "Could not import " + Path(filePath).name + ". The tag for the pdf does not exist in the document.",
                                           parent = pbWindow)
                    
                    #Remove the pdfPathPage element that was not imported
                    self.pdfPathPage.pop(i)
                    
                    #Go to next iteration
                    continue
                
                #Set the output path of the compiled report. Set the name equal to the name of the original input Pdf file.
                outputPdfPath = Path(tempFolder, searchPdfName + '_' + str(i) + '.pdf')
                
                #Insert the pdf to the page number specified by pdfPathPage
                self.insertBetweenPdf(searchPdfPath, outputPdfPath, self.pdfPathPage[i])
                
                #Set the next pdf to be merged as the putput pdf
                searchPdfPath = outputPdfPath

                #Increment the counter
                i += 1

        #Remove all the temporary files and directories
        os.remove(Path(Path(self.inputPdfPath).parent, searchPdfName + '.pdf'))
        #shutil.rmtree(tempFolder)

    def compileFinalReport(self):

         #Initialize counter
        i = 0

        #Set the inputPdfPath as the original pdf doc without any tags
        inputPdfPath = self.inputPdfPath

        #Iter through all the items in pdfPathPage 2D list 
        for item in self.pdfPathPage:

            if item == self.pdfPathPage[-1]:
                placeholder = "Final"

            else:
                placeholder = str(i)

            #Set the output path in the same directory as the reportPath
            outputPdfPath = Path(Path(self.inputPdfPath).parent, Path(self.inputPdfPath).stem + '_' + placeholder + '.pdf')

            #Insert the pdf in pdfPathPage into the final output pdf
            self.insertBetweenPdf(inputPdfPath, outputPdfPath, item)

            os.remove(inputPdfPath)
            
            #Set the input pdf as the new created pdf
            inputPdfPath = outputPdfPath

            #Increment the counter
            i+=1
