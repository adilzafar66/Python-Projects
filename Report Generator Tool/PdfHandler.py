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

# Import required libraries
import os
import PyPDF2
import img2pdf
from PIL import Image
from pathlib import Path
import pdfplumber
from PyPDF2 import PdfFileWriter, PdfFileReader
from pdf2image import convert_from_path

# Class that provides all the pdf manipulation functions
class PdfHandler:
        
    # Converts an image to pdf using PIL
    def convertImgToPdf(self, imgPath, pdfPath = None):

        # Open image
        image = Image.open(imgPath)

        # Save in the same directory as the image if pdfPath is not given
        if pdfPath == None:
            image.save(Path(imgPath).with_suffix('.pdf'))
        else:
            image.save(pdfPath)

        #Close the image
        image.close()

    # Converts an image to pdf using pdf bytes
    def convertToPdf(self, imgPath, pdfPath = None):

        # Open image
        image = Image.open(imgPath)

        # Convert image to pdf bytes
        pdfBytes = img2pdf.convert(image.filename)

        # Save in the same directory as the image if pdfPath is not given
        if pdfPath == None:
            file = open(Path(imgPath).with_suffix('.pdf'), "wb")
        else:
            file = open(pdfPath, "wb")

        # Create the pdf
        file.write(pdfBytes)

        # Close files
        image.close()
        file.close()

    # Converts a pdf to image
    def convertPdfToImg(self, pdfPath, imgDir = None):

        # Set pdf name and pdf directory
        pdfName = Path(pdfPath).stem
        pdfDir = Path(pdfPath).parent

        # Save in the same directory as the pdf if imgDir is not given
        if imgDir == None:
            imgDir = pdfDir

        # Extract all the pdf pages
        pages = convert_from_path(pdfPath, dpi = 500)

        # Loop through all the pages and save them as JPEGs
        for i in range(len(pages)):
            pages[i].save(Path(imgDir, pdfName + "-page%s.jpg" % i), 'JPEG')

    # Renames a file
    def renameFile(self, dirPath, oldFileName, newFileName):

        filePath = Path(dirPath, oldFileName)
        return filePath.rename(Path(folderPath, newFileName))

    # Combines all the pdfs in a directory into a single pdf file
    def mergePdfs(self, dirPath, outFilePath, deleteOriginalFiles = False, exclusionPaths = []):

        # exclusionPaths contains paths of pdf files that are ignored by the function
        # Initialize pdf merger
        merger = PyPDF2.PdfFileMerger()

        # Loop through the directory and search for pdf files
        for filename in os.listdir(dirPath):
            filePath = Path(dirPath, filename)
            if filename.endswith(".pdf") and filePath not in [Path(item) for item in exclusionPaths]:

                # Add the pdf to the pdf merger buffer
                merger.append(str(filePath))
                
        # Compile the pdfs into a single pdf
        merger.write(str(outFilePath))
        merger.close()
        del merger

        # If the deleteOriginalFiles variable is passed, delete the original pdfs
        if deleteOriginalFiles:
            self.deletePdfs(dirPath, exclusionPaths)

    # Splits the pages of a pdf and saves each page as a seperate pdf
    def splitPdf(self, pdfPath, deleteParent = False):

        # Set the path of the input pdf
        inputPdf = PdfFileReader(open(pdfPath, "rb"))
        pdfName = Path(pdfPath).stem
        pdfDir = Path(pdfPath).parent

        # Loop through all the pages in the input pdf
        for i in range(inputPdf.numPages):

            # Initialize pdf writer
            output = PdfFileWriter()

            # Add the current page to the pdf writer buffer
            output.addPage(inputPdf.getPage(i))

            # Save the current page as a pdf file
            with open(Path(pdfDir, pdfName + "-page%s.pdf" % i), "wb") as outputStream:
                output.write(outputStream)

        # Delete the original pdf if deleteParent is passed as True
        if deleteParent:
            os.remove(pdfPath)

    # Gets the page number by searching the keyword passed in a pdf
    def getPageFromKeyword(self, inputPdfPath, keyword, pageDelay = 0):

        # Open the pdf file
        pdfFile = pdfplumber.open(inputPdfPath)

        # Loop through the pdf pages
        for page in pdfFile.pages:

            # Skip the first pageDelay amount of pages
            if page.page_number < pageDelay:
                continue

            # Extract the text of the current page
            pageText = str(page.extract_text())

            # Check if the keyword exists in the extracted text
            if pageText.find(keyword) != -1:
                return page.page_number

        # Empty the buffer
        pdfFile.flush_cache()

        # Close file
        pdfFile.close()
        del pdfFile

        return None

    # Inserts a given pdf into another given pdf at a specific location identified by the pdfPathPage list
    def insertBetweenPdf(self, inputPdfPath, outputPdfPath, pdfPathPage = None):

        # Initialize the pdf merger
        merger = PyPDF2.PdfFileMerger()

        # Add the input file to the start
        merger.append(str(inputPdfPath))

        # If the pdfPathPage list is not empty
        if pdfPathPage != None:

            # Insert the pdf in the pdfPathPage[0] at the page number pdfPathPage[1]
            merger.merge(pdfPathPage[1], str(pdfPathPage[0]))

        # Compile and write the new pdf
        merger.write(str(outputPdfPath))

        # Close pdf merger
        merger.close()
        del merger

    def deletePdfs(self, dirPath, exclusionPaths):

        # Loop through all the files in dirPath
        for filename in os.listdir(dirPath):
            filePath = Path(dirPath, filename)
            # If the file is a pdf file and does not have an instance in excluionPaths, delete it
            if filePath not in [Path(item) for item in exclusionPaths] and filename.endswith(".pdf"):
                os.remove(filePath)
