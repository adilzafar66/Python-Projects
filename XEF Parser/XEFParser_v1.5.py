# Author: M. Adil Zafar Khan
# Date: 26th October 2021
# Description: The program takes an PLC XEF or ZEF file and creates a CSV file with all the HMI variables to be exported to Vijeo Designer.

# Importing libraries (make sure any missing libraries are installed)
import os
import csv
import time
import zipfile
import shutil
from pathlib import Path
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter.messagebox import showerror, showwarning, showinfo

#-----------------------------------------------------Functions-----------------------------------------------------

# Copies a file/directory without overwriting if a file/directory already exists with the same input name
def safe_copy(infilePath, outfilePath):

    # Directory path for the copied file
    outDir = Path(outfilePath).parent

    # Check if the file already exists
    if not Path(outfilePath).is_file():
        
        # Copy the file
        shutil.copy(infilePath, outfilePath)
        # Return destination file path
        return outfilePath
    
    # If file exists
    else:
        
        # Split file name and extension into two variables
        base, extension = os.path.splitext(os.path.basename(outfilePath))
        i = 1
        # Loop until file name does not already exist in the destination directory
        while os.path.exists(os.path.join(outDir, '{}_{}{}'.format(base, i, extension))):
            i += 1
        # Copy the file
        shutil.copy(infilePath, os.path.join(outDir, '{}_{}{}'.format(base, i, extension)))
        # Return destination file path
        return os.path.join(outDir, '{}_{}{}'.format(base, i, extension))

# Takes a 2D list and saves the contents of the list to a csv file
def savetoCSV(projectName, listItems, filename):

    # Define all the fields/columns for the csv
    fields = ["Type", "Name","Data Type", "Data Source",
              "Dimension", "Description", "Initial Value",
              "NumofBytes", "Data Sharing", "Alarm", "English ID",
              "Alarm Message", "Alarm Type", "Trigger Condition",
              "Deadband", "Target", "LoLo\\Lo\\Hi\\HiHi", "Minor",
              "Major", "Alarm Group", "Severity", "Vibration Pattern",
              "Vibration Time","Sound File","Play Mode","Scan Group",
              "Device Address", "Bit Number", "Data Format", "Signed",
              "Data Length", "Offset Bit No", "Bit Width", "InputRange",
              "Min", "Max", "DataScaling", "RawMin", "RawMax", "ScaledMin",
              "ScaledMax", "IndirectEnabled", "IndirectAddress","Retentive",
              "LoggingGroup", "LogUserOperationsOnVariable"]
    
    # Create a new csv file
    with open(filename, 'w', newline="") as csvfile:
  
        # creating a csv dict writer object
        writer = csv.writer(csvfile)
        
        # writing headers
        writer.writerow(["'5.1.0", "Vijeo-Designer 6.2.11 CSV output"])
        writer.writerow(fields)
        writer.writerow(["Folder", projectName] + [''] * 44)
  
        # writing data rows
        writer.writerows(listItems)

# Takes in string and substring parameters and returns the substring if it exists within the string
def extractWord(s, subStr):

    # Locate the position of the space character starting from the position of the first letter of the substring
    findDelimiter = s.find(' ', s.find(subStr))

    # If the space character is found
    if findDelimiter != -1:
        
        # Return the characters between starting position of substring and the position of the space character
        return s[s.find(subStr) : findDelimiter]
    
    # If space character is not found but the substring is found
    elif s.find(subStr) != -1:
        
        # Substring is the last word in the string so return the last word
        return s[s.find(subStr) : ]
    else:
        return ''

# Takes in a string, extracts tags from it and puts the results in a list
def parseMetaTags(metaTag):

    # Total number of tags possible = 18
    # Create an empty list of length 18
    tagList = [''] * 18

    # Extract all the tags from the string
    # If a tag does not exist, return an empty string 
    tagList[0] = extractWord(metaTag, '-a ')
    tagList[1] = extractWord(metaTag, '-t ')
    tagList[2] = extractWord(metaTag, '-i ')
    tagList[3] = extractWord(metaTag, '-e ')
    tagList[4] = extractWord(metaTag, '-log')
    tagList[5] = extractWord(metaTag, '-nc')
    tagList[6] = extractWord(metaTag, '-s ')
    tagList[7] = extractWord(metaTag, '-eu=')
    tagList[8] = extractWord(metaTag, '-mineu=')
    tagList[9] = extractWord(metaTag, '-maxeu=')
    tagList[10] = extractWord(metaTag, '-minraw=')
    tagList[11] = extractWord(metaTag, '-maxraw=')
    tagList[12] = extractWord(metaTag, '-logdb=')
    tagList[13] = extractWord(metaTag, '-logper=')
    tagList[14] = extractWord(metaTag, '-onmsg=')
    tagList[15] = extractWord(metaTag, '-offmsg=')
    tagList[16] = extractWord(metaTag, '-area=')
    tagList[17] = extractWord(metaTag, '-f=')

    # Return the filled list
    return tagList

# Extracts the XEF file from a given ZEF file and returns its path
def extractXefFromZef(filePath):

    p = Path(filePath)
    # Copy the file and convert it to .zip extension by calling safe_copy function and return the path of the copied file to filePath
    filePath = safe_copy(p, p.with_suffix('.zip'))
    
    # Unzip the .zip file just created
    unzipDir = os.path.join(os.path.dirname(filePath), Path(filePath).stem)
    with zipfile.ZipFile(filePath, 'r') as zip_ref:
        zip_ref.extractall(unzipDir)
        
    # Delete the .zip file
    os.remove(filePath)
    
    # Return a list containing the paths of all the XEF files in the given directory
    return list(Path(unzipDir).glob('*.xef'))

# Does all the processing on the XEFor ZEF file
# Creates a 2D list with all the information required and saves it to a csv in the given directory path
def parseXEF(filePath, alarmGroup, scanGroup, logStatus, groupLog, fileName):

    # Flag to keep track of whether an unzipped directory needs to be deleted later
    deleteLater = False
    # Stores the directory location
    dirPath = os.path.dirname(filePath)

    # Check if the file has a .zef extension
    if Path(filePath).suffix == '.zef':

        # Get the XEF file from the ZEF file and change the filePath to the newly created XEF file
        filePath = extractXefFromZef(filePath)[0]
        # The unzipped folder will have to be deleted later
        deleteLater = True

    # Check if the file can be opened
    try:
        ET.parse(filePath)
    except:
        debugger.insert(tk.END, "Error: file selected does not have a '.zef' or '.xef' extension.")
        debugger.itemconfig(tk.END, foreground="red")
        showerror(title='File Open Error', message="Cannot open the file specified.")
        return
    else:
        # Put the contents of the XEF file into tree (converts to XML)
        tree = ET.parse(filePath)

    # Empty the log window
    debugger.delete(0, tk.END)
    
    # Get the root element of the XML file
    root = tree.getroot()
    # Create an empty 2D list to store HMI Tags and their attributes
    itemList = []

    # Find the project name
    projectName = root.find("contentHeader").attrib.get('name')[-4:]

    # Variable that counts the number of tags connected to a Logging Group
    logCount = 0
    # Variable that keeps a track of the number of warnings
    debuggerIndex = 1

    # Initialize the display of the debugger
    debugger.insert(tk.END, "Catching warnings... ")
    debugger.insert(tk.END, "-----------------------------------------")
    debugger.insert(tk.END, '')

    # Find all the HMI elements and get the parent element with the child element named attribute of the HMI attribute
    variableLocation = root.iterfind("./dataBlock/variables/attribute[@name='IsVariableHMI']/..[attribute]")
    
    # Loop through all the HMI variables in the dataBlock
    for descendant in variableLocation:

        # List that stores all the information for the current variable in the iteration
        # Total number of columns in the final csv must be 46
        rowList = [''] * 46
        # Initialize an empty list of length 18 (max number of meta tags possible) to store all the meta tags
        metaTagsList = [''] * 18

        # Add the variable name apppended with the project name to the rowList
        rowList[1] = projectName + '.' + descendant.attrib.get('name')

        # Add the address to the rowList
        rowList[26] = descendant.attrib.get('topologicalAddress')

        # Check if the address is NoneType or is not a memory address
        if rowList[26] is None or rowList[26][:2] != '%M':
            # Give a warning in the debugger
            debugger.insert(tk.END, str(debuggerIndex) +'. ' + "Warning: could not import " + rowList[1] + ". No memory address found." )
            # Move debugger to the next line
            debuggerIndex += 1
            # Do not import the current variable
            continue
        
        # Add the variable type to the rowList
        rowList[2] = descendant.attrib.get('typeName')
        
        # Check if the type is EBOOL
        if rowList[2] == "EBOOL":
            # Change it to BOOL because Vijeo only lets you import BOOL and not EBOOL
            rowList[2] = "BOOL"

        # Find the comment attribute of the variable
        temp = descendant.find("comment")
        if temp is not None:
            # Add it to the rowList
            rowList[5] = descendant.find("comment").text

        # Find the meta tags of the variable
        temp = descendant.find("./attribute/[@name='CustomerString']")
        # If meta tags exist
        if temp is not None:
            # Put all the meta tags into the metaTagsList list
            metaTagsList = parseMetaTags(descendant.find("./attribute/[@name='CustomerString']").attrib.get('value'))

        # Check if the variable is an alarm, event or trip
        if metaTagsList[0] == '-a' or metaTagsList[1] == '-t' or metaTagsList[3] == '-e':

            # Fixed values for the alarm, event or trip variables
            rowList[9] = "Enable"
            # Copy the description to the alarm message
            rowList[11] = rowList[5]
            rowList[13] = "when high"
            rowList[16] = "_\\_\\_\\_"
            # Add the user given Scan Group and Logging Group names
            rowList[19] = alarmGroup
            rowList[25] = scanGroup

            # Add severity of each type of event
            # 20 for trip, 10 for alarm and 1 for event
            if '-t' in metaTagsList:
                rowList[20] = 20
                
            elif '-a' in metaTagsList:
                rowList[20] = 10
                
            elif '-e' in metaTagsList:
                rowList[20] = 1
                
            else:
                rowList[20] = ''

        else:
            # Add no default values or severity
            # Disable the alarm
            rowList[9] = "Disable"

        # Check if log status meta tag exists and the Logging Group is enabled by the user
        if metaTagsList[4] == '-log' and logStatus == 1:
            # Add the Group Logging to the rowList
            rowList[44] = groupLog
            # Increment the Group Logging variables count
            logCount += 1
        
        # Fixed Values of columns 1, 4, 5, 11, 26, 35, 36 and 42
        rowList[0] = "Variable"
        rowList[3] = "External"
        rowList[4] = 0
        rowList[10] = 1
        rowList[25] = scanGroup
        rowList[34] = metaTagsList[8][metaTagsList[8].find("=") + 1 : ]
        rowList[35] = metaTagsList[9][metaTagsList[9].find("=") + 1 : ]
        rowList[41] = "Disable"

        # If variable type is not BOOL
        if rowList[2] != 'BOOL':

            # If it is UDINT, set data length to 32 bits
            if rowList[2] == 'UDINT':
                rowList[30] = "32Bits"
                
            # If it is REAL, leave the column empty    
            elif rowList[2] == 'REAL':
                pass

            # If it is INT or UINT, set data length to 16 bits
            else:
                rowList[30] = "16Bits"

            # Disable the input range
            rowList[33] = "Disable"

        # Add the variable to the 2D list
        itemList.append(rowList)
    
    # If the user did not enter a file name, use the file name of the input file for the csv
    if len(fileName) == 0:
        fileName = Path(filePath).stem

    # If the Logging Group variables exceed 100, give a warning since Vijeo will not build the project in that case
    if(logCount > 100):
        debugger.insert(tk.END, str(debuggerIndex) + ". Critical warning: variable Logging Group [count: " + str(logCount) + "] has exceeded the maximum number [100] possible in [HMI_TRE].")
        debugger.itemconfig(tk.END, foreground='red')
        debuggerIndex += 1

    # Save the 2D list to a csv in the same directory as the input file
    savetoCSV(projectName, itemList, os.path.join(dirPath, fileName + "_Vijeo_Export.csv"))

    # Delete the spare folder created by unzipping the ZEF file
    if deleteLater:
        try:
            shutil.rmtree(os.path.dirname(filePath))
        except:
            pass

    # Debugger output at when the program succesfully executes    
    debugger.insert(tk.END, '')
    debugger.insert(tk.END, "Total warnings: " + str(debuggerIndex - 1))
    debugger.insert(tk.END, "Export Succesful.")
    debugger.insert(tk.END, "File export path: " +  os.path.join(dirPath, fileName + "_Vijeo_Export.csv"))
    debugger.itemconfig(debugger.size() - 3, foreground="red")
    debugger.itemconfig(debugger.size() - 2, foreground="green")
    debugger.itemconfig(debugger.size() - 1, foreground="green")
    debugger.itemconfig(0, foreground="black")
    debugger.itemconfig(1, foreground="black")
    debugger.select_set(debugger.size() - 5)
    debugger.see(tk.END)

    # Open the directory where the csv file has been saved
    os.startfile(dirPath)

    # Show a success message box 
    showinfo(title='Success!', message="Your file has been created successfully!")

# Handles all the user input errors    
def errorHandler(filePath, alarmGroup, scanGroup, logStatus, groupLog, fileName):

    # Input file path error
    if len(filePath) == 0:
        debugger.insert(tk.END, "Error: root returns an empty dictionary. No file selected.")
        debugger.itemconfig(tk.END, foreground="red")
        showerror(title='Choose a file', message="Please select a file from the browser.")
        
    # Alarm Group entry error
    elif len(alarmGroup) == 0:
        debugger.insert(tk.END, "Error: variable alarmGroup returns value of type 'NoneType'.")
        debugger.itemconfig(tk.END, foreground="red")
        showerror(title='Empty Field', message="Please enter an alarm group name.")
        
    # Scan Group entry error
    elif len(scanGroup) == 0:
        debugger.insert(tk.END, "Error: variable scanGroup returns value of type 'NoneType'.")
        debugger.itemconfig(tk.END, foreground="red")
        showerror(title='Empty Field', message="Please enter a scan group name.")

    # Logging Group entry error when it is enabled
    elif logStatus == 1 and len(groupLog) == 0:
        debugger.insert(tk.END, "Error: variable groupLog returns value of type 'NoneType'.")
        debugger.itemconfig(tk.END, foreground="red")
        showerror(title='Empty Field', message="Please enter a group log name.")

    # If no errors then run the program
    else:
        parseXEF(filePath, alarmGroup, scanGroup, logStatus, groupLog, fileName)

# Lets a user choose a file from a browser window    
def browse_button():

    # Open browse window
    filename = filedialog.askopenfilename()
    filePath.set(filename)

    # Limit the file path display to 19 characters because our display size is limited
    if len(filePath.get()) != 0:
        ifileName.set(filename[:19] + '...')
    else:
        ifileName.set('')

# Creates a Logging Group entry    
def showInput():

    groupLog_label.grid(column=0, row=6, sticky=tk.W, padx=40)
    groupLog_entry.grid(column=1, row=6, sticky=tk.E, ipadx = 30)

# Hides the Logging Group entry
def hideInput():

    groupLog_entry.grid_forget()
    groupLog_label.grid_forget()

#-----------------------------------------------------Main-----------------------------------------------------
    
# Root window
root = tk.Tk()
root.geometry("400x520+500+140")
root.title('Login')
root.resizable(False, True)
root.minsize(300, 400)

# Configure the grid columns
root.columnconfigure(0)
root.columnconfigure(1)
root.columnconfigure(2)
root.columnconfigure(3)

# Configure the grid rows
root.rowconfigure(0, weight=1)
root.rowconfigure(1, weight=3)
root.rowconfigure(2, weight=1)
root.rowconfigure(3, weight=1)
root.rowconfigure(4, weight=1)
root.rowconfigure(5, weight=1)
root.rowconfigure(6, weight=1)
root.rowconfigure(7, weight=1)
root.rowconfigure(8, weight=1)
root.rowconfigure(9, weight=1)
root.rowconfigure(10, weight=2)

# Window title
root.title("XEF Parser")

# Create global variables for storing input data
filePath = tk.StringVar()
ifileName = tk.StringVar()
alarmGroup = tk.StringVar()
scanGroup = tk.StringVar()
groupLog = tk.StringVar()
ofileName = tk.StringVar()
selected = tk.IntVar()

##canvas = tk.Canvas(root, width = 150, height = 20)
##canvas.grid(row=0, column=1, sticky=tk.W)
##imgLogo = tk.PhotoImage(file="Logo.PPM")
##imgLogo = imgLogo.subsample(5)
##
##canvas.create_image(20,20, anchor=tk.NW, image=imgLogo)

# Title label    
lbl1 = ttk.Label(master=root, text="XEF Parser for Vijeo")
lbl1.config(font=("Tahoma", 14))
lbl1.grid(row=0, sticky=tk.SW, padx=(120,0), column=0, columnspan = 4)

# Label that displays the path after a file is selected from the browse window
lblPath = ttk.Label(master=root, textvariable=ifileName, foreground='grey')
lblPath.grid(row=2, column=1, columnspan=4, sticky=tk.W , padx=(80,0))

# Label for the browse button
filePath_label = ttk.Label(root, text="File Path:")
filePath_label.grid(column=0, row=2, sticky=tk.W, padx=40, pady=5)

# Browse button
button2 = ttk.Button(text="Browse", command=browse_button)
button2.grid(row=2, column=1, sticky=tk.W, pady=5)

# Label for the Alarm Group entry
alarmGroup_label = ttk.Label(root, text="Alarm Group:")
alarmGroup_label.grid(column=0, row=3, sticky=tk.W, padx=40)

# Alarm Group entry
alarmGroup_entry = ttk.Entry(root, textvariable=alarmGroup)
alarmGroup_entry.grid(column=1, row=3, sticky=tk.W, ipadx = 30)

# Label for the Scan Group entry
scanGroup_label = ttk.Label(root, text="Scan Group:")
scanGroup_label.grid(column=0, row=4, sticky=tk.W, padx=40, pady=5)

# Scan Group entry
scanGroup_entry = ttk.Entry(root, textvariable=scanGroup)
scanGroup_entry.grid(column=1, row=4, sticky=tk.W, pady=5, ipadx = 30)

# Label for the Logging Group entry
groupStatus_label = ttk.Label(root, text="Logging Group:")
groupStatus_label.grid(column=0, row=5, sticky=tk.W, padx=40)

# Logging Group entry
groupLog_entry = ttk.Entry(root, textvariable=groupLog)
groupLog_label = ttk.Label(root, text="Log Name:")

# Label for the File Name entry
ofileName_label = ttk.Label(root, text="File Name:")
ofileName_label.grid(column=0, row=7, sticky=tk.W, padx=40, pady=5)

# File Name entry
ofileName_entry = ttk.Entry(root, textvariable=ofileName)
ofileName_entry.grid(column=1, row=7, sticky=tk.W, pady=5, ipadx = 30)

# Radio buttons for the Logging Group enable/disable
r1 = ttk.Radiobutton(root, text='Enable', value=1, variable=selected, command=showInput).grid(column=1, row=5, sticky=tk.W , pady=5)
r2 = ttk.Radiobutton(root, text='Disable', value=2, variable=selected, command=hideInput).grid(column=1, row=5, sticky=tk.E , padx=(0,50), pady=5)

# Disbale the Logging Group as default
selected.set(2)

# Create button which executes the program
create_button = ttk.Button(root, text="Create", command=lambda: errorHandler(filePath.get(), alarmGroup.get(), scanGroup.get(), selected.get(), groupLog.get(), ofileName.get()))
create_button.grid(column=1, row=8, sticky=tk.W , pady=5)

# Label for the log list
debugger_label = ttk.Label(root, text="Status Log:")
debugger_label.grid(column=0, row=9, sticky=tk.NW, padx = 40)

# Log list
debugger = tk.Listbox(root, height=6, width = 42, fg='orange')
debugger.grid(column=0, columnspan=4, row=9, sticky=tk.SW, padx=(40, 0), ipadx = 30, pady=(10,15))

# Horizontal scrollbar for the log list
h=ttk.Scrollbar(root, orient='horizontal', command=debugger.xview)
h.grid(row = 9, column = 0, sticky=tk.SW, columnspan=4, padx=(40, 0), ipadx=131)

# Vertical scrollbar for the log list
v=ttk.Scrollbar(root, orient='vertical', command=debugger.yview)
v.grid(row = 9, column = 3, sticky=tk.SE, ipady=24, pady = (0, 16), padx = (0, 1))

# Add the scroll bars to the log list
debugger.config(xscrollcommand=h.set, yscrollcommand=v.set)

# Loop the GUI
root.mainloop()

#-----------------------------------------------------END-----------------------------------------------------


