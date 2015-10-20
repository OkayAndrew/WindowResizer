# Description
# - This script starts programs (if not already running) and resized them to the desired
#   monitor location and size. The specific programs and location data is read from an XML
#   configuration file.

# Import Modules
import win32api
import win32gui
import os
import time
import xml.etree.ElementTree as ET

# Reused Functions
# - generate list of open windows (part 1)
def window_enum_handler(hwnd, resultList):
    if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
        resultList.append((hwnd, win32gui.GetWindowText(hwnd)))
# - generate list of open windows (part 2)
def get_app_list(handles=[]):
    mlst=[]
    handles = []
    win32gui.EnumWindows(window_enum_handler, handles)
    for handle in handles:
        mlst.append(handle)
    return mlst

# Initialize Variables
configFileName = 'StartupAndResizerConfig.xml'
sortedMonitorData = sorted(win32api.EnumDisplayMonitors(0), key=lambda xaxis: xaxis[2][0]) # list of monitor data, sorted by x-axis to ensure they're listed according to display from left to right
numberOfMonitors = len(sortedMonitorData)
programList = []

# Read Config Data
try:
    xmlDocument = ET.parse(configFileName)
    xmlRoot = xmlDocument.getroot()
    # - program data
    for i in xrange(len(xmlRoot)): # - loop through each program
        programList.append([])
        for j in xrange(len(xmlRoot[i])): # - loop through each data point for that program
            programList[i].append(xmlRoot[i][j].text)
except:
    print 'ERROR: Could not read the configuration file.\nModify inputs and run again'
    exit

# List Currently Running Programs
# - call function to list all open windows
windowList = get_app_list()

# Start Programs (if not already started)
startupTime = 0
# - loop through each listed program
for thisProgram in programList:
    thisProgramStarted = False
    # - loop throuch each open window
    for thisWindow in windowList:
        # - check if this window matches the desired program
        if thisProgram[4] in thisWindow[1]:
            thisProgramStarted = True
            break
    # - check if this program isn't found anywhere in the list of open windows
    if not thisProgramStarted:
        # - start the program
        os.startfile(thisProgram[0])
        # - check if this program's startup time is greater than the current value
        thisProgramStartupTime = int(thisProgram[5])
        if thisProgramStartupTime > startupTime:
            startupTime = thisProgramStartupTime
# - if a program has been started, wait so that it can come up before moving the windows
if startupTime <> 0:
    time.sleep(startupTime)

# Move Windows
# - loop through each listed program
for thisProgram in programList:
    # - change the monitor number string in the config file to an integer
    try:
        thisProgram[1] = int(thisProgram[1])
    except:
        print 'ERROR: The monitor number of '+thisProgram[1]+' for the program named '+thisProgram[0]+' can not be resolved as an integer.\nModify inputs and run again'
    # - change the program width string in the config file to an integer, if appropriate
    try:
        thisProgram[3] = int(thisProgram[3])
    except:
        a = 1 # no error since program width can be 'full', 'half', or an integer
    # - loop through each open window
    for thisWindow in windowList:
        # - check if this window matches the desired program
        if thisProgram[4] in thisWindow[1]:
            # - find the size/position information for the monitor of interest
            thisMonitor = thisProgram[1]
            if thisMonitor > numberOfMonitors:
                thisMonitor = numberOfMonitors
            thisMonitorXPosition = sortedMonitorData[thisMonitor-1][2][0]
            thisMonitorWidth = sortedMonitorData[thisMonitor-1][2][2] - thisMonitorXPosition
            thisMonitorHeight = sortedMonitorData[thisMonitor-1][2][3]
            # - determine the window width
            if thisProgram[3] == 'full':
                programWidth = thisMonitorWidth
            elif thisProgram[3] == 'half':
                programWidth = thisMonitorWidth/2
            elif isinstance(thisProgram[3],int):
                if thisProgram[3] > 0:
                    programWidth = thisProgram[3]
                else: # a negative width means the full width of the monitor minus this amount
                    programWidth = thisMonitorWidth + thisProgram[3]
            else:
                print 'ERROR: Invalid width value of '+thisProgram[3]+'.\nAcceptable values are "full", "half", or an integer.\nModify inputs and run again'
                exit
            # - determine the window x-position
            if thisProgram[2] == 'full' or thisProgram[2] == 'left':
                programXPosition = thisMonitorXPosition
            elif thisProgram[2] == 'right':
                programXPosition = thisMonitorXPosition + thisMonitorWidth - programWidth
            else:
                print 'ERROR: Invalid side value of '+thisProgram[2]+'.\nAcceptable values are "full", "left", or "right".\nModify inputs and run again'
                exit
            # - move the window to the desired position
            win32gui.MoveWindow(thisWindow[0], programXPosition, 0, programWidth, thisMonitorHeight, True)
            break
