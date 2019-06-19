"""
AUTHOR: Melinda Pullman
SCRIPT: ncep_stageIV_processing.py
ORGANIZATION: USACE MVK District

PURPOSE: This script takes NCEP Stage IV gridded precip in grib files and
converts the files into a DSS file using grib2xmrg.exe andgridLoadXMRG.exe
from gritUtil

DATA: The NCEP Stage IV gridded precip data can be obtained at
http://data.eol.ucar.edu/cgi-bin/codiac/fgr_form/id=21.093

HOW TO EXECUTE SCRIPT: This script creates a windows shell command that will
execute grib2xmrg.exe and gridLoadXMRG.exe. 

I. Setup: 
    1. Python should be on your Windows path; you can test by opening a command
    window and typing “python” at the prompt.
    2. 7-zip should be installed on your Corps computer, but the command used to
    uncompress the .Z file isn’t on the path. It’s probably available as
    C:\Program Files\7-zip\7z.exe though. You can look for the executable or do
    a test from the command line. 

II. Steps:
    1. Create a folder to store the required programs (gridLoadXMRG, grib2XMRG,
    unpackStage4Archive.py), the zipped GRIB files, and the extract file.
    2. Create a subfolder inside the folder you created in step 1. Name this
    folder: “DSS”. This folder is where they output DSS files will be saved.
    3. Copy the required programs into the folder you created in step 1: 
        a. gridLoadXMRG.exe file
        b. grib2XMRG.exe file
        c. unpackStage4Archive.py file
        d. extract.dat file
    4. Download NCEP Stage IV QPE Data from:
    http://data.eol.ucar.edu/cgi-bin/codiac/fgr_form/id=21.093
    5. Save the downloaded QPE Data (it should be inside a zipped folder) into the
    folder you created in step 1 (where dataZvP7ui.zip is the downloaded NCEP QPE
    Data).
    6. Open a windows terminal and navigate to the folder you created in step 1. 
        a. >>>> D:  (change drives)
        b. >>>> cd \NCEP_QPE\  (change directories) 
    7. Once you are in the folder you created in step 1 and you have all of the
    files within that folder, run the following windows terminal command: 
        a. >>>> python unpackStage4Archive.py dataZvP7ui.zip MVKExtract.dat
"""

# import packages        
import sys
import os
import subprocess
from zipfile import ZipFile

# dunno
uncompressCmd = "\"C:\\Program Files\\7-zip\\7z.exe\" x -so "

# open zipped file
with ZipFile(sys.argv[1], 'r') as zipObj:
    # get list of all archived file names from zip
    fileList = zipObj.namelist()
    # iterate over the file names
    for fileName in range(0, len(fileList)-1):
        # check filename that are 1 hr
        if fileList[fileName].startswith('ST4') and fileList[fileName].split('.')[2] == '01h':
            # extract file from zip
            zipObj.extract(fileList[fileName])
            # build shell command
            comLine = uncompressCmd + fileList[fileName] + " | .\\grib2xmrg" +
            " | .\\gridLoadXMRG CWMS=DSS" + " site=CONUS f=stage4" + " extract=" + sys.argv[2]
            # execute shell command
            subprocess.call(comLine, shell = True)
            # remove extracted file
            os.remove(fileList[fileName])
