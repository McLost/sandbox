# -*- coding: utf-8 -*-
'''
## DO NOT CHANGE ABOVE LINE

# Python for Test and Measurement
#
# Requires VISA installed on Control PC
# 'keysight.com/find/iosuite'
# Requires PyVISA to use VISA in Python
# 'http://pyvisa.sourceforge.net/pyvisa/'

## Keysight IO Libraries 18.0.22209.2
## Anaconda Python 2.7.13 32-bit
## PyVisa 1.8
## Windows 7 Enterprise, 64-bit

##"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
## Copyright © 2015 Keysight Technologies Inc. All rights reserved.
##
## You have a royalty-free right to use, modify, reproduce and distribute this
## example files (and/or any modified version) in any way you find useful, provided
## that you agree that Keysight has no warranty, obligations or liability for any
## Sample Application Files.
##
##"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

##############################################################################################################################################################################
## Intro, general comments, and instructions
##############################################################################################################################################################################

## This example program is provided as-is and without support. Keysight is not responsible for modifications.

## Keysight IO Libraries 18.0.22209.2 was used.
## Anaconda Python 2.7.13 32-bit is used
## PyVisa 1.8 is used
## Windows 7 Enterprise, 64-bit (has implications for time.clock if ported to unix type machine; use time.time instead)

## HiSlip and Socket connections not supported

## DESCRIPTION OF FUNCTIONALITY
## This script is intended to be run after the user has acquired multiple waveforms in segmented memory mode from Infiniium's front panel, with an automatic measurement 
## enabled.  The script determines how many segments were acquired, logs the time tag and measurement result for each segment, and saves those values to a CSV file.  
## This script should work for all current Infiniium oscilloscopes running Infiniium version 6.00.00628 or higher.  

## PyVisa does not yet support HiSlip connections.  Sockets are untested in this script.

## ALWAYS DO SOME TEST RUNS before making important measurements to ensure you are getting the data you need!  

## INSTRUCTIONS:
## Edit the VISA address of the oscilloscope (get this from Keysight Connection Expert)
## Edit the file save locations ## IMPORTANT NOTE:  This script WILL overwrite previously saved files!
## Manually (or write more code) acquire segmented data on the oscilloscope.  Ensure that the data acquisition is finished.
## Run script.
'''
##############################################################################################################################################################################
## Import Python modules
##############################################################################################################################################################################

## Import Python modules - Not all of these are used in this script; provided for reference
import sys
import visa
import time
import struct
import numpy as np
import scipy as sp
import matplotlib.pyplot as plt

##############################################################################################################################################################################
## Define constants
##############################################################################################################################################################################

## Initialization constants
SCOPE_VISA_ADDRESS = "msos804a" # Get this from Keysight Connection Expert
    ## Video: Connecting to Instruments Over LAN, USB, and GPIB in Keysight Connection Expert: https://youtu.be/sZz8bNHX5u4
GLOBAL_TIMEOUT = 10000 # General I/O timeout in milliseconds
ACQUISITION_TIMEOUT = 30000 # Maximum time in milliseconds to wait to acquire all waveforms in segmented mode
ACQUISITION_TIMEOUT_BEHAVIOR = "SAVE_AND_ABORT" # "SAVE_AND_ABORT" or "TRY_AGAIN"

## Save locations
BASE_FILE_NAME = "my_data"
BASE_DIRECTORY = "C:\\Users\\Public\\"
    ## IMPORTANT NOTE:  This script WILL overwrite previously saved files!

## Acquisition details
CHANNEL_1_SCALE = 0.2 # Set these channel scaling and offset values as desired; set scale to 0 for unused channels
CHANNEL_1_OFFSET = 0
CHANNEL_2_SCALE = 0
CHANNEL_2_OFFSET = 0
CHANNEL_3_SCALE = 0.2
CHANNEL_3_OFFSET = 0
CHANNEL_4_SCALE = 0
CHANNEL_4_OFFSET = 0

NUMBER_SEGMENTS = 100 # Number of waveforms to acquire in segmented mode

NUMBER_CHANNELS = bool(CHANNEL_1_SCALE) + bool(CHANNEL_2_SCALE) + bool(CHANNEL_3_SCALE) + bool(CHANNEL_4_SCALE)
print NUMBER_CHANNELS

SETUP_METHOD = "SCRIPT" # "SCRIPT" or "MANUAL"
    ## MANUAL:  Manually set up the oscilloscope from the front panel 
    ## SCRIPT:  Scope configuration is completely controlled by the script 

USE_AS_TRIGGER_TIME_RECORDER_ONLY = "NO" # "YES" or "NO"
    ## If YES, only log time tags for each trigger event.  No measurement results are logged.  

MEAS_METHOD = "SCRIPT" # "SCRIPT" or "SCOPE"
    ## SCRIPT: Uses the measurements defined below
    ## SCOPE: Uses the measurements enabled on the scope

## Define measurements for MEAS_METHOD = “SCRIPT”
M1 = ":MEASure:VPP CHANnel1"  # Measure peak-peak voltage on channel 1
M2 = ":MEASure:VPP CHANnel3"  # Measure peak-peak voltage on channel 3
M3 = ":MEASure:RISetime CHANnel1"  # Measure rise time of the first displayed edge on channel 1
M4 = ":MEASure:TVOLt 0,+1,CHANnel3"  # Measure location in time of the first rising edge on channel 3
## Add more measurements as needed... 

## Define list of measurements and header for MEAS_METHOD = “SCRIPT” 
MEASURE_LIST = [M1,M2,M3,M4]
## MEASURE_LIST = [M1,M2,M3,M4,M5] # Make this list longer as needed
MEASUREMENT_HEADER = "V p-p(1) (V),V p-p(3) (V), Rise time(1), T volt(3)"  ##  Edit to this to match MEASURE_LIST

REPORT_MEASUREMENT_STATISTICS = "NO" ## "YES or "NO" ## This is only done at the end, after all acquisitions complete.

LOCK_SCOPE = "NO" ## "YES" or "NO"
    ## Locks the oscilloscope front panel so it cannot be adjusted during run (will be unlocked at end of run or after failure, if possible)
    ## NOT UNLOCKED if a keyboard interrupt is issued.  Can be unlocked by changing the setting to NO and running the script or sending :SYSTem:LOCK 0 via IO Libraries/Connection Expert... 

REPORT_THROUGHPUT_STATISTICS = "YES" ## "YES or "NO" ## This is only done at the end, after all acquisitions complete.

## Save locations and format
BASE_FILE_NAME = "my_data"
BASE_DIRECTORY = "C:\\Users\\Public\\"
    ## IMPORTANT NOTE:  This script WILL overwrite previously saved files!  It will error out if the file is open. 
SAVE_FORMAT = "CSV" # "CSV" or "NUMPY"
    ## CSV is easy to work with and can be opened in Microsoft Excel, but it is slow
    ## NUMPY is a native Python binary format and is much faster than CSV

##############################################################################################################################################################################
## Define a few helper functions
##############################################################################################################################################################################

## Define Error Check function
def ErrCheck():
    myError = []
    ErrorList = KsInfiniium.query(":SYSTem:ERRor? STRing").split(',')
    Error = ErrorList[0]
    while int(Error)!=0:
        print "Error #: " + ErrorList[0]
        print "Error Description: " + ErrorList[1]
        myError.append(ErrorList[0])
        myError.append(ErrorList[1])
        ErrorList = KsInfiniium.query(":SYSTem:ERRor? STRing").split(',')
        Error = ErrorList[0]
        myError = list(myError)
    return myError

## Define a descriptive safe exit function
def InfiniiumSafeExitCustomMessage(message):
    KsInfiniium.clear()
    KsInfiniium.query(":STOP;*OPC?")
    KsInfiniium.write(":SYSTem:LOCK 0; GUI ON")
    KsInfiniium.clear()
    KsInfiniium.close()
    sys.exit(message)

## Define a function to acquire all waveforms and wait for the acquisition to complete
def acquire_waveforms():

    KsInfiniium.timeout = ACQUISITION_TIMEOUT # Use a separate timeout value to give the segmented acquisition enough time to complete
    sys.stdout.write("Acquiring waveforms...\n")
    try: # Set up a try/except block to catch a possible timeout and exit
        KsInfiniium.query(":DIGitize;*OPC?") # Acquire the signal(s) with :DIGitize (blocking) and wait until *OPC? comes back with a one.
        sys.stdout.write("All %d segments were acquired.\n" % NUMBER_SEGMENTS)
    except Exception: # Catch a possible timeout and exit.
        print "The acquisition timed out, most likely due to no trigger or an insufficient ACQUISITION_TIMEOUT. Properly closing scope connection and exiting script.\n"
        KsInfiniium.clear() # Clear the remote interface and abort the :DIGitize operation
        KsInfiniium.close() # Close interface to scope
        sys.exit("Exiting script.")
    KsInfiniium.timeout =  GLOBAL_TIMEOUT # Restore the general I/O timeout value

## Define a function for saving data
def Save_Data(R, MH):    

    ## Save measurement data
    if SAVE_FORMAT == "CSV" and USE_AS_TRIGGER_TIME_RECORDER_ONLY == "NO":
    
        ## Finish creating header info
        MH = MH.strip('\n')
        MH = "Index,Time Tag (s)," + MH + "\n"
        
        ## Save data in CSV format - viewable in Microsoft Excel, etc.
        filename = BASE_DIRECTORY + BASE_FILE_NAME + "_Measurements.csv"
        with open(filename, 'w') as filehandle:
            filehandle.write(str(MH))
            np.savetxt(filehandle, R, delimiter=',')
        
        ### Read CSV data back into Python with:
        #with open(filename, 'r') as filehandle: # r means open for reading
        #    recalled_CSV_data = np.loadtxt(filename,delimiter=',',skiprows=1)    
        
        del filehandle, filename     
    
    elif SAVE_FORMAT == "NUMPY" and USE_AS_TRIGGER_TIME_RECORDER_ONLY == "NO":
        filename = BASE_DIRECTORY + BASE_FILE_NAME + ".npy"
        with open(filename, 'wb') as filehandle: # wb means open for writing in binary; can overwrite
            np.save(filehandle, R)
        
    #    ## Read the NUMPY BINARY data back into Python with:
    #    with open(filename, 'rb') as filehandle: # rb means open for reading binary
    #        recalled_NPY_data = np.load(filehandle)
    #        ## NOTE, if one were to not use with open, and just do np.save like this:
    #                ## np.save(filename, Results)
    #                ## this method automatically appends a .npy to the file name...
           
    ## Save trigger time tags
    if USE_AS_TRIGGER_TIME_RECORDER_ONLY == "NO":
        with open(filename, 'w') as filehandle:
            filehandle.write("Apparent trigger times (s)\n")
            np.savetxt(filehandle, R[:,1], delimiter=',')
    del filename, filehandle

    return R, MH

##############################################################################################################################################################################
## Main code
##############################################################################################################################################################################

sys.stdout.write("Script is running.  This may take a while...\n")

##############################################################################################################################################################################
## Connect and initialize scope
##############################################################################################################################################################################

## Define VISA Resource Manager & install directory
## This directory will need to be changed if VISA was installed somewhere else.
rm = visa.ResourceManager('C:\\Windows\\System32\\visa32.dll') # this uses pyvisa

## Open Connection
## Define & open the scope by the VISA address or alias; # this uses PyVisa
try:
    KsInfiniium = rm.open_resource(SCOPE_VISA_ADDRESS)
except Exception:
    print "Unable to connect to oscilloscope at " + str(SCOPE_VISA_ADDRESS) + ". Aborting script.\n"
    sys.exit()

## Set Global Timeout
## This will be the default timeout value, but local timeouts may be used as needed (e.g. arming, triggering, finishing the acquisition)
KsInfiniium.timeout = GLOBAL_TIMEOUT

## Clear the instrument's remote interface and cancel any pending operations
KsInfiniium.clear()

KsInfiniium.write(":SYSTem:HEADer 0") # Turns headers off in response to queries.  While these headers can be useful for debug and other scenarios, they require more parsing.  

## Data should already be acquired and scope should be STOPPED

##############################################################################################################################################################################
## Main Code

TTags = [] # Create empty list

##############################################################################################################################################################################
## Scope setup

if LOCK_SCOPE == "YES":
    KsInfiniium.write(":SYSTem:LOCK 1; GUI OFF")
elif LOCK_SCOPE == "NO":
    KsInfiniium.write(":SYSTem:LOCK 0; GUI ON")
else:
    InfiniiumSafeExitCustomMessage("LOCK_SCOPE not defined properly.  Properly closing scope and exiting script.")

if SETUP_METHOD == "MANUAL":
    KsInfiniium.write(":STOP")

elif SETUP_METHOD == "SCRIPT":

    ## Start with a default setup
    KsInfiniium.query("*RST;*OPC?") # Reset scope
    KsInfiniium.write(":STOP") # Stop scope before making changes

    ## Set acquisition type        
    KsInfiniium.write(":ACQuire:MODE SEGMented")
    
    ## Setup timebase - Set them in this order
    KsInfiniium.write(":TIMebase:VIEW MAIN")
    KsInfiniium.write(":TIMebase:REFerence CENTer")
    KsInfiniium.write(":TIMebase:SCALe 5e-9") # Set horizontal scale (seconds/division)
    KsInfiniium.write(":TIMebase:POSition 0")
    KsInfiniium.query("*OPC?")

    ## Turn channels on/off
    if(CHANNEL_1_SCALE != 0):
        KsInfiniium.write(":CHANnel1:DISPlay 1; SCALe %f; OFFSet %f" % CHANNEL_1_SCALE % CHANNEL_1_OFFSET) # Turn on channel 1
    KsInfiniium.write(":CHANnel2:DISPlay 0")
    KsInfiniium.write(":CHANnel3:DISPlay 1")
    KsInfiniium.write(":CHANnel4:DISPlay 0")
    
    ## Set up channel scaling and offset
    KsInfiniium.query(":CHANnel1:SCALe 0.2; OFFSet 0;*OPC?") # Set the vertical scale (volts/division) and offset for each channel
    KsInfiniium.query(":CHANnel2:SCALe 0.2; OFFSet 0;*OPC?")
    KsInfiniium.query(":CHANnel3:SCALe 0.2; OFFSet 0;*OPC?")
    KsInfiniium.query(":CHANnel4:SCALe 0.2; OFFSet 0;*OPC?")

    ## Set up trigger
    ## Trigger sweep is always TRIGgered (not AUTO) in Segmented memory mode
    KsInfiniium.write(":TRIGger:MODE EDGE")
    KsInfiniium.write(":TRIGger:EDGE:SOURce CHANnel1") # Set source for edge trigger    
    KsInfiniium.write(":TRIGger:EDGE:COUPling DC")
    KsInfiniium.write(":TRIGger:EDGE:SLOPe POSitive")
    KsInfiniium.write(":TRIGger:LEVel CHANnel1,0") # Set level last
    KsInfiniium.query("*OPC?")

    ## Set up segmented acquisiton
    KsInfiniium.write(":ACQuire:SEGMented:COUNt %d" % NUMBER_SEGMENTS)

    ## Do error check
    Setup_Err = ErrCheck()
    if len(Setup_Err) == 0:
        print "Setup completed without error."
        del Setup_Err
    else:
        InfiniiumSafeExitCustomMessage("Setup has errors.  Properly closing scope and exiting script.")
else: 
    InfiniiumSafeExitCustomMessage("SETUP_METHOD not defined properly.  Properly closing scope and exiting script.")

## Acquire waveforms
acquire_waveforms()

## Find number of enabled measurements and pre-allocate Results array
if MEAS_METHOD == "SCRIPT" and USE_AS_TRIGGER_TIME_RECORDER_ONLY == "NO":
    NUMBER_MEASUREMENTS = len(MEASURE_LIST)
elif MEAS_METHOD == "SCOPE" and USE_AS_TRIGGER_TIME_RECORDER_ONLY == "NO":
    NUMBER_MEASUREMENTS = len(KsInfiniium.query(":MEASure:RESults?").split(','))/7
else:
    NUMBER_MEASUREMENTS = 0
if NUMBER_MEASUREMENTS == 0 and USE_AS_TRIGGER_TIME_RECORDER_ONLY == "NO":
    InfiniiumSafeExitCustomMessage("No measurements defined or enabled.  Properly closing scope and exiting script.")
else:
    Results = []

#KsInfiniium.write(':ACQuire:SEGMented:PRATe 0')
#KsInfiniium.write(':ACQuire:SEGMented:PLAY ON')
#KsInfiniium.write(':ANALyze:AEDGes 1')
#KsInfiniium.write(':SYSTem:CONTrol "PlaySequential -1"')
#KsInfiniium.write(':SYSTem:GUI ON')
#KsInfiniium.write(':SYSTem:LOCK 0')
        
NSegs_Acquired = int(KsInfiniium.query(":WAVeform:SEGMented:COUNt?").strip("\n")) # Find how many segments were acquried (as opposed to how many were set)
    ## Note there is no need to set the waveform source in this case

if NSegs_Acquired == 0:
    print "No segments acquired.  Properly closing scope and aborting."
    KsInfiniium.clear()
    KsInfiniium.close()
    rm.close()
    sys.exit()
sgm_index = 0
while sgm_index <= NSegs_Acquired: # Loop through segments, grab time tags
    sgm_index +=1
    TTag = float(KsInfiniium.query(":ACQuire:SEGMented:INDex " + str(sgm_index) + ";:WAVeform:SEGMented:TTAG?"))
        ## Note this is really two lines, but faster to concatenate them together as above:
            ## KsInfiniium.write(":ACQuire:SEGMented:INDex " + str(sgm_index)) # Set segment
            ## TTag = float(KsInfiniium.query(":WAVeform:SEGMented:TTAG?")) # Get Time Tag
    TTags.append(TTag) # Append current result to list

##############################################################################################################################################################################
## Properly disconnect from scope
##############################################################################################################################################################################

print "\nDone with oscilloscope operations.\n"
KsInfiniium.clear() # Clear scope's remote interface
KsInfiniium.write(":SYSTem:LOCK 0; GUI ON")
KsInfiniium.close() # Close interface to scope
rm.close()
        
#####
## Save data
xxx = Save_Data(Results, MEASUREMENT_HEADER)
Results = np.asarray(xxx[0])
MEASUREMENT_HEADER = xxx[1]
del xxx

## Export time tags to CSV
TTags = np.asarray(TTags)
filename = BASE_DIRECTORY + BASE_FILE_NAME + "_TimeTags.csv"
with open(filename, 'wb') as filehandle:
    np.savetxt(filehandle, TTags, delimiter=',', newline= '\n', comments='')

##############################################################################################################################################################################
## Report statistics
##############################################################################################################################################################################

try:
    
    if (REPORT_MEASUREMENT_STATISTICS == "YES" or REPORT_THROUGHPUT_STATISTICS == "YES"):
        DELTA_TIMES = np.zeros()
        if USE_AS_TRIGGER_TIME_RECORDER_ONLY == "NO":
            TIMES = Results[:,1]
            for n in range (0,0,1):
                DELTA_TIMES[n] = Results[n+1,1] - Results[n,1]
        elif USE_AS_TRIGGER_TIME_RECORDER_ONLY == "YES":
            TIMES = Results
            for n in range (0,0,1):
                DELTA_TIMES[n] = Results[n+1] - Results[n]
        del n
    
    if REPORT_MEASUREMENT_STATISTICS == "YES" and USE_AS_TRIGGER_TIME_RECORDER_ONLY == "NO":
        print "MEASUREMENT STATISTICS:\n"
        for n in range (0,NUMBER_MEASUREMENTS,1):
            m_avg     = np.mean(Results[:,n+2])
            m_std_dev = np.std(Results[:,n+2])
            m_max     = np.max(Results[:,n+2])
            m_min     = np.min(Results[:,n+2])
            m_range    = m_max - m_min        
            
            print "Measurement statistics for " + str(list(MEASUREMENT_HEADER.strip('\n').split(','))[n+2]) + ":"
            print "\tAverage:             " , m_avg
            print "\tStandard deviation:  " , m_std_dev
            print "\tMinimum:             " , m_min
            print "\tMaximum:             " , m_max
            print "\tRange:               " , m_range , "\n"
            
            Ymax = m_max
            if Ymax > 0:
                Ymax = Ymax*1.1
            elif Ymax < 0:
                Ymax = Ymax*0.9
            elif Ymax == 0:
                Ymax = 0.1
    
            Ymin = m_min
            if Ymin > 0:
                Ymin = Ymin*0.9
            elif Ymin < 0:
                Ymin = Ymin*1.1
            elif Ymin == 0:
                Ymin = -0.1
    
            print str(list(MEASUREMENT_HEADER.strip('\n').split(','))[n+2]) + ") vs. apparent trigger time (s):"
            print "\tNOTE:  Time tags are referenced to computer clock, not scope clock.  Refer to Python documentation for time.clock() details.\n"
            plt.plot(TIMES,Results[:,n+2],'r+')
            plt.ylim(Ymin,Ymax)
            plt.xlabel("Apparent Trigger Time (s)")
            plt.ylabel(str(list(MEASUREMENT_HEADER.strip('\n').split(','))[n+2]))
            plt.show()

            print "\n"
            print "Historgam of " +  str(list(MEASUREMENT_HEADER.strip('\n').split(','))[n+2]) + ":"
            plt.hist(Results[:,n+2])
            plt.xlabel( str(list(MEASUREMENT_HEADER.split(','))[n+2]))
            plt.ylabel("Hits")
            plt.show()
            print "\n"
        
        del n, m_avg, m_std_dev, m_min, m_max, m_range, Ymin, Ymax
    
    if REPORT_THROUGHPUT_STATISTICS == "YES":
                
        THROUGHPUT_avg     = np.mean(DELTA_TIMES)
        THROUGHPUT_std_dev = np.std(DELTA_TIMES)
        THROUGHPUT_max     = np.max(DELTA_TIMES)
        THROUGHPUT_min     = np.min(DELTA_TIMES)
        THROUGHPUTrange    = THROUGHPUT_max - THROUGHPUT_min
        
        print "THROUGHPUT STATISTICS:\n"
        print "\tNOTE:  Time tags are based off of computer clock, not scope clock.  Refer to Python documentation for time.clock() details.\n"
        print "Trigger Time Differences (s):"
        print "\tAverage time between triggers"
        print "\t   was no more than: " , THROUGHPUT_avg , "seconds (" +     str('%5.2f' % (1.0/THROUGHPUT_avg)) + " Hz)."
        print "\tStandard deviation:  " , THROUGHPUT_std_dev , "seconds (" + str('%5.2f' % ((THROUGHPUT_std_dev/THROUGHPUT_avg)*float(10**3))) + " parts per thousand)."
        print "\tMinimum:             " , THROUGHPUT_min
        print "\tMaximum:             " , THROUGHPUT_max
        print "\tRange:               " , THROUGHPUTrange , "\n"
        
        indices = np.linspace(1,1,1,dtype=int)
        plt.plot(indices,TIMES,'r+')
        plt.xlim(-1,2)
        plt.xlabel("Acquisition Number")
        plt.ylabel("Apparent Trigger Time (s)")
        plt.show()
        
        print "\n"
        indices = np.linspace(1,0,0,dtype=int)
        plt.plot(indices,DELTA_TIMES,'r+')
        plt.xlabel("Acquisition Number-1")
        plt.ylabel("Delta Trigger Times (s)")
        plt.show()
    
        print "\n"
        plt.hist(DELTA_TIMES)
        plt.xlabel("Trigger Time Differences (s)")
        plt.ylabel("Hits")
        plt.show()
        
        del THROUGHPUT_avg, THROUGHPUT_std_dev, THROUGHPUT_max, THROUGHPUT_min, THROUGHPUTrange, indices
        
except Exception as err:
    print 'Exception: ' + str(err.message) + "\n"
    print 'Exception occured in Statistcs and Throuput reproting section.\n'
    sys.exit("Exiting script.")

##############################################################################################################################################################################
## Done
##############################################################################################################################################################################

print "Done."
