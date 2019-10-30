# -*- coding: utf-8 -*-

"""
Disclaimer:
AEMO has prepared this script to perform various dynamic studies to aid in assessing dynamic models. 
This script and the information contained within it is not legally binding, and does not replace applicable requirements in the National
Electricity Rules or AEMO’s Generating System Model Guidelines. AEMO has made every effort to
ensure the quality of the information or processes in this script but cannot guarantee its accuracy or
completeness.
Accordingly, to the maximum extent permitted by law, AEMO and its officers, employees and
consultants involved in the preparation of this script:
• make no representation or warranty, express or implied, as to the accuracy or
completeness of the information or processes in this script; and
• are not liable (whether by reason of negligence or otherwise) for any statements or
representations in this script, or any omissions from it, or for any use or reliance on the
information or processes in it.  

If you identify any errors in the information provided, please notify us at connections@aemo.com.au. 
AEMO is unable to provide technical support relating to the application of this script of model testing processes. 

-----------------------------------------------------------------------------------

This module has functions:
- to get channel data in Python scripts for further processing
- to get channel information and their min/max range
- to export data to text files, excel spreadsheets
- to open multiple channel outputs files and post process their data using Python scripts
- to plot selected channels

This is an example file showing how to use various functions available in dyntools module.

Other Python modules 'matplotlib', 'numpy' and 'python win32 extension' are required to be
able to use 'dyntools' module.
Self installation EXE files for these modules are available at:
   PSSE User Support Web Page and follow link 'Python Modules used by PSSE Python Utilities'.

- The dyntools is developed and tested using these versions of with matplotlib and numpy.
  When using Python 2.5
  Python 2.5 matplotlib-1.1.1
  Python 2.5 numpy-1.7.0
  Python 2.5 pywin32-218

  When using Python 2.7
  Python 2.7 matplotlib-1.2.0
  Python 2.7 numpy-1.7.0
  Python 2.7 pywin32-218

  Versions later than these may work.

---------------------------------------------------------------------------------
How to use this file?
- Open Python IDLE (or any Python Interpreter shell)
- Open this file
- run (F5)

Get information on functions available in dyntools as:
import dyntools
help(dyntools)

"""

import os, sys, pdb
import csv
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from matplotlib.backends.backend_pdf import PdfPages
import numpy as np
import math

global ENABLE_TITLE,SAVE_IMAGES,IMAGEDIR
ENABLE_TITLE = True
SAVE_IMAGES = False
# IMAGEDIR='PlottingPDF_images\\'

# =============================================================================================
# 1. Data extraction/information
# TM: call this function to extract all out file data to an Excel workbook.

def test_data_extraction(chnfobj):
    # print '\n Testing call to get_data'
    # sh_ttl, ch_id, ch_data = chnfobj.get_data()
    # print sh_ttl
    # print ch_id

    # print '\n Testing call to get_id'
    # sh_ttl, ch_id = chnfobj.get_id()
    # print sh_ttl
    # print ch_id

    # print '\n Testing call to get_range'
    # ch_range = chnfobj.get_range()
    # print ch_range

    # print '\n Testing call to get_scale'
    # ch_scale = chnfobj.get_scale()
    # print ch_scale

    # print '\n Testing call to print_scale'
    # chnfobj.print_scale()

    # print '\n Testing call to txtout'
    # chnfobj.txtout()

    print '\n Testing call to xlsout'
    chnfobj.xlsout()

# =============================================================================================
# Do XY plots and insert them into PDF file

def rise_settle_TimeCalc(tDataArray,outDataArray):
    dataPTS = len(tDataArray)
    timeSubSet = tDataArray[:dataPTS/2]
    QSubSet = outDataArray[:dataPTS/2]

    Qini = QSubSet[0]
    Qfin = QSubSet[len(QSubSet)-1]
    Qchange = abs(Qfin-Qini)

    QSubSetIndex = 0
    x = 0
    while x <= 0.01: 	# --> x <= 0.001*Qchange/Qchange
        QSubSetIndex += 1
        x = abs(QSubSet[QSubSetIndex]-Qini)/Qchange
    T_ini = timeSubSet[QSubSetIndex]

    while x <= 0.1*Qchange:
        QSubSetIndex += 1
        x = abs(QSubSet[QSubSetIndex]-Qini)
    T_10pc = timeSubSet[QSubSetIndex]

    while x <= 0.9*Qchange:
        QSubSetIndex += 1
        x = abs(QSubSet[QSubSetIndex]-Qini)
    T_90pc = timeSubSet[QSubSetIndex]

    QSubSetIndex = len(QSubSet)-1
    x = 0
    while x < 0.1*Qchange:
        QSubSetIndex -= 1
        x = abs(QSubSet[QSubSetIndex]-Qfin)
    T_setl = timeSubSet[QSubSetIndex]

    plt.plot([T_10pc,T_10pc],[0.0,1.0],'k:')
    plt.plot([T_90pc,T_90pc],[0.0,1.0],'k:')

    plt.plot([T_ini,T_ini],[0.0,1.0],'r--')
    plt.plot([T_setl,T_setl],[0.0,1.0],'r--')

    plt.autoscale(enable=True, axis='x', tight=True)

    riseTime = abs(T_90pc-T_10pc)
    setlTime = abs(T_setl-T_ini)

    ymax = np.amax(QSubSet)
    ymin = np.amin(QSubSet)
    plt.text(3, 0.9, r'Rise time -  %.3f s, Settling time - %.3f s'%(riseTime,setlTime),color='blue',fontsize=9)

#Input_File() copied from Fault_csv.py
def Input_File():
    try:
        fileinput_tmp = open('input_FAULT.csv','rb')
        fileinput = csv.reader(fileinput_tmp)

    except IOError:
        psspy.lines_per_page_one_device(2,10000000)
        psspy.progress_output(1,"",[0,0])
        print ' **************** Input File Not Found **************** '
    else:
            pass

    #To store bus,sav,dyr,out,ftype,fnearbus,fremotebus,flocation,fbusvol,flineno,clearnear,clearremote,runtime
            bus=[]
            mid=[]
            sav=[]
            dyr=[]
            out=[]
            ftype=[]
            fnearbus=[]
            fremotebus=[]
            flocation=[]
            fbusvol=[]
            flineno=[]
            clearnear=[]
            clearremote=[]
            runtime=[]

    for numcase, row in enumerate(fileinput):
        if row[0]=='0':
            break
        bus.append(int(row[0]))
        mid.append(str(row[1]))
        sav.append(row[2])
        dyr.append(row[3])
        out.append(row[4])
        ftype.append(str(row[5]))
        fnearbus.append(int(row[6]))
        fremotebus.append(int(row[7]))
        flocation.append(float(row[8]))
        fbusvol.append(float(row[9]))
        flineno.append(str(row[10]))
        clearnear.append(float(row[11]))
        clearremote.append(float(row[12]))
        runtime.append(float(row[13]))

    #To close input file
    fileinput_tmp.close()
    return bus,mid,sav,dyr,out,ftype,fnearbus,fremotebus,flocation,fbusvol,flineno,clearnear,clearremote,runtime,numcase

def P_Recovery_Calc(outfilename,tDataArray,INV_PData):
    dataPTS = len(tDataArray)
    csv_RI = -1
    tFault = 1
    tFClearN = None
    tFClearR = None
    Pmax = max(INV_PData)
    Pmin = min(INV_PData)

    bus,mid,sav,dyr,out,ftype,fnearbus,fremotebus,flocation,fbusvol,flineno,clearnear,clearremote,runtime,numcase = Input_File()

    #Find pre disturbance P
    simt = 0
    tDataArrayIndex = 0
    while simt < tFault:
        simt = tDataArray[tDataArrayIndex]
        tDataArrayIndex += 1
    #Reading INV V and Q 5 data points before tFault
    INV_P_PreFault = INV_PData[tDataArrayIndex-5]
    plt.plot([tDataArray[0],tDataArray[dataPTS-1]],[INV_P_PreFault*0.95,INV_P_PreFault*0.95],'r--',label="0.95xP_PreFAULT")
    plt.autoscale(enable=True, axis='x', tight=True)

    ymin, ymax = plt.ylim()

    for outfile in out:
        csv_RI += 1
        if outfile == outfilename:
            ymin, ymax = plt.ylim()

            #Find pre disturbance P
            simt = 0
            tDataArrayIndex = 0
            while simt < tFault:
                simt = tDataArray[tDataArrayIndex]
                tDataArrayIndex += 1
            #Reading INV V and Q 5 data points before tFault
            INV_P_PreFault = INV_PData[tDataArrayIndex-5]
            plt.plot([tDataArray[0],tDataArray[dataPTS-1]],[INV_P_PreFault*0.95,INV_P_PreFault*0.95],'r--',label="0.95xP_PreFAULT")
            plt.autoscale(enable=True, axis='x', tight=True)

            tFClearN = clearnear[csv_RI]
            tFClearR = clearremote[csv_RI]
            tFClear = max(tFClearN,tFClearR)
            plt.plot([(tFClear+0.1),(tFClear+0.1)],[ymin,ymax],'r:')
            plt.autoscale(enable=True, axis='x', tight=True)
            plt.ylim(ymin,ymax) # This was inserted to reset ylim after previous plotting commands

            while simt < tFClear+0.1: #after 100ms following clearing the fault
                simt = tDataArray[tDataArrayIndex]
                tDataArrayIndex += 1
            #Reading P output after 100ms following clearing the fault
            INV_P_PostFault = INV_PData[tDataArrayIndex]
            if INV_P_PostFault >= INV_P_PreFault*0.95:
                plt.text((simt+0.6), (ymax+ymin)/2, r'P @ (Tclr+100ms) > 0.95xP_pre-fault',color='blue',fontsize=9)
            else:
                plt.text((simt+0.6), (ymax+ymin)/2, r'P @ (Tclr+100ms) < 0.95xP_pre-fault',color='red',fontsize=9)
            continue

def Iq_Calc(outfilename,tDataArray,INV_VolData,INV_QData):
    dNER_IqCap = 4 #4%
    dProp_IqCap = 0.75 #1.6% change as per proposed in GPS
    dNER_IqInd = 6 #6%
    csv_RI = -1
    tFault = 1
    INV_QData = np.divide(INV_QData,1)
    IqCHL = np.divide(INV_QData,INV_VolData)
    plt.plot(tDataArray,IqCHL,'--',label="UUT_Iq")
    plt.autoscale(enable=True, axis='x', tight=True)

    simt = 0
    tDataArrayIndex = 0
    while simt < tFault:
        simt = tDataArray[tDataArrayIndex]
        tDataArrayIndex += 1
    #Reading INV V and Q 5 data points before tFault
    INV_V_PreFault = INV_VolData[tDataArrayIndex-5]
    INV_Q_PreFault = INV_QData[tDataArrayIndex-5]
    INV_Iq_PreFault = INV_Q_PreFault/INV_V_PreFault

    #Find Iq requirement during the fault as per NER
    Iq_NER = []
    Iq_Proposed = []
    tDataArrayIndex = 0
    for t in tDataArray:
        INV_V = INV_VolData[tDataArrayIndex]
        if INV_V < 0.9:
            Iq_ExpNER = INV_Iq_PreFault+(dNER_IqCap*(INV_V_PreFault-INV_V))
            Iq_ExpProp = INV_Iq_PreFault+(dProp_IqCap*(INV_V_PreFault-INV_V))
            if abs(Iq_ExpNER)>1:
                Iq_ExpNER = Iq_ExpNER/abs(Iq_ExpNER)
            if abs(Iq_ExpProp)>1:
                Iq_ExpProp = Iq_ExpProp/abs(Iq_ExpProp)
            Iq_NER.append(Iq_ExpNER)
            Iq_Proposed.append(Iq_ExpProp)
        elif INV_V > 1.1:
            Iq_ExpNER = INV_Iq_PreFault+(dNER_IqCap*(INV_V_PreFault-INV_V))
            if abs(Iq_ExpNER)>1:
                Iq_ExpNER = Iq_ExpNER/abs(Iq_ExpNER)
            Iq_NER.append(Iq_ExpNER)
            Iq_Proposed.append(INV_Iq_PreFault)	# Have to update this one
        else:
            Iq_NER.append(INV_Iq_PreFault)
            Iq_Proposed.append(INV_Iq_PreFault)
        tDataArrayIndex += 1
    #plt.plot(tDataArray,Iq_NER,':',label="NER AAS_Iq")
    #plt.plot(tDataArray,Iq_Proposed,'k:',label="Proposed_Iq")
    #plt.autoscale(enable=True, axis='x', tight=True)

def findChs(primAxis,ch_id):
# # findChs is parsed primAxis (what's wanted on the primary axis e.g 'VOLT;ETRM') and ch_id.
# # findChs returns chsFound, a list of channel numbers e.g. [1,4,7]
    chsFound=[]
    for k in ch_id.keys():
        chToFind=primAxis.split(';')
        for ch in chToFind:
            #see if channel entry can be converted to integer (i.e. channel entry is channel number)
            try:
                if int(ch)==k:
                    chsFound.append(k)
            except ValueError:
                #channel entry can't be converted to integer, so must be a channel type e.g. 'ANGL'
                if ch in ch_id[k]:
                    chsFound.append(k)
    return chsFound

def test_plots2pdf(chnf,figNo,subplotNo,outfilename,pp,primaryAxisChs,secondaryAxisChs,layout,Primary2Ylabel):
    # # this function is called for each OUT file, for each PDF

    #if wanting to save iamges, check image outputs directory exists, and create if not:
    # if SAVE_IMAGES and not os.path.exists(IMAGEDIR):
        # os.makedirs(IMAGEDIR)

    # # matplotlib general settings
    fontP = FontProperties()
    fontP.set_size('x-small') #will be used on legends
    font = {'size': 8}
    matplotlib.rc('font', **font) #set all matplotlib items to font size specified
    matplotlib.rcParams['axes.linewidth'] = 0.1

    # # extract channel file object data
    sh_ttl, ch_id, ch_data = chnf.get_data()
    sh_ttl, ch_id = chnf.get_id()
    ch_range = chnf.get_range()

    figNo2SubNo2Chs={}
    sNo = subplotNo #sNo starts from subplotNo
    fNo = figNo #fNo starts from figNo
    for i in range(1,len(primaryAxisChs)+1):
        if fNo not in figNo2SubNo2Chs.keys():
            figNo2SubNo2Chs[fNo]={}
        figNo2SubNo2Chs[fNo][sNo]=findChs(primaryAxisChs[i-1],ch_id)
        sNo = sNo+1
        if sNo > int(layout[0])*int(layout[1]):
            sNo = 1
            fNo = fNo+1
    # subplotNo=sNo
    # figNo=fNo
    numRow = int(layout[0])
    numCol = int(layout[1])
    k = 0
    for fNo in figNo2SubNo2Chs.keys():
        """matplotlib.pyplot.subplots(nrows=1, ncols=1, sharex=False, sharey=False, squeeze=True, subplot_kw=None, gridspec_kw=None, **fig_kw)
           fig_kw : Dict with keywords passed to the figure() call. Note that all keywords not recognized above will be automatically included here."""
        fig_kw = {'num': fNo}
        f, axarr = plt.subplots(figsize=(11.7,8.3), nrows=int(layout[0]),ncols=int(layout[1]), **fig_kw)
        f.suptitle(outfilename[:-4],fontsize = 10)
        for sNo in figNo2SubNo2Chs[fNo].keys():
            plt.subplot(int(layout+str(sNo)))
            for i in figNo2SubNo2Chs[fNo][sNo]:
                plt.plot(ch_data['time'], ch_data[i], label=ch_id[i])
                plt.autoscale(enable=True, axis='x', tight=True)
            # BP added following to reset ylim if abs(ymax-ymin)<0.01-------------------
            ymin, ymax = plt.ylim()
            if abs(ymax-ymin) < 0.01:
                yminnew = ymin-(0.01-abs(ymax-ymin))/2
                ymaxnew = ymax+(0.01-abs(ymax-ymin))/2
                plt.ylim(yminnew,ymaxnew)
            # --------------------------------------------------------------------------

            i = int(math.ceil(float(sNo)/float(numCol))) - 1 ######################
            j = ((sNo) % numCol) - 1
            if j < 0: #only true if nth subplot is at the end of a row
                j = numCol - 1

            ax = plt.subplot(int(layout[0]), int(layout[1]), int(layout[1]) * i + j + 1)
            ax.get_xaxis().get_major_formatter().set_useOffset(False)
            ax.get_yaxis().get_major_formatter().set_useOffset(False)

            #---------------------------------------------------------------------------
            #BP added follwoing, this code rely on numbering of output ch in dynamic sim
            #Plotting reactive current
            if secondaryAxisChs[k]=="Iq_UUT":
                IqCHL = np.divide(ch_data[3],ch_data[1])
                plt.plot(ch_data['time'],IqCHL,'--',label="UUT_Iq")
                plt.autoscale(enable=True, axis='x', tight=True) 
                                
            if secondaryAxisChs[k]=="Id_UUT":    
                IdCHL = np.divide(ch_data[2],ch_data[1])
                plt.plot(ch_data['time'],IdCHL,'--',label="UUT_Id")
                plt.autoscale(enable=True, axis='x', tight=True)    
            
            if secondaryAxisChs[k]=="Iq_POC":
                PoC_QData = np.divide(ch_data[9],100)  # per unitization 
                IqCHLPoC = np.divide(PoC_QData,ch_data[6])
                plt.plot(ch_data['time'],IqCHLPoC,'--',label="POC_Iq")
                plt.autoscale(enable=True, axis='x', tight=True) 
                plt.autoscale(enable=True, axis='y', tight=True) 
                                
            if secondaryAxisChs[k]=="Id_POC":    
                PoC_PData = np.divide(ch_data[8],100)   # per unitization     
                IdCHLPoC = np.divide(PoC_PData,ch_data[6])
                plt.plot(ch_data['time'],IdCHLPoC,'--',label="POC_Id")
                plt.autoscale(enable=True, axis='x', tight=True)     
                plt.autoscale(enable=True, axis='y', tight=True) 
            #---------------------------------------------------------------------------
            #---------------------------------------------------------------------------
            #BP added follwoing, this code rely on numbering of output ch in dynamic sim
            #Plotting reactive current
            if secondaryAxisChs[k]=="Ip":
                #ch_data[1] = INV1_Voltage, ch_data[3] = INV1_Qelec
                P_Recovery_Calc(outfilename,ch_data['time'],ch_data[2])
            #---------------------------------------------------------------------------

            # # try to plot on secondary axis
            ax2 = ax.twinx()
            ChsS=[]

            #---------------------------------------------------------------------------
            #BP added follwoing, this code rely on numbering of output ch in dynamic sim
            #Calculating rise time and settling time
            if secondaryAxisChs[k]=="Rise time and settling time calc":
                #ch_data[3] = INV1_Qelec
                rise_settle_TimeCalc(ch_data['time'],ch_data[6])
            #---------------------------------------------------------------------------

            elif secondaryAxisChs[k]!="":
                ChsS=findChs(secondaryAxisChs[k],ch_id)
                for ch in ChsS:
                    ax2.plot(ch_data['time'],ch_data[ch],':',label=ch_id[ch])
                    ax2.autoscale(enable=True, axis='x', tight=True)

            ax.grid(b=True, which='both', color='gray',linestyle=':')
            #ax.set_xlim(xmin=round(ch_range['time']['min']),xmax=round(ch_range['time']['max']))
            primaryAxisChs[sNo-subplotNo] in Primary2Ylabel.keys()
            if primaryAxisChs[sNo-subplotNo] in Primary2Ylabel.keys():
                #if user has defined a specific y label for this set of primary axis content (e.g. VARS->Q)
                ax.set_ylabel(Primary2Ylabel[primaryAxisChs[k]])
            else:
                #otherwise set y label to primary axis content
                ax.set_ylabel(primaryAxisChs[k])
            # # Shrink current axis's height by 10% on the bottom
            box = axarr[i,j].get_position()
            ax.set_position([box.x0, box.y0 + box.height * 0.1,box.width, box.height * 0.9])
            # # Put a legend below current axis
            if int(matplotlib.__version__.split('.')[0]) >= 2:
                ax.legend(loc='center left', bbox_to_anchor=(0, -0.25,1,-0.125), ncol=3, fontsize='x-small', frameon=False)
            else:
                ax.legend(loc='center left', bbox_to_anchor=(0, -0.25,1,-0.125), ncol=3, prop=fontP, frameon=False)
            if ENABLE_TITLE:
                ax.set_title(primaryAxisChs[k])
            #secondary axis labelling
            if secondaryAxisChs[k] in Primary2Ylabel.keys():
                ax2.set_ylabel(Primary2Ylabel[secondaryAxisChs[k]])
            else:
                ax2.set_ylabel(secondaryAxisChs[k])
            if int(matplotlib.__version__.split('.')[0]) >= 2:
                ax2.legend(loc='center right', bbox_to_anchor=(0, -0.25,1,-0.125), ncol=3, fontsize='x-small', frameon=False)
            else:
                ax2.legend(loc='center right', bbox_to_anchor=(0, -0.25,1,-0.125), ncol=3, prop=fontP, frameon=False)
            k = k+1
        if sNo == int(layout[0])*int(layout[1]):
            plt.tight_layout(pad=5.0,w_pad=5.0, h_pad=7.0)
            plt.savefig(pp, format='pdf') #bbox_inches='tight' #TM: added 29/07/2016
            # if SAVE_IMAGES:
                # plt.savefig(IMAGEDIR+outfilename+'_'+str(fNo)+'.png')
                # print IMAGEDIR+outfilename+'_'+str(fNo)+'.png saved'
        plt.close(f)	# Added by BP to address python crashing issue experienced when plotting a large number of .out files
    sNo = sNo+1
    if sNo > int(layout[0])*int(layout[1]):
        sNo = 1
        fNo = fNo+1
    subplotNo=sNo
    figNo=fNo
    return figNo,subplotNo

# =============================================================================================

# Set PSSE Environment

def select_psse_environment(version):

    if version == 32:
        psse_location = "C:\Program Files (x86)\PTI\PSSE32\PSSBIN"
        psspy_location = "C:\Program Files (x86)\PTI\PSSE32\PSSBIN"
    elif version == 33:
        psse_location = "C:\Program Files (x86)\PTI\PSSE33\PSSBIN"
        psspy_location = "C:\Program Files (x86)\PTI\PSSE33\PSSPY27"
    elif version == 34:
        psse_location = "C:\Program Files (x86)\PTI\PSSE34\PSSBIN"
        psspy_location = "C:\Program Files (x86)\PTI\PSSE34\PSSPY27"
    else:
        raise ValueError("Version %d is not supported by this script." % version)

    if 'psse' not in os.environ['PATH'].lower():
        sys.path.append(psse_location)
        os.environ['PATH'] += ';' + psse_location
    sys.path.append(psspy_location)

# =============================================================================================

if __name__ == '__main__':
    import psse34
    import csv
    import sys #used for passing in the argument

    if int(matplotlib.__version__.split('.')[0]) >= 2:
        import matplotlib.style
        import matplotlib as mpl

        mpl.style.use('classic')

    file_name = "PlottingPDF.csv"
    OUTDIR = 'Results\\' #must end in \\ e.g. 'Results\\'
    if len(sys.argv)>1:
        if sys.argv[1] != None:
            file_name = sys.argv[1]
    if len(sys.argv)>2:
        if sys.argv[2] != None:
            OUTDIR = sys.argv[2]+"\\"

    SOURCEOUTS = [] #if empty list, all .out files in OUTDIR will be used
    outfiles=[]
    if SOURCEOUTS == []:
        for file in os.listdir(OUTDIR):
            if file.endswith(".out"):
                SOURCEOUTS.append(file)
    outfiles = [OUTDIR + outfile for outfile in SOURCEOUTS]
    # Specify here correct PSSBIN folder path for PSSE version you would be working with
    pssbin33 = r"C:\Program Files\PTI\PSSE33\PSSBIN"
    pssbin32 = r"C:\Program Files (x86)\PTI\PSSE32\PSSBIN"
    pssbin34 = r"C:\Program Files (x86)\PTI\PSSE34\PSSBIN"

    # Select here which PSSE version to run
    if sys.version_info[0] == 2 and sys.version_info[1] == 5:
        select_psse_environment(version=32)
    elif sys.version_info[0] == 2 and sys.version_info[1] == 7:
        select_psse_environment(version=34)
    else:
        raise ValueError("Python version is not supported by this script.")
    import dyntools #dyntools must be imported AFTER select_psse_environment() call

    f=open(file_name, 'r')
    reader = csv.reader(f)
    data = list(list(rec) for rec in csv.reader(f, delimiter=',')) #reads csv into a list of lists
    f.close() #close the csv

    PDF2Layout={}
    Primary2Ylabel={}
    PDF2OUT2Primary={}
    """PDF2OUT2Primary is a dictionary of dictionaries, with levels of PDF File -> OUT File Identifier -> Primary Axis Content
    e.g. PDF2OUT2Primary={'out_file.pdf': {'': ['ANGL', 'POWR', 'VARS'], '_Fau': ['VOLT', 'POWR', 'VARS']}}"""
    PDF2OUT2Secondary={}

    PDFidx=None
    OUTidx=None
    PrimaryIdx=None
    SecondaryIdx=None
    LayoutIdx=None
    YlabelIdx=None
    YlabelSidx=None
    for i in range(len(data[0])):
        if "PDF File" in data[0][i]:
            PDFidx = i
        elif "OUT File" in data[0][i]:
            OUTidx = i
        elif "Primary Axis" in data[0][i]:
            PrimaryIdx = i
        elif "Secondary Axis" in data[0][i]:
            SecondaryIdx = i
        elif "Layout" in data[0][i]:
            LayoutIdx = i
        elif "Y-Axis Primary Label" in data[0][i]:
            YlabelIdx = i
        elif "Y-Axis Secondary Label" in data[0][i]:
            YlabelSidx = i
    for row in data[1:]:
        # set PDF's layout if not already done:
        if row[PDFidx] not in PDF2Layout.keys() and row[LayoutIdx]!="":
            PDF2Layout[row[PDFidx]]=row[LayoutIdx]
        # set Primary axis content's y-label if not already done:
        # try:
            # row[PrimaryIdx] not in Primary2Ylabel.keys() and row[YlabelIdx]!=""
        # except TypeError:
            # pdb.set_trace()
        if row[PrimaryIdx] not in Primary2Ylabel.keys() and row[YlabelIdx]!="":
            Primary2Ylabel[row[PrimaryIdx]]=row[YlabelIdx]
        if row[SecondaryIdx] not in Primary2Ylabel.keys() and row[SecondaryIdx]!="":
            Primary2Ylabel[row[SecondaryIdx]]=row[YlabelSidx]
        # Look at PDF -> OUT.
        if row[PDFidx] not in PDF2OUT2Primary.keys():
            # PDF not in PDF layer. Add PDF -> OUT -> [Primary] to dict.
            PDF2OUT2Primary[row[PDFidx]] = {row[OUTidx]:[row[PrimaryIdx]]}
            PDF2OUT2Secondary[row[PDFidx]] = {row[OUTidx]:[row[SecondaryIdx]]}
        else:
            # PDF is in PDF layer. Look at OUT -> Primary.
            if row[OUTidx] not in PDF2OUT2Primary[row[PDFidx]].keys():
                # OUT not in this PDF's OUT keys. Add OUT -> Primary to this PDF.
                PDF2OUT2Primary[row[PDFidx]][row[OUTidx]] = [row[PrimaryIdx]]
                PDF2OUT2Secondary[row[PDFidx]][row[OUTidx]] = [row[SecondaryIdx]]
            else:
                # OUT is in this PDF's OUT keys. Add Primary to existing list.
                PDF2OUT2Primary[row[PDFidx]][row[OUTidx]].append(row[PrimaryIdx])
                PDF2OUT2Secondary[row[PDFidx]][row[OUTidx]].append(row[SecondaryIdx])
    Count=0

    figNo = 1 #integer incremented within test_plots2pdf() with each new figure, and reset at end of PDF loop below
    for pdfName in PDF2OUT2Primary.keys():
        plotCount = 0	# Variable to count number of plots in the current .pdf file (BP added)
        maxPlots = 250	# max number of plots per .pdf

        layout = PDF2Layout[pdfName]
        pp = PdfPages(OUTDIR+pdfName) #pp is the PDF where all plots will go
        subplotNo = 1
        for outfile in outfiles:
            #set list of channels to plot for this PDF, and this OUT file
            primaryAxisChs = [] #will be list of channels to plot for this PDF, and this OUT file. TM: moved here 2016/07/29
            secondaryAxisChs = []
            found = False
            Length=int(len(outfiles))
            #output progress
            print '\r Plotting:{', Count, '/' ,Length,'}\r',
            Count = Count+1
            for outId in PDF2OUT2Primary[pdfName].keys():
                if outId in outfile:
                    #if filename identifier (outId) is empty string, only set primary axis channels if no match has been found yet:
                    if (outId=='' and not found) or outId!='':
                        primaryAxisChs=PDF2OUT2Primary[pdfName][outId]
                        secondaryAxisChs=PDF2OUT2Secondary[pdfName][outId]
                        found = True
            if not found:
                print("WARNING: No CSV entry found for OUT file "+outfile+" in PDF file "+pdfName)
            # create object
            if found:
                #BP added following to limit # of plots to maxPlots in a pdf 20171003*************************
                plotCount += 1
                if (plotCount) > maxPlots and ((plotCount-1)/maxPlots == 1) and ((plotCount-1)%maxPlots == 0):
                    pp.close() #close the PDF
                    dst_file = OUTDIR+pdfName
                    pdfNameNew = pdfName[:-4]+'_'+str((plotCount-1)/maxPlots)+'.pdf'
                    new_dst_file = OUTDIR+pdfNameNew
                    os.rename(dst_file, new_dst_file)

                    pdfNameNew = pdfName[:-4]+'_'+str(((plotCount-1)/maxPlots)+1)+'.pdf'
                    pp = PdfPages(OUTDIR+pdfNameNew) #pp is the PDF where all plots will go

                elif (plotCount) > maxPlots and ((plotCount-1)%maxPlots == 0):
                    pp.close() #close the PDF
                    #plt.close()
                    pdfNameNew = pdfName[:-4]+'_'+str(((plotCount-1)/maxPlots)+1)+'.pdf'
                    pp = PdfPages(OUTDIR+pdfNameNew) #pp is the PDF where all plots will go
                #*********************************************************************************************
                outlst = [outfile]
                chnf = dyntools.CHNF(outlst)
                outfilename = outlst[0].split('\\')[-1] #title for plots created with this .out file will be the filename
                # test_data_extraction(chnf)
                figNo,subplotNo = test_plots2pdf(chnf,figNo,subplotNo,outfilename,pp,primaryAxisChs,secondaryAxisChs,layout,Primary2Ylabel) #returns next figure and subplot number
        if subplotNo != 1:#== int(layout[0])*int(layout[1]):
            #plt.tight_layout(pad=5.0,w_pad=5.0, h_pad=7.0)
            plt.savefig(pp, format='pdf') #bbox_inches='tight' #TM: added 29/07/2016
            if SAVE_IMAGES:
                plt.savefig(IMAGEDIR+outfilename+'_'+str(figNo)+'.png')
                print IMAGEDIR+outfilename+'_'+str(figNo)+'.png saved'
        pp.close() #close the PDF
        print OUTDIR+pdfName,"saved"
        figNo = 1

# =============================================================================================
