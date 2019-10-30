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

------------------------------------------------------------------------------------------------

This tool is designed to aid model pre-endorsement by testing
a given model under varying conditions in a SMIB case.

"""

import glob, os, sys, math, csv, time, logging, traceback, exceptions
import shutil, psutil
from win32com import client
from subprocess import Popen
from multiprocessing import Process, current_process
from time import sleep
WorkingFolder = os.getcwd()

PSSE_LOCATION = r"C:\Program Files (x86)\PTI\PSSE34\PSSBIN"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['path'] + ';' + PSSE_LOCATION

PSSE_LOCATION = r"C:\Program Files (x86)\PTI\PSSE34\PSSPY27"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['path'] + ';' + PSSE_LOCATION

import psspy
import redirect
import multiprocessing
import subprocess
import glob
import shutil
import psse_env_manager as em
import time

_i = psspy.getdefaultint()
_f = psspy.getdefaultreal()
_s = psspy.getdefaultchar()

file_name = WorkingFolder + "\DATA.txt"             
f = open(file_name,'r')
lines = f.readlines()
f.close()

# ====================================== Setup ====================================== #
SAV_File = str(lines[[i for i, s in enumerate(lines) if 'SAV_File =' in s][0]].replace("SAV_File =", "").replace("\t", "").replace("\n", "").replace(" ", ""))
DYR_File_GNCLS = str(lines[[i for i, s in enumerate(lines) if 'DYR_File_GNCLS =' in s][0]].replace("DYR_File_GNCLS =", "").replace("\t", "").replace("\n", "").replace(" ", ""))
DYR_File_ZINGEN = str(lines[[i for i, s in enumerate(lines) if 'DYR_File =' in s][0]].replace("DYR_File =", "").replace("\t", "").replace("\n", "").replace(" ", ""))

CON_File = 'Conv.sav'
SNP_File = 'Snap.snp'
BAT_File = 'ExecuteBatFile.bat'

ModelAccept_delete_create_savs = int(lines[[i for i, s in enumerate(lines) if 'ModelAccept_delete_create_savs =' in s][0]].replace("ModelAccept_delete_create_savs =", "").replace("\t", "").replace("\n", "")) #flag to allow creation of new and deletion of old sav cases (will also delete results)

# There's an assumption that the base case originally entered is running at maximum output
Pmax_actual = float(lines[[i for i, s in enumerate(lines) if 'Pmax_actual =' in s][0]].replace("Pmax_actual =", "").replace("\t", "").replace("\n", ""))  # Actual PMAX for the UUT in MW
Qmax_pu = float(lines[[i for i, s in enumerate(lines) if 'Qmax_pu =' in s][0]].replace("Qmax_pu =", "").replace("\t", "").replace("\n", ""))        # Qpu on MBASE of UUT
SBASE = float(lines[[i for i, s in enumerate(lines) if 'SBASE =' in s][0]].replace("SBASE =", "").replace("\t", "").replace("\n", ""))                # System base = 100 MVA

POC_VCtrl_Tgt = float(lines[[i for i, s in enumerate(lines) if 'POC_VCtrl_Tgt =' in s][0]].replace("POC_VCtrl_Tgt =", "").replace("\t", "").replace("\n", ""))            # POC voltage control target


# Bus Numbers:
Bus_Search = True
count_bus = 1
bus_mch_all = []
while Bus_Search == True:
    try:
        search_string = "INV" + str(count_bus) + "_Bus ="
        bus_mch_all.append(int(lines[[i for i, s in enumerate(lines) if search_string in s][0]].replace(search_string, "").replace("\t", "").replace("\n", "")))
        count_bus += 1
    except:
        Bus_Search = False

mID_all = []        
for indx, bus_num in enumerate(bus_mch_all):
    try:
        search_string = "mID" + str(indx+1) + " ="
        mID_all.append(str(lines[[i for i, s in enumerate(lines) if search_string in s][0]].replace(search_string, "").replace("\t", "").replace("\n", "").replace(" ", "")))
    except:
        print "NOT ALL REQUIRED MACHINE IDs HAVE BEEN INCLUDED OR INCORRECT NAMING FORMAT USED IN DATA.TXT\n"
        print "\n", mID_all, "       ", bus_mch_all
        exit()

print 
bus_mch1 = int(lines[[i for i, s in enumerate(lines) if 'INV1_Bus =' in s][0]].replace("INV1_Bus =", "").replace("\t", "").replace("\n", ""))                 # unit under test (UUT) bus
INV1_Bus = int(lines[[i for i, s in enumerate(lines) if 'INV1_Bus =' in s][0]].replace("INV1_Bus =", "").replace("\t", "").replace("\n", ""))                 # unit under test (UUT) bus
mID = str(lines[[i for i, s in enumerate(lines) if 'mID1 =' in s][0]].replace("mID1 =", "").replace("\t", "").replace("\n", "").replace(" ", ""))           # machine ID
bus_PCC = int(lines[[i for i, s in enumerate(lines) if 'POC =' in s][0]].replace("POC =", "").replace("\t", "").replace("\n", ""))               # point of common connection bus
POC = int(lines[[i for i, s in enumerate(lines) if 'POC =' in s][0]].replace("POC =", "").replace("\t", "").replace("\n", ""))               # point of common connection bus
bus_flt = int(lines[[i for i, s in enumerate(lines) if 'bus_flt =' in s][0]].replace("bus_flt =", "").replace("\t", "").replace("\n", ""))                # fault bus (must be added to case between PCC and infinite buses)
bus_inf = int(lines[[i for i, s in enumerate(lines) if 'SMIB =' in s][0]].replace("SMIB =", "").replace("\t", "").replace("\n", ""))                # infinite bus
SMIB = int(lines[[i for i, s in enumerate(lines) if 'SMIB =' in s][0]].replace("SMIB =", "").replace("\t", "").replace("\n", ""))                # infinite bus
POC_frBus = int(lines[[i for i, s in enumerate(lines) if 'POC_frBus =' in s][0]].replace("POC_frBus =", "").replace("\t", "").replace("\n", ""))
POC_toBus = int(lines[[i for i, s in enumerate(lines) if 'POC_toBus =' in s][0]].replace("POC_toBus =", "").replace("\t", "").replace("\n", ""))
bus_IDTRF = int(lines[[i for i, s in enumerate(lines) if 'bus_IDTRF =' in s][0]].replace("bus_IDTRF =", "").replace("\t", "").replace("\n", ""))            # bus number of the new bus added to include a transformer during angle_step test

# Network conditions at point of common connection - get this from doing a ACCC study at the connection bus
#Rs_PCC = 0.071643
#Xs_PCC = 0.122000

# Defined X/R and SCR values for case set-up
# Weak grid conditions should be listed first
#XR_ratio = [2.8,3.0]
XR_line = str(lines[[i for i, s in enumerate(lines) if 'XR_ratio =' in s][0]].replace("XR_ratio =", "").replace("\t", "").replace("\n", "").replace(" ", ""))
XR_ratio = [float(XX) for XX in XR_line[XR_line.find("[")+1:XR_line.find("]")].split(",")]
SCR_line = str(lines[[i for i, s in enumerate(lines) if 'SCR =' in s][0]].replace("SCR =", "").replace("\t", "").replace("\n", "").replace(" ", ""))
SCR = [float(XX) for XX in SCR_line[SCR_line.find("[")+1:SCR_line.find("]")].split(",")]
#SCR = [6.0,5.0]                    # include this in MBASE of one UUT - 65 MVA (BP added)

Vflt_Tflt = [(0.01,0.220),(0.7,0.430)]
t_step_and_acc_fact = [(0.001,0.3),(0.001,1.0),(0.002,0.3)]
        
# Following not used by BP so far
sim_prnt_stp = 9
sim_rt = 20.0                  # make sure to update plotting.py too

# ALSO Need to update any function that has change_var()/change_con() in it!!
# For multiple models, use DOCU to find out where the VARS and CONS begin, then add the offset from the datasheet

#PPCmodelName = str(lines[[i for i, s in enumerate(lines) if 'PPCmodelName =' in s][0]].replace("PPCmodelName =", "").replace("\t", "").replace("\n", "").replace(" ", ""))
CON_PPCPref = int(lines[[i for i, s in enumerate(lines) if 'CON_PPCPref =' in s][0]].replace("CON_PPCPref =", "").replace("\t", "").replace("\n", ""))
CON_PPCVref = int(lines[[i for i, s in enumerate(lines) if 'CON_PPCVref =' in s][0]].replace("CON_PPCVref =", "").replace("\t", "").replace("\n", ""))
VAR_PPCVini = int(lines[[i for i, s in enumerate(lines) if 'VAR_PPCVini =' in s][0]].replace("VAR_PPCVini =", "").replace("\t", "").replace("\n", ""))	# (L+1)
VAR_PPCPini = int(lines[[i for i, s in enumerate(lines) if 'VAR_PPCPini =' in s][0]].replace("VAR_PPCPini =", "").replace("\t", "").replace("\n", ""))	# (L+2)



#CON_irrad = 1       # CON for irradiance  SMASC model
#CON_Pref = 97       # CON for real power reference - CLSF uses the PPC to control this, located at 79+18 = 97
##CON_Pref = 2
#VAR_Qref = 90 #L+89 in smasc modelL+3 195+3
#
#CON_PFref = 88      # CON for PF setpoint (79+9 = 88)
#CON_Vref=89                    # CON for PF setpoint (79+10 = 89)
#CON_PFext = 18      # CON for over or under excitation (PF direction) - not used in CLSF
#
#CON_QVARmod = 79    # CON for switching between reactive power control modes
#CON_VoltRef = 89    # CON for setting the voltage setpoint (77+10 = 87)

# Dynamic Solution Parameters --------------------------------------------------------
max_solns = 200            # network solution maximum number of iterations
sfactor = 0.3              # acceleration sfactor used in the network solution
con_tol = 0.0001           # convergence tolerance used in the network solution
dT = 0.001                 # simulation step time
frq_filter = 0.04          # filter time constant used in calculating bus frequancy deviations    
int_delta_thrsh = 0.06     # intermediate simulation mode time step threshold used in extended term simulations
islnd_delta_thrsh = 0.14   # large (island frequency) simulation mode time step threshold used in extended term simulations
islnd_sfactor = 1.0        # large (island frequency) simulation mode acceleration factor used in extended term simulations
islnd_con_tol = 0.0005     # large (island frequency) simulation mode convergence tolerance used in extended term simulations

# ZINGEN.xlsx data format ------------------------------------------------------------
rIstart = 3                    #first data row in ZINGEN1.xlsx
cIstart = 1                    #first data column in ZINGEN1.xlsx
# ==================================================================================== #
def setup_logging_to_file(filename):
            logging.basicConfig( filename=filename,filemode='w',level=logging.DEBUG,format= '%(asctime)s - %(levelname)s - %(message)s',)

def extract_function_name():
            tb = sys.exc_info()[-1]
            stk = traceback.extract_tb(tb, 1)
            fname = stk[0][3]
            return fname

def log_exception(e):
            logging.error(
            "Function {function_name} raised {exception_class} ({exception_docstring}): {exception_message}".format(
            function_name = extract_function_name(), #this is optional
            exception_class = e.__class__,
            exception_docstring = e.__doc__,
            exception_message = e.message))

def clean_fort_files(directory):
    for subdir, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.fort'):
                os.remove(os.path.join(subdir, file))
            # if file.startswith('fort.'):
            #     os.remove(os.path.join(subdir, file))     

def run_LoadFlow():
            psspy.fdns([1,0,0,1,1,1,0,0])
            psspy.fdns([1,0,0,1,1,1,0,0])
            psspy.fdns([1,0,0,1,1,0,0,0])
            psspy.fdns([1,0,0,1,1,0,0,0])
            psspy.fnsl([1,0,0,1,1,0,0,0])            #Full Newton-Raphson
            psspy.fnsl([1,0,0,1,1,0,0,0])
            psspy.fnsl([1,0,0,1,1,0,0,0])
        
            blownup = psspy.solved()
            if blownup == 0:
                return 0
            else:
                return 1

def dirCreateClean(path,fileTypes):
    def subCreateClean(subpath,fileTypes):
        try:
            os.mkdir(subpath)
        except OSError:
            pass
                
        #delete all files listed in fileTypes in the directory            
        os.chdir(subpath)
        for type in fileTypes:
            filelist = glob.glob(type)
            for f in filelist:
                os.remove(f)
    currentDir = os.getcwd()
    subCreateClean(path,fileTypes)
    subCreateClean(path+'\\Results',fileTypes)
    subCreateClean(path+'\\PSSEOut',fileTypes)
    os.chdir(currentDir)
        
def intialise_PSSE(path,Case,Run):
            redirect.psse2py()
            psspy.psseinit()     #initialise PSSE so psspy commands can be called
            psspy.throwPsseExceptions = True

            Al_Prom = path + '\\' + Case[:3] +'_'+ Run + '_AlandProm.dat'
            Progress = path + '\\' + Case[:3] +'_'+ Run + '_Prog.dat'
            Report = path + '\\' + Case[:3] +'_'+ Run + '_Rep.dat'
            
            #! Clean PSS/e report files in path------------------
            Outf=open(Al_Prom,'w+')
            Outf.close()
            Outf=open(Progress,'w+')
            Outf.close()
            Outf=open(Report,'w+')
            Outf.close()
            psspy.prompt_output(2,Al_Prom,[2,0])
            psspy.alert_output(2,Al_Prom,[2,0])
            psspy.progress_output(2,Progress,[2,0])
            psspy.report_output(2,Report,[2,0])

def set_ZINGEN1_DataSets(ws,cI,Case):
    global pathTestFiles
    global rIstart
    rI = rIstart
    x = -0.004

    DynStudyCase = ws.Cells(rI-2, cI).Value
    if DynStudyCase != None:
        psspy.case(pathTestFiles+'\\'+Case)

        psspy.bsys(11,0,[0.0,0.0],0,[],1,[bus_inf],0,[],0,[])
        busData = psspy.abusreal(11,2,["BASE","PU"])
        busPUVolt = busData[1][1][0]
        busBasekVolt = busData[1][0][0]/math.sqrt(3)

        ZINGEN1=open(WorkingFolder + '\\' + 'ZINGEN1.dat','w+')
        while ws.Cells(rI, cI).Value != None:
            x = ws.Cells(rI, cI).Value
            y = ws.Cells(rI, cI+1).Value
            #! Following is valid only if flat start data is included in up to 6th row in ZINGEN1.xlsx
            if rI <= 6:
                z = (ws.Cells(rI, cI+2).Value)*busBasekVolt*busPUVolt            # Setting up actual Line-Neutral ZINGEN terminal voltage
            else:
                z = (ws.Cells(rI, cI+2).Value)*busBasekVolt*busPUVolt            # Setting up actual Line-Neutral ZINGEN terminal voltage
            #!----------------------------------------------------------------------------------------
            ZINGEN1.write("%.3f\t%.4f\t%.4f\n" %(x,y,z))
            rI +=1
        ZINGEN1.write("%.3f\t%.4f\t%.4f\n" %((x+dT),y,z))
        ZINGEN1.close()
        
    return [x,DynStudyCase]        #x is the run time as of ZINGEN1.xlsx

def plotResults(TestDir,pathTestDir,pyPlot,csvPlot):
    global pathTestFiles
            
    PlotScript = "\""+WorkingFolder+'\\'+pyPlot+"\""
    PlotConfig = "\""+WorkingFolder+'\\'+csvPlot+"\""
    pathPlotResults = "\""+pathTestDir+ '\\' + "Results"+"\""
    os.system("C:/Python27/python "+PlotScript+" "+PlotConfig+" "+pathPlotResults)

    pathPlotResults = pathTestDir+'\\'+"Results"
    for file in os.listdir(pathPlotResults):
        if (not file.endswith(".pdf")):
            continue            
        dst_file = pathPlotResults+'\\'+file
        new_dst_file = pathPlotResults+'\\'+TestDir+'_'+file
        os.rename(dst_file, new_dst_file)
        dst_dir= pathTestFiles+'\\Results'
        src_file = new_dst_file
        shutil.copy(src_file,dst_dir)

def dordie(StudyIq,pathTestDir,Case,dT,sfactor,Vflt,Tflt,Zs_inf_mach_100MVA):
    global pathTestFiles 
    """ do or die: this thing launches a process and kills it if unresponsive """
    arguments = (pathTestDir,Case,dT,sfactor,Vflt,Tflt,Zs_inf_mach_100MVA)
    timeout   =  300*60
    
    p = multiprocessing.Process(target=StudyIq, args=arguments)    
    p.start()
    p.join(timeout)
    if p.is_alive():
        p.terminate()  
def active_pythons():
    
    n = 0
    for proc in psutil.process_iter():
    
        process = psutil.Process(proc.pid)
        pname   = process.name
    
        if pname == 'python.exe':
            
            n = n + 1                           
    return n
    
def active_processes(JOBS):
    
    n = 0    
    for job in JOBS:
        if job.is_alive():
            n = n+1
    return n      
            
def Fault_Study_Test():
    '''
    Numerical stability of the model is tested by applying faults with different fault 
    impedances and clearing times.
            Different acceleration factors (sfactor) and solution time steps are used.
            Assumes weak grid conditions are listed first in XR_ratio and SCR variables
            '''
    global pathTestFiles  
    print("\nFault study test started ...")
    TestDir  = "Fault_Study_Test"
    pyPlot = "PlottingPDF.py"
    csvPlot = "PlottingPDF_Gen_VPQ.csv"            
    Zs_inf_mach_100MVA = 0            # this can be updated to read from the model
    caseFilt1 = "scr_"+str(round(SCR[0],2))
    caseFilt2 = "xr_"+str(XR_ratio[0])               

    pathTestDir = pathTestFiles+'\\'+TestDir
    dirCreateClean(pathTestDir,["*.out","*.sav","*.DAT","*.pdf"]) # Note that .pdf files are deleted as well
    psspy.close_powerflow()

    def runFault(pathTestDir,Case,dTNew,sfactorNew,Vflt,Tflt,Zs_inf_mach_100MVA):
        Run = "Vflt"+str(Vflt)+"_Tflt"+str(Tflt)+"_sFac"+str(sfactorNew)+"_dT"+str(dTNew)
        intialise_PSSE((pathTestDir+"\\PSSEOut"),Case,Run)
        psspy.case(pathTestFiles+'\\'+Case)
        #
        # Calculate fault admittance needed to create specified voltage dip
        ierr, Zsys1 = psspy.brndt2(bus_flt,bus_inf,'1','RX')
        Zsys1 = Zsys1 + Zs_inf_mach_100MVA
        fault_B = ((1-Vflt)/(Vflt*Zsys1)).imag * SBASE 
        fault_G = -fault_B / 10.0 # locked XR ratio of fault to 10 (GB ratio = -10). Phase angle jump if XR_fault <> XR_system.
        # fault_G = ((1-Vflt)/(Vflt*Zsys1)).real
        
        print("\nDynamic study case - "+"\""+Run+"\""+" executing on "+Case+" ...")
        print "G is %s, B is %s" %(fault_G,fault_B)
        
        dyn_setup(DYR_File_GNCLS)
        #! Reset dynamic simulation parameters
        psspy.dynamics_solution_param_2(realar1=sfactorNew, realar3=dTNew)
        psspy.snap(sfile=pathTestFiles+'\\'+SNP_File)                                            
        
        OUTPUT_name = pathTestDir + '\\' + "Results" + '\\' + Case +"_"+ Run + '.out'
        # Initalise
        psspy.case(pathTestFiles+'\\'+CON_File)
        psspy.rstr(pathTestFiles+'\\'+SNP_File)
        psspy.strt(0,OUTPUT_name)
        psspy.strt(0,OUTPUT_name)
        psspy.strt(0,OUTPUT_name)
        # Run dymanic simulations
        psspy.run(0,1.0,5,5,5)
        psspy.dist_bus_fault(bus_flt,1, 0.0,[fault_G, fault_B])
        psspy.change_channel_out_file(OUTPUT_name)
        psspy.run(0,(1.0+Tflt),5,5,5)
        psspy.dist_clear_fault(1)
        psspy.change_channel_out_file(OUTPUT_name)
        psspy.run(0,10,5,5,5)
        #psspy.pssehalt_2()
        
    for file in os.listdir(pathTestFiles):
        if (not file.endswith(".sav") or file.endswith(CON_File)):
            continue
            
        Case = file
        
        for Vflt,Tflt in Vflt_Tflt:
            if (Case.find(caseFilt1) != -1) and (Case.find(caseFilt2) != -1):
                for dTNew,sfactorNew in t_step_and_acc_fact:
                    runFault(pathTestDir,Case,dTNew,sfactorNew,Vflt,Tflt,Zs_inf_mach_100MVA)
            else:
                runFault(pathTestDir,Case,dT,sfactor,Vflt,Tflt,Zs_inf_mach_100MVA)                  

    plotResults(TestDir,pathTestDir,pyPlot,csvPlot)
   
def FRT_Iq_Response_Test():
    '''
    Reactive current response during FRT is tested.  
            '''
    global pathTestFiles  
    print("\nIq response test started ...")
    TestDir  = "FRT_Iq_Response_Test"
    pyPlot = "PlottingPDF.py"
    csvPlot = "PlottingPDF_Gen_VPQ.csv"            
    Zs_inf_mach_100MVA = 0            # this can be updated to read from the model

    Vflt_Tflt_Local = [(0.91,0.220),(0.875,0.220),(0.85,0.220),(0.825,0.220),(0.81,0.220),(0.775,0.220),(0.75,0.220),(0.725,0.220),(0.71,0.220),(0.675,0.220),(0.65,0.220),(0.625,0.220),(0.61,0.220),(0.575,0.220),(0.55,0.220),(0.525,0.220),(0.51,0.220),(0.475,0.220),(0.45,0.220),(0.425,0.220),(0.41,0.220),(0.375,0.220),(0.35,0.220),(0.325,0.220),(0.31,0.220),(0.275,0.220),(0.25,0.220),(0.225,0.220),(0.21,0.220),(0.175,0.220),(0.15,0.220),(0.125,0.220),(0.11,0.220)]

    pathTestDir = pathTestFiles+'\\'+TestDir
    dirCreateClean(pathTestDir,["*.out","*.sav","*.DAT","*.pdf"]) # Note that .pdf files are deleted as well
    psspy.close_powerflow()
    for file in os.listdir(pathTestFiles):
        if (not file.endswith(".sav") or file.endswith(CON_File)):
            continue        
        Case = file
        jobs = []   
        for Vflt,Tflt in Vflt_Tflt_Local:
            arguments = (StudyIq,pathTestDir,Case,dT,sfactor,Vflt,Tflt,Zs_inf_mach_100MVA)
            # runFault(pathTestDir,Case,dT,sfactor,Vflt,Tflt,Zs_inf_mach_100MVA)  
            p = multiprocessing.Process(target=dordie, args =arguments)
            p.start()
            sleep(1)
            jobs.append(p)                      # list of jobs
            while active_processes(jobs)>1:     # prevent from flooding
                # time.sleep(5)
                jobs[len(jobs)-1].join(5)       # wait 5s so that we do not penalize code with while loop                         
    plotResults(TestDir,pathTestDir,pyPlot,csvPlot)  
    
def StudyIq(pathTestDir,Case,dTNew,sfactorNew,Vflt,Tflt,Zs_inf_mach_100MVA):
    global pathTestFiles  
    pathTestFiles  = WorkingFolder+'\\'+"Test Files"    
    DYR_File_GNCLS = 'C:\\UserData\\z003sarm\\Documents\\Emma Wang\\Liying_Wang_PTI\\03_Projects\\02_GidginbungSolarFarm\\MAT\\SMA146_GNCLS.dyr'   
    Run = "Vflt"+str(Vflt)+"_Tflt"+str(Tflt)
    intialise_PSSE((pathTestDir+"\\PSSEOut"),Case,Run)
    psspy.case(pathTestFiles+'\\'+Case)
    #
    # Calculate fault admittance needed to create specified voltage dip
    ierr, Zsys1 = psspy.brndt2(bus_flt,bus_inf,'1','RX')
    Zsys1 = Zsys1 + Zs_inf_mach_100MVA
    fault_B = ((1-Vflt)/(Vflt*Zsys1)).imag * 100 # Assume SBASE = 100
    fault_G = -fault_B / 10.0 # locked XR ratio of fault to 10 (GB ratio = -10). Phase angle jump if XR_fault <> XR_system.
    # fault_G = ((1-Vflt)/(Vflt*Zsys1)).real
    
    print("\nDynamic study case - "+"\""+Run+"\""+" executing on "+Case+" ...")
    print "G is %s, B is %s" %(fault_G,fault_B)
    
    dyn_setup(DYR_File_GNCLS)
    #! Reset dynamic simulation parameters
    psspy.dynamics_solution_param_2(realar1=sfactorNew, realar3=dTNew)
    psspy.snap(sfile=pathTestFiles+'\\'+SNP_File)                                            
    
    OUTPUT_name = pathTestDir + '\\' + "Results" + '\\' + Case +"_"+ Run + '.out'
            # Initalise
    psspy.case(pathTestFiles+'\\'+CON_File)
    psspy.rstr(pathTestFiles+'\\'+SNP_File)
    psspy.strt(0,OUTPUT_name)
    psspy.strt(0,OUTPUT_name)
    psspy.strt(0,OUTPUT_name)
    # Run dymanic simulations
    psspy.run(0,1.0,5,5,5)
    psspy.dist_bus_fault(bus_flt,1, 0.0,[fault_G, fault_B])
    psspy.change_channel_out_file(OUTPUT_name)
    psspy.run(0,(1.0+Tflt),5,5,5)
    psspy.dist_clear_fault(1)
    psspy.change_channel_out_file(OUTPUT_name)
    psspy.run(0,10,5,5,5)
    #psspy.pssehalt_2()
    
    
    
def Voltage_Angle_Step_Test():
    '''
            Create a new bus, IDTRF near bus_flt and adds a 1:1 transformer
    ANG1 of this transformer is changed during model testing            
            '''
    global pathTestFiles
    print("\nVoltage angle step test started ...")
    TestDir  = "Voltage_Angle_Step_Test"
    Run = "Voltage_Angle_Step_Test"
    pyPlot = "PlottingPDF.py"
    csvPlot = "PlottingPDF_Gen_VPQ.csv"            
    pathTestDir = pathTestFiles+'\\'+TestDir
    dirCreateClean(pathTestDir,["*.out","*.sav","*.DAT","*.pdf"]) # Note that .pdf files are deleted as well

    psspy.close_powerflow()
    for file in os.listdir(pathTestFiles):
        if (not file.endswith(".sav") or file.endswith(CON_File)):
            continue
            
        Case = file
        intialise_PSSE((pathTestDir+"\\PSSEOut"),Case,Run)
        psspy.case(pathTestFiles+'\\'+Case)
        psspy.ltap(bus_flt,bus_inf,r"""1""", 0.0001,bus_IDTRF,r"""IDTRF""", _f)
        psspy.purgbrn(bus_IDTRF,bus_flt,r"""1""")
        psspy.two_winding_data_3(bus_flt,bus_IDTRF,r"""1""",[1,bus_flt,1,0,0,0,33,0,bus_flt,0,1,0,1,1,1],[0.0, 0.0001, 100.0, 1.0,0.0,0.0, 1.0,0.0,0.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0,0.0,0.0, 1.1, 0.9, 1.1, 0.9,0.0,0.0,0.0],r"""IDTRF""")
        err = run_LoadFlow()  
    
        print("\nDynamic study case - "+"\""+Run+"\""+" executing on "+Case+" ...")
        dyn_setup(DYR_File_GNCLS)  
        OUTPUT_name = pathTestDir + '\\' + "Results" + '\\' + Case +"_"+ Run + '.out'
    
                # Initalise
        psspy.case(pathTestFiles+'\\'+CON_File)
        psspy.rstr(pathTestFiles+'\\'+SNP_File)
        psspy.strt(0,OUTPUT_name)
        psspy.strt(0,OUTPUT_name)
        psspy.strt(0,OUTPUT_name)
        # Run dymanic simulations
        psspy.run(0, 1.0,5,5,5)
        psspy.two_winding_data_3(bus_flt,bus_IDTRF,r"""1""",realari6 = 20)
        psspy.run(0, 6.0,5,5,5)
        psspy.two_winding_data_3(bus_flt,bus_IDTRF,r"""1""",realari6 = 0)
        psspy.run(0, 11.0,5,5,5)
        psspy.two_winding_data_3(bus_flt,bus_IDTRF,r"""1""",realari6 = -20)
        psspy.run(0, 16.0,5,5,5)
        psspy.two_winding_data_3(bus_flt,bus_IDTRF,r"""1""",realari6 = 0)
        psspy.run(0, 21.0,5,5,5)

    plotResults(TestDir,pathTestDir,pyPlot,csvPlot)
            
def POC_Pref_Step_Test():
    '''
    Pref of PPC of UUT changed by dPref.            
            '''
    global pathTestFiles
    global MBASE
    dPref = 0.2
    print("\nPOC Pref step test started ...")
    TestDir  = "POC_Pref_Step_Test"
    Run = "POC_Pref_Step_Test"
    pyPlot = "PlottingPDF.py"
    csvPlot = "PlottingPDF_Gen_VPQ.csv"            
    pathTestDir = pathTestFiles+'\\'+TestDir
    dirCreateClean(pathTestDir,["*.out","*.sav","*.DAT","*.pdf"]) # Note that .pdf files are deleted as well

    psspy.close_powerflow()
    for file in os.listdir(pathTestFiles):
        if (not file.endswith(".sav") or file.endswith(CON_File)):
            continue
            
        Case = file
        intialise_PSSE((pathTestDir+"\\PSSEOut"),Case,Run)
        psspy.case(pathTestFiles+'\\'+Case)

        # Read POC initial MW magnitude
        psspy.bsys(11,0,[0.0,0.0],0,[],2,[POC_frBus,POC_toBus],0,[],0,[]) # - Select unit connecting
        POC_Br_Data = psspy.aflowcplx(11, 1, 1, 2, 'PQ')
        POCPUMW = abs(POC_Br_Data[1][0][1].real)/MBASE
        
        print("\nDynamic study case - "+"\""+Run+"\""+" executing on "+Case+" ...")

        dyn_setup(DYR_File_GNCLS)  
        OUTPUT_name = pathTestDir + '\\' + "Results" + '\\' + Case +"_"+ Run + '.out'
    
                # Initalise
        psspy.case(pathTestFiles+'\\'+CON_File)
        psspy.rstr(pathTestFiles+'\\'+SNP_File)
        psspy.strt(0,OUTPUT_name)
        psspy.strt(0,OUTPUT_name)
        psspy.strt(0,OUTPUT_name)
        
        J_vals = []
        L_vals = []
        Pset_vals = []
        for indx, bus_num in enumerate(bus_mch_all):
            ierr, J1 = psspy.mdlind(bus_num, mID_all[indx], 'EXC', 'CON')
            ierr, L1 = psspy.mdlind(bus_num, mID_all[indx], 'EXC', 'VAR')
            ierr, Pset1 = psspy.dsrval('VAR', L1+VAR_PPCPini)
            J_vals.append(J1)
            L_vals.append(L1)
            Pset_vals.append(Pset1)
                    
        # Run dymanic simulations
        psspy.run(0, 1.0,5,5,5)
        for indx, J1 in enumerate(J_vals):
            psspy.change_con(J1+CON_PPCPref, Pset_vals[indx]*(1-dPref))
        psspy.run(0, 30.0,5,5,5)
        for indx, J1 in enumerate(J_vals):
            psspy.change_con(J1+CON_PPCPref, Pset_vals[indx])
        psspy.run(0, 60.0,5,5,5)
        
    plotResults(TestDir,pathTestDir,pyPlot,csvPlot)
    
def Long_Run_Test():
    '''
    Flat run for tRun period of time.            
            '''
    global pathTestFiles
    tRun = 600
            
    print("\nLong run test started ...")
    TestDir  = "Long_Run_Test"
    Run = "Long_Run_Test"
    pyPlot = "PlottingPDF.py"
    csvPlot = "PlottingPDF_Gen_VPQ.csv"            
    pathTestDir = pathTestFiles+'\\'+TestDir
    dirCreateClean(pathTestDir,["*.out","*.sav","*.DAT","*.pdf"]) # Note that .pdf files are deleted as well

    psspy.close_powerflow()
    for file in os.listdir(pathTestFiles):
        if (not file.endswith(".sav") or file.endswith(CON_File)):
            continue
            
        Case = file
        intialise_PSSE((pathTestDir+"\\PSSEOut"),Case,Run)
        psspy.case(pathTestFiles+'\\'+Case)
  
        print("\nDynamic study case - "+"\""+Run+"\""+" executing on "+Case+" ...")

        dyn_setup(DYR_File_GNCLS)  
        OUTPUT_name = pathTestDir + '\\' + "Results" + '\\' + Case +"_"+ Run + '.out'
    
                # Initalise
        psspy.case(pathTestFiles+'\\'+CON_File)
        psspy.rstr(pathTestFiles+'\\'+SNP_File)
        psspy.strt(0,OUTPUT_name)
        psspy.strt(0,OUTPUT_name)
        psspy.strt(0,OUTPUT_name)
        # Run dymanic simulations
        psspy.run(0, tRun,100,100,100)

    plotResults(TestDir,pathTestDir,pyPlot,csvPlot)

def POC_Vref_Step_Test():
    '''
    Vref of PPC of UUT changed by dVref.            
            '''
    global pathTestFiles
    dVref = 0.02
    print("\nPOC Vref step test started ...")
    TestDir  = "POC_Vref_Step_Test"
    Run = "POC_Vref_Step_Test"
    pyPlot = "PlottingPDF.py"
    csvPlot = "PlottingPDF_Gen_VPQ.csv"            
    pathTestDir = pathTestFiles+'\\'+TestDir
    dirCreateClean(pathTestDir,["*.out","*.sav","*.DAT","*.pdf"]) # Note that .pdf files are deleted as well

    psspy.close_powerflow()
    for file in os.listdir(pathTestFiles):
        if (not file.endswith(".sav") or file.endswith(CON_File)):
            continue
            
        Case = file
        intialise_PSSE((pathTestDir+"\\PSSEOut"),Case,Run)
        psspy.case(pathTestFiles+'\\'+Case)

        # Read POC initial voltage
        psspy.bsys(11,0,[0.0,0.0],0,[],1,[bus_PCC],0,[],0,[])
        busData = psspy.abusreal(11,2,["BASE","PU"])
        POCbusPUVolt = busData[1][1][0]        
    
        print("\nDynamic study case - "+"\""+Run+"\""+" executing on "+Case+" ...")

        dyn_setup(DYR_File_GNCLS)  
        OUTPUT_name = pathTestDir + '\\' + "Results" + '\\' + Case +"_"+ Run + '.out'
    
                # Initalise
        psspy.case(pathTestFiles+'\\'+CON_File)
        psspy.rstr(pathTestFiles+'\\'+SNP_File)
        psspy.strt(0,OUTPUT_name)
        psspy.strt(0,OUTPUT_name)
        psspy.strt(0,OUTPUT_name)

        J_vals = []
        L_vals = []
        Vset_vals = []
        for indx, bus_num in enumerate(bus_mch_all):
            ierr, J1 = psspy.mdlind(bus_num, str(mID_all[indx]), 'EXC', 'CON')
            ierr, L1 = psspy.mdlind(bus_num, str(mID_all[indx]), 'EXC', 'VAR')
            ierr, vset1 = psspy.dsrval('VAR', L1+VAR_PPCVini)
            J_vals.append(J1)
            L_vals.append(L1)
            Vset_vals.append(vset1)
        
        # Run dymanic simulations
        psspy.run(0, 1.0,5,5,5)
        for indx, J1 in enumerate(J_vals):
            psspy.change_con(J1+CON_PPCVref, float(Vset_vals[indx])+dVref)
            print "\n", Vset_vals[indx]+dVref
        psspy.run(0, 10.0,5,5,5)
        for indx, J1 in enumerate(J_vals):
            psspy.change_con(J1+CON_PPCVref, float(Vset_vals[indx]))
            print "\n", Vset_vals[indx]
        psspy.run(0, 20.0,5,5,5)
        for indx, J1 in enumerate(J_vals):
            psspy.change_con(J1+CON_PPCVref, float(Vset_vals[indx])-dVref)
            print "\n", Vset_vals[indx]-dVref
        psspy.run(0, 30.0,5,5,5)
        for indx, J1 in enumerate(J_vals):
            psspy.change_con(J1+CON_PPCVref, float(Vset_vals[indx]))
            print "\n", Vset_vals[indx]
        psspy.run(0, 40.0,5,5,5)

    plotResults(TestDir,pathTestDir,pyPlot,csvPlot)

def Voltage_Step_Test():
    '''
    Network voltage at SMIB changed as of "ZINGEN1_Voltage_Step_Test.xlsx".
    ZINGEN used            
            '''
    global pathTestFiles
    global cIstart
    print("\nNetwork voltage step test started ...")
    TestDir  = "Voltage_Step_Test"
    ZINGEN1_Exdata = "ZINGEN1_Voltage_Step_Test.xlsx"
    pyPlot = "PlottingPDF.py"
    csvPlot = "PlottingPDF_Gen_VPQ.csv"            

    pathTestDir = pathTestFiles+'\\'+TestDir
    dirCreateClean(pathTestDir,["*.out","*.sav","*.DAT","*.pdf"]) # Note that .pdf files are deleted as well

    #! Open Excel ZINGEN1 data file ----------------------------------
    excel = client.Dispatch('Excel.Application')
    try:
        excel.Visible = False
    except:
        pass
    wb = excel.Workbooks.Open(WorkingFolder + '\\' + ZINGEN1_Exdata)
    ws = wb.Worksheets("ZINGEN1")
            
    psspy.close_powerflow()
    for file in os.listdir(pathTestFiles):
        if (not file.endswith(".sav") or file.endswith(CON_File)):
            continue
            
        Case = file
        intialise_PSSE((pathTestFiles+"\\PSSEOut"),Case,"Setup")

        cI = cIstart        #first data column in ZINGEN1.xlsx
        STOP = 0                    #flag to identify the last data column in ZINGEN1.xlsx
        while STOP == 0:
            #! Creating a batch file to execute        
            outReturn = set_ZINGEN1_DataSets(ws,cI,Case)
            runTime = outReturn[0]
            DynStudyCase = outReturn[1]
            Run = DynStudyCase
            if DynStudyCase != None:
                ExecuteBatFile = open(WorkingFolder + '\\' + BAT_File,'w+')                    
                ExecuteBatFile.write("C:/Python27/python %s %s %.3f %s %s\n" % (("\"DynamicSim.py\""),("\""+Case+"\""),runTime,("\""+Run+"\""),("\""+pathTestDir+"\"")))
                cI += 3
                ExecuteBatFile.close()
                print("\nDynamic study case - "+"\""+Run+"\""+" executing on "+Case+" ...")                
                #! Execute the batch file created
                p = Popen(BAT_File, cwd=WorkingFolder)
                stdout, stderr = p.communicate()                    
            else:
                print("\n")
                STOP = 1        

    excel.Application.Quit()            
    plotResults(TestDir,pathTestDir,pyPlot,csvPlot)

def Under_Voltage_Trip_Test():
    '''
    Network voltage at SMIB changed as of "ZINGEN1_Under_Voltage_Trip_Test.xlsx".
    ZINGEN used            
            '''
    global pathTestFiles
    global cIstart
    print("\nUnder voltage trip test started ...")
    TestDir  = "Under_Voltage_Trip_Test"
    ZINGEN1_Exdata = "ZINGEN1_Under_Voltage_Trip_Test.xlsx"
    pyPlot = "PlottingPDF.py"
    csvPlot = "PlottingPDF_Gen_VPQ.csv"            

    pathTestDir = pathTestFiles+'\\'+TestDir
    dirCreateClean(pathTestDir,["*.out","*.sav","*.DAT","*.pdf"]) # Note that .pdf files are deleted as well

    #! Open Excel ZINGEN1 data file ----------------------------------
    excel = client.Dispatch('Excel.Application')
    try:
        excel.Visible = False
    except:
        pass
    wb = excel.Workbooks.Open(WorkingFolder + '\\' + ZINGEN1_Exdata)
    ws = wb.Worksheets("ZINGEN1")
            
    psspy.close_powerflow()
    for file in os.listdir(pathTestFiles):
        if (not file.endswith(".sav") or file.endswith(CON_File)):
            continue
            
        Case = file
        intialise_PSSE((pathTestFiles+"\\PSSEOut"),Case,"Setup")

        cI = cIstart        #first data column in ZINGEN1.xlsx
        STOP = 0                    #flag to identify the last data column in ZINGEN1.xlsx
        while STOP == 0:
            #! Creating a batch file to execute        
            outReturn = set_ZINGEN1_DataSets(ws,cI,Case)
            runTime = outReturn[0]
            DynStudyCase = outReturn[1]
            Run = DynStudyCase
            if DynStudyCase != None:
                ExecuteBatFile = open(WorkingFolder + '\\' + BAT_File,'w+')                    
                ExecuteBatFile.write("C:/Python27/python %s %s %.3f %s %s\n" % (("\"DynamicSim.py\""),("\""+Case+"\""),runTime,("\""+Run+"\""),("\""+pathTestDir+"\"")))
                cI += 3
                ExecuteBatFile.close()
                print("\nDynamic study case - "+"\""+Run+"\""+" executing on "+Case+" ...")                
                #! Execute the batch file created
                p = Popen(BAT_File, cwd=WorkingFolder)
                stdout, stderr = p.communicate()                    
            else:
                print("\n")
                STOP = 1        

    excel.Application.Quit()            
    plotResults(TestDir,pathTestDir,pyPlot,csvPlot)

def Over_Voltage_Trip_Test():
    '''
    Network voltage at SMIB changed as of "ZINGEN1_Over_Voltage_Trip_Test.xlsx".
    ZINGEN used            
            '''
    global pathTestFiles
    global cIstart
    print("\nOver voltage trip test started ...")
    TestDir  = "Over_Voltage_Trip_Test"
    ZINGEN1_Exdata = "ZINGEN1_Over_Voltage_Trip_Test.xlsx"
    pyPlot = "PlottingPDF.py"
    csvPlot = "PlottingPDF_Gen_VPQ.csv"            

    pathTestDir = pathTestFiles+'\\'+TestDir
    dirCreateClean(pathTestDir,["*.out","*.sav","*.DAT","*.pdf"]) # Note that .pdf files are deleted as well

    #! Open Excel ZINGEN1 data file ----------------------------------
    excel = client.Dispatch('Excel.Application')
    try:
        excel.Visible = False
    except:
        pass
    wb = excel.Workbooks.Open(WorkingFolder + '\\' + ZINGEN1_Exdata)
    ws = wb.Worksheets("ZINGEN1")
            
    psspy.close_powerflow()
    for file in os.listdir(pathTestFiles):
        if (not file.endswith(".sav") or file.endswith(CON_File)):
            continue
            
        Case = file
        intialise_PSSE((pathTestFiles+"\\PSSEOut"),Case,"Setup")

        cI = cIstart        #first data column in ZINGEN1.xlsx
        STOP = 0                    #flag to identify the last data column in ZINGEN1.xlsx
        while STOP == 0:
            #! Creating a batch file to execute        
            outReturn = set_ZINGEN1_DataSets(ws,cI,Case)
            runTime = outReturn[0]
            DynStudyCase = outReturn[1]
            Run = DynStudyCase
            if DynStudyCase != None:
                ExecuteBatFile = open(WorkingFolder + '\\' + BAT_File,'w+')                    
                ExecuteBatFile.write("C:/Python27/python %s %s %.3f %s %s\n" % (("\"DynamicSim.py\""),("\""+Case+"\""),runTime,("\""+Run+"\""),("\""+pathTestDir+"\"")))
                cI += 3
                ExecuteBatFile.close()
                print("\nDynamic study case - "+"\""+Run+"\""+" executing on "+Case+" ...")                
                #! Execute the batch file created
                p = Popen(BAT_File, cwd=WorkingFolder)
                stdout, stderr = p.communicate()                    
            else:
                print("\n")
                STOP = 1        

    excel.Application.Quit()            
    plotResults(TestDir,pathTestDir,pyPlot,csvPlot)

def Under_Frequency_Trip_Test():
    '''
    System frequency is changed as of "ZINGEN1_Under_Frequency_Trip_Test.xlsx".
    ZINGEN used            
            '''
    global pathTestFiles
    global cIstart
    print("\nUnder Frequency trip test started ...")
    TestDir  = "Under_Frequency_Trip_Test"
    ZINGEN1_Exdata = "ZINGEN1_Under_Frequency_Trip_Test.xlsx"
    pyPlot = "PlottingPDF.py"
    csvPlot = "PlottingPDF_Gen_VPQdFreq.csv"            

    pathTestDir = pathTestFiles+'\\'+TestDir
    dirCreateClean(pathTestDir,["*.out","*.sav","*.DAT","*.pdf"]) # Note that .pdf files are deleted as well

    #! Open Excel ZINGEN1 data file ----------------------------------
    excel = client.Dispatch('Excel.Application')
    try:
        excel.Visible = False
    except:
        pass
    wb = excel.Workbooks.Open(WorkingFolder + '\\' + ZINGEN1_Exdata)
    ws = wb.Worksheets("ZINGEN1")
            
    psspy.close_powerflow()
    for file in os.listdir(pathTestFiles):
        if ((not file.endswith(".sav")) or (not file.endswith("Qzero.sav")) or file.endswith(CON_File)):
            continue
            
        Case = file
        intialise_PSSE((pathTestFiles+"\\PSSEOut"),Case,"Setup")

        cI = cIstart        #first data column in ZINGEN1.xlsx
        STOP = 0                    #flag to identify the last data column in ZINGEN1.xlsx
        while STOP == 0:
            #! Creating a batch file to execute        
            outReturn = set_ZINGEN1_DataSets(ws,cI,Case)
            runTime = outReturn[0]
            DynStudyCase = outReturn[1]
            Run = DynStudyCase
            if DynStudyCase != None:
                ExecuteBatFile = open(WorkingFolder + '\\' + BAT_File,'w+')                    
                ExecuteBatFile.write("C:/Python27/python %s %s %.3f %s %s\n" % (("\"DynamicSim.py\""),("\""+Case+"\""),runTime,("\""+Run+"\""),("\""+pathTestDir+"\"")))
                cI += 3
                ExecuteBatFile.close()
                print("\nDynamic study case - "+"\""+Run+"\""+" executing on "+Case+" ...")                
                #! Execute the batch file created
                p = Popen(BAT_File, cwd=WorkingFolder)
                stdout, stderr = p.communicate()                    
            else:
                print("\n")
                STOP = 1        

    excel.Application.Quit()            
    plotResults(TestDir,pathTestDir,pyPlot,csvPlot)

def Over_Frequency_Trip_Test():
    '''
    System frequency is changed as of "ZINGEN1_Over_Frequency_Trip_Test.xlsx".
    ZINGEN used            
            '''
    global pathTestFiles
    global cIstart
    print("\nOver Frequency trip test started ...")
    TestDir  = "Over_Frequency_Trip_Test"
    ZINGEN1_Exdata = "ZINGEN1_Over_Frequency_Trip_Test.xlsx"
    pyPlot = "PlottingPDF.py"
    csvPlot = "PlottingPDF_Gen_VPQdFreq.csv"            

    pathTestDir = pathTestFiles+'\\'+TestDir
    dirCreateClean(pathTestDir,["*.out","*.sav","*.DAT","*.pdf"]) # Note that .pdf files are deleted as well

    #! Open Excel ZINGEN1 data file ----------------------------------
    excel = client.Dispatch('Excel.Application')
    try:
        excel.Visible = False
    except:
        pass
    wb = excel.Workbooks.Open(WorkingFolder + '\\' + ZINGEN1_Exdata)
    ws = wb.Worksheets("ZINGEN1")
            
    psspy.close_powerflow()
    for file in os.listdir(pathTestFiles):
        if ((not file.endswith(".sav")) or (not file.endswith("Qzero.sav")) or file.endswith(CON_File)):
            continue
            
        Case = file
        intialise_PSSE((pathTestFiles+"\\PSSEOut"),Case,"Setup")

        cI = cIstart        #first data column in ZINGEN1.xlsx
        STOP = 0                    #flag to identify the last data column in ZINGEN1.xlsx
        while STOP == 0:
            #! Creating a batch file to execute        
            outReturn = set_ZINGEN1_DataSets(ws,cI,Case)
            runTime = outReturn[0]
            DynStudyCase = outReturn[1]
            Run = DynStudyCase
            if DynStudyCase != None:
                ExecuteBatFile = open(WorkingFolder + '\\' + BAT_File,'w+')                    
                ExecuteBatFile.write("C:/Python27/python %s %s %.3f %s %s\n" % (("\"DynamicSim.py\""),("\""+Case+"\""),runTime,("\""+Run+"\""),("\""+pathTestDir+"\"")))
                cI += 3
                ExecuteBatFile.close()
                print("\nDynamic study case - "+"\""+Run+"\""+" executing on "+Case+" ...")                
                #! Execute the batch file created
                p = Popen(BAT_File, cwd=WorkingFolder)
                stdout, stderr = p.communicate()                    
            else:
                print("\n")
                STOP = 1        

    excel.Application.Quit()            
    plotResults(TestDir,pathTestDir,pyPlot,csvPlot) 
    
def create_saved_cases(path,num,Xs,Rs,Xsys,Rsys,Pgen,case_id):
    '''
    This function has to be updated as sutable to set P and Q for any type of gen model
            Currently, this sets a value to WPF to gain a desired Q output
            '''
    global MBASE
    global WMOD
    MisMatch_Tol = 0.001
    Qgenmax = min(abs(MBASE*Qmax_pu),abs(math.sqrt(MBASE**2-Pmax_actual**2)))

    def find_Vsched():
        ierr, Vinf = psspy.busdat(bus_inf,'PU')
        ierr, Vpcc = psspy.busdat(bus_PCC,'PU')
        dV_PCC = POC_VCtrl_Tgt - Vpcc
        Vsch = Vpcc
        while abs(dV_PCC) > MisMatch_Tol:
            Vsch = Vinf + dV_PCC/2
            psspy.plant_data(bus_inf,realar1=Vsch)      # set the infinite bus scheduled voltage to the estimated voltage for this condtion
            err = run_LoadFlow()
            ierr, Vinf = psspy.busdat(bus_inf,'PU')
            ierr, Vpcc = psspy.busdat(bus_PCC,'PU')
            dV_PCC = POC_VCtrl_Tgt - Vpcc

        ierr, Vmch = psspy.busdat(bus_mch1,'PU')
        ierr, Vpcc = psspy.busdat(bus_PCC,'PU')
        print ">>> UUT V %1.3f pu, PCC V %1.3f pu" % (round(Vmch,3),round(Vpcc,3))       
        return Vsch
    
    for indx, bus_num in enumerate(bus_mch_all):
        psspy.machine_data_2(bus_num,mID_all[indx],realar1=Pgen)  # Update the UUT active power output
    #psspy.machine_data_2(bus_mch,'1',realar1=Pgen,realar8=Rs,realar9=Xs)       # This is needed if the dynamics source impedance is different to load flow
    
    #psspy.machine_data_2(bus_inf,'1',realar8=Rsys,realar9=Xsys)                 # Update the infinite bus equivalent impedance
    #psspy.seq_machine_data(bus_inf,r"""1""",[ Rsys, Xsys, Rsys, Xsys,_f,_f])    # Update the infinite bus equivalent impedance in sequence info
    
    psspy.machine_data_2(bus_inf,'1',realar8=0.0,realar9=0.0001)                 # Update the infinite bus equivalent impedance
    psspy.seq_machine_data(bus_inf,'1',[ 0.0, 0.0001, 0.0, 0.0001,_f,_f])    # Update the infinite bus equivalent impedance in sequence info
    
    psspy.branch_data(bus_PCC,bus_flt,'1', realar1=Rsys*.1,realar2=Xsys*.1)
    psspy.branch_data(bus_flt,bus_inf,'1', realar1=Rsys*.9,realar2=Xsys*.9)
    #psspy.branch_data(bus_flt,bus_inf,'1', realar1=Rsys,realar2=Xsys)
    
    err = run_LoadFlow()

    # Qzero
    Q_target = 0.
    WPF = 1
    if WMOD == 0:
        for indx, bus_num in enumerate(bus_mch_all):
            psspy.machine_data_2(bus_num,mID_all[indx],realar2=Q_target,realar3=Q_target,realar4=Q_target)
    else:
        for indx, bus_num in enumerate(bus_mch_all):
            psspy.machine_data_2(bus_num,mID_all[indx],realar17=WPF)
    err = run_LoadFlow()            
    Vsched_for_Qzero = find_Vsched()
    print "Vsched_for_Qzero is %s" %Vsched_for_Qzero

    if Pgen == 0:
        num = num + 1
        psspy.save(path+'\\'+("%03d"%num)+str(case_id)+'_'+'Qzero'+'.sav')
        Vsched_for_Qlag = Vsched_for_Qzero
        num = num + 1
        psspy.save(path+'\\'+("%03d"%num)+str(case_id)+'_'+'Qlead'+'.sav')
        Vsched_for_Qlead = Vsched_for_Qzero
        num = num + 1
        psspy.save(path+'\\'+("%03d"%num)+str(case_id)+'_'+'Qlag'+'.sav')
    else:
        # Qzero
        num = num + 1
        psspy.save(path+'\\'+("%03d"%num)+str(case_id)+'_'+'Qzero'+'.sav')
        
        # Qlead = UUT exporting reactive power to the grid
        Q_target = Pmax_actual*0.1      # set the UUT target Q to 0.1*Pmax_actual (export)
        Q_target = min(abs(Q_target),Qgenmax)*Q_target/abs(Q_target)
        WPF = abs(Pgen/math.sqrt(Pgen**2+Q_target**2))
        if WMOD == 0:
            for indx, bus_num in enumerate(bus_mch_all):
                psspy.machine_data_2(bus_num,mID_all[indx],realar2=Q_target,realar3=Q_target,realar4=Q_target)
        else:
            for indx, bus_num in enumerate(bus_mch_all):
                psspy.machine_data_2(bus_num,mID_all[indx],realar17=WPF)
        err = run_LoadFlow()
        Vsched_for_Qlead = find_Vsched()
        num = num + 1
        psspy.save(path+'\\'+("%03d"%num)+str(case_id)+'_'+'Qlead'+'.sav')
        print "Vsched_for_Qlead is %s, target Q is %s" %(Vsched_for_Qlead,round(Q_target*2,3))
        
        # Qlag = UUT importing reactive power from the grid
        Q_target = -Pmax_actual*0.1      # set the UUT target Q to -0.1*Pmax_actual (import)
        Q_target = min(abs(Q_target),Qgenmax)*Q_target/abs(Q_target)
        WPF = -abs(Pgen/math.sqrt(Pgen**2+Q_target**2))
        if WMOD == 0:
            for indx, bus_num in enumerate(bus_mch_all):
                psspy.machine_data_2(bus_num,mID_all[indx],realar2=Q_target,realar3=Q_target,realar4=Q_target)
        else:
            for indx, bus_num in enumerate(bus_mch_all):
                psspy.machine_data_2(bus_num,mID_all[indx],realar17=WPF)
        err = run_LoadFlow()
        Vsched_for_Qlag = find_Vsched()
        num = num + 1
        psspy.save(path+'\\'+("%03d"%num)+str(case_id)+'_'+'Qlag'+'.sav')
        print "Vsched_for_Qlag is %s, target Q is %s" %(Vsched_for_Qlag,round(Q_target*2,3))

    print "-----------------------------------"
            
    Vsched = {'Qzero':Vsched_for_Qzero, 'Qlag':Vsched_for_Qlag, 'Qlead':Vsched_for_Qlead}
    return Vsched,num 
    
def dyn_setup(DYR_File):
            global pathTestFiles
            global CON_File
            global SNP_File

            #! Convert the network ---------------------------------------        
            psspy.cong(0)
            psspy.conl(0,1,1,[0,0],[ 100.0,0.0,0.0, 100.0])
            psspy.conl(0,1,2,[0,0],[ 100.0,0.0,0.0, 100.0])
            psspy.conl(0,1,3,[0,0],[ 100.0,0.0,0.0, 100.0])
            
            psspy.ordr(0)    #! Order the matrix: ORDR
            psspy.fact()     #! Factorize the matrix: FACT
            psspy.tysl(0)    #! TYSL
            
            #! Linking libraries ----------------------------------------- 
            for file in os.listdir(WorkingFolder):
                if file.endswith(".dll"):
                    psspy.addmodellibrary(file)
    
            psspy.dyre_new([1,1,1,1],(WorkingFolder+'\\'+DYR_File),(pathTestFiles+'\\'+"Conec.flx"),(pathTestFiles+'\\'+"Conet.flx"),(pathTestFiles+'\\'+"Compile.bat"))            #! Read in the dynamic data file        
            psspy.save(pathTestFiles+'\\'+CON_File)  #! Save the converted case
            
            #! Setup Dynamic Simulation parameters     --------------Update for Every Project ------------------------
            psspy.dynamics_solution_param_2(intgar1=max_solns, realar1=sfactor, realar2 =con_tol, realar3=dT, realar4 =frq_filter,
                                            realar5=int_delta_thrsh, realar6=islnd_delta_thrsh, realar7=islnd_sfactor, realar8=islnd_con_tol)
            psspy.set_netfrq(1)
            psspy.set_relang(1,0,"")
            
            #! Set recording channels ------------------------------------
            #! UPDATE as required in each study --------------------------
            chani = 1
            for indx, bus_num in enumerate(bus_mch_all):
                psspy.machine_array_channel([chani,4,bus_num],mID_all[indx],r"""UUT_Voltage""") #1
                chani += 1
                psspy.machine_array_channel([chani,2,bus_num],mID_all[indx],r"""UUT_Pelec""") #2
                chani += 1
                psspy.machine_array_channel([chani,3,bus_num],mID_all[indx],r"""UUT_Qelec""") #3
                chani += 1
                psspy.machine_array_channel([chani,1,bus_num],mID_all[indx],r"""UUT_ANGL""") #4
                chani += 1
                psspy.bus_frequency_channel([chani,bus_num],r"""FREELEC""") #5
                chani += 1
            psspy.voltage_and_angle_channel([chani,-1,-1,POC],['POC_Voltage','POC_ANGL']) #6,7
            chani += 2
            psspy.branch_p_and_q_channel([chani,-1, -1, POC_frBus,POC_toBus], r"""1""", ['P_POC','Q_POC'])            #8,9
            chani += 2
            psspy.bus_frequency_channel([chani,POC],r"""FRE_POC""") #10
            chani += 1
            psspy.voltage_channel([chani,-1,-1,SMIB],r"""SMIB_Voltage""") #11
            chani += 1
            
            psspy.snap(sfile=pathTestFiles+'\\'+SNP_File)
    
def main():
    '''Main'''
    global PMAX         # Store PMAX of the UUT as of the SAV_File
    global MBASE        # for MVA base of UUT
    global WMOD                    # Default control mode set for UUT in LF model
    global pathTestFiles            # path to all test files (including .sav files)
            
    os.system('cls')            
    start = time.time()            #Read current system time

    pathTestFiles = WorkingFolder+'\\'+"Test Files"
    if ModelAccept_delete_create_savs == 1: 
        dirCreateClean(pathTestFiles,["*.out","*.sav","*.DAT","*.pdf"]) # Note that .pdf files are deleted as well
    else:
        base_files = [os.path.join(pathTestFiles, "PSSEOut"), os.path.join(pathTestFiles, "Results")]
        for f in base_files:
            if not os.path.isdir(f):
                os.makedirs(f)
#    
    intialise_PSSE((pathTestFiles+"\\PSSEOut"),SAV_File,"Setup")    
    ierr = psspy.case(SAV_File)
        
    # Create multiple saved cases
    XR_RATIO_MACHINE = 3 # Fixed
    #ierr, Zsource_inf_mach = psspy.macdt2(bus_inf, '1', 'ZSORCE')
    #ierr, MBASE_inf_mach = psspy.macdat(bus_inf, '1', 'MBASE')
    #Zs_inf_mach_100MVA = abs(Zsource_inf_mach)*100/MBASE_inf_mach
    Zs_inf_mach_100MVA = 0

    ierr, PMAX = psspy.macdat(bus_mch1,mID,'PMAX')
    ierr, MBASE = psspy.macdat(bus_mch1,mID,'MBASE')
    ierr, WMOD = psspy.macint(bus_mch1,mID,'WMOD')
    Pgen = [round(PMAX,3),round(0.5*PMAX,3), round(0.1*PMAX,3)]      # do tests at 100%, 50% and 10% of power output: round(PMAX,3),
    
    SCR_location = 'POC'    # The location where the SCR is specified
                            # 'POC' = at the point of connection
                            # 'Terminals' = at the machine terminals
                            
    Z_RETICULATION = 0.11   # Reticulation impedance (per unit on SBASE) between the POC 
                            # and machine terminals. If the plant has multiple machines 
                            # the net impedance should be entered including any parallel
                            # or series connections.
                            # (Only required if SCR is specified at the machine terminals)
                            # (i.e. if SCR_location - 'Terminals')
                            # (if SCR_location = 'POC' this value is not used by the script)
    
    id = 0  
    for scr in SCR:
        for xr in XR_ratio:
            # Rs = math.sqrt(((1.0/scr)**2)/((XR_RATIO_MACHINE)**2+1.0))
            # Xs = Rs * XR_RATIO_MACHINE
            
            Rs = 0      # UUT source impedance
            Xs = 10000          
            
            Rsys = 0 # Initialise Rsys
           
            if SCR_location == 'POC':
                # Direct conversion of SCR to impedance
                # The impedance is on MBASE
                Rsys = math.sqrt(((1.0/scr)**2)/(xr**2+1.0)) #make sure that the division is forced to be a floating point number
            elif SCR_location == 'Terminals':
                # Convert reticulation impedance from SBASE to MBASE for
                # translating the SCR from the terminals to the POC
                Z_RETICULATION_MBASE = Z_RETICULATION*(MBASE/SBASE)
                # The resulting Rsys and Xsys will also be on MBASE
                Rsys = math.sqrt(((1.0/scr-(Z_RETICULATION_MBASE))**2)/(xr**2+1.0))
            
            Xsys = Rsys * xr
            
#            # Convert impedances from MBASE to SBASE for entry in PSS/E
#            Rsys = Rsys*(SBASE/MBASE)
#            Xsys = Xsys*(SBASE/MBASE)
# #         
            if ModelAccept_delete_create_savs == 1: 
                print "Rsys_new: %s, Xsys_new: %s, Zsys_new: %.4f" %(round(Rsys,4),round(Xsys,4),math.sqrt(Rsys**2+Xsys**2))
                print "-----------------------------------"
                for P in Pgen:
                    if len(XR_ratio) > 1:
                        if xr == XR_ratio[1] and P == Pgen[1]:
                            continue # skip this particular case set-up (otherwise too many cases)
                    case_id = "_scr_"+str(round(scr,2))+"_xr_"+str(xr)+"_P_"+str(P)
                    Vsched,id = create_saved_cases(pathTestFiles,id,Xs,Rs,Xsys,Rsys,P,case_id)

#--------------------------------------------------------------------------------------------------------------------                            
#-------------EK added below functionality to allow error logging and skipping of error-ridden functions-------------
#--------------------------------------------------------------------------------------------------------------------            
#----------------------NOTE: This functionality uses EK added functions 'setup_logging_to_file',---------------------
#----------------------------------'extract_function_name' and 'log_exception'---------------------------------------
#--------------------------------------------------------------------------------------------------------------------
    file_name = WorkingFolder + "\RUN_STUDY_CHECKLIST.txt"             
            
    run_functions = []
    sc = open(file_name,'r')
    lines = sc.readlines()
    sc.close()

    for lindx, line in enumerate(lines):
        line = line.replace("\n", "")
        if "~" in line and "Y" in line[len(line)-1]:
            hold = line.split("~")
            run_functions.append(hold[0])
            
            Summary_Runs = ""
    if any("Voltage_Step_Test" in Functions for Functions in run_functions):
        #os.system("taskkill /im EXCEL.EXE") #this fully ensures the excel process is ended
        try:
            Voltage_Step_Test()
            print ">>> ", "Voltage_Step_Test()", " EXECUTED"
            #Popen("taskkill /F /im EXCEL.EXE",shell=True)
        except exceptions.Exception as e:
            print ">>> ", extract_function_name(), " FAILED"
            log_exception(e)
    #Popen("taskkill /F /im EXCEL.EXE",shell=True)
    clean_fort_files(WorkingFolder)
            
    if any("POC_Vref_Step_Test" in Functions for Functions in run_functions):
        #os.system("taskkill /im EXCEL.EXE") #this fully ensures the excel process is ended
        try:
            POC_Vref_Step_Test()
            print ">>> ", "POC_Vref_Step_Test()", " EXECUTED"
            #Popen("taskkill /F /im EXCEL.EXE",shell=True)
        except exceptions.Exception as e:
            print ">>> ", extract_function_name(), " FAILED"
            log_exception(e)
    #Popen("taskkill /F /im EXCEL.EXE",shell=True)
    clean_fort_files(WorkingFolder)
            
    if any("Voltage_Angle_Step_Test" in Functions for Functions in run_functions):
        #os.system("taskkill /im EXCEL.EXE") #this fully ensures the excel process is ended
        try:
            Voltage_Angle_Step_Test()
            print ">>> ", "Voltage_Angle_Step_Test()", " EXECUTED"
            #Popen("taskkill /F /im EXCEL.EXE",shell=True)
        except exceptions.Exception as e:
            print ">>> ", extract_function_name(), " FAILED"
            log_exception(e)
    #Popen("taskkill /F /im EXCEL.EXE",shell=True)
    clean_fort_files(WorkingFolder)

    if any("POC_Pref_Step_Test" in Functions for Functions in run_functions):
        #os.system("taskkill /im EXCEL.EXE") #this fully ensures the excel process is ended
        try:
            POC_Pref_Step_Test()
            print ">>> ", "POC_Pref_Step_Test()", " EXECUTED"
            #Popen("taskkill /F /im EXCEL.EXE",shell=True)
        except exceptions.Exception as e:
            print ">>> ", extract_function_name(), " FAILED"
            log_exception(e)
    #Popen("taskkill /F /im EXCEL.EXE",shell=True)
    clean_fort_files(WorkingFolder)
            
    if any("Fault_Study_Test" in Functions for Functions in run_functions):
        #os.system("taskkill /im EXCEL.EXE") #this fully ensures the excel process is ended
        try:
            Fault_Study_Test()
            print ">>> ", "Fault_Study_Test()", " EXECUTED"
            #Popen("taskkill /F /im EXCEL.EXE",shell=True)
        except exceptions.Exception as e:
            print ">>> ", extract_function_name(), " FAILED"
            log_exception(e)
    #Popen("taskkill /F /im EXCEL.EXE",shell=True)
    clean_fort_files(WorkingFolder)
            
    if any("FRT_Iq_Response_Test" in Functions for Functions in run_functions):
        #os.system("taskkill /im EXCEL.EXE") #this fully ensures the excel process is ended
        try:
            FRT_Iq_Response_Test()
            print ">>> ", "FRT_Iq_Response_Test()", " EXECUTED"
            #Popen("taskkill /F /im EXCEL.EXE",shell=True)
        except exceptions.Exception as e:
            print ">>> ", extract_function_name(), " FAILED"
            log_exception(e)
    #Popen("taskkill /F /im EXCEL.EXE",shell=True)
    clean_fort_files(WorkingFolder)
            
    if any("Under_Voltage_Trip_Test" in Functions for Functions in run_functions):
        #os.system("taskkill /im EXCEL.EXE") #this fully ensures the excel process is ended
        try:
            Under_Voltage_Trip_Test()
            print ">>> ", "Under_Voltage_Trip_Test()", " EXECUTED"
            #Popen("taskkill /F /im EXCEL.EXE",shell=True)
        except exceptions.Exception as e:
            print ">>> ", extract_function_name(), " FAILED"
            log_exception(e)
    #Popen("taskkill /F /im EXCEL.EXE",shell=True)
    clean_fort_files(WorkingFolder)
            
    if any("Over_Voltage_Trip_Test" in Functions for Functions in run_functions):
        #os.system("taskkill /im EXCEL.EXE") #this fully ensures the excel process is ended
        try:
            Over_Voltage_Trip_Test()
            print ">>> ", "Over_Voltage_Trip_Test()", " EXECUTED"
            #Popen("taskkill /F /im EXCEL.EXE",shell=True)
        except exceptions.Exception as e:
            print ">>> ", extract_function_name(), " FAILED"
            log_exception(e)
    #Popen("taskkill /F /im EXCEL.EXE",shell=True)
    clean_fort_files(WorkingFolder)
            
    if any("Under_Frequency_Trip_Test" in Functions for Functions in run_functions):
        #os.system("taskkill /im EXCEL.EXE") #this fully ensures the excel process is ended
        try:
            Under_Frequency_Trip_Test()
            print ">>> ", "Under_Frequency_Trip_Test()", " EXECUTED"
            #Popen("taskkill /F /im EXCEL.EXE",shell=True)
        except exceptions.Exception as e:
            print ">>> ", extract_function_name(), " FAILED"
            log_exception(e)
    #Popen("taskkill /F /im EXCEL.EXE",shell=True)
    clean_fort_files(WorkingFolder)
                
    if any("Over_Frequency_Trip_Test" in Functions for Functions in run_functions):
        #os.system("taskkill /im EXCEL.EXE") #this fully ensures the excel process is ended
        try:
            Over_Frequency_Trip_Test()
            print ">>> ", "Over_Frequency_Trip_Test()", " EXECUTED"
            #Popen("taskkill /F /im EXCEL.EXE",shell=True)
        except exceptions.Exception as e:
            print ">>> ", extract_function_name(), " FAILED"
            log_exception(e)
    #Popen("taskkill /F /im EXCEL.EXE",shell=True)
    clean_fort_files(WorkingFolder)
            
    if any("Long_Run_Test" in Functions for Functions in run_functions):
        #os.system("taskkill /im EXCEL.EXE") #this fully ensures the excel process is ended
        try:
            Long_Run_Test()
            print ">>> ", "Long_Run_Test()", " EXECUTED"
        except exceptions.Exception as e:
            print ">>> ", extract_function_name(), " FAILED"
            log_exception(e)
    #Popen("taskkill /F /im EXCEL.EXE",shell=True)
    clean_fort_files(WorkingFolder)

    sec = time.time()-start
    hrs = "%02d"%(sec//3600)
    min = "%02d"%((sec%3600)//60)
    sec = "%02d"%((sec%3600)%60)
    Outf=open(WorkingFolder + '\\' + 'Run Time.txt','w+')
    Outf.write("Run time - "+(hrs)+":"+(min)+":"+(sec)+" hrs")
    Outf.close()
    print("Run time - "+(hrs)+":"+(min)+":"+(sec)+" hrs")
    
#Python boilerplate
if __name__ == '__main__':
    main()