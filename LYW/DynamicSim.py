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

"""

import os, sys, math, csv, time
from win32com import client

PSSE_LOCATION = r"C:\Program Files (x86)\PTI\PSSE34\PSSBIN"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['path'] + ';' + PSSE_LOCATION

PSSE_LOCATION = r"C:\Program Files (x86)\PTI\PSSE34\PSSPY27"
sys.path.append(PSSE_LOCATION)
os.environ['PATH'] = os.environ['path'] + ';' + PSSE_LOCATION

import psse34
import psspy
import redirect		#redirects popups

WorkingFolder = os.getcwd()

file_name = WorkingFolder + "\DATA.txt" 	
f = open(file_name,'r')
lines = f.readlines()
f.close()

# ====================================== Setup ====================================== #
DYR_File_GNCLS = str(lines[[i for i, s in enumerate(lines) if 'DYR_File_GNCLS =' in s][0]].replace("DYR_File_GNCLS =", "").replace("\t", "").replace("\n", "").replace(" ", ""))
DYR_File_ZINGEN = str(lines[[i for i, s in enumerate(lines) if 'DYR_File =' in s][0]].replace("DYR_File =", "").replace("\t", "").replace("\n", "").replace(" ", ""))
CON_File = 'Conv.sav'
SNP_File = 'Snap.snp'

DYR_File = str(lines[[i for i, s in enumerate(lines) if 'DYR_File =' in s][0]].replace("DYR_File =", "").replace("\t", "").replace("\n", "").replace(" ", ""))

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
		exit()


SMIB = int(lines[[i for i, s in enumerate(lines) if 'SMIB =' in s][0]].replace("SMIB =", "").replace("\t", "").replace("\n", ""))			#SMIB bus number\
POC = int(lines[[i for i, s in enumerate(lines) if 'POC =' in s][0]].replace("POC =", "").replace("\t", "").replace("\n", ""))				#POC bus number
INV1_Bus = int(lines[[i for i, s in enumerate(lines) if 'INV1_Bus =' in s][0]].replace("INV1_Bus =", "").replace("\t", "").replace("\n", ""))			#INV1_Bus
POC_frBus = int(lines[[i for i, s in enumerate(lines) if 'POC_frBus =' in s][0]].replace("POC_frBus =", "").replace("\t", "").replace("\n", ""))		#POC_frBus
POC_toBus = int(lines[[i for i, s in enumerate(lines) if 'POC_toBus =' in s][0]].replace("POC_toBus =", "").replace("\t", "").replace("\n", ""))		#POC_toBus
bus_mch1 = int(lines[[i for i, s in enumerate(lines) if 'INV1_Bus =' in s][0]].replace("INV1_Bus =", "").replace("\t", "").replace("\n", ""))		# unit under test (UUT) bus
mID = str(lines[[i for i, s in enumerate(lines) if 'mID1 =' in s][0]].replace("mID1 =", "").replace("\t", "").replace("\n", "").replace(" ", ""))		# machine ID
bus_inf = int(lines[[i for i, s in enumerate(lines) if 'SMIB =' in s][0]].replace("SMIB =", "").replace("\t", "").replace("\n", ""))			#SMIB bus number\
bus_PCC = int(lines[[i for i, s in enumerate(lines) if 'POC =' in s][0]].replace("POC =", "").replace("\t", "").replace("\n", ""))				#POC bus number
#EK ADDED - find common folder - Get Bus Data --------------------------------------END

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

global pathTestFiles
pathTestFiles = WorkingFolder+'\\'+"Test Files"
# ==================================================================================== #

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

def run_Dynamics(Case,runTime,Run,pathTestDir):
	global pathTestFiles
	global CON_File
	global SNP_File

	intialise_PSSE((pathTestDir+"\\PSSEOut"),Case,Run)
	psspy.case(pathTestFiles+'\\'+Case)
	
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
	
	psspy.dyre_new([1,1,1,1],(WorkingFolder+'\\'+DYR_File_ZINGEN),(pathTestFiles+'\\'+"Conec.flx"),(pathTestFiles+'\\'+"Conet.flx"),(pathTestFiles+'\\'+"Compile.bat"))	#! Read in the dynamic data file		
	psspy.save(pathTestFiles+'\\'+CON_File)  #! Save the converted case
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
	psspy.branch_p_and_q_channel([chani,-1, -1, POC_frBus,POC_toBus], r"""1""", ['P_POC','Q_POC'])	#8,9
	chani += 2
	psspy.bus_frequency_channel([chani,POC],r"""FRE_POC""") #10
	chani += 1
	psspy.voltage_channel([chani,-1,-1,SMIB],r"""SMIB_Voltage""") #11
	chani += 1
	# psspy.voltage_channel([12,-1,-1,41400],r"""4TVZ132A""") #11
	# psspy.voltage_channel([13,-1,-1,41401],r"""4TVZ132B""") #11
	

	psspy.dynamics_solution_param_2(intgar1=max_solns, realar1=sfactor, realar2 =con_tol, realar3=dT, realar4 =frq_filter,
									realar5=int_delta_thrsh, realar6=islnd_delta_thrsh, realar7=islnd_sfactor, realar8=islnd_con_tol)
	#! Setup Dynamic Simulation parameters     --------------Update for Every Project ------------------------
	if 'sFac0.3' in Run:
		psspy.dynamics_solution_param_2(realar1=0.3)
	if 'sFac1.0' in Run:
		psspy.dynamics_solution_param_2(realar1=1.0)	
		
	#! Update simulation step according to Run name										
	if 'dT0.001' in Run:									
		psspy.dynamics_solution_param_2(realar3=0.001)
	if 'dT0.002' in Run:									
		psspy.dynamics_solution_param_2(realar3=0.002)	
	psspy.set_netfrq(1)
	psspy.set_relang(1,0,"")
	psspy.snap(sfile=pathTestFiles+'\\'+SNP_File)	
	OUTPUT_name = pathTestDir + '\\' + "Results" + '\\' + Case +"_"+ Run + '.out'
	# Initalise
	psspy.case(pathTestFiles+'\\'+CON_File)
	psspy.rstr(pathTestFiles+'\\'+SNP_File)
	psspy.strt(0,OUTPUT_name)
	psspy.strt(0,OUTPUT_name)
	psspy.strt(0,OUTPUT_name)
	# Run dymanic simulations
	psspy.run(0, runTime,5,5,5)
	psspy.pssehalt_2()
def main():
	'''Main'''
	Case = str(sys.argv[1])
	runTime = float(sys.argv[2])
	Run = str(sys.argv[3])
	pathTestDir = str(sys.argv[4])
	run_Dynamics(Case,runTime,Run,pathTestDir)

#Python boilerplate
if __name__ == '__main__':
	main() 

	
