from plxscripting.easy import *
import subprocess, time
import math
import pandas as pd
import os
import xlsxwriter

###############################################
PORT_i = 10000  # Define a port number.
PORT_o = 10001
PASSWORD = 'SxDBR<TYKRAX834~'  # Define a password.
# Start the scripting server.
s_i, g_i = new_server('localhost', PORT_i, password=PASSWORD)
s_o, g_o = new_server('localhost', PORT_o, password=PASSWORD)

###############################################
#Elements:
EXCEL_PATH=r'C:\Users\phtsang\Desktop\PLAXIS_V22\Script\\'
EXCEL_NAME='plate_output.xlsx'
FILENAME=EXCEL_PATH+EXCEL_NAME
###############################################
#Inputs:
plate_input=['Plate_1']
phase_input=['Third excavation stage [Phase_5]']

plate=[plt for plt in g_o.Plates[:]] # Loop through the plate object in existing model
phase=[p for p in g_o.Phases[:]] # Loop through all available phases

###############################################
#Get Plate results:
def get_plate(plate_o,phase_o):

    plateY=g_o.getresults(plate_o,phase_o,g_o.ResultTypes.Plate.Y, "node")
    plateQ=g_o.getresults(plate_o,phase_o,g_o.ResultTypes.Plate.Q2D, "node")
    plateM=g_o.getresults(plate_o,phase_o,g_o.ResultTypes.Plate.M2D, "node")
    plateAxial=g_o.getresults(plate_o,phase_o,g_o.ResultTypes.Plate.Nx2D, "node")    


    phasename=str(phase_o.Identification).split('[')[0]
    col1='Shear [kN]'+'_'+phasename
    col2='Bending Moment [kNm/m]'+'_'+phasename
    col3='Axial Force [kN]'+'_'+phasename
 

    results = {'Y': plateY, col1: plateQ,col2: plateM,col3: plateAxial}

    plateresults=pd.DataFrame(results)
    plateresults = plateresults.sort_values(by=['Y'],ascending=False)

    return plateresults
###############################################
#Export Excel:
def export_excel(plate_input,phase_input,filename):
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')

    name=str(phase[5].Identification).split(' [')[1]
    name=name.split(']')[0]
    sheet_name = "%s_%s" % (plate[0].Name, name)
    results = get_plate(plate[0], phase[5])
    results.to_excel(writer,sheet_name=sheet_name,index=False)
    writer.save()


export_excel(plate_input,phase_input,FILENAME)

