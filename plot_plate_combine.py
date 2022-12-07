from plxscripting.easy import *
import math
import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import (
    ScatterChart,
    Reference,
    Series,
)
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
EXCEL_NAME='Plate_y.xlsx'

FILENAME=EXCEL_PATH+EXCEL_NAME

plate=[plt for plt in g_o.Plates[:]]
phase=[p for p in g_o.Phases[:]]

###############################################
#Inputs:
plate_input=['Plate_1']
phase_input=['Installation of strut [Phase_3]','Second (submerged) excavation stage [Phase_4]','Third excavation stage [Phase_5]']

###############################################
#Get Plate results:
def get_plate(interface_o,phase_o):

    plateY=g_o.getresults(interface_o,phase_o,g_o.ResultTypes.Plate.Y, "node")
    plateM=g_o.getresults(interface_o,phase_o,g_o.ResultTypes.Plate.M2D, "node")


    phasename=str(phase_o.Identification).split('[')[0]
    col1='Bending Moment [kNm/m]'+'_'+phasename

    results = {'Y': plateY,col1: plateM}

    plateresults=pd.DataFrame(results)
    plateresults = plateresults.sort_values(by=['Y'],ascending=False)

    return plateresults