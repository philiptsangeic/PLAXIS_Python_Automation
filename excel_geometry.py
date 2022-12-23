from plxscripting.easy import *
import subprocess, time
import pandas as pd
import openpyxl

###############################################
PLAXIS_PATH = r'C:\Program Files\Bentley\Geotechnical\PLAXIS 2D CONNECT Edition V22\\Plaxis2DXInput.exe'  # Specify PLAXIS path on server.
PORT_i = 10000  # Define a port number.
PORT_o = 10001
PASSWORD = 'SxDBR<TYKRAX834~'  # Define a password.
subprocess.Popen([PLAXIS_PATH, f'--AppServerPassword={PASSWORD}', f'--AppServerPort={PORT_i}'], shell=False)  # Start the PLAXIS remote scripting service.
time.sleep(5)  # Wait for PLAXIS to boot before sending commands to the scripting service.

# Start the scripting server.
s_i, g_i = new_server('localhost', PORT_i, password=PASSWORD)
s_o, g_o = new_server('localhost', PORT_o, password=PASSWORD)

s_i.new()

g_i.gotostructures()
###############################################
#This section reads the properties of structures from excel template

source=r"C:\Users\phtsang\Desktop\PLAXIS_V22\Python_automation"
file="Input_param"+".xlsx"
geomsheet="Geometry"
df_geom = pd.read_excel(file, sheet_name = geomsheet,engine="openpyxl")

geom=[]
for i in range(2):
    geom1=df_geom[[df_geom.columns[i*2],df_geom.columns[i*2+1]]].set_index(df_geom.columns[i*2]).to_dict()
    geom.append(geom1) #Gives a dictionary of geometry properties

###############################################
#Plates:
platesheet="Plates"
df_plate = pd.read_excel(file, sheet_name = platesheet,engine="openpyxl")

#Material
material=list(dict.fromkeys(df_plate.iloc[:,2].to_list()))

for i in range(len(material)):
    g_i.platemat('Identification',material[i])

platematerials = [mat for mat in g_i.Materials[:] if mat.TypeName.value == 'PlateMat']

for i in range(df_plate.count()[0]):
    #Create Plate
    plate=g_i.plate(
        (geom[0]['X value'][df_plate.iloc[i,3]],geom[1]['Y value'][df_plate.iloc[i,4]]),
        (geom[0]['X value'][df_plate.iloc[i,5]],geom[1]['Y value'][df_plate.iloc[i,6]]),
        )
    plate1=plate[-1]
    #Rename Plate
    plate1.rename(df_plate.iloc[i,0])
    #Create Interface
    if df_plate.iloc[i,1] == 'Y':
        plate2=plate[-2]
        g_i.posinterface(plate2)
        g_i.neginterface(plate2)
    #Rename Interface
        plate2.PositiveInterface.rename(df_plate.iloc[i,0]+'_PosInterface')
        plate2.NegativeInterface.rename(df_plate.iloc[i,0]+'_NegInterface')
    #Set Material
    for j in range(len(material)):
        if df_plate.iloc[i,2] == platematerials[j].Identification:
            plate1.setmaterial(platematerials[j])

###############################################
#Material Database:

#Plate:
platematsheet="PlateMatName"
df_platemat = pd.read_excel(file, sheet_name = platematsheet,engine="openpyxl")

for i in range(df_platemat.count()[0]):
    EA = df_platemat.iloc[i,1]
    EI = df_platemat.iloc[i,2]
    w = df_platemat.iloc[i,3]
    nu = df_platemat.iloc[i,4]
    if df_platemat.iloc[i,5] == 'Y':
        Punch=True
    else:
        Punch=False

    platematerials = [mat for mat in g_i.Materials[:] if mat.TypeName.value == 'PlateMat']

    for j in range(len(platematerials)):
        if df_platemat.iloc[i,0] == platematerials[j].Identification:
            platematerials[j].setproperties("MaterialType","Elastic", "w", w, "EA1", EA, 
            "EI", EI, "StructNu", nu, 'PreventPunching', Punch)

