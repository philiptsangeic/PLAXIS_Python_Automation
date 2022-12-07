from plxscripting.easy import *
import subprocess, time

PLAXIS_PATH = r'C:\Program Files\Bentley\Geotechnical\PLAXIS 2D CONNECT Edition V22\\Plaxis2DXInput.exe'  # Specify PLAXIS path on server.

PORT_i = 10000  # Define a port number.
PORT_o = 10001

PASSWORD = 'SxDBR<TYKRAX834~'  # Define a password (up to user choice).

subprocess.Popen([PLAXIS_PATH, f'--AppServerPassword={PASSWORD}', f'--AppServerPort={PORT_i}'], shell=False)  # Start the PLAXIS remote scripting service.

time.sleep(5)  # Wait for PLAXIS to boot before sending commands to the scripting service.


# Start the scripting server.

s_i, g_i = new_server('localhost', PORT_i, password=PASSWORD)
s_o, g_o = new_server('localhost', PORT_o, password=PASSWORD)


s_i.new()

g_i.gotostructures()

###############
# Plate

#Material name and geometry
material=['Concrete']

df_plate=[{'x1':-10,'y1':0,'x2':10,'y2':0},{'x1':-10,'y1':10,'x2':-10,'y2':0},{'x1':-10,'y1':10,'x2':10,'y2':10},{'x1':10,'y1':10,'x2':10,'y2':0}]

# Create material
for i in range(len(material)):
    g_i.platemat('Identification',material[i])

platematerials = [mat for mat in g_i.Materials[:] if mat.TypeName.value == 'PlateMat']

#Create Plate
for i in range(len(df_plate)):
    plate=g_i.plate(
        (df_plate[i]['x1'],df_plate[i]['y1']),
        (df_plate[i]['x2'],df_plate[i]['y2']),
        )
    plate1=plate[-1]

#Create Interface

    plate2=plate[-2]
    g_i.posinterface(plate2)
    g_i.neginterface(plate2)
#Set Material

    plate1.setmaterial(platematerials[0])
