import os
import time
import sys
import logging
import uno
import Archivo

'''logging.basicConfig()
log = logging.getLogger("Holapp")
log.setLevel(logging.DEBUG) #set verbosity to show all messages of severity >= DEBUG
log.info("Starting my app")'''

librebin = os.popen('which libreoffice | tr -d \'\\n\'').read()
libre_param = "--accept=\"socket,host=localhost,port=2002;urp;\" --norestore &"

os.system(librebin + " " + libre_param)
#os.system('/opt/openoffice4/program/soffice -accept=\"socket,host=localhost,port=2002\;urp\;\" &')
#time.sleep(6)

'''oDoc = Archivo.importCSV("file:///media/linux/csv/E01/1888.csv","_blank",True)
if not oDoc:
    print("Fallo al abrir el archivo")'''

oNuevoDocumento = Archivo.planillaNueva()
quit()


