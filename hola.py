import os
import time
import sys
import logging
import uno
import Archivo
import Hojas

librebin = os.popen('which libreoffice | tr -d \'\\n\'').read()
libre_param = "--accept=\"socket,host=localhost,port=2002;urp;\" --norestore &"
#os.system(librebin + " " + libre_param)
#os.system('/opt/openoffice4/program/soffice -accept=\"socket,host=localhost,port=2002\;urp\;\" &')
#time.sleep(6)

#oDoc = Archivo.importCSV("file:///media/linux/csv/E01/1888.csv","_blank",True)
#if not oDoc:
#    print("Fallo al abrir el archivo")'''

archivo_origen = "/media/linux/csv/E01/resumen2.csv"
if not Archivo.ArchivoExiste(archivo_origen):
    print("Archivo " + archivo_origen + " no encontrado")
    quit()

oNuevoDocumento = Archivo.planillaNueva()
Archivo.PestanaDesdeCSV(archivo_origen, "Resumenes", oNuevoDocumento)


quit()


