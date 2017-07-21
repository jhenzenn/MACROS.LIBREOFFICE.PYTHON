import os
import uno
import time

def InicializarLODesktop(nIntentos):
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)

    for i in range(1,nIntentos):
        try:
            time.sleep(1)
            context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
            if context:
                oDesktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
            return oDesktop
            break
        except:
            print("Esperando inicio de LibreOffice - Intento: " + str(i))

def planillaNueva():
    from com.sun.star.beans import PropertyValue
    oDesktop = InicializarLODesktop(10)
    oDataDoc = oDesktop.loadComponentFromURL("private:factory/scalc","_blank", 0, ())
    return oDataDoc

def importCSV(sCSVFileURL, sFrame, bVisible):
    from com.sun.star.beans import PropertyValue

    tDataProperties = []
    p = PropertyValue()

    p.value.Name = "Hidden"
    p.value.Value = not bVisible
    tDataProperties.append(p)

    p.value.Name = "FilterName"
    p.value.Value = "Text - txt - csv (StarCalc)"
    tDataProperties.append(p)

    p.value.Name = "FilterOptions"
    p.value.Value = "59,34,IBMPC_850,1,,0,false,true"
    tDataProperties.append(p)

    oDesktop = InicializarLODesktop(6)
    oDataDoc = oDesktop.loadComponentFromURL(sCSVFileURL, sFrame, 0, tDataProperties)
    if not oDataDoc:
        #print("Fallo al abrir archivo")
        return None
    else:
        return oDataDoc

'''def Main():
    oDoc = importCSV("file:///media/linux/csv/E01/resumen1.csv","_blank",True)
    if not oDoc:
        print("Fallo al abrir el archivo")
    else:
        oDoc.dispose()
        quit()

Main()'''
