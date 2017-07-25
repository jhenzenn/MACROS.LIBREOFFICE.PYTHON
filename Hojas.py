import uno


def BorrarHoja(oDoc, sSheet):
    try:
        oSheets = oDoc.getSheets()
        oSheets.removeByName(sSheet)
        return True
    except:
        print("No pude borrar la hoja " + sSheet)
        return False

def RangoCopiar(oDoc, oSheet, sRange):
    try:
        localcontext = uno.getComponentContext()
        resolver = localcontext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localcontext)
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        oDispatcher = smgr.createInstance("com.sun.star.frame.DispatchHelper")
        #oDispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
        #seleccionamos el rango
        oDoc.CurrentController.select(oSheet.getCellRangeByName(sRange))
        #lo copiamos
        oDispatcher.executeDispatch(oDoc, ".uno:Copy", "", 0, ())
        return True
    except:
        print("No pude copiar")
        return False

def RangoPegar(oDoc, oSheet, sRange):
    try:
        localcontext = uno.getComponentContext()
        resolver = localcontext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localcontext)
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        oDispatcher = smgr.createInstance("com.sun.star.frame.DispatchHelper")
        #oDispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
        #seleccionamos el rango
        oDoc.CurrentController.select(oSheet.getCellRangeByName(sRange))
        #lo copiamos
        oDispatcher.executeDispatch(oDoc, ".uno:Paste", "", 0, ())
        return True
    except:
        print("No pude pegar")
        return False

def CopiarHoja(oDocOrigen, oSheetOrigen, oDocDestino, sSheetDestino):
    try:
        oSheets = oDocDestino.getSheets()
        oSheets.insertNewByName(sSheetDestino, oSheets.Count)
        oSheetDestino = oSheets.getByName(sSheetDestino)
        RangoCopiar(oDocOrigen, oSheetOrigen, oSheetOrigen.getCellRangeByName("A1:AMJ1048576"))
        RangoPegar(oDocDestino, oSheetDestino, "A1")
        return True
    except:
        print("No pude pegar la hoja " + sSheetDestino)
        return False


