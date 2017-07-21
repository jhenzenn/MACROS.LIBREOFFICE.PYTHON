import uno

def BorrarHoja(oDoc, sSheet):
    oSheets = oDoc.getSheets()
    oSheets.removeByName(sSheet)
    return True

