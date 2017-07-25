import os
import uno
import LOHelper

archivo_origen = "/media/linux/csv/E01/resumen2.csv"

oHelper = LOHelper.OfficeHelper()
oDocCSV = oHelper.importCSV(oHelper.convertToURL(archivo_origen))
oDocNuevo = oHelper.createNewScalc()

oDocNuevo.Sheets.importSheet(oDocCSV, oDocCSV.Sheets.getByIndex(0).Name, oDocNuevo.Sheets.Count)
oDocCSV.close(0)
oDocNuevo.Sheets.removeByName(oDocNuevo.Sheets.getByIndex(0).Name)

quit()
