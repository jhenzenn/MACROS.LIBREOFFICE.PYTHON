import os
import uno
import Archivo
import LOHelper
import Tablas
from com.sun.star.table import TableBorder
from com.sun.star.table import BorderLine

archivo_origen = "/media/linux/csv/E01/resumen2.csv"

oHelper = LOHelper.OfficeHelper()
#oDocCSV = oHelper.importCSV(oHelper.convertToURL(archivo_origen))
oDocNuevo = oHelper.createNewScalc()

#oDocNuevo.Sheets.importSheet(oDocCSV, oDocCSV.Sheets.getByIndex(0).Name, oDocNuevo.Sheets.Count)
#oDocCSV.close(0)
Archivo.PestanaDesdeCSV(archivo_origen,"Resumenes",oDocNuevo)
oDocNuevo.Sheets.removeByName(oDocNuevo.Sheets.getByIndex(0).Name)

oHojaActiva = oDocNuevo.getCurrentController().getActiveSheet()

mNombresRegiones =  ["CUENTAS CORRIENTES BANCARIAS","CHEQUES LIBRADOS","VENCIMIENTOS","COMPROBANTES PENDIENTES","CHEQUES EN CARTERA","RESUMEN DE MONEDAS","PROYECCION COBRANZAS"]
mColoresRegiones =  [10079487                      ,16750848          ,-1            ,-1                       ,10079487            ,10079487            ,-1                    ]
mColumnasRegiones = [[1,2,3,4]                     ,[2,3]             ,[2]           ,[2,3]                    ,[1,2]               ,[1,2,3,4]           ,[1]                   ]
mFilasTotalizadas = [True                          ,True              ,True          ,True                     ,False               ,True                ,True                  ]
mColumnasConNombre =[True                          ,True              ,True          ,True                     ,True                ,True                ,True                  ]

for sRegion in mNombresRegiones:
    i = mNombresRegiones.index(sRegion)
    oSD = oHojaActiva.createSearchDescriptor()
    oSD.setSearchString(sRegion)
    oSD.SearchWords = True
    oSD.SearchCaseSensitive = True
    oSD.SearchType = 1
    oBuscarEn = oHojaActiva.getCellRangeByName("A1:Z254")
    oEncontrado = oBuscarEn.findFirst(oSD)
    if not oEncontrado is None:
        oCursor = oHojaActiva.createCursorByRange(oEncontrado)
        oCursor.collapseToCurrentRegion()
        Tablas.FormatearRegion(oCursor,mColoresRegiones[i], mColumnasRegiones[i], mFilasTotalizadas[i], mColumnasConNombre[i])

oTableBorder = TableBorder()
oBorderLine = BorderLine()
# Cuadrito de resumen
oBorderLine.OuterLineWidth = 5
oTableBorder.IsBottomLineValid = True
oTableBorder.IsLeftLineValid = True
oTableBorder.IsRightLineValid = True
oTableBorder.IsTopLineValid = True
oTableBorder.IsHorizontalLineValid = False
oTableBorder.IsVerticalLineValid = False
oTableBorder.LeftLine = oBorderLine
oTableBorder.RightLine = oBorderLine
oTableBorder.TopLine = oBorderLine
oTableBorder.BottomLine = oBorderLine

oSD = oHojaActiva.createSearchDescriptor()
oSD.setSearchString("Resumen")
oSD.SearchWords = True
oSD.SearchCaseSensitive = True
oSD.SearchType = 1
oBuscarEn = oHojaActiva.getCellRangeByName("A1:Z254")
oEncontrado = oBuscarEn.findFirst(oSD)
oCursor = oHojaActiva.createCursorByRange(oEncontrado)
oCuadroResumen = oHojaActiva.getCellRangeByPosition(oCursor.RangeAddress.StartColumn,oCursor.RangeAddress.StartRow,oCursor.RangeAddress.StartColumn+1,oCursor.RangeAddress.StartRow+8)
oCuadroResumen.TableBorder = oTableBorder
oCuadroResumen.getCellRangeByPosition(0,0,oCuadroResumen.Columns.Count-1,0).CellBackColor = 0
oCuadroResumen.getCellRangeByPosition(0,0,oCuadroResumen.Columns.Count-1,0).CharFontName = "Arial"
oCuadroResumen.getCellRangeByPosition(0,0,oCuadroResumen.Columns.Count-1,0).CharHeight = 11
oCuadroResumen.getCellRangeByPosition(0,0,oCuadroResumen.Columns.Count-1,0).CharWeight = 150
oCuadroResumen.getCellRangeByPosition(0,0,oCuadroResumen.Columns.Count-1,0).CharColor = LOHelper.RGB(255,255,255)
oCuadroResumen.getCellRangeByPosition(0,oCuadroResumen.Rows.Count-3,oCuadroResumen.Columns.Count-1,oCuadroResumen.Rows.Count-3).CellBackColor = LOHelper.RGB(153,255,153)

oHojaActiva.getCellRangeByName("F1").Columns.Width = 800
oHojaActiva.getCellRangeByName("J1").Columns.Width = 800

quit()
