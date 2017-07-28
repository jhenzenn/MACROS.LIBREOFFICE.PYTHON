import uno
import LOHelper


def FormatearRegion(oRango, color, ColumnasNumericas, UltimaFilaEsTotal, SegundaFilaEsTitulo):
    from com.sun.star.table import TableBorder
    from com.sun.star.table import BorderLine
    oBorderLine = BorderLine()
    oBorderLine.OuterLineWidth = 5

    oTableBorder = TableBorder()
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

    oSubRango = oRango.getCellRangeByPosition(0, 0, oRango.Columns.Count - 1, 0)
    oSubRango.CharFontName = "Arial"
    oSubRango.CharHeight = 12
    oSubRango.CharWeight = 150
    oSubRango.CellBackColor = color
    #oSubRango.CharColor = RGB(0,0,0)
        #RGB(255,255,204)
        #.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
        #.Rows.Height = 900
        #.TableBorder = oTableBorder

    # si la 1º fila es ..... sacamos las negritas
    if oRango.getCellRangeByPosition(0, 0, 0, 0).FormulaArray[0][0] == "Saldo Deudores x Vtas":
        oRango.getCellRangeByPosition(0, 0, oRango.Columns.Count - 1, 0).CharWeight = 100
        oRango.getCellRangeByPosition(1, 0, 1, 0).NumberFormat = 104
    # aca clavar el if si la 2ª fila es de titulo, entonces poner negritas
    if SegundaFilaEsTitulo:
        oRango.getCellRangeByPosition(0, 1, oRango.Columns.Count - 1, 1).CharWeight = 150
        NumFilasTitulo = 2
    else:
        NumFilasTitulo = 1

    oSubRango = oRango.getCellRangeByPosition(0, 1, oRango.Columns.Count - 1, oRango.Rows.Count - 2)
    oSubRango.CharFontName = "Arial"
    oSubRango.CharHeight = 10

    if UltimaFilaEsTotal:
        oSubRango = oRango.getCellRangeByPosition(0, oRango.Rows.Count - 1, oRango.Columns.Count - 1, oRango.Rows.Count - 1)
        oSubRango.CharFontName = "Arial"
        oSubRango.CharHeight = 10
        oSubRango.CharWeight = 150
        oSubRango.TableBorder = oTableBorder

    for i in ColumnasNumericas:
        oSubRango = oRango.getCellRangeByPosition(i,NumFilasTitulo,i,oRango.Rows.Count-1)
        oSubRango.NumberFormat = 4
        oSubRango.Columns.OptimalWidth = True


