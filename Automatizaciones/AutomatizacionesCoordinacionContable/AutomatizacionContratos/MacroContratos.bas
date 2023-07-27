Attribute VB_Name = "Módulo1"
Option Explicit

Dim col, fila As Integer
Dim colLetra As String

Sub MacroContratos()
'
' MacroContratos Macro
'

'
    col = Sheets("Hoja1").Range("A1").CurrentRegion.Columns.Count
    colLetra = Cells(1, col).Address
    fila = Sheets("Hoja1").Range("A1").CurrentRegion.Rows.Count

    Range("A4").Select
    Application.CutCopyMode = False
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Sheets("Hoja1").Range("A" & 1 & ":" & colLetra & fila), Version:=8).CreatePivotTable TableDestination:= _
        "Hoja2!R4C1", TableName:="TablaDinámica3", DefaultVersion:=8
    Sheets("Hoja2").Select
    Cells(4, 1).Select
    With ActiveSheet.PivotTables("TablaDinámica3")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("TablaDinámica3").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("TablaDinámica3").RepeatAllLabels xlRepeatLabels
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Acreedor").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Ps").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Nombre 1").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Población").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Rg").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Nº ident.fis.1"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Texto").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Doc.compr.").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("T").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Cl.").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("B").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("S").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Creado el").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Creado por").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("CPag").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("OrgC").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("GCp").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Mon.").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Tp.cambio").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Fecha doc.").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("IniPerVa").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("FinPerVal").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Gr.").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Estr.").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Lib").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("EstadLib").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Creado el2").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("char255").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Creado el3").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields( _
        "Periodicidad de la Evaluación").Subtotals = Array(False, False, False, False, _
        False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Nivel de riesgo"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Denominación"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("B2").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Texto breve").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Material").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Ce.").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Grupo art.").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Valor neto").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Valor neto2").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("EFi").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("FaF").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("RF").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Ce.gestor").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Pos.presup.").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Denom.clase-documen"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Fe.entrega").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Denom.gr.artíc."). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields( _
        "Denominación 2 del gr.artículo").Subtotals = Array(False, False, False, False, _
        False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Pos.").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("TablaDinámica3").RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables("TablaDinámica3").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Texto")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Nº ident.fis.1")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Acreedor")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Nombre 1")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Nivel de riesgo")
        .Orientation = xlRowField
        .Position = 5
    End With
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Denominación")
        .Orientation = xlRowField
        .Position = 6
    End With
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields("GCp")
        .Orientation = xlRowField
        .Position = 7
    End With
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Doc.compr.")
        .Orientation = xlRowField
        .Position = 8
    End With
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields( _
        "Denom.clase-documen")
        .Orientation = xlRowField
        .Position = 9
    End With
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Creado por")
        .Orientation = xlRowField
        .Position = 10
    End With
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Fecha doc.")
        .Orientation = xlRowField
        .Position = 11
    End With
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Fecha doc.").AutoGroup
    Range("K4").Select
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Años").Orientation = _
        xlHidden
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Trimestres"). _
        Orientation = xlHidden
    Selection.Ungroup
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields("IniPerVa")
        .Orientation = xlRowField
        .Position = 12
    End With
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("IniPerVa").AutoGroup
    ActiveWindow.ScrollColumn = 5
    Range("L4").Select
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Años").Orientation = _
        xlHidden
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Trimestres"). _
        Orientation = xlHidden
    Selection.Ungroup
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields("FinPerVal")
        .Orientation = xlRowField
        .Position = 13
    End With
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("FinPerVal").AutoGroup
    ActiveWindow.ScrollColumn = 6
    Range("M4").Select
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Años").Orientation = _
        xlHidden
    ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Trimestres"). _
        Orientation = xlHidden
    Selection.Ungroup
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields("char255")
        .Orientation = xlRowField
        .Position = 14
    End With
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Texto breve")
        .Orientation = xlRowField
        .Position = 15
    End With
    ActiveWindow.ScrollColumn = 7
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Mon.")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWorkbook.ShowPivotTableFieldList = False
End Sub


