Attribute VB_Name = "Módulo1"
Option Explicit
Dim final As Boolean
Dim i As Integer
Sub MacroCFR09()
Attribute MacroCFR09.VB_ProcData.VB_Invoke_Func = " \n14"

i = 14
final = False

Do Until final = True

    If Cells(i, 5).Value = "TOTAL CUENTAS NO ASIGNADAS" Then
        final = True
    ElseIf IsNumeric(Cells(i, 5).Value) Then
        If Cells(i, 11).Value >= 0 Then
            Cells(i, 21).Value = Cells(i, 11).Value
        Else
            Cells(i, 21).Value = "Revisar"
        End If
    End If
    i = i + 1
Loop
    
Call Macro3
    
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Range("J6:L7").Select
    Selection.Cut
    Range("K6").Select
    ActiveSheet.Paste
    Columns("I:J").Select
    Range("I4").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("J:J").Select
    Range("J4").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("K:K").Select
    Range("K4").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("L:L").Select
    Range("L4").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("F:G").Select
    Range("F4").Activate
    Selection.Delete Shift:=xlToLeft
    Range("I10").Select
    ActiveWindow.SmallScroll Down:=-3
    
    Range("A6:M6").Select
    Selection.AutoFilter
    
    Cells(6, 12).Value = "COMENTARIOS CUENTA NATURALEZA BANCARIA"
    Cells(6, 13).Value = "CUENTAS NUEVAS"
    
    Range("L6:M6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 12611584
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("L6:M6").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("L11").Select
    
End Sub

