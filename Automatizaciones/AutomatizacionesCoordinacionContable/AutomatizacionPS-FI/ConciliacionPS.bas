Attribute VB_Name = "Módulo1"
Option Explicit

Dim ps, fi, co As Worksheet
Dim psName, fiName, coName, cuenta, valor, denominacion As String
Dim filasCO, i, j, filasFI, hojas, k As Integer
Dim contador, valorSuma As Integer
Dim rango As Range

Sub conciliacion()

Set ps = Sheets(2)
Set fi = Sheets(1)
Set co = ActiveSheet
psName = ps.Name
fiName = fi.Name
coName = co.Name
'filasAF = Sheets("AF").Range("C9", Range("C9").End(xlDown)).Rows.Count
filasFI = 500
filasCO = Range("A2", Range("A2").End(xlDown)).Rows.Count

Sheets(coName).Range("A" & 1).Value = "CUENTA"
Sheets(coName).Range("B" & 1).Value = "DENMINACION"
Sheets(coName).Range("C" & 1).Value = "SALDO PS"
Sheets(coName).Range("D" & 1).Value = "SALDO FI"
Sheets(coName).Range("E" & 1).Value = "DIFERENCIA"

For i = 2 To filasCO + 1
    cuenta = Sheets(coName).Range("A" & i).Value
    For j = 6 To filasCO + 8
        If Sheets(psName).Range("F" & j).Value = cuenta Then
            valor = Sheets(psName).Range("G" & j).Value
            Sheets(coName).Range("C" & i).Value = valor
            Exit For
        End If
    Next j

    For k = 13 To filasFI
        If Sheets(fiName).Range("E" & k).Value = cuenta Then
            valor = Sheets(fiName).Range("K" & k).Value
            Sheets(coName).Range("D" & i).Value = valor
            denominacion = Sheets(fiName).Range("H" & k).Value
            Sheets(coName).Range("B" & i).Value = denominacion
            Exit For
        End If
    Next k
Next i

For i = 2 To filasCO + 1
    Sheets(coName).Range("E" & i).Value = Sheets(coName).Range("D" & i).Value - Sheets(coName).Range("C" & i).Value
Next i


End Sub
