Attribute VB_Name = "Módulo1"
Option Explicit

Dim UltimaFilaC, UltimaFilaD As Long
Dim Cont As Long
Dim i, j, k, z, col, anterior As Integer
Dim Dinero As String
Dim HojaBalance, HojaSAP, FechaInforme, MesInforme, Texto, HojaPYG, HojaComparar As String
Dim Descripcion As String
Dim Encontrado As Boolean
Dim rng As Range

Sub Balance()

j = 2
HojaSAP = ActiveSheet.Name
UltimaFilaC = Sheets(HojaSAP).Range("C" & Rows.Count).End(xlUp).Row
HojaBalance = Sheets(1).Name
HojaComparar = Sheets(4).Name
UltimaFilaD = Sheets(HojaComparar).Range("C" & Rows.Count).End(xlUp).Row

FechaInforme = Sheets(HojaSAP).Cells(7, 10)
MesInforme = Mid(FechaInforme, 10, 2)

For i = 9 To UltimaFilaC
Encontrado = False

    If Sheets(HojaSAP).Cells(i, 5) <> Sheets(HojaBalance).Cells(j, 1) Then
        For k = 1 To UltimaFilaD
            If Sheets(HojaComparar).Cells(k, 5) = Sheets(HojaSAP).Cells(i, 5) Then
                Encontrado = True
                anterior = k - 1
                k = UltimaFilaD + 5
            End If
        Next k
        
        If Encontrado = False Then
            Sheets(HojaBalance).Activate
            Sheets(HojaBalance).Cells(j, 1).Select
            Sheets(HojaBalance).Cells(j, 1).EntireRow.Insert
        End If
        
    End If
    
    If IsEmpty(Sheets(HojaSAP).Cells(i + 1, 5)) And IsEmpty(Sheets(HojaSAP).Cells(i + 2, 5)) Then
        Dinero = Sheets(HojaSAP).Cells(i, 11)
        Descripcion = Sheets(HojaSAP).Cells(i, 8)
        Texto = Sheets(HojaSAP).Cells(i, 5)
        
        If Encontrado = True Then
            Dinero = 0
            Descripcion = Sheets(HojaComparar).Cells(anterior, 8)
            Texto = Sheets(HojaComparar).Cells(anterior, 5)
            i = i - 11
        End If
        
        Sheets(HojaBalance).Cells(j, 1) = Texto
        Sheets(HojaBalance).Cells(j, 2) = Descripcion
        Select Case MesInforme
            Case "01"
                Sheets(HojaBalance).Cells(j, 14) = Dinero
                Sheets(HojaBalance).Cells(j, 14) = VBA.Format(Sheets(HojaBalance).Cells(j, 14), "#.##0")
                Sheets(HojaBalance).Cells(j, 14) = Dinero
            Case "02"
                Sheets(HojaBalance).Cells(j, 13) = Dinero
                Sheets(HojaBalance).Cells(j, 13) = VBA.Format(Sheets(HojaBalance).Cells(j, 13), "#.##0")
                Sheets(HojaBalance).Cells(j, 13) = Dinero
            Case "03"
                Sheets(HojaBalance).Cells(j, 12) = Dinero
                Sheets(HojaBalance).Cells(j, 12) = VBA.Format(Sheets(HojaBalance).Cells(j, 12), "#.##0")
                Sheets(HojaBalance).Cells(j, 12) = Dinero
            Case "04"
                Sheets(HojaBalance).Cells(j, 11) = Dinero
                Sheets(HojaBalance).Cells(j, 11) = VBA.Format(Sheets(HojaBalance).Cells(j, 11), "#.##0")
                Sheets(HojaBalance).Cells(j, 11) = Dinero
            Case "05"
                Sheets(HojaBalance).Cells(j, 10) = Dinero
                Sheets(HojaBalance).Cells(j, 10) = VBA.Format(Sheets(HojaBalance).Cells(j, 10), "#.##0")
                Sheets(HojaBalance).Cells(j, 10) = Dinero
            Case "06"
                Sheets(HojaBalance).Cells(j, 9) = Dinero
                Sheets(HojaBalance).Cells(j, 9) = VBA.Format(Sheets(HojaBalance).Cells(j, 9), "#.##0")
                Sheets(HojaBalance).Cells(j, 9) = Dinero
            Case "07"
                Sheets(HojaBalance).Cells(j, 8) = Dinero
                Sheets(HojaBalance).Cells(j, 8) = VBA.Format(Sheets(HojaBalance).Cells(j, 8), "#.##0")
                Sheets(HojaBalance).Cells(j, 8) = Dinero
            Case "08"
                Sheets(HojaBalance).Cells(j, 7) = Dinero
                Sheets(HojaBalance).Cells(j, 7) = VBA.Format(Sheets(HojaBalance).Cells(j, 7), "#.##0")
                Sheets(HojaBalance).Cells(j, 7) = Dinero
            Case "09"
                Sheets(HojaBalance).Cells(j, 6) = Dinero
                Sheets(HojaBalance).Cells(j, 6) = VBA.Format(Sheets(HojaBalance).Cells(j, 6), "#.##0")
                Sheets(HojaBalance).Cells(j, 6) = Dinero
            Case "10"
                Sheets(HojaBalance).Cells(j, 5) = Dinero
                Sheets(HojaBalance).Cells(j, 5) = VBA.Format(Sheets(HojaBalance).Cells(j, 5), "#.##0")
                Sheets(HojaBalance).Cells(j, 5) = Dinero
            Case "11"
                Sheets(HojaBalance).Cells(j, 4) = Dinero
                Sheets(HojaBalance).Cells(j, 4) = VBA.Format(Sheets(HojaBalance).Cells(j, 4), "#.##0")
                Sheets(HojaBalance).Cells(j, 4) = Dinero
            Case "12"
                Sheets(HojaBalance).Cells(j, 3) = Dinero
                Sheets(HojaBalance).Cells(j, 3) = VBA.Format(Sheets(HojaBalance).Cells(j, 3), "#.##0")
                Sheets(HojaBalance).Cells(j, 3) = Dinero
        End Select
        
        j = j + 1
        i = i + 11
    End If
    
    Dinero = Sheets(HojaSAP).Cells(i, 11)
    Descripcion = Sheets(HojaSAP).Cells(i, 8)
    Texto = Sheets(HojaSAP).Cells(i, 5)
    
'    If Texto = "ACTIVO" Then
'            Sheets(HojaSAP).Cells(i, 5).Offset(1, 0).EntireRow.Insert
'    End If
'
    If Texto = "ESTADO DE RESULTADOS" Then
        Call PYG
        Exit Sub
    End If
    
    If Encontrado = True Then
            Dinero = 0
            Descripcion = Sheets(HojaComparar).Cells(anterior, 8)
            Texto = Sheets(HojaComparar).Cells(anterior, 5)
            i = i - 1
    End If
    
    Sheets(HojaBalance).Cells(j, 1) = Texto
    Sheets(HojaBalance).Cells(j, 2) = Descripcion
    Select Case MesInforme
        Case "01"
            Sheets(HojaBalance).Cells(j, 14) = Dinero
            Sheets(HojaBalance).Cells(j, 14) = VBA.Format(Sheets(HojaBalance).Cells(j, 14), "#.##0")
            Sheets(HojaBalance).Cells(j, 14) = Dinero
        Case "02"
            Sheets(HojaBalance).Cells(j, 13) = Dinero
            Sheets(HojaBalance).Cells(j, 13) = VBA.Format(Sheets(HojaBalance).Cells(j, 13), "#.##0")
            Sheets(HojaBalance).Cells(j, 13) = Dinero
        Case "03"
            Sheets(HojaBalance).Cells(j, 12) = Dinero
            Sheets(HojaBalance).Cells(j, 12) = VBA.Format(Sheets(HojaBalance).Cells(j, 12), "#.##0")
            Sheets(HojaBalance).Cells(j, 12) = Dinero
        Case "04"
            Sheets(HojaBalance).Cells(j, 11) = Dinero
            Sheets(HojaBalance).Cells(j, 11) = VBA.Format(Sheets(HojaBalance).Cells(j, 11), "#.##0")
            Sheets(HojaBalance).Cells(j, 11) = Dinero
        Case "05"
            Sheets(HojaBalance).Cells(j, 10) = Dinero
            Sheets(HojaBalance).Cells(j, 10) = VBA.Format(Sheets(HojaBalance).Cells(j, 10), "#.##0")
            Sheets(HojaBalance).Cells(j, 10) = Dinero
        Case "06"
            Sheets(HojaBalance).Cells(j, 9) = Dinero
            Sheets(HojaBalance).Cells(j, 9) = VBA.Format(Sheets(HojaBalance).Cells(j, 9), "#.##0")
            Sheets(HojaBalance).Cells(j, 9) = Dinero
        Case "07"
            Sheets(HojaBalance).Cells(j, 8) = Dinero
            Sheets(HojaBalance).Cells(j, 8) = VBA.Format(Sheets(HojaBalance).Cells(j, 8), "#.##0")
            Sheets(HojaBalance).Cells(j, 8) = Dinero
        Case "08"
            Sheets(HojaBalance).Cells(j, 7) = Dinero
            Sheets(HojaBalance).Cells(j, 7) = VBA.Format(Sheets(HojaBalance).Cells(j, 7), "#.##0")
            Sheets(HojaBalance).Cells(j, 7) = Dinero
        Case "09"
            Sheets(HojaBalance).Cells(j, 6) = Dinero
            Sheets(HojaBalance).Cells(j, 6) = VBA.Format(Sheets(HojaBalance).Cells(j, 6), "#.##0")
            Sheets(HojaBalance).Cells(j, 6) = Dinero
        Case "10"
            Sheets(HojaBalance).Cells(j, 5) = Dinero
            Sheets(HojaBalance).Cells(j, 5) = VBA.Format(Sheets(HojaBalance).Cells(j, 5), "#.##0")
            Sheets(HojaBalance).Cells(j, 5) = Dinero
        Case "11"
            Sheets(HojaBalance).Cells(j, 4) = Dinero
            Sheets(HojaBalance).Cells(j, 4) = VBA.Format(Sheets(HojaBalance).Cells(j, 4), "#.##0")
            Sheets(HojaBalance).Cells(j, 4) = Dinero
        Case "12"
            Sheets(HojaBalance).Cells(j, 3) = Dinero
            Sheets(HojaBalance).Cells(j, 3) = VBA.Format(Sheets(HojaBalance).Cells(j, 3), "#.##0")
            Sheets(HojaBalance).Cells(j, 3) = Dinero
    End Select
    j = j + 1
    
Next i


End Sub

Sub PYG()

i = i + 2
j = 2
HojaPYG = Sheets(2).Name

If MesInforme = "01" Then
    col = 11
Else
    col = 15
End If

For z = i To UltimaFilaC
    If Sheets(HojaSAP).Cells(z, 5) <> Sheets(HojaPYG).Cells(j, 1) Then
        Sheets(HojaPYG).Activate
        Sheets(HojaPYG).Cells(j, 1).Select
        Sheets(HojaPYG).Cells(j, 1).EntireRow.Insert
    End If
    If IsEmpty(Sheets(HojaSAP).Cells(z + 1, 5)) And IsEmpty(Sheets(HojaSAP).Cells(z + 2, 5)) Then
        Dinero = Sheets(HojaSAP).Cells(z, col)
        Descripcion = Sheets(HojaSAP).Cells(z, 8)
        Texto = Sheets(HojaSAP).Cells(z, 5)
        
        If IsNumeric(Texto) Then
            
        End If
        
'        If Texto = "ACTIVO" Then
'            ActiveCell.Offset(1, 0).EntireRow.Insert
'        End If
        
        Sheets(HojaPYG).Cells(j, 1) = Texto
        Sheets(HojaPYG).Cells(j, 2) = Descripcion
        Select Case MesInforme
            Case "01"
                Sheets(HojaPYG).Cells(j, 14) = Dinero
                Sheets(HojaPYG).Cells(j, 14) = VBA.Format(Sheets(HojaPYG).Cells(j, 14), "#.##0")
                Sheets(HojaPYG).Cells(j, 14) = Dinero
            Case "02"
                Sheets(HojaPYG).Cells(j, 13) = Dinero
                Sheets(HojaPYG).Cells(j, 13) = VBA.Format(Sheets(HojaPYG).Cells(j, 13), "#.##0")
                Sheets(HojaPYG).Cells(j, 13) = Dinero
            Case "03"
                Sheets(HojaPYG).Cells(j, 12) = Dinero
                Sheets(HojaPYG).Cells(j, 12) = VBA.Format(Sheets(HojaPYG).Cells(j, 12), "#.##0")
                Sheets(HojaPYG).Cells(j, 12) = Dinero
            Case "04"
                Sheets(HojaPYG).Cells(j, 11) = Dinero
                Sheets(HojaPYG).Cells(j, 11) = VBA.Format(Sheets(HojaPYG).Cells(j, 11), "#.##0")
                Sheets(HojaPYG).Cells(j, 11) = Dinero
            Case "05"
                Sheets(HojaPYG).Cells(j, 10) = Dinero
                Sheets(HojaPYG).Cells(j, 10) = VBA.Format(Sheets(HojaPYG).Cells(j, 10), "#.##0")
                Sheets(HojaPYG).Cells(j, 10) = Dinero
            Case "06"
                Sheets(HojaPYG).Cells(j, 9) = Dinero
                Sheets(HojaPYG).Cells(j, 9) = VBA.Format(Sheets(HojaPYG).Cells(j, 9), "#.##0")
                Sheets(HojaPYG).Cells(j, 9) = Dinero
            Case "07"
                Sheets(HojaPYG).Cells(j, 8) = Dinero
                Sheets(HojaPYG).Cells(j, 8) = VBA.Format(Sheets(HojaPYG).Cells(j, 8), "#.##0")
                Sheets(HojaPYG).Cells(j, 8) = Dinero
            Case "08"
                Sheets(HojaPYG).Cells(j, 7) = Dinero
                Sheets(HojaPYG).Cells(j, 7) = VBA.Format(Sheets(HojaPYG).Cells(j, 7), "#.##0")
                Sheets(HojaPYG).Cells(j, 7) = Dinero
            Case "09"
                Sheets(HojaPYG).Cells(j, 6) = Dinero
                Sheets(HojaPYG).Cells(j, 6) = VBA.Format(Sheets(HojaPYG).Cells(j, 6), "#.##0")
                Sheets(HojaPYG).Cells(j, 6) = Dinero
            Case "10"
                Sheets(HojaPYG).Cells(j, 5) = Dinero
                Sheets(HojaPYG).Cells(j, 5) = VBA.Format(Sheets(HojaPYG).Cells(j, 5), "#.##0")
                Sheets(HojaPYG).Cells(j, 5) = Dinero
            Case "11"
                Sheets(HojaPYG).Cells(j, 4) = Dinero
                Sheets(HojaPYG).Cells(j, 4) = VBA.Format(Sheets(HojaPYG).Cells(j, 4), "#.##0")
                Sheets(HojaPYG).Cells(j, 4) = Dinero
            Case "12"
                Sheets(HojaPYG).Cells(j, 3) = Dinero
                Sheets(HojaPYG).Cells(j, 3) = VBA.Format(Sheets(HojaPYG).Cells(j, 3), "#.##0")
                Sheets(HojaPYG).Cells(j, 3) = Dinero
        End Select
        
        If Texto = "ESTADO DE RESULTADOS" Then
            Exit Sub
        End If
        
        j = j + 1
        z = z + 11
    End If
    
    
    Dinero = Sheets(HojaSAP).Cells(z, col)
    Descripcion = Sheets(HojaSAP).Cells(z, 8)
    Texto = Sheets(HojaSAP).Cells(z, 5)
    
'    If Texto = "ACTIVO" Then
'            Sheets(HojaSAP).Cells(i, 5).Offset(1, 0).EntireRow.Insert
'    End If
'
    
    Sheets(HojaPYG).Cells(j, 1) = Texto
    Sheets(HojaPYG).Cells(j, 2) = Descripcion
    Select Case MesInforme
        Case "01"
            Sheets(HojaPYG).Cells(j, 14) = Dinero
            Sheets(HojaPYG).Cells(j, 14) = VBA.Format(Sheets(HojaPYG).Cells(j, 14), "#.##0")
            Sheets(HojaPYG).Cells(j, 14) = Dinero
        Case "02"
            Sheets(HojaPYG).Cells(j, 13) = Dinero
            Sheets(HojaPYG).Cells(j, 13) = VBA.Format(Sheets(HojaPYG).Cells(j, 13), "#.##0")
            Sheets(HojaPYG).Cells(j, 13) = Dinero
        Case "03"
            Sheets(HojaPYG).Cells(j, 12) = Dinero
            Sheets(HojaPYG).Cells(j, 12) = VBA.Format(Sheets(HojaPYG).Cells(j, 12), "#.##0")
            Sheets(HojaPYG).Cells(j, 12) = Dinero
        Case "04"
            Sheets(HojaPYG).Cells(j, 11) = Dinero
            Sheets(HojaPYG).Cells(j, 11) = VBA.Format(Sheets(HojaPYG).Cells(j, 11), "#.##0")
            Sheets(HojaPYG).Cells(j, 11) = Dinero
        Case "05"
            Sheets(HojaPYG).Cells(j, 10) = Dinero
            Sheets(HojaPYG).Cells(j, 10) = VBA.Format(Sheets(HojaPYG).Cells(j, 10), "#.##0")
            Sheets(HojaPYG).Cells(j, 10) = Dinero
        Case "06"
            Sheets(HojaPYG).Cells(j, 9) = Dinero
            Sheets(HojaPYG).Cells(j, 9) = VBA.Format(Sheets(HojaPYG).Cells(j, 9), "#.##0")
            Sheets(HojaPYG).Cells(j, 9) = Dinero
        Case "07"
            Sheets(HojaPYG).Cells(j, 8) = Dinero
            Sheets(HojaPYG).Cells(j, 8) = VBA.Format(Sheets(HojaPYG).Cells(j, 8), "#.##0")
            Sheets(HojaPYG).Cells(j, 8) = Dinero
        Case "08"
            Sheets(HojaPYG).Cells(j, 7) = Dinero
            Sheets(HojaPYG).Cells(j, 7) = VBA.Format(Sheets(HojaPYG).Cells(j, 7), "#.##0")
            Sheets(HojaPYG).Cells(j, 7) = Dinero
        Case "09"
            Sheets(HojaPYG).Cells(j, 6) = Dinero
            Sheets(HojaPYG).Cells(j, 6) = VBA.Format(Sheets(HojaPYG).Cells(j, 6), "#.##0")
            Sheets(HojaPYG).Cells(j, 6) = Dinero
        Case "10"
            Sheets(HojaPYG).Cells(j, 5) = Dinero
            Sheets(HojaPYG).Cells(j, 5) = VBA.Format(Sheets(HojaPYG).Cells(j, 5), "#.##0")
            Sheets(HojaPYG).Cells(j, 5) = Dinero
        Case "11"
            Sheets(HojaPYG).Cells(j, 4) = Dinero
            Sheets(HojaPYG).Cells(j, 4) = VBA.Format(Sheets(HojaPYG).Cells(j, 4), "#.##0")
            Sheets(HojaPYG).Cells(j, 4) = Dinero
        Case "12"
            Sheets(HojaPYG).Cells(j, 3) = Dinero
            Sheets(HojaPYG).Cells(j, 3) = VBA.Format(Sheets(HojaPYG).Cells(j, 3), "#.##0")
            Sheets(HojaPYG).Cells(j, 3) = Dinero
    End Select
    j = j + 1
    
    If Texto = "ESTADO DE RESULTADOS" Then
        Exit Sub
    End If
Next z

End Sub

Sub InsertarFila()

HojaBalance = Sheets(1).Name

Sheets(HojaBalance).Activate
Sheets(HojaBalance).Cells(2, 1).Select
Sheets(HojaBalance).Cells(2, 1).EntireRow.Insert
'rng.Offset(1, 0).EntireRow.Insert shift:=xlDown

End Sub

