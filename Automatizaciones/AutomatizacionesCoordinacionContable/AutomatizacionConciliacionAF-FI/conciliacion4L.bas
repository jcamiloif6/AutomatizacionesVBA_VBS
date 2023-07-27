Attribute VB_Name = "Módulo1"
Option Explicit

Dim af, af2, fi, co As Worksheet
Dim afName, af2Name, fiName, coName, cuenta, valor, denominacion As String
Dim filasAF, filasCO, i, j, filasFI, hojas, k, z As Integer
Dim rango As Range

'depreciacion
Dim dep As Worksheet
Dim depName As String
Dim valorSuma, contador As Long
Dim valorSumaAF2, valorSumaAF1 As Currency

Sub conciliacion()

hojas = Sheets.Count

If hojas = 3 Then
    Set af = Sheets(2)
    Set fi = Sheets(1)
    Set co = ActiveSheet
    afName = af.Name
    fiName = fi.Name
    coName = co.Name
    'filasAF = Sheets("AF").Range("C9", Range("C9").End(xlDown)).Rows.Count
    filasFI = 500
    filasCO = Range("A2", Range("A2").End(xlDown)).Rows.Count
    
    Sheets(coName).Range("A" & 1).Value = "Cuenta"
    Sheets(coName).Range("B" & 1).Value = "Denominación"
    Sheets(coName).Range("C" & 1).Value = "AF"
    Sheets(coName).Range("D" & 1).Value = "FI"
    Sheets(coName).Range("E" & 1).Value = "DIF"
    
    For i = 2 To filasCO + 1
        cuenta = Sheets(coName).Range("A" & i).Value
        For j = 9 To filasCO + 8
            If Sheets(afName).Range("C" & j).Value = cuenta Then
                valor = Sheets(afName).Range("G" & j).Value
                denominacion = Sheets(afName).Range("D" & j).Value
                Sheets(coName).Range("B" & i).Value = denominacion
                Sheets(coName).Range("C" & i).Value = valor
                Exit For
            End If
        Next j
        
        For k = 14 To filasFI
            If Sheets(fiName).Range("E" & k).Value = cuenta Then
                valor = Sheets(fiName).Range("K" & k).Value
                Sheets(coName).Range("D" & i).Value = valor
                Exit For
            End If
        Next k
    Next i
    
    For i = 2 To filasCO + 1
        Sheets(coName).Range("E" & i).Value = Sheets(coName).Range("C" & i).Value - Sheets(coName).Range("D" & i).Value
    Next i

ElseIf hojas = 4 And (Sheets(2).Name = "21AF" Or Sheets(2).Name = "23AF") Then
    Set af = Sheets(2)
    Set af2 = Sheets(3)
    Set fi = Sheets(1)
    Set co = ActiveSheet
    afName = af.Name
    af2Name = af2.Name
    fiName = fi.Name
    coName = co.Name
    'filasAF = Sheets("AF").Range("C9", Range("C9").End(xlDown)).Rows.Count
    filasFI = 500
    filasCO = Range("A2", Range("A2").End(xlDown)).Rows.Count

    Sheets(coName).Range("B" & 1).Value = "Denominación"
    Sheets(coName).Range("C" & 1).Value = "AF"
    Sheets(coName).Range("D" & 1).Value = "FI"
    Sheets(coName).Range("E" & 1).Value = "DIF"

    For i = 2 To filasCO + 1
        cuenta = Sheets(coName).Range("A" & i).Value
        For j = 9 To filasCO + 8
            If Sheets(afName).Range("C" & j).Value = cuenta Then
                valor = Sheets(afName).Range("G" & j).Value
                denominacion = Sheets(afName).Range("D" & j).Value
                Sheets(coName).Range("B" & i).Value = denominacion
                Sheets(coName).Range("C" & i).Value = valor
                Exit For
            End If
        Next j
        
        For j = 9 To filasCO + 8
            If Sheets(af2Name).Range("C" & j).Value = cuenta Then
                valor = Sheets(af2Name).Range("G" & j).Value
                denominacion = Sheets(af2Name).Range("D" & j).Value
                Sheets(coName).Range("B" & i).Value = denominacion
                Sheets(coName).Range("C" & i).Value = valor
                Exit For
            End If
        Next j

        For k = 14 To filasFI
            If Sheets(fiName).Range("E" & k).Value = cuenta Then
                valor = Sheets(fiName).Range("K" & k).Value
                Sheets(coName).Range("D" & i).Value = valor
                Exit For
            End If
        Next k
    Next i

    For i = 2 To filasCO + 1
        Sheets(coName).Range("E" & i).Value = Sheets(coName).Range("C" & i).Value - Sheets(coName).Range("D" & i).Value
    Next i

ElseIf hojas = 4 Then
    Set af = Sheets(2)
    Set af2 = Sheets(3)
    Set fi = Sheets(1)
    Set co = ActiveSheet
    afName = af.Name
    af2Name = af2.Name
    fiName = fi.Name
    coName = co.Name
    'filasAF = Sheets("AF").Range("C9", Range("C9").End(xlDown)).Rows.Count
    filasFI = 500
    filasCO = Range("A2", Range("A2").End(xlDown)).Rows.Count
    
    Sheets(coName).Range("B" & 1).Value = "Denominación"
    Sheets(coName).Range("C" & 1).Value = "AF"
    Sheets(coName).Range("D" & 1).Value = "FI"
    Sheets(coName).Range("E" & 1).Value = "DIF"
    
    For i = 2 To filasCO + 1
        cuenta = Sheets(coName).Range("A" & i).Value
        For j = 9 To filasCO + 8
            If Sheets(afName).Range("C" & j).Value = cuenta Then
                valor = Sheets(afName).Range("G" & j).Value + Sheets(af2Name).Range("G" & j).Value
                denominacion = Sheets(afName).Range("D" & j).Value
                Sheets(coName).Range("B" & i).Value = denominacion
                Sheets(coName).Range("C" & i).Value = valor
                Exit For
            End If
        Next j
        
        For k = 14 To filasFI
            If Sheets(fiName).Range("E" & k).Value = cuenta Then
                valor = Sheets(fiName).Range("K" & k).Value
                Sheets(coName).Range("D" & i).Value = valor
                Exit For
            End If
        Next k
    Next i
    
    For i = 2 To filasCO + 1
        Sheets(coName).Range("E" & i).Value = Sheets(coName).Range("C" & i).Value - Sheets(coName).Range("D" & i).Value
    Next i

End If

Sheets(coName).Range("A" & i + 3).Value = "Depreciación"
i = Sheets(coName).Range("A" & i + 3).Row
Sheets(coName).Range("A" & i + 1).Value = "Cuenta"
Sheets(coName).Range("B" & i + 1).Value = "Denominación"
Sheets(coName).Range("C" & i + 1).Value = "AF"
Sheets(coName).Range("D" & i + 1).Value = "FI"
Sheets(coName).Range("E" & i + 1).Value = "DIF"


End Sub

Sub depreciacion()

k = 9
i = 2
z = 2
j = 9
valorSuma = 0
valorSumaAF1 = 0
valorSumaAF2 = 0

hojas = Sheets.Count

If hojas = 4 Then

    Set af = Sheets(2)
    Set fi = Sheets(1)
    Set dep = Sheets(4)
    Set co = ActiveSheet
    afName = af.Name
    fiName = fi.Name
    coName = co.Name
    depName = dep.Name
    
    Do Until Sheets(coName).Cells(i, 1).Value = "Depreciación"
        i = i + 1
    Loop
    
    i = i + 2
    
    Do Until IsEmpty(Sheets(coName).Cells(i, 1).Value)
        valorSuma = 0
        Do Until IsEmpty(Sheets(fiName).Cells(k, 5).Value)
            If Sheets(coName).Cells(i, 1).Value = Sheets(fiName).Cells(k, 5).Value Then
                Sheets(coName).Cells(i, 2).Value = Sheets(fiName).Cells(k, 8).Value
                Sheets(coName).Cells(i, 4).Value = Sheets(fiName).Cells(k, 11).Value
                k = 9
                Exit Do
            End If
            k = k + 1
        Loop
        
        Do Until IsEmpty(Sheets(depName).Cells(z, 1).Value)
            If Sheets(depName).Cells(z, 2).Value = Sheets(coName).Cells(i, 1).Value Then
                contador = Sheets(depName).Cells(z, 1).Value
                Do Until IsEmpty(Sheets(afName).Cells(j, 3).Value)
                    If Sheets(afName).Cells(j, 3).Value = contador Then
                        valorSuma = valorSuma + Sheets(afName).Cells(j, 8).Value
                    End If
                    j = j + 1
                Loop
                
                j = 9
                
            End If
            
            z = z + 1
        Loop
        
        Sheets(coName).Cells(i, 3).Value = valorSuma
        
        z = 2
        
        Sheets(coName).Cells(i, 5).Value = Sheets(coName).Cells(i, 3).Value - Sheets(coName).Cells(i, 4).Value
        
        i = i + 1
    Loop

ElseIf hojas = 5 Then

    Set af = Sheets(2)
    Set af2 = Sheets(3)
    Set fi = Sheets(1)
    Set dep = Sheets(5)
    Set co = ActiveSheet
    afName = af.Name
    af2Name = af2.Name
    fiName = fi.Name
    coName = co.Name
    depName = dep.Name
    
    Do Until Sheets(coName).Cells(i, 1).Value = "Depreciación"
        i = i + 1
    Loop
    
    i = i + 2
    
    Do Until IsEmpty(Sheets(coName).Cells(i, 1).Value)
        valorSuma = 0
        valorSumaAF1 = 0
        valorSumaAF2 = 0
        Do Until IsEmpty(Sheets(fiName).Cells(k, 5).Value)
            If Sheets(coName).Cells(i, 1).Value = Sheets(fiName).Cells(k, 5).Value Then
                Sheets(coName).Cells(i, 2).Value = Sheets(fiName).Cells(k, 8).Value
                Sheets(coName).Cells(i, 4).Value = Sheets(fiName).Cells(k, 11).Value
                k = 9
                Exit Do
            End If
            k = k + 1
        Loop
        
        Do Until IsEmpty(Sheets(depName).Cells(z, 1).Value)
            If Sheets(depName).Cells(z, 2).Value = Sheets(coName).Cells(i, 1).Value Then
                contador = Sheets(depName).Cells(z, 1).Value
                Do Until IsEmpty(Sheets(afName).Cells(j, 3).Value)
                    If Sheets(afName).Cells(j, 3).Value = contador Then
                        valorSumaAF1 = valorSumaAF1 + Sheets(afName).Cells(j, 8).Value
                    End If
                    j = j + 1
                Loop
                
                j = 9
                
                Do Until IsEmpty(Sheets(af2Name).Cells(j, 3).Value)
                    If Sheets(af2Name).Cells(j, 3).Value = contador Then
                        valorSumaAF2 = valorSumaAF2 + Sheets(af2Name).Cells(j, 8).Value
                    End If
                    j = j + 1
                Loop
                
                j = 9
                
            End If
            
            z = z + 1
        Loop
        
        valorSuma = valorSumaAF1 + valorSumaAF2
        
        Sheets(coName).Cells(i, 3).Value = valorSuma
        
        z = 2
        
        Sheets(coName).Cells(i, 5).Value = Sheets(coName).Cells(i, 3).Value - Sheets(coName).Cells(i, 4).Value
        
        i = i + 1
    Loop
    
    End If



End Sub
