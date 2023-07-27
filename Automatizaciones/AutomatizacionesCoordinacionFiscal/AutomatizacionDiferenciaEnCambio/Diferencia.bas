Attribute VB_Name = "Módulo1"
Option Explicit

Dim decision1, valorEntrada, valorSalida, fila, i, filaSubtotal, j, k, valorFinal As Integer
Dim z, cont, subir As Integer
Dim salidaParcial As Long
Dim filaSalida As Integer
Dim texto, textoValorEntrada, textoValorSalida, fecha, trm, concepto, trasladoComparar, traslado As String
Dim subtotal, subtotalComparar, hojaBanco, nombreBanco As String
Dim mes, fechaMovimiento As String
Dim final, encontradoSubtotal, encontradoControl, finalBanco, encontrarCelda As Boolean

Dim finalEntrada As Boolean
Dim y, saldoFiscalCOP As Integer

'Datos de Control
Dim a, colControl As Integer
Dim valorSumatoria
Dim totalBanco, totalBancoComparar As String

Sub diferencia()

hojaBanco = Sheets(3).Name
nombreBanco = Sheets(hojaBanco).Cells(1, 1).Value

nombreBanco = Mid(nombreBanco, 1, 3)

If nombreBanco = "JPM" Then
    Call JPM
ElseIf nombreBanco = "BNP" Then
    Call BNP
ElseIf nombreBanco = "CIT" Then
End If

End Sub


Sub controlJP()

colControl = 27
a = 6
encontradoControl = False
totalBanco = "TOTAL BANCO JP"

Do Until encontradoControl = True
    totalBancoComparar = Cells(a, 1).Value
    totalBancoComparar = Mid(totalBancoComparar, 1, 14)
    totalBancoComparar = UCase(totalBancoComparar)
    If totalBancoComparar = totalBanco Then
        Do Until IsEmpty(Cells(a, colControl))
            valorSumatoria = valorSumatoria + (Cells(a, colControl).Value * Cells(6, colControl).Value)
            colControl = colControl + 1
        Loop
        Cells(a + 1, 22).Value = Cells(a, 22).Value - valorSumatoria
        encontradoControl = True
    End If
    a = a + 1
Loop

End Sub

Sub JPM()

finalBanco = False
encontradoSubtotal = False
i = 4
k = 4
z = 6
y = 1
cont = 0
final = False
subtotal = "SUBTOTAL"
traslado = "TRASLADO"

Do Until finalBanco = True
    nombreBanco = Cells(z, 1).Value
    nombreBanco = Mid(nombreBanco, 8, 2)
    If nombreBanco = "JP" Then
        finalBanco = True
        cont = z
    End If
    z = z + 1
Loop

'loop que recorre la hoja del banco
Do Until IsEmpty(Sheets(hojaBanco).Cells(k, 1))
    fecha = Sheets(hojaBanco).Cells(k, 1).Value
    fechaMovimiento = fecha
    encontrarCelda = False
    z = cont
    
    If Len(fecha) = 9 Then
        fecha = Mid(fecha, 3, 2)
    ElseIf Len(fecha) = 10 Then
        fecha = Mid(fecha, 4, 2)
    End If
    
    'loop para encontrar celda vacía donde se pondra dato
    Do Until encontrarCelda = True
        subtotalComparar = Sheets("MonExtranjera").Cells(z, 1).Value
        mes = subtotalComparar
        subtotalComparar = Mid(subtotalComparar, 1, 8)
        subtotalComparar = UCase(subtotalComparar)
        
        'if para encontrar la palabra subtotal
        If subtotalComparar = subtotal Then
            Cells(z, 1).Select
            ActiveCell.EntireRow.Insert
            z = z + 1
            mes = Mid(mes, 10)
            If fecha = "01" And mes = "ENERO" Then
                
                subir = z - 1
                Do Until Not IsEmpty(Cells(subir, 2).Value)
                    If Mid(Cells(subir, 1).Value, 1, 8) = subtotal Then
                        subir = subir + 1
                        Exit Do
                    End If
                    subir = subir - 1
                Loop
                subir = subir + 1
                encontrarCelda = True
                
                Cells(subir, 2) = Sheets(hojaBanco).Cells(k, 2).Value
                trasladoComparar = Mid(Sheets(hojaBanco).Cells(k, 2).Value, 1, 8)
                trasladoComparar = UCase(trasladoComparar)
                Cells(subir, 5) = Sheets(hojaBanco).Cells(k, 1).Value
                i = 4
                Do Until final = True
                    If (Sheets("TRM BanRep").Cells(i, 1).Value = "" And Sheets("TRM BanRep").Cells(i + 1, 1).Value = "") Or Sheets("TRM BanRep").Cells(i, 1).Value = fechaMovimiento Then
                        final = True
                        Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                        Sheets("MonExtranjera").Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                        Exit Do
                    End If
                    i = i + 1
                Loop
                
                final = False
                
                'if para determinar si es salida o si es entrada
                If Not IsEmpty(Sheets(hojaBanco).Cells(k, 5)) Then 'if salida
                    valorSalida = Sheets(hojaBanco).Cells(k, 5).Value
                    
                    For j = 27 To 200
                        If Cells(z, j).Value > 0 Then
                            salidaParcial = Cells(z, j).Value
                            If valorSalida > salidaParcial Then
                                Cells(subir, j).Value = salidaParcial * -1
                                Cells(subir, 10).Value = 1
                                Cells(subir, 11).Value = 1
                                valorSalida = valorSalida - salidaParcial
                                Cells(z, j).Value = 0
                                Cells(subir, 16).Value = salidaParcial
                                Cells(subir, 17).Value = salidaParcial
                                Cells(subir, 18).Value = Cells(6, j).Value
                                
                                If trasladoComparar = traslado Then
                                    Cells(subir, 19).Value = Cells(subir, 18).Value
                                Else
                                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                                End If
                                
                                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                                
                                'Saldo Fiscal Mon/Ex
                                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
                '
                                'Saldo Fiscal COP
                                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                                   + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                                   * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                                
                                Cells(subir + 1, j).Select
                                ActiveCell.EntireRow.Insert
                                Cells(subir + 1, 1).Value = Cells(subir, 1).Value
                                Cells(subir + 1, 2).Value = Cells(subir, 2).Value
                                Cells(subir + 1, 3).Value = Cells(subir, 3).Value
                                Cells(subir + 1, 5).Value = Cells(subir, 5).Value
                                z = z + 1
                                
                                valorFinal = j
                            Else
                                Cells(subir, j).Value = valorSalida * -1
                                Cells(subir, 10).Value = 1
                                Cells(subir, 11).Value = 1
                                Cells(z, j).Value = Cells(z, j) - valorSalida
                                Cells(subir, 16).Value = valorSalida
                                Cells(subir, 17).Value = valorSalida
                                Cells(subir, 18).Value = Cells(6, j).Value
                                
                                If trasladoComparar = traslado Then
                                    Cells(subir, 19).Value = Cells(subir, 18).Value
                                Else
                                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                                End If
                                
                                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                                
                                'Saldo Fiscal Mon/Ex
                                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
                '
                                'Saldo Fiscal COP
                                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                                    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                                    * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                                
                                valorSalida = 0
                                
                                Exit For
                                
                            End If
                            subir = subir + 1
                        End If
                    Next j
                    
                    If valorSalida > 0 Then
                    
                        Cells(z, valorFinal).Value = -valorSalida
                        MsgBox "No hay dinero suficiente para realizar la salida de " & valorSalida & "USD", _
                            vbOKOnly + vbExclamation, "Alerta"
                        k = 100
                    End If
                    
                ElseIf Not IsEmpty(Sheets(hojaBanco).Cells(k, 4)) Then
                    valorEntrada = Sheets(hojaBanco).Cells(k, 4).Value
                    Cells(subir, 10).Value = 1
                    Cells(subir, 11).Value = 1
                    Cells(subir, 12).Value = valorEntrada
                    Cells(subir, 13).Value = valorEntrada
                    Cells(subir, 15).Value = Cells(subir, 12).Value * Cells(subir, 14).Value
                    
                    Do Until finalEntrada = True
                        If Cells(5, y).Value = fechaMovimiento Then
                            Cells(subir, y).Value = valorEntrada
                            Cells(z, y).Value = Cells(z, y).Value + valorEntrada
                            finalEntrada = True
                        End If
                        y = y + 1
                    Loop
                    
                    y = 1
                    finalEntrada = False
                    
                    Cells(subir, 21).Value = Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value
                     Cells(subir, 22).Value = Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                    * Cells(subir, 10).Value * Cells(subir, 8).Value
                    
'                    Cells(subir, 22).Value = saldoFiscalCOP
                                
                End If 'if para determinar si es salida o si es entrada
                
            ElseIf fecha = "02" And mes = "FEBRERO" Then
                Call febreroJP
            ElseIf fecha = "03" And mes = "MARZO" Then
                Call marzoJP
            ElseIf fecha = "04" And mes = "ABRIL" Then
                Call abrilJP
            ElseIf fecha = "05" And mes = "MAYO" Then
                Call mayoJP
            ElseIf fecha = "06" And mes = "JUNIO" Then
                Call junioJP
            ElseIf fecha = "07" And mes = "JULIO" Then
                Call julioJP
            ElseIf fecha = "08" And mes = "AGOSTO" Then
                Call agostoJP
            ElseIf fecha = "09" And mes = "SEPTIEMBRE" Then
                Call septiembreJP
            ElseIf fecha = "10" And mes = "OCTUBRE" Then
                Call octubreJP
            ElseIf fecha = "11" And mes = "NOVIEMBRE" Then
                Call noviembreJP
            ElseIf fecha = "12" And mes = "DICIEMBRE" Then
                Call diciembreJP
            End If
        End If 'if para encontrar la palabra subtotal
        
        z = z + 1
        
    Loop 'loop para encontrar celda vacía donde se pondra dato
    
    k = k + 1
Loop 'loop que recorre la hoja del banco





'decision1 = InputBox(texto)
'fila = ActiveCell.Row
'filaSubtotal = fila
'concepto = Cells(fila, 2).Value
'trasladoComparar = Mid(concepto, 1, 8)
'
'Do Until encontradoSubtotal = True
'    subtotalComparar = Cells(filaSubtotal, 1).Value
'    subtotalComparar = Mid(subtotalComparar, 1, 8)
'    subtotalComparar = UCase(subtotalComparar)
'    If subtotalComparar = subtotal Then
'        encontradoSubtotal = True
'        filaSalida = filaSubtotal
'    End If
'    filaSubtotal = filaSubtotal + 1
'Loop
'
'trasladoComparar = UCase(trasladoComparar)
'
'Do Until final = True
'    If (Sheets("TRM BanRep").Cells(i, 1).Value = "" And Sheets("TRM BanRep").Cells(i + 1, 1).Value = "") Or Sheets("TRM BanRep").Cells(i, 1).Value = fecha Then
'        final = True
'        Sheets("MonExtranjera").Cells(fila, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
'        Sheets("MonExtranjera").Cells(fila, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
'        Exit Do
'    End If
'    i = i + 1
'Loop
'
''Entradas
'If decision1 = 1 Then
'    i = 1
'    final = False
'
'    valorEntrada = InputBox(textoValorEntrada)
'    Cells(fila, 12).Value = valorEntrada
'    Cells(fila, 13).Value = valorEntrada
'    Cells(fila, 15).Value = Cells(fila, 12).Value * Cells(fila, 14).Value
'
'    Do Until final = True
'        If Cells(5, i).Value = fecha Then
'            Cells(fila, i).Value = valorEntrada
'            Cells(filaSalida, i).Value = Cells(filaSalida, i).Value + valorEntrada
'            final = True
'        End If
'        i = i + 1
'    Loop
'
'    'Saldo Fiscal Mon/Ex
'    Cells(fila, 21).Value = Cells(fila, 12).Value + Cells(fila, 16).Value + Cells(fila, 6).Value
''
''    'Saldo Fiscal COP
''    Cells(fila, 22).Value = Cells(fila, 16).Value * Cells(fila, 10).Value * Cells(fila, 18).Value _
''        + Cells(fila, 12).Value * Cells(fila, 10).Value * Cells(fila, 14).Value + Cells(fila, 6).Value _
''        * Cells(fila, 10).Value * Cells(fila, 8).Value
'
'End If
'
''Salidas
'If decision1 = 2 Then
'
'    valorSalida = InputBox(textoValorSalida)
'
'    For j = 27 To 1000
'        If Cells(filaSalida, j).Value > 0 Then
'            salidaParcial = Cells(filaSalida, j).Value
'            If valorSalida > salidaParcial Then
'                Cells(fila, j).Value = salidaParcial * -1
'                valorSalida = valorSalida - salidaParcial
'                Cells(filaSalida, j).Value = 0
'                Cells(fila, 16).Value = salidaParcial
'                Cells(fila, 17).Value = salidaParcial
'                Cells(fila, 18).Value = Cells(6, j).Value
'
'                If trasladoComparar = traslado Then
'                    Cells(fila, 19).Value = Cells(fila, 18).Value
'                Else
'                    Cells(fila, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
'                End If
'
'                Cells(fila, 20).Value = Cells(fila, 19).Value * Cells(fila, 16).Value
'                Sheets("MonExtranjera").Cells(fila, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
'
'                'Saldo Fiscal Mon/Ex
'                Cells(fila, 21).Value = Cells(fila, 12).Value + Cells(fila, 16).Value + Cells(fila, 6).Value
''
'                'Saldo Fiscal COP
'                Cells(fila, 22).Value = Cells(fila, 16).Value * Cells(fila, 10).Value * Cells(fila, 18).Value _
'                   + Cells(fila, 12).Value * Cells(fila, 10).Value * Cells(fila, 14).Value + Cells(fila, 6).Value _
'                   * Cells(fila, 10).Value * Cells(fila, 8).Value
'
'                Cells(fila + 1, j).Select
'                ActiveCell.EntireRow.Insert
'                Cells(fila + 1, 1).Value = Cells(fila, 1).Value
'                Cells(fila + 1, 2).Value = Cells(fila, 2).Value
'                Cells(fila + 1, 3).Value = Cells(fila, 3).Value
'                Cells(fila + 1, 5).Value = Cells(fila, 5).Value
'                filaSalida = filaSalida + 1
'
'                valorFinal = j
'            Else
'                Cells(fila, j).Value = valorSalida * -1
'                Cells(filaSalida, j).Value = Cells(filaSalida, j) - valorSalida
'                Cells(fila, 16).Value = valorSalida
'                Cells(fila, 17).Value = valorSalida
'                Cells(fila, 18).Value = Cells(6, j).Value
'
'                If trasladoComparar = traslado Then
'                    Cells(fila, 19).Value = Cells(fila, 18).Value
'                Else
'                    Cells(fila, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
'                End If
'
'                Cells(fila, 20).Value = Cells(fila, 19).Value * Cells(fila, 16).Value
'                Sheets("MonExtranjera").Cells(fila, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
'
'                'Saldo Fiscal Mon/Ex
'                Cells(fila, 21).Value = Cells(fila, 12).Value + Cells(fila, 16).Value + Cells(fila, 6).Value
''
'                'Saldo Fiscal COP
'                Cells(fila, 22).Value = Cells(fila, 16).Value * Cells(fila, 10).Value * Cells(fila, 18).Value _
'                    + Cells(fila, 12).Value * Cells(fila, 10).Value * Cells(fila, 14).Value + Cells(fila, 6).Value _
'                    * Cells(fila, 10).Value * Cells(fila, 8).Value
'
'                valorSalida = 0
'
'                Exit For
'
'            End If
'            fila = fila + 1
'        End If
'    Next j
'
'    If valorSalida > 0 Then
'
'        Cells(filaSalida, valorFinal).Value = -valorSalida
'
'        MsgBox "No hay dinero suficiente para realizar la salida de " & valorSalida & "USD", _
'            vbOKOnly + vbExclamation, "Alerta"
'    End If
'
'End If

'Call control

End Sub

Sub BNP()



End Sub

Sub febreroJP()

subir = z - 1
Do Until Not IsEmpty(Cells(subir, 2).Value)
    If Mid(Cells(subir, 1).Value, 1, 8) = subtotal Then
        subir = subir + 1
        Exit Do
    End If
    subir = subir - 1
Loop
subir = subir + 1
encontrarCelda = True

Cells(subir, 2) = Sheets(hojaBanco).Cells(k, 2).Value
trasladoComparar = Mid(Sheets(hojaBanco).Cells(k, 2).Value, 1, 8)
trasladoComparar = UCase(trasladoComparar)
Cells(subir, 5) = Sheets(hojaBanco).Cells(k, 1).Value
i = 4
Do Until final = True
    If (Sheets("TRM BanRep").Cells(i, 1).Value = "" And Sheets("TRM BanRep").Cells(i + 1, 1).Value = "") Or Sheets("TRM BanRep").Cells(i, 1).Value = fechaMovimiento Then
        final = True
        Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Sheets("MonExtranjera").Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Exit Do
    End If
    i = i + 1
Loop

final = False

'if para determinar si es salida o si es entrada
If Not IsEmpty(Sheets(hojaBanco).Cells(k, 5)) Then 'if salida
    valorSalida = Sheets(hojaBanco).Cells(k, 5).Value
    
    For j = 27 To 200
        If Cells(z, j).Value > 0 Then
            salidaParcial = Cells(z, j).Value
            If valorSalida > salidaParcial Then
                Cells(subir, j).Value = salidaParcial * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                valorSalida = valorSalida - salidaParcial
                Cells(z, j).Value = 0
                Cells(subir, 16).Value = salidaParcial
                Cells(subir, 17).Value = salidaParcial
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                   + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                   * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                Cells(subir + 1, j).Select
                ActiveCell.EntireRow.Insert
                Cells(subir + 1, 1).Value = Cells(subir, 1).Value
                Cells(subir + 1, 2).Value = Cells(subir, 2).Value
                Cells(subir + 1, 3).Value = Cells(subir, 3).Value
                Cells(subir + 1, 5).Value = Cells(subir, 5).Value
                z = z + 1
                
                valorFinal = j
            Else
                Cells(subir, j).Value = valorSalida * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                Cells(z, j).Value = Cells(z, j) - valorSalida
                Cells(subir, 16).Value = valorSalida
                Cells(subir, 17).Value = valorSalida
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                    * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                valorSalida = 0
                
                Exit For
                
            End If
            subir = subir + 1
        End If
    Next j
    
    If valorSalida > 0 Then
    
        Cells(z, valorFinal).Value = -valorSalida
        
        MsgBox "No hay dinero suficiente para realizar la salida de " & valorSalida & " USD", _
            vbOKOnly + vbExclamation, "Alerta"
            
        k = 100
    End If
    
ElseIf Not IsEmpty(Sheets(hojaBanco).Cells(k, 4)) Then
    valorEntrada = Sheets(hojaBanco).Cells(k, 4).Value
    Cells(subir, 10).Value = 1
    Cells(subir, 11).Value = 1
    Cells(subir, 12).Value = valorEntrada
    Cells(subir, 13).Value = valorEntrada
    Cells(subir, 15).Value = Cells(subir, 12).Value * Cells(subir, 14).Value
    
    Do Until finalEntrada = True
        If Cells(5, y).Value = fechaMovimiento Then
            Cells(subir, y).Value = valorEntrada
            Cells(z, y).Value = Cells(z, y).Value + valorEntrada
            finalEntrada = True
        End If
        y = y + 1
    Loop
    
    y = 1
    finalEntrada = False
    
    Cells(subir, 21).Value = Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value
     Cells(subir, 22).Value = Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
    * Cells(subir, 10).Value * Cells(subir, 8).Value
    
'                    Cells(subir, 22).Value = saldoFiscalCOP
                
End If 'if para determinar si es salida o si es entrada

End Sub

Sub marzoJP()

subir = z - 1
Do Until Not IsEmpty(Cells(subir, 2).Value)
    If Mid(Cells(subir, 1).Value, 1, 8) = subtotal Then
        subir = subir + 1
        Exit Do
    End If
    subir = subir - 1
Loop
subir = subir + 1
encontrarCelda = True

Cells(subir, 2) = Sheets(hojaBanco).Cells(k, 2).Value
trasladoComparar = Mid(Sheets(hojaBanco).Cells(k, 2).Value, 1, 8)
trasladoComparar = UCase(trasladoComparar)
Cells(subir, 5) = Sheets(hojaBanco).Cells(k, 1).Value
i = 4
Do Until final = True
    If (Sheets("TRM BanRep").Cells(i, 1).Value = "" And Sheets("TRM BanRep").Cells(i + 1, 1).Value = "") Or Sheets("TRM BanRep").Cells(i, 1).Value = fechaMovimiento Then
        final = True
        Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Sheets("MonExtranjera").Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Exit Do
    End If
    i = i + 1
Loop

final = False

'if para determinar si es salida o si es entrada
If Not IsEmpty(Sheets(hojaBanco).Cells(k, 5)) Then 'if salida
    valorSalida = Sheets(hojaBanco).Cells(k, 5).Value
    
    For j = 27 To 200
        If Cells(z, j).Value > 0 Then
            salidaParcial = Cells(z, j).Value
            If valorSalida > salidaParcial Then
                Cells(subir, j).Value = salidaParcial * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                valorSalida = valorSalida - salidaParcial
                Cells(z, j).Value = 0
                Cells(subir, 16).Value = salidaParcial
                Cells(subir, 17).Value = salidaParcial
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                   + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                   * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                Cells(subir + 1, j).Select
                ActiveCell.EntireRow.Insert
                Cells(subir + 1, 1).Value = Cells(subir, 1).Value
                Cells(subir + 1, 2).Value = Cells(subir, 2).Value
                Cells(subir + 1, 3).Value = Cells(subir, 3).Value
                Cells(subir + 1, 5).Value = Cells(subir, 5).Value
                z = z + 1
                
                valorFinal = j
            Else
                Cells(subir, j).Value = valorSalida * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                Cells(z, j).Value = Cells(z, j) - valorSalida
                Cells(subir, 16).Value = valorSalida
                Cells(subir, 17).Value = valorSalida
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                    * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                valorSalida = 0
                
                Exit For
                
            End If
            subir = subir + 1
        End If
    Next j
    
    If valorSalida > 0 Then
    
        Cells(z, valorFinal).Value = -valorSalida
        
        MsgBox "No hay dinero suficiente para realizar la salida de " & valorSalida & " USD", _
            vbOKOnly + vbExclamation, "Alerta"
            
        k = 100
    End If
    
ElseIf Not IsEmpty(Sheets(hojaBanco).Cells(k, 4)) Then
    valorEntrada = Sheets(hojaBanco).Cells(k, 4).Value
    Cells(subir, 10).Value = 1
    Cells(subir, 11).Value = 1
    Cells(subir, 12).Value = valorEntrada
    Cells(subir, 13).Value = valorEntrada
    Cells(subir, 15).Value = Cells(subir, 12).Value * Cells(subir, 14).Value
    
    Do Until finalEntrada = True
        If Cells(5, y).Value = fechaMovimiento Then
            Cells(subir, y).Value = valorEntrada
            Cells(z, y).Value = Cells(z, y).Value + valorEntrada
            finalEntrada = True
        End If
        y = y + 1
    Loop
    
    y = 1
    finalEntrada = False
    
    Cells(subir, 21).Value = Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value
     Cells(subir, 22).Value = Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
    * Cells(subir, 10).Value * Cells(subir, 8).Value
    
'                    Cells(subir, 22).Value = saldoFiscalCOP
                
End If 'if para determinar si es salida o si es entrada

End Sub

Sub abrilJP()

subir = z - 1
Do Until Not IsEmpty(Cells(subir, 2).Value)
    If Mid(Cells(subir, 1).Value, 1, 8) = subtotal Then
        subir = subir + 1
        Exit Do
    End If
    subir = subir - 1
Loop
subir = subir + 1
encontrarCelda = True

Cells(subir, 2) = Sheets(hojaBanco).Cells(k, 2).Value
trasladoComparar = Mid(Sheets(hojaBanco).Cells(k, 2).Value, 1, 8)
trasladoComparar = UCase(trasladoComparar)
Cells(subir, 5) = Sheets(hojaBanco).Cells(k, 1).Value
i = 4
Do Until final = True
    If (Sheets("TRM BanRep").Cells(i, 1).Value = "" And Sheets("TRM BanRep").Cells(i + 1, 1).Value = "") Or Sheets("TRM BanRep").Cells(i, 1).Value = fechaMovimiento Then
        final = True
        Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Sheets("MonExtranjera").Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Exit Do
    End If
    i = i + 1
Loop

final = False

'if para determinar si es salida o si es entrada
If Not IsEmpty(Sheets(hojaBanco).Cells(k, 5)) Then 'if salida
    valorSalida = Sheets(hojaBanco).Cells(k, 5).Value
    
    For j = 27 To 200
        If Cells(z, j).Value > 0 Then
            salidaParcial = Cells(z, j).Value
            If valorSalida > salidaParcial Then
                Cells(subir, j).Value = salidaParcial * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                valorSalida = valorSalida - salidaParcial
                Cells(z, j).Value = 0
                Cells(subir, 16).Value = salidaParcial
                Cells(subir, 17).Value = salidaParcial
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                   + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                   * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                Cells(subir + 1, j).Select
                ActiveCell.EntireRow.Insert
                Cells(subir + 1, 1).Value = Cells(subir, 1).Value
                Cells(subir + 1, 2).Value = Cells(subir, 2).Value
                Cells(subir + 1, 3).Value = Cells(subir, 3).Value
                Cells(subir + 1, 5).Value = Cells(subir, 5).Value
                z = z + 1
                
                valorFinal = j
            Else
                Cells(subir, j).Value = valorSalida * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                Cells(z, j).Value = Cells(z, j) - valorSalida
                Cells(subir, 16).Value = valorSalida
                Cells(subir, 17).Value = valorSalida
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                    * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                valorSalida = 0
                
                Exit For
                
            End If
            subir = subir + 1
        End If
    Next j
    
    If valorSalida > 0 Then
    
        Cells(z, valorFinal).Value = -valorSalida
        
        MsgBox "No hay dinero suficiente para realizar la salida de " & valorSalida & " USD", _
            vbOKOnly + vbExclamation, "Alerta"
            
        k = 100
    End If
    
ElseIf Not IsEmpty(Sheets(hojaBanco).Cells(k, 4)) Then
    valorEntrada = Sheets(hojaBanco).Cells(k, 4).Value
    Cells(subir, 10).Value = 1
    Cells(subir, 11).Value = 1
    Cells(subir, 12).Value = valorEntrada
    Cells(subir, 13).Value = valorEntrada
    Cells(subir, 15).Value = Cells(subir, 12).Value * Cells(subir, 14).Value
    
    Do Until finalEntrada = True
        If Cells(5, y).Value = fechaMovimiento Then
            Cells(subir, y).Value = valorEntrada
            Cells(z, y).Value = Cells(z, y).Value + valorEntrada
            finalEntrada = True
        End If
        y = y + 1
    Loop
    
    y = 1
    finalEntrada = False
    
    Cells(subir, 21).Value = Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value
     Cells(subir, 22).Value = Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
    * Cells(subir, 10).Value * Cells(subir, 8).Value
    
'                    Cells(subir, 22).Value = saldoFiscalCOP
                
End If 'if para determinar si es salida o si es entrada

End Sub

Sub mayoJP()

subir = z - 1
Do Until Not IsEmpty(Cells(subir, 2).Value)
    If Mid(Cells(subir, 1).Value, 1, 8) = subtotal Then
        subir = subir + 1
        Exit Do
    End If
    subir = subir - 1
Loop
subir = subir + 1
encontrarCelda = True

Cells(subir, 2) = Sheets(hojaBanco).Cells(k, 2).Value
trasladoComparar = Mid(Sheets(hojaBanco).Cells(k, 2).Value, 1, 8)
trasladoComparar = UCase(trasladoComparar)
Cells(subir, 5) = Sheets(hojaBanco).Cells(k, 1).Value
i = 4
Do Until final = True
    If (Sheets("TRM BanRep").Cells(i, 1).Value = "" And Sheets("TRM BanRep").Cells(i + 1, 1).Value = "") Or Sheets("TRM BanRep").Cells(i, 1).Value = fechaMovimiento Then
        final = True
        Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Sheets("MonExtranjera").Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Exit Do
    End If
    i = i + 1
Loop

final = False

'if para determinar si es salida o si es entrada
If Not IsEmpty(Sheets(hojaBanco).Cells(k, 5)) Then 'if salida
    valorSalida = Sheets(hojaBanco).Cells(k, 5).Value
    
    For j = 27 To 200
        If Cells(z, j).Value > 0 Then
            salidaParcial = Cells(z, j).Value
            If valorSalida > salidaParcial Then
                Cells(subir, j).Value = salidaParcial * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                valorSalida = valorSalida - salidaParcial
                Cells(z, j).Value = 0
                Cells(subir, 16).Value = salidaParcial
                Cells(subir, 17).Value = salidaParcial
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                   + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                   * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                Cells(subir + 1, j).Select
                ActiveCell.EntireRow.Insert
                Cells(subir + 1, 1).Value = Cells(subir, 1).Value
                Cells(subir + 1, 2).Value = Cells(subir, 2).Value
                Cells(subir + 1, 3).Value = Cells(subir, 3).Value
                Cells(subir + 1, 5).Value = Cells(subir, 5).Value
                z = z + 1
                
                valorFinal = j
            Else
                Cells(subir, j).Value = valorSalida * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                Cells(z, j).Value = Cells(z, j) - valorSalida
                Cells(subir, 16).Value = valorSalida
                Cells(subir, 17).Value = valorSalida
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                    * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                valorSalida = 0
                
                Exit For
                
            End If
            subir = subir + 1
        End If
    Next j
    
    If valorSalida > 0 Then
    
        Cells(z, valorFinal).Value = -valorSalida
        
        MsgBox "No hay dinero suficiente para realizar la salida de " & valorSalida & " USD", _
            vbOKOnly + vbExclamation, "Alerta"
            
        k = 100
    End If
    
ElseIf Not IsEmpty(Sheets(hojaBanco).Cells(k, 4)) Then
    valorEntrada = Sheets(hojaBanco).Cells(k, 4).Value
    Cells(subir, 10).Value = 1
    Cells(subir, 11).Value = 1
    Cells(subir, 12).Value = valorEntrada
    Cells(subir, 13).Value = valorEntrada
    Cells(subir, 15).Value = Cells(subir, 12).Value * Cells(subir, 14).Value
    
    Do Until finalEntrada = True
        If Cells(5, y).Value = fechaMovimiento Then
            Cells(subir, y).Value = valorEntrada
            Cells(z, y).Value = Cells(z, y).Value + valorEntrada
            finalEntrada = True
        End If
        y = y + 1
    Loop
    
    y = 1
    finalEntrada = False
    
    Cells(subir, 21).Value = Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value
     Cells(subir, 22).Value = Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
    * Cells(subir, 10).Value * Cells(subir, 8).Value
    
'                    Cells(subir, 22).Value = saldoFiscalCOP
                
End If 'if para determinar si es salida o si es entrada

End Sub

Sub junioJP()

subir = z - 1
Do Until Not IsEmpty(Cells(subir, 2).Value)
    If Mid(Cells(subir, 1).Value, 1, 8) = subtotal Then
        subir = subir + 1
        Exit Do
    End If
    subir = subir - 1
Loop
subir = subir + 1
encontrarCelda = True

Cells(subir, 2) = Sheets(hojaBanco).Cells(k, 2).Value
trasladoComparar = Mid(Sheets(hojaBanco).Cells(k, 2).Value, 1, 8)
trasladoComparar = UCase(trasladoComparar)
Cells(subir, 5) = Sheets(hojaBanco).Cells(k, 1).Value
i = 4
Do Until final = True
    If (Sheets("TRM BanRep").Cells(i, 1).Value = "" And Sheets("TRM BanRep").Cells(i + 1, 1).Value = "") Or Sheets("TRM BanRep").Cells(i, 1).Value = fechaMovimiento Then
        final = True
        Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Sheets("MonExtranjera").Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Exit Do
    End If
    i = i + 1
Loop

final = False

'if para determinar si es salida o si es entrada
If Not IsEmpty(Sheets(hojaBanco).Cells(k, 5)) Then 'if salida
    valorSalida = Sheets(hojaBanco).Cells(k, 5).Value
    
    For j = 27 To 200
        If Cells(z, j).Value > 0 Then
            salidaParcial = Cells(z, j).Value
            If valorSalida > salidaParcial Then
                Cells(subir, j).Value = salidaParcial * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                valorSalida = valorSalida - salidaParcial
                Cells(z, j).Value = 0
                Cells(subir, 16).Value = salidaParcial
                Cells(subir, 17).Value = salidaParcial
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                   + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                   * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                Cells(subir + 1, j).Select
                ActiveCell.EntireRow.Insert
                Cells(subir + 1, 1).Value = Cells(subir, 1).Value
                Cells(subir + 1, 2).Value = Cells(subir, 2).Value
                Cells(subir + 1, 3).Value = Cells(subir, 3).Value
                Cells(subir + 1, 5).Value = Cells(subir, 5).Value
                z = z + 1
                
                valorFinal = j
            Else
                Cells(subir, j).Value = valorSalida * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                Cells(z, j).Value = Cells(z, j) - valorSalida
                Cells(subir, 16).Value = valorSalida
                Cells(subir, 17).Value = valorSalida
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                    * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                valorSalida = 0
                
                Exit For
                
            End If
            subir = subir + 1
        End If
    Next j
    
    If valorSalida > 0 Then
    
        Cells(z, valorFinal).Value = -valorSalida
        
        MsgBox "No hay dinero suficiente para realizar la salida de " & valorSalida & " USD", _
            vbOKOnly + vbExclamation, "Alerta"
            
        k = 100
    End If
    
ElseIf Not IsEmpty(Sheets(hojaBanco).Cells(k, 4)) Then
    valorEntrada = Sheets(hojaBanco).Cells(k, 4).Value
    Cells(subir, 10).Value = 1
    Cells(subir, 11).Value = 1
    Cells(subir, 12).Value = valorEntrada
    Cells(subir, 13).Value = valorEntrada
    Cells(subir, 15).Value = Cells(subir, 12).Value * Cells(subir, 14).Value
    
    Do Until finalEntrada = True
        If Cells(5, y).Value = fechaMovimiento Then
            Cells(subir, y).Value = valorEntrada
            Cells(z, y).Value = Cells(z, y).Value + valorEntrada
            finalEntrada = True
        End If
        y = y + 1
    Loop
    
    y = 1
    finalEntrada = False
    
    Cells(subir, 21).Value = Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value
     Cells(subir, 22).Value = Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
    * Cells(subir, 10).Value * Cells(subir, 8).Value
    
'                    Cells(subir, 22).Value = saldoFiscalCOP
                
End If 'if para determinar si es salida o si es entrada

End Sub

Sub julioJP()

subir = z - 1
Do Until Not IsEmpty(Cells(subir, 2).Value)
    If Mid(Cells(subir, 1).Value, 1, 8) = subtotal Then
        subir = subir + 1
        Exit Do
    End If
    subir = subir - 1
Loop
subir = subir + 1
encontrarCelda = True

Cells(subir, 2) = Sheets(hojaBanco).Cells(k, 2).Value
trasladoComparar = Mid(Sheets(hojaBanco).Cells(k, 2).Value, 1, 8)
trasladoComparar = UCase(trasladoComparar)
Cells(subir, 5) = Sheets(hojaBanco).Cells(k, 1).Value
i = 4
Do Until final = True
    If (Sheets("TRM BanRep").Cells(i, 1).Value = "" And Sheets("TRM BanRep").Cells(i + 1, 1).Value = "") Or Sheets("TRM BanRep").Cells(i, 1).Value = fechaMovimiento Then
        final = True
        Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Sheets("MonExtranjera").Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Exit Do
    End If
    i = i + 1
Loop

final = False

'if para determinar si es salida o si es entrada
If Not IsEmpty(Sheets(hojaBanco).Cells(k, 5)) Then 'if salida
    valorSalida = Sheets(hojaBanco).Cells(k, 5).Value
    
    For j = 27 To 200
        If Cells(z, j).Value > 0 Then
            salidaParcial = Cells(z, j).Value
            If valorSalida > salidaParcial Then
                Cells(subir, j).Value = salidaParcial * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                valorSalida = valorSalida - salidaParcial
                Cells(z, j).Value = 0
                Cells(subir, 16).Value = salidaParcial
                Cells(subir, 17).Value = salidaParcial
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                   + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                   * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                   
                
                Cells(subir + 1, j).Select
                ActiveCell.EntireRow.Insert
                Cells(subir + 1, 1).Value = Cells(subir, 1).Value
                Cells(subir + 1, 2).Value = Cells(subir, 2).Value
                Cells(subir + 1, 3).Value = Cells(subir, 3).Value
                Cells(subir + 1, 5).Value = Cells(subir, 5).Value
                z = z + 1
                
                valorFinal = j
            Else
                Cells(subir, j).Value = valorSalida * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                Cells(z, j).Value = Cells(z, j) - valorSalida
                Cells(subir, 16).Value = valorSalida
                Cells(subir, 17).Value = valorSalida
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                    * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                valorSalida = 0
                
                Exit For
                
            End If
            subir = subir + 1
        End If
    Next j
    
    If valorSalida > 0 Then
    
        Cells(z, valorFinal).Value = -valorSalida
        
        MsgBox "No hay dinero suficiente para realizar la salida de " & valorSalida & " USD", _
            vbOKOnly + vbExclamation, "Alerta"
            
        k = 100
    End If
    
ElseIf Not IsEmpty(Sheets(hojaBanco).Cells(k, 4)) Then
    valorEntrada = Sheets(hojaBanco).Cells(k, 4).Value
    Cells(subir, 10).Value = 1
    Cells(subir, 11).Value = 1
    Cells(subir, 12).Value = valorEntrada
    Cells(subir, 13).Value = valorEntrada
    Cells(subir, 15).Value = Cells(subir, 12).Value * Cells(subir, 14).Value
    
    Do Until finalEntrada = True
        If Cells(5, y).Value = fechaMovimiento Then
            Cells(subir, y).Value = valorEntrada
            Cells(z, y).Value = Cells(z, y).Value + valorEntrada
            finalEntrada = True
        End If
        y = y + 1
    Loop
    
    y = 1
    finalEntrada = False
    
    Cells(subir, 21).Value = Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value
     Cells(subir, 22).Value = Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
    * Cells(subir, 10).Value * Cells(subir, 8).Value
    
'                    Cells(subir, 22).Value = saldoFiscalCOP
                
End If 'if para determinar si es salida o si es entrada

End Sub

Sub agostoJP()

subir = z - 1
Do Until Not IsEmpty(Cells(subir, 2).Value)
    If Mid(Cells(subir, 1).Value, 1, 8) = subtotal Then
        subir = subir + 1
        Exit Do
    End If
    subir = subir - 1
Loop
subir = subir + 1
encontrarCelda = True

Cells(subir, 2) = Sheets(hojaBanco).Cells(k, 2).Value
trasladoComparar = Mid(Sheets(hojaBanco).Cells(k, 2).Value, 1, 8)
trasladoComparar = UCase(trasladoComparar)
Cells(subir, 5) = Sheets(hojaBanco).Cells(k, 1).Value
i = 4
Do Until final = True
    If (Sheets("TRM BanRep").Cells(i, 1).Value = "" And Sheets("TRM BanRep").Cells(i + 1, 1).Value = "") Or Sheets("TRM BanRep").Cells(i, 1).Value = fechaMovimiento Then
        final = True
        Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Sheets("MonExtranjera").Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Exit Do
    End If
    i = i + 1
Loop

final = False

'if para determinar si es salida o si es entrada
If Not IsEmpty(Sheets(hojaBanco).Cells(k, 5)) Then 'if salida
    valorSalida = Sheets(hojaBanco).Cells(k, 5).Value
    
    For j = 27 To 200
        If Cells(z, j).Value > 0 Then
            salidaParcial = Cells(z, j).Value
            If valorSalida > salidaParcial Then
                Cells(subir, j).Value = salidaParcial * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                valorSalida = valorSalida - salidaParcial
                Cells(z, j).Value = 0
                Cells(subir, 16).Value = salidaParcial
                Cells(subir, 17).Value = salidaParcial
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                   + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                   * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                Cells(subir + 1, j).Select
                ActiveCell.EntireRow.Insert
                Cells(subir + 1, 1).Value = Cells(subir, 1).Value
                Cells(subir + 1, 2).Value = Cells(subir, 2).Value
                Cells(subir + 1, 3).Value = Cells(subir, 3).Value
                Cells(subir + 1, 5).Value = Cells(subir, 5).Value
                z = z + 1
                
                valorFinal = j
            Else
                Cells(subir, j).Value = valorSalida * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                Cells(z, j).Value = Cells(z, j) - valorSalida
                Cells(subir, 16).Value = valorSalida
                Cells(subir, 17).Value = valorSalida
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                    * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                valorSalida = 0
                
                Exit For
                
            End If
            subir = subir + 1
        End If
    Next j
    
    If valorSalida > 0 Then
    
        Cells(z, valorFinal).Value = -valorSalida
        
        MsgBox "No hay dinero suficiente para realizar la salida de " & valorSalida & " USD", _
            vbOKOnly + vbExclamation, "Alerta"
            
        k = 100
    End If
    
ElseIf Not IsEmpty(Sheets(hojaBanco).Cells(k, 4)) Then
    valorEntrada = Sheets(hojaBanco).Cells(k, 4).Value
    Cells(subir, 10).Value = 1
    Cells(subir, 11).Value = 1
    Cells(subir, 12).Value = valorEntrada
    Cells(subir, 13).Value = valorEntrada
    Cells(subir, 15).Value = Cells(subir, 12).Value * Cells(subir, 14).Value
    
    Do Until finalEntrada = True
        If Cells(5, y).Value = fechaMovimiento Then
            Cells(subir, y).Value = valorEntrada
            Cells(z, y).Value = Cells(z, y).Value + valorEntrada
            finalEntrada = True
        End If
        y = y + 1
    Loop
    
    y = 1
    finalEntrada = False
    
    Cells(subir, 21).Value = Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value
     Cells(subir, 22).Value = Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
    * Cells(subir, 10).Value * Cells(subir, 8).Value
    
'                    Cells(subir, 22).Value = saldoFiscalCOP
                
End If 'if para determinar si es salida o si es entrada

End Sub

Sub septiembreJP()

subir = z - 1
Do Until Not IsEmpty(Cells(subir, 2).Value)
    If Mid(Cells(subir, 1).Value, 1, 8) = subtotal Then
        subir = subir + 1
        Exit Do
    End If
    subir = subir - 1
Loop
subir = subir + 1
encontrarCelda = True

Cells(subir, 2) = Sheets(hojaBanco).Cells(k, 2).Value
trasladoComparar = Mid(Sheets(hojaBanco).Cells(k, 2).Value, 1, 8)
trasladoComparar = UCase(trasladoComparar)
Cells(subir, 5) = Sheets(hojaBanco).Cells(k, 1).Value
i = 4
Do Until final = True
    If (Sheets("TRM BanRep").Cells(i, 1).Value = "" And Sheets("TRM BanRep").Cells(i + 1, 1).Value = "") Or Sheets("TRM BanRep").Cells(i, 1).Value = fechaMovimiento Then
        final = True
        Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Sheets("MonExtranjera").Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Exit Do
    End If
    i = i + 1
Loop

final = False

'if para determinar si es salida o si es entrada
If Not IsEmpty(Sheets(hojaBanco).Cells(k, 5)) Then 'if salida
    valorSalida = Sheets(hojaBanco).Cells(k, 5).Value
    
    For j = 27 To 200
        If Cells(z, j).Value > 0 Then
            salidaParcial = Cells(z, j).Value
            If valorSalida > salidaParcial Then
                Cells(subir, j).Value = salidaParcial * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                valorSalida = valorSalida - salidaParcial
                Cells(z, j).Value = 0
                Cells(subir, 16).Value = salidaParcial
                Cells(subir, 17).Value = salidaParcial
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                   + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                   * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                Cells(subir + 1, j).Select
                ActiveCell.EntireRow.Insert
                Cells(subir + 1, 1).Value = Cells(subir, 1).Value
                Cells(subir + 1, 2).Value = Cells(subir, 2).Value
                Cells(subir + 1, 3).Value = Cells(subir, 3).Value
                Cells(subir + 1, 5).Value = Cells(subir, 5).Value
                z = z + 1
                
                valorFinal = j
            Else
                Cells(subir, j).Value = valorSalida * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                Cells(z, j).Value = Cells(z, j) - valorSalida
                Cells(subir, 16).Value = valorSalida
                Cells(subir, 17).Value = valorSalida
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                    * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                valorSalida = 0
                
                Exit For
                
            End If
            subir = subir + 1
        End If
    Next j
    
    If valorSalida > 0 Then
    
        Cells(z, valorFinal).Value = -valorSalida
        
        MsgBox "No hay dinero suficiente para realizar la salida de " & valorSalida & " USD", _
            vbOKOnly + vbExclamation, "Alerta"
            
        k = 100
    End If
    
ElseIf Not IsEmpty(Sheets(hojaBanco).Cells(k, 4)) Then
    valorEntrada = Sheets(hojaBanco).Cells(k, 4).Value
    Cells(subir, 10).Value = 1
    Cells(subir, 11).Value = 1
    Cells(subir, 12).Value = valorEntrada
    Cells(subir, 13).Value = valorEntrada
    Cells(subir, 15).Value = Cells(subir, 12).Value * Cells(subir, 14).Value
    
    Do Until finalEntrada = True
        If Cells(5, y).Value = fechaMovimiento Then
            Cells(subir, y).Value = valorEntrada
            Cells(z, y).Value = Cells(z, y).Value + valorEntrada
            finalEntrada = True
        End If
        y = y + 1
    Loop
    
    y = 1
    finalEntrada = False
    
    Cells(subir, 21).Value = Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value
     Cells(subir, 22).Value = Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
    * Cells(subir, 10).Value * Cells(subir, 8).Value
    
'                    Cells(subir, 22).Value = saldoFiscalCOP
                
End If 'if para determinar si es salida o si es entrada

End Sub

Sub octubreJP()

subir = z - 1
Do Until Not IsEmpty(Cells(subir, 2).Value)
    If Mid(Cells(subir, 1).Value, 1, 8) = subtotal Then
        subir = subir + 1
        Exit Do
    End If
    subir = subir - 1
Loop
subir = subir + 1
encontrarCelda = True

Cells(subir, 2) = Sheets(hojaBanco).Cells(k, 2).Value
trasladoComparar = Mid(Sheets(hojaBanco).Cells(k, 2).Value, 1, 8)
trasladoComparar = UCase(trasladoComparar)
Cells(subir, 5) = Sheets(hojaBanco).Cells(k, 1).Value
i = 4
Do Until final = True
    If (Sheets("TRM BanRep").Cells(i, 1).Value = "" And Sheets("TRM BanRep").Cells(i + 1, 1).Value = "") Or Sheets("TRM BanRep").Cells(i, 1).Value = fechaMovimiento Then
        final = True
        Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Sheets("MonExtranjera").Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Exit Do
    End If
    i = i + 1
Loop

final = False

'if para determinar si es salida o si es entrada
If Not IsEmpty(Sheets(hojaBanco).Cells(k, 5)) Then 'if salida
    valorSalida = Sheets(hojaBanco).Cells(k, 5).Value
    
    For j = 27 To 200
        If Cells(z, j).Value > 0 Then
            salidaParcial = Cells(z, j).Value
            If valorSalida > salidaParcial Then
                Cells(subir, j).Value = salidaParcial * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                valorSalida = valorSalida - salidaParcial
                Cells(z, j).Value = 0
                Cells(subir, 16).Value = salidaParcial
                Cells(subir, 17).Value = salidaParcial
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                   + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                   * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                Cells(subir + 1, j).Select
                ActiveCell.EntireRow.Insert
                Cells(subir + 1, 1).Value = Cells(subir, 1).Value
                Cells(subir + 1, 2).Value = Cells(subir, 2).Value
                Cells(subir + 1, 3).Value = Cells(subir, 3).Value
                Cells(subir + 1, 5).Value = Cells(subir, 5).Value
                z = z + 1
                
                valorFinal = j
            Else
                Cells(subir, j).Value = valorSalida * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                Cells(z, j).Value = Cells(z, j) - valorSalida
                Cells(subir, 16).Value = valorSalida
                Cells(subir, 17).Value = valorSalida
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                    * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                valorSalida = 0
                
                Exit For
                
            End If
            subir = subir + 1
        End If
    Next j
    
    If valorSalida > 0 Then
    
        Cells(z, valorFinal).Value = -valorSalida
        
        MsgBox "No hay dinero suficiente para realizar la salida de " & valorSalida & " USD", _
            vbOKOnly + vbExclamation, "Alerta"
            
        k = 100
    End If
    
ElseIf Not IsEmpty(Sheets(hojaBanco).Cells(k, 4)) Then
    valorEntrada = Sheets(hojaBanco).Cells(k, 4).Value
    Cells(subir, 10).Value = 1
    Cells(subir, 11).Value = 1
    Cells(subir, 12).Value = valorEntrada
    Cells(subir, 13).Value = valorEntrada
    Cells(subir, 15).Value = Cells(subir, 12).Value * Cells(subir, 14).Value
    
    Do Until finalEntrada = True
        If Cells(5, y).Value = fechaMovimiento Then
            Cells(subir, y).Value = valorEntrada
            Cells(z, y).Value = Cells(z, y).Value + valorEntrada
            finalEntrada = True
        End If
        y = y + 1
    Loop
    
    y = 1
    finalEntrada = False
    
    Cells(subir, 21).Value = Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value
     Cells(subir, 22).Value = Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
    * Cells(subir, 10).Value * Cells(subir, 8).Value
    
'                    Cells(subir, 22).Value = saldoFiscalCOP
                
End If 'if para determinar si es salida o si es entrada

End Sub

Sub noviembreJP()

subir = z - 1
Do Until Not IsEmpty(Cells(subir, 2).Value)
    If Mid(Cells(subir, 1).Value, 1, 8) = subtotal Then
        subir = subir + 1
        Exit Do
    End If
    subir = subir - 1
Loop
subir = subir + 1
encontrarCelda = True

Cells(subir, 2) = Sheets(hojaBanco).Cells(k, 2).Value
trasladoComparar = Mid(Sheets(hojaBanco).Cells(k, 2).Value, 1, 8)
trasladoComparar = UCase(trasladoComparar)
Cells(subir, 5) = Sheets(hojaBanco).Cells(k, 1).Value
i = 4
Do Until final = True
    If (Sheets("TRM BanRep").Cells(i, 1).Value = "" And Sheets("TRM BanRep").Cells(i + 1, 1).Value = "") Or Sheets("TRM BanRep").Cells(i, 1).Value = fechaMovimiento Then
        final = True
        Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Sheets("MonExtranjera").Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Exit Do
    End If
    i = i + 1
Loop

final = False

'if para determinar si es salida o si es entrada
If Not IsEmpty(Sheets(hojaBanco).Cells(k, 5)) Then 'if salida
    valorSalida = Sheets(hojaBanco).Cells(k, 5).Value
    
    For j = 27 To 200
        If Cells(z, j).Value > 0 Then
            salidaParcial = Cells(z, j).Value
            If valorSalida > salidaParcial Then
                Cells(subir, j).Value = salidaParcial * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                valorSalida = valorSalida - salidaParcial
                Cells(z, j).Value = 0
                Cells(subir, 16).Value = salidaParcial
                Cells(subir, 17).Value = salidaParcial
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                   + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                   * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                Cells(subir + 1, j).Select
                ActiveCell.EntireRow.Insert
                Cells(subir + 1, 1).Value = Cells(subir, 1).Value
                Cells(subir + 1, 2).Value = Cells(subir, 2).Value
                Cells(subir + 1, 3).Value = Cells(subir, 3).Value
                Cells(subir + 1, 5).Value = Cells(subir, 5).Value
                z = z + 1
                
                valorFinal = j
            Else
                Cells(subir, j).Value = valorSalida * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                Cells(z, j).Value = Cells(z, j) - valorSalida
                Cells(subir, 16).Value = valorSalida
                Cells(subir, 17).Value = valorSalida
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                    * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                valorSalida = 0
                
                Exit For
                
            End If
            subir = subir + 1
        End If
    Next j
    
    If valorSalida > 0 Then
    
        Cells(z, valorFinal).Value = -valorSalida
        
        MsgBox "No hay dinero suficiente para realizar la salida de " & valorSalida & " USD", _
            vbOKOnly + vbExclamation, "Alerta"
            
        k = 100
    End If
    
ElseIf Not IsEmpty(Sheets(hojaBanco).Cells(k, 4)) Then
    valorEntrada = Sheets(hojaBanco).Cells(k, 4).Value
    Cells(subir, 10).Value = 1
    Cells(subir, 11).Value = 1
    Cells(subir, 12).Value = valorEntrada
    Cells(subir, 13).Value = valorEntrada
    Cells(subir, 15).Value = Cells(subir, 12).Value * Cells(subir, 14).Value
    
    Do Until finalEntrada = True
        If Cells(5, y).Value = fechaMovimiento Then
            Cells(subir, y).Value = valorEntrada
            Cells(z, y).Value = Cells(z, y).Value + valorEntrada
            finalEntrada = True
        End If
        y = y + 1
    Loop
    
    y = 1
    finalEntrada = False
    
    Cells(subir, 21).Value = Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value
     Cells(subir, 22).Value = Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
    * Cells(subir, 10).Value * Cells(subir, 8).Value
    
'                    Cells(subir, 22).Value = saldoFiscalCOP
                
End If 'if para determinar si es salida o si es entrada

End Sub

Sub diciembreJP()

subir = z - 1
Do Until Not IsEmpty(Cells(subir, 2).Value)
    If Mid(Cells(subir, 1).Value, 1, 8) = subtotal Then
        subir = subir + 1
        Exit Do
    End If
    subir = subir - 1
Loop
subir = subir + 1
encontrarCelda = True

Cells(subir, 2) = Sheets(hojaBanco).Cells(k, 2).Value
trasladoComparar = Mid(Sheets(hojaBanco).Cells(k, 2).Value, 1, 8)
trasladoComparar = UCase(trasladoComparar)
Cells(subir, 5) = Sheets(hojaBanco).Cells(k, 1).Value
i = 4
Do Until final = True
    If (Sheets("TRM BanRep").Cells(i, 1).Value = "" And Sheets("TRM BanRep").Cells(i + 1, 1).Value = "") Or Sheets("TRM BanRep").Cells(i, 1).Value = fechaMovimiento Then
        final = True
        Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Sheets("MonExtranjera").Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
        Exit Do
    End If
    i = i + 1
Loop

final = False

'if para determinar si es salida o si es entrada
If Not IsEmpty(Sheets(hojaBanco).Cells(k, 5)) Then 'if salida
    valorSalida = Sheets(hojaBanco).Cells(k, 5).Value
    
    For j = 27 To 200
        If Cells(z, j).Value > 0 Then
            salidaParcial = Cells(z, j).Value
            If valorSalida > salidaParcial Then
                Cells(subir, j).Value = salidaParcial * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                valorSalida = valorSalida - salidaParcial
                Cells(z, j).Value = 0
                Cells(subir, 16).Value = salidaParcial
                Cells(subir, 17).Value = salidaParcial
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                   + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                   * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                Cells(subir + 1, j).Select
                ActiveCell.EntireRow.Insert
                Cells(subir + 1, 1).Value = Cells(subir, 1).Value
                Cells(subir + 1, 2).Value = Cells(subir, 2).Value
                Cells(subir + 1, 3).Value = Cells(subir, 3).Value
                Cells(subir + 1, 5).Value = Cells(subir, 5).Value
                z = z + 1
                
                valorFinal = j
            Else
                Cells(subir, j).Value = valorSalida * -1
                Cells(subir, 10).Value = 1
                Cells(subir, 11).Value = 1
                Cells(z, j).Value = Cells(z, j) - valorSalida
                Cells(subir, 16).Value = valorSalida
                Cells(subir, 17).Value = valorSalida
                Cells(subir, 18).Value = Cells(6, j).Value
                
                If trasladoComparar = traslado Then
                    Cells(subir, 19).Value = Cells(subir, 18).Value
                Else
                    Cells(subir, 19).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                End If
                
                Cells(subir, 20).Value = Cells(subir, 19).Value * Cells(subir, 16).Value
                Sheets("MonExtranjera").Cells(subir, 14).Value = Sheets("TRM BanRep").Cells(i, 7).Value
                
                'Saldo Fiscal Mon/Ex
                Cells(subir, 21).Value = (Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value) * (-1)
'
                'Saldo Fiscal COP
                Cells(subir, 22).Value = (Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
                    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
                    * Cells(subir, 10).Value * Cells(subir, 8).Value) * (-1)
                
                valorSalida = 0
                
                Exit For
                
            End If
            subir = subir + 1
        End If
    Next j
    
    If valorSalida > 0 Then
    
        Cells(z, valorFinal).Value = -valorSalida
        
        MsgBox "No hay dinero suficiente para realizar la salida de " & valorSalida & " USD", _
            vbOKOnly + vbExclamation, "Alerta"
            
        k = 100
    End If
    
ElseIf Not IsEmpty(Sheets(hojaBanco).Cells(k, 4)) Then
    valorEntrada = Sheets(hojaBanco).Cells(k, 4).Value
    Cells(subir, 10).Value = 1
    Cells(subir, 11).Value = 1
    Cells(subir, 12).Value = valorEntrada
    Cells(subir, 13).Value = valorEntrada
    Cells(subir, 15).Value = Cells(subir, 12).Value * Cells(subir, 14).Value
    
    Do Until finalEntrada = True
        If Cells(5, y).Value = fechaMovimiento Then
            Cells(subir, y).Value = valorEntrada
            Cells(z, y).Value = Cells(z, y).Value + valorEntrada
            finalEntrada = True
        End If
        y = y + 1
    Loop
    
    y = 1
    finalEntrada = False
    
    Cells(subir, 21).Value = Cells(subir, 12).Value + Cells(subir, 16).Value + Cells(subir, 6).Value
     Cells(subir, 22).Value = Cells(subir, 16).Value * Cells(subir, 10).Value * Cells(subir, 18).Value _
    + Cells(subir, 12).Value * Cells(subir, 10).Value * Cells(subir, 14).Value + Cells(subir, 6).Value _
    * Cells(subir, 10).Value * Cells(subir, 8).Value
    
'                    Cells(subir, 22).Value = saldoFiscalCOP
                
End If 'if para determinar si es salida o si es entrada

End Sub


