Attribute VB_Name = "Módulo1"
Option Explicit

Dim sociedad, asterisco, fecha, mes As String
Dim i, j, sociedadCodigo, sociedadCodigoDigito, sociedadComparar As Integer
Dim valor As Currency
Dim final As Boolean

Dim porcentajeVariacion, variacion, mesAnterior, valor5905 As Long
Dim finalEncontrado As Boolean
Dim k, z, sociedadCodigoComparar, cont As Integer
Dim sociedadCompararFaltante, hojaBalance, hojaResultado, hojaDatos As String

Sub Balance()

k = 2
j = 2
i = 1
final = False
finalEncontrado = False


Do Until final = True
    j = 2
    asterisco = Sheets(hojaDatos).Cells(i, 2).Value
    finalEncontrado = False
    If asterisco = "**" Then
        final = True
    ElseIf asterisco = "*" Then
        sociedad = Sheets(hojaDatos).Cells(i, 4).Value
        sociedadCodigo = Mid(sociedad, 12, 4)
        sociedadCodigoDigito = Mid(sociedadCodigo, 1, 1)
        If sociedadCodigoDigito = "1" Or sociedadCodigoDigito = "2" Or sociedadCodigoDigito = "3" Then
            Do Until IsEmpty(Sheets(hojaBalance).Cells(j, 1))
                sociedadComparar = Sheets(hojaBalance).Cells(j, 1).Value
                If sociedadComparar = sociedadCodigo Then
                    valor = Sheets(hojaDatos).Cells(i, 20).Value
                    
                    If mes = "05" Then
                        Sheets(hojaBalance).Cells(j, 8).Value = valor / 1000000
                        Sheets(hojaBalance).Cells(j, 16).Value = Sheets(hojaBalance).Cells(j, 8).Value - _
                            Sheets(hojaBalance).Cells(j, 7).Value
                        mesAnterior = Sheets(hojaBalance).Cells(j, 7).Value
                        If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 7))) And Sheets(hojaBalance).Cells(j, 16).Value > 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 7))) And Sheets(hojaBalance).Cells(j, 16).Value < 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = -1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 7))) And Sheets(hojaBalance).Cells(j, 16).Value = 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 0
                        Else
                            Sheets(hojaBalance).Cells(j, 17).Value = Sheets(hojaBalance).Cells(j, 16).Value / Sheets(hojaBalance).Cells(j, 7).Value
                        End If
                        
                    ElseIf mes = "01" Then
                        Sheets(hojaBalance).Cells(j, 4).Value = valor / 1000000
                        Sheets(hojaBalance).Cells(j, 16).Value = Sheets(hojaBalance).Cells(j, 4).Value - _
                            Sheets(hojaBalance).Cells(j, 3).Value
                        mesAnterior = Sheets(hojaBalance).Cells(j, 3).Value
                        If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 3))) And Sheets(hojaBalance).Cells(j, 16).Value > 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 3))) And Sheets(hojaBalance).Cells(j, 16).Value < 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = -1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 3))) And Sheets(hojaBalance).Cells(j, 16).Value = 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 0
                        Else
                            Sheets(hojaBalance).Cells(j, 17).Value = Sheets(hojaBalance).Cells(j, 16).Value / Sheets(hojaBalance).Cells(j, 3).Value
                        End If
                        
                    ElseIf mes = "02" Then
                        Sheets(hojaBalance).Cells(j, 5).Value = valor / 1000000
                        Sheets(hojaBalance).Cells(j, 16).Value = Sheets(hojaBalance).Cells(j, 5).Value - _
                            Sheets(hojaBalance).Cells(j, 4).Value
                        mesAnterior = Sheets(hojaBalance).Cells(j, 4).Value
                        If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 4))) And Sheets(hojaBalance).Cells(j, 16).Value > 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 4))) And Sheets(hojaBalance).Cells(j, 16).Value < 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = -1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 4))) And Sheets(hojaBalance).Cells(j, 16).Value = 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 0
                        Else
                            Sheets(hojaBalance).Cells(j, 17).Value = Sheets(hojaBalance).Cells(j, 16).Value / Sheets(hojaBalance).Cells(j, 4).Value
                        End If
                        
                    ElseIf mes = "03" Then
                        Sheets(hojaBalance).Cells(j, 6).Value = valor / 1000000
                        Sheets(hojaBalance).Cells(j, 16).Value = Sheets(hojaBalance).Cells(j, 6).Value - _
                            Sheets(hojaBalance).Cells(j, 5).Value
                        mesAnterior = Sheets(hojaBalance).Cells(j, 5).Value
                        If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 5))) And Sheets(hojaBalance).Cells(j, 16).Value > 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 5))) And Sheets(hojaBalance).Cells(j, 16).Value < 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = -1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 5))) And Sheets(hojaBalance).Cells(j, 16).Value = 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 0
                        Else
                            Sheets(hojaBalance).Cells(j, 17).Value = Sheets(hojaBalance).Cells(j, 16).Value / Sheets(hojaBalance).Cells(j, 5).Value
                        End If
                        
                    ElseIf mes = "04" Then
                        Sheets(hojaBalance).Cells(j, 7).Value = valor / 1000000
                        Sheets(hojaBalance).Cells(j, 16).Value = Sheets(hojaBalance).Cells(j, 7).Value - _
                            Sheets(hojaBalance).Cells(j, 6).Value
                        mesAnterior = Sheets(hojaBalance).Cells(j, 6).Value
                        If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 6))) And Sheets(hojaBalance).Cells(j, 16).Value > 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 6))) And Sheets(hojaBalance).Cells(j, 16).Value < 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = -1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 6))) And Sheets(hojaBalance).Cells(j, 16).Value = 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 0
                        Else
                            Sheets(hojaBalance).Cells(j, 17).Value = Sheets(hojaBalance).Cells(j, 16).Value / Sheets(hojaBalance).Cells(j, 6).Value
                        End If
                        
                    ElseIf mes = "06" Then
                        Sheets(hojaBalance).Cells(j, 9).Value = valor / 1000000
                        Sheets(hojaBalance).Cells(j, 16).Value = Sheets(hojaBalance).Cells(j, 9).Value - _
                            Sheets(hojaBalance).Cells(j, 8).Value
                        mesAnterior = Sheets(hojaBalance).Cells(j, 8).Value
                        If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 8))) And Sheets(hojaBalance).Cells(j, 16).Value > 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 8))) And Sheets(hojaBalance).Cells(j, 16).Value < 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = -1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 8))) And Sheets(hojaBalance).Cells(j, 16).Value = 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 0
                        Else
                            Sheets(hojaBalance).Cells(j, 17).Value = Sheets(hojaBalance).Cells(j, 16).Value / Sheets(hojaBalance).Cells(j, 8).Value
                        End If
                        
                    ElseIf mes = "07" Then
                        Sheets(hojaBalance).Cells(j, 10).Value = valor / 1000000
                        Sheets(hojaBalance).Cells(j, 16).Value = Sheets(hojaBalance).Cells(j, 10).Value - _
                            Sheets(hojaBalance).Cells(j, 9).Value
                        mesAnterior = Sheets(hojaBalance).Cells(j, 9).Value
                        If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 9))) And Sheets(hojaBalance).Cells(j, 16).Value > 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 9))) And Sheets(hojaBalance).Cells(j, 16).Value < 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = -1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 9))) And Sheets(hojaBalance).Cells(j, 16).Value = 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 0
                        Else
                            Sheets(hojaBalance).Cells(j, 17).Value = Sheets(hojaBalance).Cells(j, 16).Value / Sheets(hojaBalance).Cells(j, 9).Value
                        End If
                        
                    ElseIf mes = "08" Then
                        Sheets(hojaBalance).Cells(j, 11).Value = valor / 1000000
                        Sheets(hojaBalance).Cells(j, 16).Value = Sheets(hojaBalance).Cells(j, 11).Value - _
                            Sheets(hojaBalance).Cells(j, 10).Value
                        mesAnterior = Sheets(hojaBalance).Cells(j, 10).Value
                        If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 10))) And Sheets(hojaBalance).Cells(j, 16).Value > 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 10))) And Sheets(hojaBalance).Cells(j, 16).Value < 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = -1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 10))) And Sheets(hojaBalance).Cells(j, 16).Value = 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 0
                        Else
                            Sheets(hojaBalance).Cells(j, 17).Value = Sheets(hojaBalance).Cells(j, 16).Value / Sheets(hojaBalance).Cells(j, 10).Value
                        End If
                        
                    ElseIf mes = "09" Then
                        Sheets(hojaBalance).Cells(j, 12).Value = valor / 1000000
                        Sheets(hojaBalance).Cells(j, 16).Value = Sheets(hojaBalance).Cells(j, 12).Value - _
                            Sheets(hojaBalance).Cells(j, 11).Value
                        mesAnterior = Sheets(hojaBalance).Cells(j, 11).Value
                        If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 11))) And Sheets(hojaBalance).Cells(j, 16).Value > 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 11))) And Sheets(hojaBalance).Cells(j, 16).Value < 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = -1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 11))) And Sheets(hojaBalance).Cells(j, 16).Value = 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 0
                        Else
                            Sheets(hojaBalance).Cells(j, 17).Value = Sheets(hojaBalance).Cells(j, 16).Value / Sheets(hojaBalance).Cells(j, 11).Value
                        End If
                        
                    ElseIf mes = "10" Then
                        Sheets(hojaBalance).Cells(j, 13).Value = valor / 1000000
                        Sheets(hojaBalance).Cells(j, 16).Value = Sheets(hojaBalance).Cells(j, 13).Value - _
                            Sheets(hojaBalance).Cells(j, 12).Value
                        mesAnterior = Sheets(hojaBalance).Cells(j, 12).Value
                        If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 12))) And Sheets(hojaBalance).Cells(j, 16).Value > 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 12))) And Sheets(hojaBalance).Cells(j, 16).Value < 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = -1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 12))) And Sheets(hojaBalance).Cells(j, 16).Value = 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 0
                        Else
                            Sheets(hojaBalance).Cells(j, 17).Value = Sheets(hojaBalance).Cells(j, 16).Value / Sheets(hojaBalance).Cells(j, 12).Value
                        End If
                        
                    ElseIf mes = "11" Then
                        Sheets(hojaBalance).Cells(j, 14).Value = valor / 1000000
                        Sheets(hojaBalance).Cells(j, 16).Value = Sheets(hojaBalance).Cells(j, 14).Value - _
                            Sheets(hojaBalance).Cells(j, 13).Value
                        mesAnterior = Sheets(hojaBalance).Cells(j, 13).Value
                        If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 13))) And Sheets(hojaBalance).Cells(j, 16).Value > 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 13))) And Sheets(hojaBalance).Cells(j, 16).Value < 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = -1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 13))) And Sheets(hojaBalance).Cells(j, 16).Value = 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 0
                        Else
                            Sheets(hojaBalance).Cells(j, 17).Value = Sheets(hojaBalance).Cells(j, 16).Value / Sheets(hojaBalance).Cells(j, 13).Value
                        End If
                    
                    ElseIf mes = "12" Then
                        Sheets(hojaBalance).Cells(j, 15).Value = valor / 1000000
                        Sheets(hojaBalance).Cells(j, 16).Value = Sheets(hojaBalance).Cells(j, 15).Value - _
                            Sheets(hojaBalance).Cells(j, 14).Value
                        mesAnterior = Sheets(hojaBalance).Cells(j, 14).Value
                        If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 14))) And Sheets(hojaBalance).Cells(j, 16).Value > 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 14))) And Sheets(hojaBalance).Cells(j, 16).Value < 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = -1
                        ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 14))) And Sheets(hojaBalance).Cells(j, 16).Value = 0 Then
                            Sheets(hojaBalance).Cells(j, 17).Value = 0
                        Else
                            Sheets(hojaBalance).Cells(j, 17).Value = Sheets(hojaBalance).Cells(j, 16).Value / Sheets(hojaBalance).Cells(j, 14).Value
                        End If
                    
                    End If
                    
                    finalEncontrado = True
                    Exit Do
                End If
                j = j + 1
            Loop
            
            If finalEncontrado = False Then
                Sheets(hojaBalance).Cells(j, 8).EntireRow.Insert
                Sheets(hojaBalance).Cells(j, 1).Value = sociedadCodigo
                valor = Sheets(hojaDatos).Cells(i, 20).Value
                If mes = "05" Then
                    Sheets(hojaBalance).Cells(j, 8).Value = valor / 1000000
                    Sheets(hojaBalance).Cells(j, 16).Value = Sheets(hojaBalance).Cells(j, 8).Value - _
                        Sheets(hojaBalance).Cells(j, 7).Value
                     mesAnterior = Sheets(hojaBalance).Cells(j, 7).Value
                    If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 7))) And Sheets(hojaBalance).Cells(j, 16).Value > 0 Then
                        Sheets(hojaBalance).Cells(j, 17).Value = 1
                    ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 7))) And Sheets(hojaBalance).Cells(j, 16).Value < 0 Then
                        Sheets(hojaBalance).Cells(j, 17).Value = -1
                    ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(j, 7))) And Sheets(hojaBalance).Cells(j, 16).Value = 0 Then
                        Sheets(hojaBalance).Cells(j, 17).Value = 0
                    Else
                        Sheets(hojaBalance).Cells(j, 17).Value = Sheets(hojaBalance).Cells(j, 16).Value / Sheets(hojaBalance).Cells(j, 7).Value
                    End If
                End If
            End If
            
        End If
    End If
    i = i + 1
    
Loop

j = 2

Do Until Sheets(hojaResultado).Cells(j, 2).Value = "TOTAL"
    j = j + 1
Loop
valor = Sheets(hojaResultado).Cells(j, 19).Value

Do Until IsEmpty(Sheets(hojaBalance).Cells(k, 1))
    If Sheets(hojaBalance).Cells(k, 1).Value = 3230 Then
        cont = cont + 1
        If cont = 2 Then
            valor5905 = Sheets(hojaBalance).Cells(k, 3).Value
            
            If mes = "01" Then
                Sheets(hojaBalance).Cells(k, 4).Value = valor
                Sheets(hojaBalance).Cells(k, 16).Value = Sheets(hojaBalance).Cells(k, 4).Value - _
                    Sheets(hojaBalance).Cells(k, 3).Value
                mesAnterior = Sheets(hojaBalance).Cells(k, 3).Value
                If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 3))) And Sheets(hojaBalance).Cells(k, 16).Value > 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 3))) And Sheets(hojaBalance).Cells(k, 16).Value < 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = -1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 3))) And Sheets(hojaBalance).Cells(k, 16).Value = 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 0
                Else
                    Sheets(hojaBalance).Cells(k, 17).Value = Sheets(hojaBalance).Cells(k, 16).Value / Sheets(hojaBalance).Cells(k, 3).Value
                End If
                
            ElseIf mes = "02" Then
                Sheets(hojaBalance).Cells(k, 5).Value = valor
                Sheets(hojaBalance).Cells(k, 16).Value = Sheets(hojaBalance).Cells(k, 5).Value - _
                    Sheets(hojaBalance).Cells(k, 4).Value
                mesAnterior = Sheets(hojaBalance).Cells(k, 4).Value
                If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 4))) And Sheets(hojaBalance).Cells(k, 16).Value > 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 4))) And Sheets(hojaBalance).Cells(k, 16).Value < 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = -1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 4))) And Sheets(hojaBalance).Cells(k, 16).Value = 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 0
                Else
                    Sheets(hojaBalance).Cells(k, 17).Value = Sheets(hojaBalance).Cells(k, 16).Value / Sheets(hojaBalance).Cells(k, 4).Value
                End If
                
            ElseIf mes = "03" Then
                Sheets(hojaBalance).Cells(k, 6).Value = valor
                Sheets(hojaBalance).Cells(k, 16).Value = Sheets(hojaBalance).Cells(k, 6).Value - _
                    Sheets(hojaBalance).Cells(k, 5).Value
                mesAnterior = Sheets(hojaBalance).Cells(k, 5).Value
                If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 5))) And Sheets(hojaBalance).Cells(k, 16).Value > 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 5))) And Sheets(hojaBalance).Cells(k, 16).Value < 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = -1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 5))) And Sheets(hojaBalance).Cells(k, 16).Value = 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 0
                Else
                    Sheets(hojaBalance).Cells(k, 17).Value = Sheets(hojaBalance).Cells(k, 16).Value / Sheets(hojaBalance).Cells(k, 5).Value
                End If
                
            ElseIf mes = "04" Then
                Sheets(hojaBalance).Cells(k, 7).Value = valor
                 Sheets(hojaBalance).Cells(k, 16).Value = Sheets(hojaBalance).Cells(k, 7).Value - _
                    Sheets(hojaBalance).Cells(k, 6).Value
                mesAnterior = Sheets(hojaBalance).Cells(k, 6).Value
                If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 6))) And Sheets(hojaBalance).Cells(k, 16).Value > 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 6))) And Sheets(hojaBalance).Cells(k, 16).Value < 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = -1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 6))) And Sheets(hojaBalance).Cells(k, 16).Value = 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 0
                Else
                    Sheets(hojaBalance).Cells(k, 17).Value = Sheets(hojaBalance).Cells(k, 16).Value / Sheets(hojaBalance).Cells(k, 6).Value
                End If
                
            ElseIf mes = "05" Then
                Sheets(hojaBalance).Cells(k, 8).Value = valor
                Sheets(hojaBalance).Cells(k, 16).Value = Sheets(hojaBalance).Cells(k, 8).Value - _
                    Sheets(hojaBalance).Cells(k, 7).Value
                mesAnterior = Sheets(hojaBalance).Cells(k, 7).Value
                If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 7))) And Sheets(hojaBalance).Cells(k, 16).Value > 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 7))) And Sheets(hojaBalance).Cells(k, 16).Value < 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = -1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 7))) And Sheets(hojaBalance).Cells(k, 16).Value = 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 0
                Else
                    Sheets(hojaBalance).Cells(k, 17).Value = Sheets(hojaBalance).Cells(k, 16).Value / Sheets(hojaBalance).Cells(k, 7).Value
                End If
                
            ElseIf mes = "06" Then
                Sheets(hojaBalance).Cells(k, 9).Value = valor
                Sheets(hojaBalance).Cells(k, 16).Value = Sheets(hojaBalance).Cells(k, 9).Value - _
                    Sheets(hojaBalance).Cells(k, 8).Value
                mesAnterior = Sheets(hojaBalance).Cells(k, 8).Value
                If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 8))) And Sheets(hojaBalance).Cells(k, 16).Value > 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 8))) And Sheets(hojaBalance).Cells(k, 16).Value < 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = -1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 8))) And Sheets(hojaBalance).Cells(k, 16).Value = 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 0
                Else
                    Sheets(hojaBalance).Cells(k, 17).Value = Sheets(hojaBalance).Cells(k, 16).Value / Sheets(hojaBalance).Cells(k, 8).Value
                End If
                
            ElseIf mes = "07" Then
                Sheets(hojaBalance).Cells(k, 10).Value = valor
                Sheets(hojaBalance).Cells(k, 16).Value = Sheets(hojaBalance).Cells(k, 10).Value - _
                    Sheets(hojaBalance).Cells(k, 9).Value
                mesAnterior = Sheets(hojaBalance).Cells(k, 9).Value
                If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 9))) And Sheets(hojaBalance).Cells(k, 16).Value > 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 9))) And Sheets(hojaBalance).Cells(k, 16).Value < 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = -1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 9))) And Sheets(hojaBalance).Cells(k, 16).Value = 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 0
                Else
                    Sheets(hojaBalance).Cells(k, 17).Value = Sheets(hojaBalance).Cells(k, 16).Value / Sheets(hojaBalance).Cells(k, 9).Value
                End If
                
            ElseIf mes = "08" Then
                Sheets(hojaBalance).Cells(k, 11).Value = valor
                Sheets(hojaBalance).Cells(k, 16).Value = Sheets(hojaBalance).Cells(k, 11).Value - _
                    Sheets(hojaBalance).Cells(k, 10).Value
                mesAnterior = Sheets(hojaBalance).Cells(k, 10).Value
                If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 10))) And Sheets(hojaBalance).Cells(k, 16).Value > 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 10))) And Sheets(hojaBalance).Cells(k, 16).Value < 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = -1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 10))) And Sheets(hojaBalance).Cells(k, 16).Value = 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 0
                Else
                    Sheets(hojaBalance).Cells(k, 17).Value = Sheets(hojaBalance).Cells(k, 16).Value / Sheets(hojaBalance).Cells(k, 10).Value
                End If
                
            ElseIf mes = "09" Then
                Sheets(hojaBalance).Cells(k, 12).Value = valor
                Sheets(hojaBalance).Cells(k, 16).Value = Sheets(hojaBalance).Cells(k, 12).Value - _
                    Sheets(hojaBalance).Cells(k, 11).Value
                mesAnterior = Sheets(hojaBalance).Cells(k, 11).Value
                If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 11))) And Sheets(hojaBalance).Cells(k, 16).Value > 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 11))) And Sheets(hojaBalance).Cells(k, 16).Value < 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = -1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 11))) And Sheets(hojaBalance).Cells(k, 16).Value = 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 0
                Else
                    Sheets(hojaBalance).Cells(k, 17).Value = Sheets(hojaBalance).Cells(k, 16).Value / Sheets(hojaBalance).Cells(k, 11).Value
                End If
                
            ElseIf mes = "10" Then
                Sheets(hojaBalance).Cells(k, 13).Value = valor
                Sheets(hojaBalance).Cells(k, 16).Value = Sheets(hojaBalance).Cells(k, 13).Value - _
                    Sheets(hojaBalance).Cells(k, 12).Value
                mesAnterior = Sheets(hojaBalance).Cells(k, 12).Value
                If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 12))) And Sheets(hojaBalance).Cells(k, 16).Value > 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 12))) And Sheets(hojaBalance).Cells(k, 16).Value < 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = -1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 12))) And Sheets(hojaBalance).Cells(k, 16).Value = 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 0
                Else
                    Sheets(hojaBalance).Cells(k, 17).Value = Sheets(hojaBalance).Cells(k, 16).Value / Sheets(hojaBalance).Cells(k, 12).Value
                End If
                
            ElseIf mes = "11" Then
                Sheets(hojaBalance).Cells(k, 14).Value = valor
                Sheets(hojaBalance).Cells(k, 16).Value = Sheets(hojaBalance).Cells(k, 14).Value - _
                    Sheets(hojaBalance).Cells(k, 13).Value
                mesAnterior = Sheets(hojaBalance).Cells(k, 13).Value
                If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 13))) And Sheets(hojaBalance).Cells(k, 16).Value > 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 13))) And Sheets(hojaBalance).Cells(k, 16).Value < 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = -1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 13))) And Sheets(hojaBalance).Cells(k, 16).Value = 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 0
                Else
                    Sheets(hojaBalance).Cells(k, 17).Value = Sheets(hojaBalance).Cells(k, 16).Value / Sheets(hojaBalance).Cells(k, 13).Value
                End If
                
            ElseIf mes = "12" Then
                Sheets(hojaBalance).Cells(k, 15).Value = valor
                Sheets(hojaBalance).Cells(k, 16).Value = Sheets(hojaBalance).Cells(k, 15).Value - _
                    Sheets(hojaBalance).Cells(k, 14).Value
                mesAnterior = Sheets(hojaBalance).Cells(k, 14).Value
                If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 14))) And Sheets(hojaBalance).Cells(k, 16).Value > 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 14))) And Sheets(hojaBalance).Cells(k, 16).Value < 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = -1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 14))) And Sheets(hojaBalance).Cells(k, 16).Value = 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 0
                Else
                    Sheets(hojaBalance).Cells(k, 17).Value = Sheets(hojaBalance).Cells(k, 16).Value / Sheets(hojaBalance).Cells(k, 14).Value
                End If
            End If
        End If
    End If
    
    If Sheets(hojaBalance).Cells(k, 1).Value = 5905 Then
        If mes = "01" Then
            Sheets(hojaBalance).Cells(k, 4).Value = valor5905
            Sheets(hojaBalance).Cells(k, 16).Value = Sheets(hojaBalance).Cells(k, 4).Value - _
                Sheets(hojaBalance).Cells(k, 3).Value
            mesAnterior = Sheets(hojaBalance).Cells(k, 3).Value
            If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 3))) And Sheets(hojaBalance).Cells(k, 16).Value > 0 Then
                Sheets(hojaBalance).Cells(k, 17).Value = 1
            ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 3))) And Sheets(hojaBalance).Cells(k, 16).Value < 0 Then
                Sheets(hojaBalance).Cells(k, 17).Value = -1
            ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 3))) And Sheets(hojaBalance).Cells(k, 16).Value = 0 Then
                Sheets(hojaBalance).Cells(k, 17).Value = 0
            Else
                Sheets(hojaBalance).Cells(k, 17).Value = Sheets(hojaBalance).Cells(k, 16).Value / Sheets(hojaBalance).Cells(k, 3).Value
            End If
                
        ElseIf mes = "02" Then
            Sheets(hojaBalance).Cells(k, 5).Value = valor5905
            Sheets(hojaBalance).Cells(k, 16).Value = Sheets(hojaBalance).Cells(k, 5).Value - _
                    Sheets(hojaBalance).Cells(k, 4).Value
                mesAnterior = Sheets(hojaBalance).Cells(k, 4).Value
                If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 4))) And Sheets(hojaBalance).Cells(k, 16).Value > 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 4))) And Sheets(hojaBalance).Cells(k, 16).Value < 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = -1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 4))) And Sheets(hojaBalance).Cells(k, 16).Value = 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 0
                Else
                    Sheets(hojaBalance).Cells(k, 17).Value = Sheets(hojaBalance).Cells(k, 16).Value / Sheets(hojaBalance).Cells(k, 4).Value
                End If
                
        ElseIf mes = "03" Then
            Sheets(hojaBalance).Cells(k, 6).Value = valor5905
            Sheets(hojaBalance).Cells(k, 16).Value = Sheets(hojaBalance).Cells(k, 6).Value - _
                    Sheets(hojaBalance).Cells(k, 5).Value
                mesAnterior = Sheets(hojaBalance).Cells(k, 5).Value
                If (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 5))) And Sheets(hojaBalance).Cells(k, 16).Value > 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 5))) And Sheets(hojaBalance).Cells(k, 16).Value < 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = -1
                ElseIf (mesAnterior = 0 Or IsEmpty(Sheets(hojaBalance).Cells(k, 5))) And Sheets(hojaBalance).Cells(k, 16).Value = 0 Then
                    Sheets(hojaBalance).Cells(k, 17).Value = 0
                Else
                    Sheets(hojaBalance).Cells(k, 17).Value = Sheets(hojaBalance).Cells(k, 16).Value / Sheets(hojaBalance).Cells(k, 5).Value
                End If
        Else
            Sheets(hojaBalance).Cells(k, 16).Value = 0
            Sheets(hojaBalance).Cells(k, 17).Value = 0
        End If
        
    End If
    
    If (Sheets(hojaBalance).Cells(k, 17).Value > 10 Or Sheets(hojaBalance).Cells(k, 17).Value < -10) _
        And (Sheets(hojaBalance).Cells(k, 16).Value > 500 Or Sheets(hojaBalance).Cells(k, 16).Value < -500) Then
            Sheets(hojaBalance).Range("Q" & k).Interior.Color = RGB(252, 213, 180)
    Else
        Sheets(hojaBalance).Range("Q" & k).Interior.Color = RGB(255, 255, 255)
    End If
    k = k + 1
Loop

End Sub

Sub resultado()

hojaDatos = Sheets(1).Name
hojaBalance = Sheets(2).Name
hojaResultado = Sheets(3).Name

Sheets(hojaResultado).Activate

Sheets(hojaResultado).Columns("D:D").EntireColumn.AutoFit
Sheets(hojaResultado).Columns("C:C").EntireColumn.AutoFit
Sheets(hojaResultado).Columns("O:O").EntireColumn.AutoFit
Sheets(hojaResultado).Columns("N:N").EntireColumn.AutoFit
Sheets(hojaResultado).Columns("M:M").EntireColumn.AutoFit
Sheets(hojaResultado).Columns("L:L").EntireColumn.AutoFit
Sheets(hojaResultado).Columns("K:K").EntireColumn.AutoFit
Sheets(hojaResultado).Columns("J:J").EntireColumn.AutoFit
Sheets(hojaResultado).Columns("I:I").EntireColumn.AutoFit
Sheets(hojaResultado).Columns("H:H").EntireColumn.AutoFit

Sheets(hojaResultado).Columns("C:C").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False

Sheets(hojaResultado).Columns("D:D").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False

Sheets(hojaResultado).Columns("E:E").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False
ActiveWindow.SmallScroll Down:=0

Sheets(hojaResultado).Columns("F:F").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False

Sheets(hojaResultado).Columns("G:G").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False

Sheets(hojaResultado).Columns("H:H").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False

Sheets(hojaResultado).Columns("I:I").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False

Sheets(hojaResultado).Columns("J:J").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False

Sheets(hojaResultado).Columns("K:K").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False

Sheets(hojaResultado).Columns("L:L").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False

Sheets(hojaResultado).Columns("M:M").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False

Sheets(hojaResultado).Columns("N:N").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False

Sheets(hojaResultado).Columns("O:O").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False
ActiveWindow.ScrollColumn = 2
ActiveWindow.ScrollColumn = 3

Sheets(hojaResultado).Columns("S:S").Select
Selection.Copy

Sheets(hojaResultado).Range("R1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False

j = 2
i = 1
final = False
fecha = Sheets(hojaDatos).Cells(3, 1).Value
mes = Mid(fecha, 53, 2)
cont = 0


Do Until final = True
    j = 2
    asterisco = Sheets(hojaDatos).Cells(i, 2).Value
    finalEncontrado = False
    If asterisco = "**" Then
        final = True
    ElseIf asterisco = "*" Then
        sociedad = Sheets(hojaDatos).Cells(i, 4).Value
        sociedadCodigo = Mid(sociedad, 12, 4)
        sociedadCodigoDigito = Mid(sociedadCodigo, 1, 1)
        If sociedadCodigoDigito = "4" Or sociedadCodigoDigito = "5" Or sociedadCodigoDigito = "6" _
            Or sociedadCodigoDigito = "7" Then
                    Do Until IsEmpty(Sheets(hojaResultado).Cells(j, 1))
                        sociedadComparar = Sheets(hojaResultado).Cells(j, 1).Value
                        If sociedadComparar = sociedadCodigo Then
                            valor = Sheets(hojaDatos).Cells(i, 20).Value
                            Sheets(hojaResultado).Cells(j, 19).Value = valor / 1000000
                            finalEncontrado = True
                            Exit Do
                        End If
                        j = j + 1
                    Loop
                    
                    If finalEncontrado = False Then
                        Sheets(hojaResultado).Cells(j, 8).EntireRow.Insert
                        Sheets(hojaResultado).Cells(j, 1).Value = sociedadCodigo
                        valor = Sheets(hojaDatos).Cells(i, 20).Value
                        Sheets(hojaResultado).Cells(j, 19).Value = valor / 1000000
                    End If
'
        End If
    End If
    i = i + 1
Loop

j = 2

Do Until IsEmpty(Sheets(hojaResultado).Cells(j, 1))
    If mes = "01" Then
        Sheets(hojaResultado).Cells(j, 4).Value = Sheets(hojaResultado).Cells(j, 19).Value - Sheets(hojaResultado).Cells(j, 18).Value
        Sheets(hojaResultado).Cells(j, 16).Value = Sheets(hojaResultado).Cells(j, 4).Value - Sheets(hojaResultado).Cells(j, 3).Value
        variacion = Sheets(hojaResultado).Cells(j, 16).Value
        mesAnterior = Sheets(hojaResultado).Cells(j, 3).Value
        If mesAnterior = 0 And variacion > 0 Then
            porcentajeVariacion = 1
        ElseIf mesAnterior = 0 And variacion < 0 Then
            porcentajeVariacion = -1
        ElseIf mesAnterior = 0 And variacion = 0 Then
            porcentajeVariacion = 0
        Else
            porcentajeVariacion = (variacion / mesAnterior)
        End If
        Sheets(hojaResultado).Cells(j, 17).Value = porcentajeVariacion
        
    ElseIf mes = "02" Then
        Sheets(hojaResultado).Cells(j, 5).Value = Sheets(hojaResultado).Cells(j, 19).Value - Sheets(hojaResultado).Cells(j, 18).Value
        Sheets(hojaResultado).Cells(j, 16).Value = Sheets(hojaResultado).Cells(j, 5).Value - Sheets(hojaResultado).Cells(j, 4).Value
        variacion = Sheets(hojaResultado).Cells(j, 16).Value
        mesAnterior = Sheets(hojaResultado).Cells(j, 4).Value
        If mesAnterior = 0 And variacion > 0 Then
            porcentajeVariacion = 1
        ElseIf mesAnterior = 0 And variacion < 0 Then
            porcentajeVariacion = -1
        ElseIf mesAnterior = 0 And variacion = 0 Then
            porcentajeVariacion = 0
        Else
            porcentajeVariacion = (variacion / mesAnterior)
        End If
        Sheets(hojaResultado).Cells(j, 17).Value = porcentajeVariacion
        
    ElseIf mes = "03" Then
        Sheets(hojaResultado).Cells(j, 6).Value = Sheets(hojaResultado).Cells(j, 19).Value - Sheets(hojaResultado).Cells(j, 18).Value
        Sheets(hojaResultado).Cells(j, 16).Value = Sheets(hojaResultado).Cells(j, 6).Value - Sheets(hojaResultado).Cells(j, 5).Value
        variacion = Sheets(hojaResultado).Cells(j, 16).Value
        mesAnterior = Sheets(hojaResultado).Cells(j, 5).Value
        If mesAnterior = 0 And variacion > 0 Then
            porcentajeVariacion = 1
        ElseIf mesAnterior = 0 And variacion < 0 Then
            porcentajeVariacion = -1
        ElseIf mesAnterior = 0 And variacion = 0 Then
            porcentajeVariacion = 0
        Else
            porcentajeVariacion = (variacion / mesAnterior)
        End If
        Sheets(hojaResultado).Cells(j, 17).Value = porcentajeVariacion
        
    ElseIf mes = "04" Then
        Sheets(hojaResultado).Cells(j, 7).Value = Sheets(hojaResultado).Cells(j, 19).Value - Sheets(hojaResultado).Cells(j, 18).Value
        Sheets(hojaResultado).Cells(j, 16).Value = Sheets(hojaResultado).Cells(j, 7).Value - Sheets(hojaResultado).Cells(j, 6).Value
        variacion = Sheets(hojaResultado).Cells(j, 16).Value
        mesAnterior = Sheets(hojaResultado).Cells(j, 6).Value
        If mesAnterior = 0 And variacion > 0 Then
            porcentajeVariacion = 1
        ElseIf mesAnterior = 0 And variacion < 0 Then
            porcentajeVariacion = -1
        ElseIf mesAnterior = 0 And variacion = 0 Then
            porcentajeVariacion = 0
        Else
            porcentajeVariacion = (variacion / mesAnterior)
        End If
        Sheets(hojaResultado).Cells(j, 17).Value = porcentajeVariacion
        
    ElseIf mes = "05" Then
        Sheets(hojaResultado).Cells(j, 8).Value = Sheets(hojaResultado).Cells(j, 19).Value - Sheets(hojaResultado).Cells(j, 18).Value
        Sheets(hojaResultado).Cells(j, 16).Value = Sheets(hojaResultado).Cells(j, 8).Value - Sheets(hojaResultado).Cells(j, 7).Value
        variacion = Sheets(hojaResultado).Cells(j, 16).Value
        mesAnterior = Sheets(hojaResultado).Cells(j, 7).Value
        If mesAnterior = 0 And variacion > 0 Then
            porcentajeVariacion = 1
        ElseIf mesAnterior = 0 And variacion < 0 Then
            porcentajeVariacion = -1
        ElseIf mesAnterior = 0 And variacion = 0 Then
            porcentajeVariacion = 0
        Else
            porcentajeVariacion = (variacion / mesAnterior)
        End If
        Sheets(hojaResultado).Cells(j, 17).Value = porcentajeVariacion
        
    ElseIf mes = "06" Then
        Sheets(hojaResultado).Cells(j, 9).Value = Sheets(hojaResultado).Cells(j, 19).Value - Sheets(hojaResultado).Cells(j, 18).Value
        Sheets(hojaResultado).Cells(j, 16).Value = Sheets(hojaResultado).Cells(j, 9).Value - Sheets(hojaResultado).Cells(j, 8).Value
        variacion = Sheets(hojaResultado).Cells(j, 16).Value
        mesAnterior = Sheets(hojaResultado).Cells(j, 8).Value
        If mesAnterior = 0 And variacion > 0 Then
            porcentajeVariacion = 1
        ElseIf mesAnterior = 0 And variacion < 0 Then
            porcentajeVariacion = -1
        ElseIf mesAnterior = 0 And variacion = 0 Then
            porcentajeVariacion = 0
        Else
            porcentajeVariacion = (variacion / mesAnterior)
        End If
        Sheets(hojaResultado).Cells(j, 17).Value = porcentajeVariacion
        
    ElseIf mes = "07" Then
        Sheets(hojaResultado).Cells(j, 10).Value = Sheets(hojaResultado).Cells(j, 19).Value - Sheets(hojaResultado).Cells(j, 18).Value
        Sheets(hojaResultado).Cells(j, 16).Value = Sheets(hojaResultado).Cells(j, 10).Value - Sheets(hojaResultado).Cells(j, 9).Value
        variacion = Sheets(hojaResultado).Cells(j, 16).Value
        mesAnterior = Sheets(hojaResultado).Cells(j, 9).Value
        If mesAnterior = 0 And variacion > 0 Then
            porcentajeVariacion = 1
        ElseIf mesAnterior = 0 And variacion < 0 Then
            porcentajeVariacion = -1
        ElseIf mesAnterior = 0 And variacion = 0 Then
            porcentajeVariacion = 0
        Else
            porcentajeVariacion = (variacion / mesAnterior)
        End If
        Sheets(hojaResultado).Cells(j, 17).Value = porcentajeVariacion
        
    ElseIf mes = "08" Then
        Sheets(hojaResultado).Cells(j, 11).Value = Sheets(hojaResultado).Cells(j, 19).Value - Sheets(hojaResultado).Cells(j, 18).Value
        Sheets(hojaResultado).Cells(j, 16).Value = Sheets(hojaResultado).Cells(j, 11).Value - Sheets(hojaResultado).Cells(j, 10).Value
        variacion = Sheets(hojaResultado).Cells(j, 16).Value
        mesAnterior = Sheets(hojaResultado).Cells(j, 10).Value
        If mesAnterior = 0 And variacion > 0 Then
            porcentajeVariacion = 1
        ElseIf mesAnterior = 0 And variacion < 0 Then
            porcentajeVariacion = -1
        ElseIf mesAnterior = 0 And variacion = 0 Then
            porcentajeVariacion = 0
        Else
            porcentajeVariacion = (variacion / mesAnterior)
        End If
        Sheets(hojaResultado).Cells(j, 17).Value = porcentajeVariacion
        
    ElseIf mes = "09" Then
        Sheets(hojaResultado).Cells(j, 12).Value = Sheets(hojaResultado).Cells(j, 19).Value - Sheets(hojaResultado).Cells(j, 18).Value
        Sheets(hojaResultado).Cells(j, 16).Value = Sheets(hojaResultado).Cells(j, 12).Value - Sheets(hojaResultado).Cells(j, 11).Value
        variacion = Sheets(hojaResultado).Cells(j, 16).Value
        mesAnterior = Sheets(hojaResultado).Cells(j, 11).Value
        If mesAnterior = 0 And variacion > 0 Then
            porcentajeVariacion = 1
        ElseIf mesAnterior = 0 And variacion < 0 Then
            porcentajeVariacion = -1
        ElseIf mesAnterior = 0 And variacion = 0 Then
            porcentajeVariacion = 0
        Else
            porcentajeVariacion = (variacion / mesAnterior)
        End If
        Sheets(hojaResultado).Cells(j, 17).Value = porcentajeVariacion
        
    ElseIf mes = "10" Then
        Sheets(hojaResultado).Cells(j, 13).Value = Sheets(hojaResultado).Cells(j, 19).Value - Sheets(hojaResultado).Cells(j, 18).Value
        Sheets(hojaResultado).Cells(j, 16).Value = Sheets(hojaResultado).Cells(j, 13).Value - Sheets(hojaResultado).Cells(j, 12).Value
        variacion = Sheets(hojaResultado).Cells(j, 16).Value
        mesAnterior = Sheets(hojaResultado).Cells(j, 12).Value
        If mesAnterior = 0 And variacion > 0 Then
            porcentajeVariacion = 1
        ElseIf mesAnterior = 0 And variacion < 0 Then
            porcentajeVariacion = -1
        ElseIf mesAnterior = 0 And variacion = 0 Then
            porcentajeVariacion = 0
        Else
            porcentajeVariacion = (variacion / mesAnterior)
        End If
        Sheets(hojaResultado).Cells(j, 17).Value = porcentajeVariacion
        
    ElseIf mes = "11" Then
        Sheets(hojaResultado).Cells(j, 14).Value = Sheets(hojaResultado).Cells(j, 19).Value - Sheets(hojaResultado).Cells(j, 18).Value
        Sheets(hojaResultado).Cells(j, 16).Value = Sheets(hojaResultado).Cells(j, 14).Value - Sheets(hojaResultado).Cells(j, 13).Value
        variacion = Sheets(hojaResultado).Cells(j, 16).Value
        mesAnterior = Sheets(hojaResultado).Cells(j, 13).Value
        If mesAnterior = 0 And variacion > 0 Then
            porcentajeVariacion = 1
        ElseIf mesAnterior = 0 And variacion < 0 Then
            porcentajeVariacion = -1
        ElseIf mesAnterior = 0 And variacion = 0 Then
            porcentajeVariacion = 0
        Else
            porcentajeVariacion = (variacion / mesAnterior)
        End If
        Sheets(hojaResultado).Cells(j, 17).Value = porcentajeVariacion
        
    ElseIf mes = "12" Then
        Sheets(hojaResultado).Cells(j, 15).Value = Sheets(hojaResultado).Cells(j, 19).Value - Sheets(hojaResultado).Cells(j, 18).Value
        Sheets(hojaResultado).Cells(j, 16).Value = Sheets(hojaResultado).Cells(j, 15).Value - Sheets(hojaResultado).Cells(j, 14).Value
        variacion = Sheets(hojaResultado).Cells(j, 16).Value
        mesAnterior = Sheets(hojaResultado).Cells(j, 14).Value
        If mesAnterior = 0 And variacion > 0 Then
            porcentajeVariacion = 1
        ElseIf mesAnterior = 0 And variacion < 0 Then
            porcentajeVariacion = -1
        ElseIf mesAnterior = 0 And variacion = 0 Then
            porcentajeVariacion = 0
        Else
            porcentajeVariacion = (variacion / mesAnterior)
        End If
        Sheets(hojaResultado).Cells(j, 17).Value = porcentajeVariacion
        
    End If
    
    If (Sheets(hojaResultado).Cells(j, 17).Value > 10 Or Sheets(hojaResultado).Cells(j, 17).Value < -10) _
        And (Sheets(hojaResultado).Cells(j, 16).Value > 500 Or Sheets(hojaResultado).Cells(j, 16).Value < -500) Then
            Sheets(hojaResultado).Range("Q" & j).Interior.Color = RGB(252, 213, 180)
    Else
        Sheets(hojaResultado).Range("Q" & j).Interior.Color = RGB(255, 255, 255)
    End If
    
    j = j + 1
Loop

Sheets(hojaResultado).Cells(1, 18).Value = "Acumulado mes anterior"
Sheets(hojaResultado).Cells(1, 19).Value = "Acumulado mes actual"
    
Call Balance
    
End Sub



