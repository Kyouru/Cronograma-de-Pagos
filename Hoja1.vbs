
Private Sub btAjustar_Click()
    Dim arreglo As Range
    Dim cuota As Range
    Dim ncuotas As Range
    Dim monto As Range
    Dim objetivo As Range
    Dim ajuste, anterior As Double
    Dim limite, limite2 As Integer
    limite = 20
    
    Set arreglo = Me.Range("ARREGLO")
    Set cuota = Me.Range("CUOTA")
    Set objetivo = Me.Range("OBJETIVO")
    Set monto = Me.Range("MONTO")
    Set ncuotas = Me.Range("NCUOTAS")
        
    If objetivo <> 0 Then
        If [TOTALCUOTAS] - [areaFija] - [areaFija2] = 1 Then
            cuota = cuota + objetivo
        Else
            arreglo = 0
            
            'cuota = Round(monto / (ncuotas + [CUOTAADICIONAL]), 2)
            cuota = Round(monto / (ncuotas), 2)
            
            ajuste = 1
            'For i = 2 To Len(CStr(Round(monto, 0)))
            For i = 2 To Len(CStr(Round(Abs(objetivo), 0)))
                ajuste = ajuste * 10
            Next i
            If objetivo < 0 Then
                Do While ajuste > 0
                    Do While objetivo <= 0 And cuota >= 0
                        anterior = cuota
                        cuota = cuota - ajuste
                    Loop
                    cuota = anterior
                    ajuste = Round(ajuste / 10, 2)
                Loop
            End If
            If objetivo > 0 Then
                Do While ajuste > 0
                    Dim countLimit As Integer
                    countLimit = 0
                    Do While objetivo >= 0 And cuota >= 0 And countLimit < limite
                        anterior = cuota
                        cuota = cuota + ajuste
                        countLimit = countLimit + 1
                    Loop
                    If countLimit >= limite Then
                        MsgBox "Error Loop"
                        Exit Sub
                    End If
                    cuota = anterior
                    ajuste = Round(ajuste / 10, 2)
                Loop
                If objetivo <> 0 Then
                    cuota = cuota + 0.01
                End If
            End If
        End If
    End If
    
    arreglo = -(objetivo - arreglo)
    'Hoja1.ScrollArea = "A1:W" & 15 + Me.Range(ncuotas).Value
End Sub

Private Sub btnLimpiar_Click()
    Range(Cells(Me.Range("InicioLimpiar").Row, Me.Range("InicioLimpiar").Column), Cells(2000, Me.Range("FinLimpiar").Column)).ClearContents
End Sub

Private Sub cbUltimoDia_Change()
    If cbUltimoDia.Value Then
        Hoja1.Range("ULTIMODIA").Value = "1"
    Else
        Hoja1.Range("ULTIMODIA").Value = "0"
    End If
End Sub
