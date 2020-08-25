Sub GenerarCSV()
    
    If Hoja1.Range("OBJETIVO") <> 0 Then
        MsgBox "Cronograma no ajustado"
        Exit Sub
    End If
    
    Dim i As Integer
    i = Hoja1.Range("INICIO_REF").Row
    
    While Hoja1.Range("INICIO_REF").Offset(i + 1 - Hoja1.Range("INICIO_REF").Row, 0) <> ""
        i = i + 1
    Wend
    Application.ScreenUpdating = False
    
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    wb.Sheets(1).Range("A1").Value = 1
    wb.Sheets(1).Range("B1").Value = 2
    wb.Sheets(1).Range("C1").Value = 3
    wb.Sheets(1).Range("D1").Value = 4
    wb.Sheets(1).Range("E1").Value = 5
    wb.Sheets(1).Range("F1").Value = 6
    wb.Sheets(1).Range("G1").Value = 7
    wb.Sheets(1).Range("H1").Value = 8
    wb.Sheets(1).Range("I1").Value = 9
    wb.Sheets(1).Range("J1").Value = 10
    wb.Sheets(1).Range("K1").Value = 11
    
    wb.Sheets(1).Range("A2").Value = "Cuotas"
    wb.Sheets(1).Range("B2").Value = "Fecha"
    wb.Sheets(1).Range("C2").Value = "Dias"
    wb.Sheets(1).Range("D2").Value = "SaldoInicial"
    wb.Sheets(1).Range("E2").Value = "TotalCuota"
    wb.Sheets(1).Range("F2").Value = "Amortizacion"
    wb.Sheets(1).Range("G2").Value = "Interes"
    wb.Sheets(1).Range("H2").Value = "Seguro Des."
    wb.Sheets(1).Range("I2").Value = "Seg.Bien"
    wb.Sheets(1).Range("J2").Value = "Apor."
    wb.Sheets(1).Range("K2").Value = "SaldoFinal"
    
    Hoja1.Range(Hoja1.Cells(Hoja1.Range("INICIO_REF").Row, Hoja1.Range("INICIO_REF").Column), Hoja1.Cells(i, Hoja1.Range("FIN_REF").Column)).Copy
    wb.Sheets(1).Range("A3").PasteSpecial xlPasteValues
    
    wb.Sheets(1).Range(wb.Sheets(1).Range("B2"), wb.Sheets(1).Range("B2").End(xlDown)).NumberFormat = "DD/MM/YYYY"
    
    If wb.Sheets(1).Range("K1").End(xlDown).Value < 0.001 Then
        wb.Sheets(1).Range("K1").End(xlDown).Value = 0
    End If
    
    If Not dirExists("C:\Cronogramas") Then
        MkDir "C:\Cronogramas"
    End If
    
    ActiveWorkbook.SaveAs Filename:="C:\Cronogramas" & "\" & Format(Now(), "YYYYMMDD-HHMMSS") & ".csv", FileFormat:=xlCSV, CreateBackup:=False
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
End Sub

Public Function dirExists(s_directory As String) As Boolean

    Set OFSO = CreateObject("Scripting.FileSystemObject")
    dirExists = OFSO.FolderExists(s_directory)
    Set OFSO = Nothing
    
End Function
