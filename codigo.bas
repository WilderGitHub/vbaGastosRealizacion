Dim MensajeConcepto As String
Dim NombreArchivoMadre As String
Dim nombre As String
Dim UltimaFila As Integer
Dim miRango As Integer
Dim nombreArchivoNuevo As String
Sub trigercin(control As IRibbonControl) ' Cuerpo principal
   If MsgBox("Empezamos el proceso no jefe?", vbOKCancel, "Aviso") = vbOK Then
        Application.ScreenUpdating = False
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        NombreArchivoMadre = ActiveWindow.Caption
        Procesar ("Aduanas")
    End If
End Sub

Sub Procesar(AbrirEsto As String)
    MensajeConcepto = "Donde esta la información de " & AbrirEsto & " ??"
    ActiveSheet.Unprotect
    DataFile = Application.GetOpenFilename("All files (*.*), *.*", , MensajeConcepto, "comues", False)
    If DataFile = False Then
        End
    Else
        Workbooks.Open DataFile
        obtenerDatos
        LlenarFormulas
        'Sheets("Data").Copy After:=Workbooks(nombre).Sheets(Workbooks(nombre).Sheets.Count)
        'MsgBox (nombre)
       'Sheets("Data").Copy
       'Sheets("Data").Move
           
        Sheets("Data").Select
       'Sheets(ActiveSheet).Move Before:=Workbooks("nombre.xlsx").Sheets(1)
        hacerResumen
        
        exportarPestañas
        ' Windows(nombre).Close
        MsgBox ("Ya hemos terminado jefe. El archivo generado se llama """ & nombreArchivoNuevo & ".xls""")
    End If
End Sub
Sub obtenerDatos()

         
        'nombre del archivo que vamos a abrir
        nombre = ActiveWorkbook.Name
        'Nos ponemos en A2 y marcamos todingo y copiamos
         Range("A1").Select
         Range(Selection, Selection.End(xlToRight)).Select
         Range(Selection, Selection.End(xlDown)).Select
         Selection.Copy
         'volvemos al archivo madre y pegamos en "Data"
         Windows(NombreArchivoMadre).Activate
         'creamos una pestaña "data"
         Sheets.Add(After:=Sheets("Diccionario")).Name = "Data"
         Sheets.Add(After:=Sheets("Data")).Name = "Resumen"
         Sheets("Data").Select
         Range("A1").Select
         Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
         
         
         'Windows(NombreArchivoMadre).Activate
         
         
          Windows(nombre).Close
    

End Sub
Sub LlenarFormulas()
    'obtenemos el rango
    getMiRango
    
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "CodigoBCB"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5]/10,diccionario,2,FALSE),0)"
    
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "Producto"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=+IFERROR(VLOOKUP(RC[-6]/10,diccionario,3,FALSE),0)"
    
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "FacturaCorregida"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = _
        "=+IF(RC[-3]=minerales," & Chr(10) & "               IF(RC[-10]>=RC[1],IF(RC[-3]=minerales,RC[-10]/COUNTIF(R2C1:R" & miRango & "C1,RC[-18]),0),RC[1])," & Chr(10) & "          0)"
        
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "FobAux"
    Range("T2").Select
    ActiveCell.FormulaR1C1 = _
        "=+IFERROR(IF(RC[-4]=minerales," & Chr(10) & "        IF(AND(RC[-13]=sanCristobal,RC[-2]=zinc)," & Chr(10) & "               RC[-11]*ratioCasoEspecial," & Chr(10) & "               RC[-5])," & Chr(10) & "          0),0)"
    
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "FobCorregido"
    Range("U2").Select
    ActiveCell.FormulaR1C1 = _
        "=+IFERROR(+IF(SUMIF(R2C1:R" & miRango & "C20,RC[-20],R2C20:R" & miRango & "C20)/RC[-12]<umbralFOB,RC[-2],RC[-1]),0)"
        
    Range("V1").Select
    ActiveCell.FormulaR1C1 = "GastoRealizacion"
    Range("V2").Select
    ActiveCell.FormulaR1C1 = "=+RC[-3]-RC[-1]"
    
    'rellenamos hacia abajo
    Range("Q2:V2").Select
    Selection.AutoFill Destination:=Range("Q2:V" & miRango & "")
    Range("Q2:V" & miRango & "").Select
    Range("q1").Select
    'llenamos como valor desde la segunda fila, pero puedes cambiar para ver la formula
    Range("Q3:V3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("Q3").Select
    
End Sub

Sub getMiRango()

    Range("w1").Select
    ActiveCell.FormulaR1C1 = "=+COUNTA(C[-7])"
    miRango = Range("w1").Value
    Range("W1").Select
    Selection.ClearContents
End Sub

Sub hacerResumen()
    'Sheets("Data").Select
    'Range("Z2").Select
    Sheets("Resumen").Select
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Resumen"
    Range("A1").Select
    Selection.Font.Size = 18
    Range("A3").Select
    
    ActiveCell.FormulaR1C1 = "Valor Factura Corregida"
    Range("A4").Select
    
    ActiveCell.FormulaR1C1 = "Fob Corregido"
    Range("A5").Select
    Columns("A:A").EntireColumn.AutoFit
    ActiveCell.FormulaR1C1 = "Gastos de Realización"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "=SUM(Data!C[17])"
    Range("B3").Select
    Selection.NumberFormat = "#,##0.00"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "=SUM(Data!C[19])"
    Range("B4").Select
    Selection.NumberFormat = "#,##0.00"
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "=SUM(Data!C[20])"
    Range("B5").Select
    Selection.NumberFormat = "#,##0.00"
    Range("B6").Select
    ActiveCell.FormulaR1C1 = "=+R[-3]C-R[-2]C"
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "Check"
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "Ratio"
    Range("B7").Select
    ActiveCell.FormulaR1C1 = "=+R[-2]C/R[-4]C"
    Range("B7").Select
    Selection.NumberFormat = "0%"
    Selection.NumberFormat = "0.0%"
    Selection.NumberFormat = "0.00%"
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("B5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("a1:d20").Select
    'Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("a1").Select
    
End Sub

Sub exportarPestañas()
    
    Sheets(Array("Data", "Resumen")).Select
    Sheets(Array("Data", "Resumen")).Copy
    
    nombreMes = Format(Now, "mmm")
    nombreArchivoNuevo = "Aduanas " & Day(Now) & nombreMes & Hour(Now) & Minute(Now)
    
    ActiveWorkbook.SaveAs Filename:=nombreArchivoNuevo, _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    'ActiveWorkbook.SaveAs Filename:= _
    '    "Aduanas " & Day(Now) & nombreMes & Hour(Now) & Minute(Now), _
    '    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    ActiveWindow.Close
    
    Sheets(Array("Data", "Resumen")).Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Parametros").Select
    Range("a1").Select
    
End Sub