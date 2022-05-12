Attribute VB_Name = "ProcedimientoSalidas"

Sub GenerarReporteMaster()
    
    Call acelerarProceso("")
    Call eliminaFila
    Call ArchivoExiste("Reporte MASTER tiempos de despacho HYPERION AICM.xlsx")
    Call insertarHoja("Etapas")
    Call cambiarNomHoja("Hoja1", "Reporte")
    Call cpEtapas("Reporte.xlsx", "Reporte MASTER tiempos de despacho HYPERION AICM.xlsx", "A1:", 2, 1, 1)
    Call cpEtapas("Libro1.xlsx", "Reporte MASTER tiempos de despacho HYPERION AICM.xlsx", "A6:", 1, 6, 1)
    Call activarLibro("Reporte MASTER tiempos de despacho HYPERION AICM.xlsx")
    Call activarHoja("Reporte")
    Call insertarColumna(6, 3, xlLeft)
    Call colocarCol("BANCO", 1, 6)
    Call borrarRango(1, "E1:")
    Call seleccionRango(1, "A6:")
    Call ordenar1Criterio("D7")
    Call colocarLeyendas
    Call colocarFechas
    Call colocarTitReport
    Range("A6:I6").Select
    Call filaGris
    Call ajustarCeldas("")
    Call formatoReporte
    Call cambioNombreBanco("")
    Call cerrarLibros("Reporte.xlsx")
    Call cerrarLibros("Libro1.xlsx")
    Call configOriginal("")
    Call seleccionRango(1, "A6:")
    Call ordenar1Criterio("A7")
    Call colocaColor
    Call guardarLibro(x)
    
End Sub

Private Sub eliminaFila()

    Call activarLibro("Libro1.xlsx")
    If extraerCadena(Cells(6, 1), "Aduanas:") <> "" Then
    
        Cells(6, 1).EntireRow.Delete
    
    End If
    

End Sub
Private Sub filaGris()
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -9.99786370433668E-02
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    
End Sub

Private Sub filaAzul()
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
End Sub

Private Sub colocarFechas()
    
    Dim filasReporte As Long: filasReporte = ultimaFila(1)
    Dim filasEtapas As Long: filasEtapas = ultimaFila(2)
    Dim pivote      As Long: pivote = 6
    Dim referencia  As String
    Dim listaRef(2) As Long
    Dim fecRep(2)   As String
    
    Call activarHoja("Etapas")
    
    listaRef(0) = buscarColumna(2, "FECHA DE ENTRADA AL PAÍS", 5)
    listaRef(1) = buscarColumna(2, "FECHA DE REVALIDACION", 5)
    listaRef(2) = buscarColumna(2, "FECHA DE MEC. DE S. AUTOMATIZADA", 5)
    
    Call activarHoja("Reporte")
    
    For x = 7 To filasReporte
        
        referencia = Cells(x, 1).Value
        
        Call activarHoja("Etapas")
        
        For y = pivote To filasEtapas
            
            If UCase(referencia) = UCase(Cells(y, 1).Value) Then
                
                fecRep(0) = Cells(y, listaRef(0)).Value
                fecRep(1) = Cells(y, listaRef(1)).Value
                fecRep(2) = Cells(y, listaRef(2)).Value
                
                Exit For
                
            Else
                
                fecRep(0) = ""
                fecRep(1) = ""
                fecRep(2) = ""
                
            End If
            
        Next y
        
        Call activarHoja("Reporte")
        Call fechas(x, fecRep)
        Call aresta(x)
        Cells(x, 2) = CLng(extraerCadena(Cells(x, 2).Value, "[0-9]{6,10}"))
    Next x
    
    Erase listaRef()
    Erase fecRep()
    
End Sub

Private Sub colocaColor()

Dim finFila As Long: finFila = ultimaFila(1)

For x = 7 To finFila

        If x Mod 2 <> 0 Then
            
            Range("A" & x & ":" & "I" & x).Select
            Call filaAzul
        End If
Next x

End Sub
Sub aresta(x As Variant)
    
    If extraerCadena(Cells(x, 1).Value, "I") <> "" And Cells(x, 5) <> "" And Cells(x, 6) <> "" And Cells(x, 7) <> "" Then
        
        Cells(x, 8) = Cells(x, 7).Value - Cells(x, 5).Value
        Cells(x, 9) = Cells(x, 7).Value - Cells(x, 6).Value
        
    ElseIf extraerCadena(Cells(x, 1).Value, "I") <> "" And Cells(x, 5) = "" And Cells(x, 7) <> "" Then
        
        Cells(x, 8) = "NO SE REG FEC ENTRADA"
        Cells(x, 9) = Cells(x, 7).Value - Cells(x, 6).Value
        
    ElseIf extraerCadena(Cells(x, 1).Value, "^MXE") <> "" And Cells(x, 7) <> "" And Cells(x, 5) <> "" Then
        
        Cells(x, 8) = Cells(x, 7).Value - Cells(x, 5).Value
        Cells(x, 9) = "EXP"
    Else
        If extraerCadena(Cells(x, 1).Value, "^R") <> "" Then
            
            Cells(x, 9) = "RECTI-EXP"
            
            If Cells(x, 7) <> "" Then
                Cells(x, 8) = Cells(x, 7).Value - Cells(x, 5).Value
            Else
                Cells(x, 8) = "NO SE REG FEC DESP"
            End If
        Else
            
            Cells(x, 8) = "NO SE REG FEC DESP"
            Cells(x, 9) = "NO SE REG FEC DESP"
            
        End If
        
    End If
    
End Sub

Private Sub colocarTitReport()
    
    Cells(6, 1) = "Referencia"
    Cells(6, 2) = "Pedimento"
    Cells(6, 3) = "Banco"
    Cells(6, 4) = "Fec pago"
    Cells(6, 5) = "Fec entrada"
    Cells(6, 6) = "Fec Revalidación"
    Cells(6, 7) = "Fec Despacho"
    Cells(6, 8) = "Despacho vs Entrada"
    Cells(6, 9) = "Despacho vs Revalida"
    
End Sub

Sub ELEGIR(x As Variant)
    
    If extraerCadena(Cells(x, 1).Value, "R[0-9]?S?[0-9]?MXE") <> "" Then
        
        Cells(x, 9) = "RECTI-EXP"
        
    ElseIf extraerCadena(Cells(x, 1).Value, "MXE") <> "" Then
        
        Cells(x, 9) = "EXP"
        
    Else
        
        Call UltimaComprobacion(x)
        
    End If
    
End Sub

Sub UltimaComprobacion(x As Variant)
    
    If Cells(x, 7) = "" Then
        
        Cells(x, 9) = "NO SE REG FEC DESP"
        Cells(x, 9) = "NO SE REG FEC DESP"
        
    ElseIf Cells(x, 5) = "" Then
        
        Cells(x, 9) = "NO SE REG FEC ENTRADA"
        Cells(x, 9) = "NO SE REG FEC ENTRADA"
        
    Else
        
        Cells(x, 9) = Cells(x, 7) - Cells(x, 5)
        
    End If
    
End Sub

Sub formatoFecha(filas As Long)
    filas = ultimaFila(1)
    For x = 8 To filas
        
        Cells(x, 49) = CDate(Cells(x, 49))
        Cells(x, 49).NumberFormat = "dd/mm/yyyy"
        
    Next x
    
End Sub
Sub fechas(x As Variant, fecRep() As String)
    
    If fecRep(0) <> "" Then
        
        Cells(x, 5) = CDbl(CDate(fecRep(0)))
        Cells(x, 5).NumberFormat = "dd/mm/yyyy"
        
    End If
    
    If fecRep(1) <> "" Then
        
        Cells(x, 6) = CDbl(CDate(fecRep(1)))
        Cells(x, 6).NumberFormat = "dd/mm/yyyy"
        
    End If
    
    If fecRep(2) <> "" Then
        
        Cells(x, 7) = CDbl(CDate(fecRep(2)))
        Cells(x, 7).NumberFormat = "dd/mm/yyyy"
        
    End If
    
End Sub

Sub colocarCol(dato As String, hoja As Integer, fila As Long)
    
    Dim columna     As Long: columna = buscarColumna(hoja, dato, fila)
    
    Cells(fila, columna).EntireColumn.Select
    Selection.Cut Destination:=Worksheets(hoja).Cells(1, 3)
    
End Sub

Function buscarColumna(hoja As Integer, nombreCol As String, fila As Long)
    
    Dim columna     As String: columna = ultimaColunma(hoja)
    
    For x = 1 To columna
        
        If UCase(Cells(fila, x).Value) = UCase(nombreCol) Then
            
            buscarColumna = x
            Exit For
            
        End If
    Next x
    
End Function

Sub borrarRango(hoja As Integer, rango As String)
    
    Call seleccionRango(hoja, rango)
    Selection.Delete
    
End Sub

'
Sub cpEtapas(libro1 As String, libro2 As String, celda1 As String, hoja As Integer, fila As Long, col As Long)
    
    Call activarLibro(libro1)
    Call seleccionRango(1, celda1)
    Call cp(hoja, fila, col, libro2)
    Call activarLibro(libro2)
    
End Sub

Function ExtraerFechaOut(x As Long, y As Long)
    
    Dim fecha       As String
    Dim dia         As String
    Dim mes         As String
    Dim anio        As String
    
    fecha = extraerCadena(extraerCadena(Cells(4, 1).Value, "Fecha Final: [0-9]{2}\/[0-9]{2}\/[0-9]{4}"), "[0-9]{2}\/[0-9]{2}\/[0-9]{4}")
    fecha = Replace(fecha, "/", "")
    dia = Mid(fecha, 1, 2)
    mes = Mid(fecha, 3, 2)
    anio = Mid(fecha, 5, 4)
    mes = ObtenerMesOut(CInt(mes))
    
    fecha = dia & " de " & mes & " del " & anio
    ExtraerFechaOut = fecha
    
End Function

Function ObtenerMesOut(mes As Integer)
    
    Dim meses(11)   As String
    
    Call listaMesesCom(meses)
    
    For x = 0 To 11
        
        If x = mes - 1 Then
            
            ObtenerMesOut = meses(x)
            Exit For
            
        End If
        
    Next x
    
End Function

Private Sub colocarLeyendas()
    
    Dim fecha       As String
    
    Call activarLibro("Reporte.xlsx")
    
    fecha = ExtraerFechaOut(4, 1)
    
    Call activarLibro("Reporte MASTER tiempos de despacho HYPERION AICM.xlsx")
    Cells(1, 1) = "DIAS DE DESPACHO"
    Cells(2, 1) = "Cliente: HYPERION"
    Cells(3, 1) = "Periodo: 01 al " & fecha
    Cells(4, 1) = "Aduana: 470"
    
End Sub

'Se selecciona un rango de datos apartir de de una celda dada hasta la ultima fila y columna
Sub seleccionRango(hoja As Integer, celda1 As String)
    
    Dim filas       As Long: filas = ultimaFila(hoja)
    Dim columna     As Long: columna = ultimaColunma(hoja)
    Dim rango       As String: rango = Cells(filas, columna).Address
    
    rango = Replace(rango, "$", "")
    
    Call SeleccionarRango(celda1 & rango)
    
End Sub

'Se selecciona un rango de datos apartir de de una celda dada hasta la ultima fila y columna dada
Sub seleccionRangoDet(hoja As Integer, celda1 As String, columna As Long)
    
    Dim filas       As Long: filas = ultimaFila(hoja)
    Dim rango       As String: rango = Cells(filas, columna).Address
    
    rango = Replace(rango, "$", "")
    
    Call SeleccionarRango(celda1 & rango)
    
End Sub

Private Sub reportes()
    
    Call activarLibro("Libro1.xlsx")
    Dim filas       As Long: filas = ultimaFila(1)
    Dim columna     As Long: columna = ultimaColunma(1)
    
    seleccionar ("A6")
    
End Sub

Sub cambioNombreBanco(x)
    Dim filas       As Long: filas = ultimaFila(1)
    
    For x = 7 To filas
        
        If UCase(Cells(x, 3).Value) = UCase("BBVA Bancomer, S.A.") Then
            Cells(x, 3).Value = "BBVA"
            
        End If
        
    Next x
    
End Sub

Private Sub formatoReporte()
    
    Range("A1:I1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("E6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
    End With
    Columns("H:I").Select
    Range("H2").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
End Sub
