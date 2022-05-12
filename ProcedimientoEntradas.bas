Attribute VB_Name = "ProcedimientoEntradas"
Sub GenerarReporteIN()
    
    Dim nombreLibro  As String
    Dim libroReporte As String
    Dim mesStri     As String
    
    'Acelerar proceso
    Call acelerarProceso("")
    
    'Se obtiene la fecha para el renombrado del reporte
    Call activarLibro("Anexo 24 – Imp y Expo.xlsx")
    Call ObtenerMes(mesStri)
    
    nombreLibro = "Reporte Hyperion IN Anexo24"
    libroReporte = "ReporteAnexo24EntradaHyperion" & mesStri
    
    'Se crea el libro de entradas
    Call ArchivoExiste(nombreLibro & ".xlsx")
    Call cambiarNombreHoja(1, "REPORTE")
    Call insertarHoja("COMPLEMENTO")
    Call activarLibro("Anexo 24 – Imp y Expo.xlsx")
    Call ordenar("A7", "B8", "CI8", "CI8", 7)
    Call eliminarCelda("Pedimentos Normales")
    Call copiarHoja(nombreLibro & ".xlsx", "COMPLEMENTO")
    Call activarLibro("Reporte Anexo 24 - Grecargo.xlsx")
    Call eliminarCelda("Pedimentos Normales")
    Call ordenar("A7", "A8", "Y8", "Y8", 7)
    Call copiarHoja(nombreLibro & ".xlsx", "REPORTE")
    Call activarLibro(nombreLibro & ".xlsx")
    
    'Se trabaja con la hoja de complemento
    Call complemento
    'Se trabaja con la hoja de reporte
    Call reporte
    'Se  copia la hoja de reporte a un nuevo libro
    
    Call ArchivoExiste(libroReporte & ".xlsx")
    Call cambiarNombreHoja(1, "REPORTE")
    Call activarLibro(nombreLibro & ".xlsx")
    'Call ordenar("A7", "W8", "S8", "A8", 7)
    Call copiarHoja(libroReporte & ".xlsx", "REPORTE")
    
    'Se complementa el libro del reportefinal
    Call activarLibro(libroReporte & ".xlsx")
    Call concatenaFraccion(7)
    
    'Se ordena la hoja del query
    Call activarLibro("Q.xls")
    Call ordenar("A1", "C2", "N2", "N2", 1)
    
    Call activarLibro(libroReporte & ".xlsx")
    Call añadirColumnas(libroReporte & ".xlsx")
    
    Call activarLibro(libroReporte & ".xlsx")
    
    Call colocarValorUnit(libroReporte & ".xlsx")
    
    Call activarLibro(libroReporte & ".xlsx")
    Call formatoCelda
    'Call ordenar("A7", "A8", "AR8", "W8", 7)
    
    'Cerramos libros innecesarios
    
    Call activarLibro("Anexo 24 – Imp y Expo.xlsx")
    Call cerrarLibro("")
    Call activarLibro("Reporte Anexo 24 - Grecargo.xlsx")
    Call cerrarLibro("")
    Call activarLibro("Q.xls")
    Call cerrarLibro("")
    
    'Se guardan los libros creados
    Call activarLibro(nombreLibro & ".xlsx")
    Call guardarLibro("")
    Call cerrarLibro("")
    Call activarLibro(libroReporte & ".xlsx")
    Call formatoFecha(1)
    Call configOriginal("")
    
    Call guardarLibro("")

End Sub

Private Sub colocarValorUnit(libro As String)
    
    Dim vUnit       As Long
    Dim filas       As Long: filas = ultimaFila(1)
    Dim q           As Long: q = ultimaFila(1)
    Dim cont        As Long: cont = 2
    Dim temp1, temp2 As Double
    
    vUnit = busquedaValor(7, "PRECIO UNITARIO", 1)
    
    Call activarLibro("Q.xls")
    q = busquedaValor(1, "VAL_UNIT", 1)
    
    For x = 8 To filas
        
        Call activarLibro("Q.xls")
        temp1 = Cells(cont, q)
        
        Call activarLibro(libro)
        Cells(x, vUnit) = temp1
        
        cont = cont + 1
        
    Next x
    
    Cells(7, vUnit) = "PRECIO UNITARIO EN MONEDA FACTURA"
    Call ajustarCeldas("")
    
End Sub

Private Sub complemento()
    
    Call activarHoja("COMPLEMENTO")
    Call Concatena("")
    Call ajustarCeldas("")
    
End Sub

Private Sub eliminarCelda(valor As String)
    
    Dim posicion    As Long
    
    posicion = busquedaValorCol(1, valor)
    
    If posicion <> 0 Then
        
        Cells(posicion, 1).EntireRow.Delete
        
    End If
    
End Sub

Private Sub reporte()
    
    Call activarHoja("REPORTE")
    Call GeneraReporte("")
    
End Sub

Sub ExtraerFecha(fecha() As Integer)
    
    Dim posicion    As Integer
    Dim fechaInicial As String
    Dim fechaFinal  As String
    
    posicion = busquedaValorColExp(1, "Fecha Inicial:")
    'Se obtiene fecha inicial y final
    fechaInicial = extraerCadena(Cells(posicion, 1).Value, "Fecha Inicial: [0-9]{2}/[0-9]{2}/[0-9]{4}")
    fechaFinal = extraerCadena(Cells(posicion, 1).Value, "Fecha Final: [0-9]{2}/[0-9]{2}/[0-9]{4}")
    
    'Se quitan las /
    fechaInicial = Replace(fechaInicial, "/", "")
    fechaFinal = Replace(fechaFinal, "/", "")
    
    'Se extraen los digitos de cada fecha
    fechaInicial = extraerCadena(fechaInicial, "[0-9]+")
    fechaFinal = extraerCadena(fechaFinal, "[0-9]+")
    
    'dia
    fecha(0) = Mid(fechaInicial, 1, 2)
    fecha(3) = Mid(fechaFinal, 1, 2)
    
    'Mes
    fecha(1) = Mid(fechaInicial, 3, 2)
    fecha(4) = Mid(fechaFinal, 3, 2)
    
    'Año
    fecha(2) = Mid(fechaInicial, 5, 4)
    fecha(5) = Mid(fechaFinal, 5, 4)
    
End Sub

Private Sub formatoCelda()
    
    Dim col         As Long: col = ultimaColunma(1)
    
    Range("A7:" & Replace(Cells(7, col).Address, "$", "")).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    
End Sub
