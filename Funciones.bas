Attribute VB_Name = "Funciones"
' Funciones y procedimientos para trabajar con proyectos

Private Sub irACeldas()
    
    Dim x           As Integer
    Dim y           As Integer
    
    x = InputBox("x: ")
    y = InputBox("y: ")
    
    Cells(x, y).Select
    
End Sub

''Ordenamiento

Sub ordenar1Criterio(keyuno As String)
    
    Selection.Sort key1:=Range(keyuno), Order1:=xlAscending, Header:=xlYes
    
End Sub

' 3 criterios
Sub ordenar(c1 As String, c2 As String, c3 As String, c4 As String, fila As Long)
    
    Dim columnas    As Long: columnas = ultimaColunma(1)
    Dim filas       As Long: filas = ultimaFila(1)
    Dim finalCol    As String
    
    finalCol = extraerCadena(Cells(fila, columnas).Address, "\w+") & filas
    
    Range(c1 & ":" & finalCol).Select
    
    Range(c1 & ":" & finalCol).Sort _
             key1:=Range(c2), Order1:=xlAscending, Header:=xlYes, _
             Key2:=Range(c2), Order2:=xlAscending, Header:=xlYes, _
             Key3:=Range(c3), Order3:=xlAscending, Header:=xlYes
    
End Sub

Sub ordenarDosCriterios(rango As String, c1 As String, c2 As String, fila As Long)
    
    Dim columnas    As Long: columnas = ultimaColunma(1)
    Dim filas       As Long: filas = ultimaFila(1)
    Dim finalCol    As String
    
    finalCol = extraerCadena(Cells(fila, columnas).Address, "\w+") & filas
    
    Range(rango & ":" & finalCol).Select
    
    Range(rango & ":" & finalCol).Sort _
                key1:=Range(c1), Order1:=xlAscending, Header:=xlYes, _
                Key2:=Range(c2), Order2:=xlAscending, Header:=xlYes
    
End Sub

Function busquedaValor(fila As Long, valor As Variant, hoja As Integer)
    
    'Se devuelve la posicion del dato buscado de acuerdo a una fila dada.
    
    Dim totalColumnas As Long: totalColumnas = ultimaColunma(hoja)
    
    For i = 1 To totalColumnas
        
        If UCase(Cells(fila, i).Value) = UCase(valor) Then
            
            busquedaValor = i
            Exit For
            
        Else
            
            busquedaValor = 0
            
        End If
        
    Next i
    
End Function

Function busquedaValorNveces(fila As Long, valor As Variant, n As Long)
    
    'Se devuelve la posicion del dato buscado de acuerdo a una fila dada.
    
    Dim totalColumnas As Long: totalColumnas = ultimaColunma(1)
    Dim cont        As Long: cont = 0
    
    For i = 1 To totalColumnas
        
        If UCase(Cells(fila, i).Value) = UCase(valor) Then
            
            cont = cont + 1
            
            If cont = n Then
                busquedaValorNveces = i
                Exit For
                
            End If
            
        Else
            
            busquedaValorNveces = 0
            
        End If
        
    Next i
    
End Function

Function busquedaValorCol(columna As Long, valor As Variant)
    
    'Se devuelve la posicion del dato buscado de acuerdo a una columna dada.
    
    Dim totaFilas   As Long: totalFilas = ultimaFila(1)
    
    For i = 1 To totalFilas
        
        If UCase(Cells(i, columna).Value) = UCase(valor) Then
            
            busquedaValorCol = i
            Exit For
            
        Else
            
            busquedaValorCol = 0
            
        End If
    Next i
    
End Function

'Devuelve la posición de acuerdo a la busqueda con regex
Function busquedaValorColExp(columna As Long, valor As String)
    
    'Se devuelve la posicion del dato buscado de acuerdo a una fila dada.
    
    Dim totaFilas   As Long: totalFilas = ultimaFila(1)
    
    For i = 1 To totalFilas
        
        If extraerCadena(Cells(i, columna).Value, valor) <> "" Then
            
            busquedaValorColExp = i
            Exit For
            
        Else
            
            busquedaValorColExp = 0
            
        End If
        
    Next i
    
End Function

Function ultimaColunma(numeroHoja As Integer)
    
    ultimaColunma = Worksheets(numeroHoja).UsedRange.SpecialCells(xlCellTypeLastCell).Column
    
End Function

Function ultimaFila(numeroHoja As Integer)
    
    'Se obtiene la ultima fila con datos
    
    ultimaFila = Worksheets(numeroHoja).UsedRange.SpecialCells(xlCellTypeLastCell).Row
    
End Function

Function extraerCadena(cadena As Variant, patron As String)
    
    'Se extrae la cadena de acuerdo a un patron
    
    Dim cadenaObj       As New RegExp
    
    cadenaObj.Pattern = patron
    
    If cadenaObj.Test(cadena) Then
        extraerCadena = cadenaObj.Execute(cadena)(0)
    Else
        extraerCadena = ""
    End If
    
End Function

' Macro para quitar error de tipo de dato no reconocido
Sub formatearCelda(columna As String)
    
    Selection.TextToColumns Destination:=Range(columna & "1"), DataType:=xlDelimited, _
                            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
                            :=Array(1, 1), TrailingMinusNumbers:=True
    
End Sub

'Se coloca un estilo a la celda o una selección
Sub formatoCelda(colorFondo As Long, colorLetra As Long, negrita As Boolean, centrado As Variant, alText As Boolean)
    
    With Selection
        .Interior.ColorIndex = colorFondo
        .Font.ColorIndex = colorLetra
        .Font.Bold = negrita
        .HorizontalAlignment = centrado
        .WrapText = alText
        .Font.Name = "Calibri"
        .Font.Size = 11
    End With
    
End Sub

'Devuelve la direccion de la celda de acuerdo a una coordenada 1,2 = B
Function posicionCelda(x As Long, y As Long)
    
    posicionCelda = Cells(x, y).Address
    posicionCelda = Replace(posicionCelda, "$", "")
    posicionCelda = extraerCadena(posicionCelda, "\D+")
    
End Function

Sub InsertarColumnasNveces(columna As Long, fin As Long)
    
    For x = 1 To fin
        
        Call insertarColumna(1, columna, xlLeft)
        
    Next x
    
End Sub

Sub InsertarFilasNveces(fila As Long, fin As Long)
    
    For x = 1 To fin
        
        Call insertarFila(1, fila, xlLeft)
        
    Next x
    
End Sub

Sub guardarLibro(x)
    
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
    
End Sub

Sub cerrarLibro(x)
    
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    
End Sub

Sub cerrarLibros(nombreLibro As String)
    
    Call activarLibro(nombreLibro)
    Call cerrarLibro("")
    
End Sub

Sub ArchivoExiste(nombreLibro As String)
    
    If Dir(ActiveWorkbook.Path & "\" & nombreLibro) <> "" Then
        
        Kill (ActiveWorkbook.Path & "\" & nombreLibro)
        
    End If
    
    Call crearLibro(nombreLibro)
    
End Sub

'Se selecciona una posicion y apartir de ahi se hace una seleccion hasta encontrar un espacio en blanco
Sub rangoDelimitado(x As Long, y As Long)
    
    Cells(x, y).Select
    Range(Selection, Selection.End(xlDown)).Select
    
End Sub

'Se guarda el libro actual con el nombre que se le de
Sub guardarComo(x)
    
    ActiveWorkbook.SaveAs Filename:=Application.ActiveWorkbook.Path & "\" & InputBox("Nombre para el nuevo libro: ", "Guardar")
    
End Sub

Sub combinarCentrar(rango As String)
    
    Range(rango).Select
    
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    
End Sub

Sub ajustarCeldas(x)
    
    Cells.EntireRow.AutoFit
    Cells.EntireColumn.AutoFit
    
End Sub

Sub crearLibro(nombreArchivo As String)
    
    Dim nombre      As String
    
    'Se obtiene la ruta del libro activo para guardar el libro nuevo
    nombre = Application.ActiveWorkbook.Path + "\" + nombreArchivo        '+ ".xlsx"
    
    'Se crea y guarda el libro en la ruta obtenida
    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:=nombre
    
End Sub

Sub limpiarColumna(x As Long, y As Long)
    
    Cells(x, y).EntireColumn.Clear
    
End Sub

Sub insertarHoja(nombreHoja As String)
    
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = nombreHoja
    
End Sub

Sub borrarHoja(nombreHoja As String)
    
    Worksheets(nombreHoja).Delete
    
End Sub

Sub cambiarNombreHoja(indice As Integer, nombreHoja As String)
    
    Worksheets(indice).Name = nombreHoja
    
End Sub

Sub activarLibro(nombreLibro As String)
    
    Workbooks(nombreLibro).Activate
    
End Sub

Sub copiarHoja(libroDestino As String, hojalibroDestino As Variant)
    
    Cells.Select
    Selection.Copy Destination:=Workbooks(libroDestino).Worksheets(hojalibroDestino).Cells(1, 1)
    
End Sub

Sub copiarA(hoja As Integer, x As Long, y As Long)
    
    Selection.Copy Destination:=Sheets(hoja).Cells(x, y)
    
End Sub

Sub cortarA(hoja As Integer, x As Long, y As Long)
    
    Selection.Cut Destination:=Sheets(hoja).Cells(x, y)
    
End Sub
Sub copiarADestino(libro As String, hoja As Integer, x As Long, y As Long)
    
    Selection.Copy Destination:=Workbooks(libro).Sheets(hoja).Cells(x, y)
    
End Sub

Sub activarHoja(nombreHoja As Variant)
    
    Sheets(nombreHoja).Activate
    
End Sub

Sub seleccionarColumnaPosicion(x As Long, y As Long)
    
    Cells(x, y).EntireColumn.Select
    
End Sub

Sub EliminaColumna(x As Long, y As Long)
    
    Cells(x, y).EntireColumn.Delete
    
End Sub

Sub seleccionarFilaPosicion(x As Long, y As Long)
    
    Cells(x, y).EntireRow.Select
    
End Sub

Sub borrarColumna(x As Long, y As Long)
    
    Cells(x, y).EntireColumn.Delete
    
End Sub

Sub borrarFilas(fila As Long, columna As Long)
    
    Cells(fila, columna).EntireRow.Delete
    
End Sub

Sub CopiarColumna(x)
    
    Selection.Copy
    
End Sub

Sub CortaColumna(x)
    
    Selection.Copy
    
End Sub

Sub PegaColumna(x)
    
    ActiveSheet.Paste
    
End Sub

Sub EliminarDuplicados(x)
    
    Selection.RemoveDuplicates Columns:=1, Header:= _
                               xlYes
    
End Sub

Sub SeleccionaColumna(columna As String)
    
    Range(columna & ":" & columna).Select
    
End Sub

Sub EliminarColumna(x)
    
    Selection.Delete
    
End Sub

Sub seleccionaRango(columnaIni As String, columnaFin As String, x As Long, y As Long)

    Range(columnaIni & x & ":" & columnaFin & y).Select
    
End Sub

Sub insertarFila(fila As Long, columna As Long)
    
    Cells(fila, columna).EntireRow.Select
    Selection.Insert Shift:=xlDown
    
End Sub

Sub insertarColumna(fila As Long, columna As Long, direccion As Variant)
    
    Cells(fila, columna).EntireColumn.Select
    Selection.Insert Shift:=direccion
    
End Sub

Sub SeleccionarRango(rango As String)
    
    Range(rango).Select
    
End Sub

Sub cp(hoja As Integer, fila As Long, columna As Long, libro As String)
    
    Selection.Copy Destination:=Workbooks(libro).Worksheets(hoja).Cells(fila, columna)
    
End Sub

Sub cambiarNomHoja(hoja As String, nuevoNombre As String)
    
    Sheets(hoja).Name = nuevoNombre
    
End Sub

Sub acelerarProceso(x)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.Calculation = xlCalculationManual
    
End Sub

Sub configOriginal(x)
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub

