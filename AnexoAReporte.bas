Attribute VB_Name = "AnexoAReporte"

Sub concatenaFraccion(fila As Long)
    
    'Se concatenan las columnas de FRACCION IMPORTACION con la de NICO y se elimina esta ultima junto con la de TIPO BIEN.
    
    Dim F_I, NICO   As Long
    Dim conteoFilas As Long
    
    F_I = busquedaValor(fila, "FRACCION IMPORTACION", 1)
    NICO = busquedaValor(fila, "NICO", 1)
    conteoFilas = ultimaFila(1)
    
    For i = 8 To conteoFilas
        
        Cells(i, F_I).Value = Cells(i, F_I).Value & "." & Cells(i, NICO).Value
        
    Next i
    
    Call borrarColumna(fila, NICO)
    Call borrarColumna(fila, busquedaValor(7, "TIPO BIEN", 1))
    
End Sub

Sub añadirColumnas(nombreLibro As String)
    
    Dim columnas(10) As String
    Dim conteoFilas As Long: conteoFilas = ultimaFila(1)
    Dim indice, indAux As Integer: indice = 0
    Dim IVAPRV      As Long
    Dim valor       As Variant
    Dim indAux1     As Integer
    
    IVAPRV = busquedaValor(7, "IVA/PRV", 1) + 1
    indAux = 2
    indAux1 = 8
    
    Call listaColumnas(columnas)
    
    For i = IVAPRV To (IVAPRV + UBound(columnas))
        
        Cells(7, i) = columnas(indice)
        
        For x = 8 To conteoFilas
            
            If indice = 0 Then
                
                Call activarLibro("Q.xls")
                valor = Cells(indAux, busquedaValor(1, "VAL_MONEFAC", 1))
                
                Call activarLibro(nombreLibro)
                Cells(x, i) = valor
                
            ElseIf indice = 1 Then
                
                Call activarLibro("Q.xls")
                valor = Cells(indAux, busquedaValor(1, "VAL_EXTR", 1))
                
                Call activarLibro(nombreLibro)
                Cells(x, i) = valor
                
            ElseIf indice = 2 Then
                
                Call activarLibro("Anexo 24 – Imp y Expo.xlsx")
                valor = Cells(x, busquedaValor(7, "VALOR DOLARES", 1))
                
                Call activarLibro(nombreLibro)
                Cells(x, i) = valor
                
            ElseIf indice = 10 Then
                
                Call activarLibro("Q.xls")
                valor = Cells(indAux, busquedaValor(1, "NUM_REFE", 1))
                
                Call activarLibro(nombreLibro)
                Cells(x, i) = valor
                
            Else
                Cells(x, i) = 0
            End If
            
            indAux = indAux + 1
            indAux1 = indAux1 + 1
        Next x
        indAux = 2
        indAux1 = 9
        indice = indice + 1
        
    Next i
    
    Erase columnas
    
End Sub


Sub ObtenerMes(mess As String)
    
    Dim meses(11)   As String
    Dim fecha(5)    As Integer
    Dim mesStri(1)  As String
    
    Call ExtraerFecha(fecha)
    Call listaMeses(meses)
    
    mesStri(0) = mes(meses, fecha(1))
    mesStri(1) = mes(meses, fecha(4))
    
    If fecha(2) = fecha(5) Then
        
        mess = fecha(0) & mesStri(0) & "-" & fecha(3) & mesStri(1) & fecha(2)
    Else
        mess = fecha(0) & mesStri(0) & fecha(2) & "-" & fecha(3) & mesStri(1) & fecha(5)
    End If
    
    Erase meses
    Erase fecha
    Erase mesStri
    
End Sub

Function mes(meses() As String, mesDig As Integer)
    
    For x = 0 To 11
        
        If x = mesDig - 1 Then
            
            mes = meses(x)
            
            Exit For
            
        End If
        
    Next x
    
End Function


Sub listaColumnas(lista() As String)
    
    lista(0) = "Moneda"
    lista(1) = "Valor Moneda Factura"
    lista(2) = "Valor Dolares"
    lista(3) = "IGI MN Pedimento"
    lista(4) = "Identificar MS"
    lista(5) = "Transporte Decrementables"
    lista(6) = "Seguro Decrementables"
    lista(7) = "Carga"
    lista(8) = "Descarga"
    lista(9) = "Otros Decrementables"
    lista(10) = "REFERENCIA"
    
End Sub

Sub listaMeses(lista() As String)
    
    lista(0) = "ene"
    lista(1) = "feb"
    lista(2) = "mar"
    lista(3) = "abr"
    lista(4) = "may"
    lista(5) = "jun"
    lista(6) = "jul"
    lista(7) = "ago"
    lista(8) = "sep"
    lista(9) = "oct"
    lista(10) = "nov"
    lista(11) = "dic"
    
End Sub

Sub listaMesesCom(lista() As String)
    
    lista(0) = "enero"
    lista(1) = "febrero"
    lista(2) = "marzo"
    lista(3) = "abril"
    lista(4) = "mayo"
    lista(5) = "junio"
    lista(6) = "julio"
    lista(7) = "agosto"
    lista(8) = "septiembre"
    lista(9) = "octubre"
    lista(10) = "noviembre"
    lista(11) = "diciembre"
    
End Sub
