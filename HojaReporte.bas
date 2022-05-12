Attribute VB_Name = "HojaReporte"

'HOJA DE REPORTE

Private Sub TituloReporte()
    Range("A1").Select
    Range("A1") = "ANEXO 24 - HYPERION IN"
    Range("A1:A6").Select
    Selection.Insert Shift:=xlToRight
End Sub
Private Sub CortaColumna()
    
    ActiveCell.EntireColumn.Select
    ActiveCell.EntireColumn.Cut
End Sub
Private Sub CopiaColumna()
    ActiveCell.EntireColumn.Select
    ActiveCell.EntireColumn.Copy
End Sub

Private Sub InsertaColumna()
    
    ActiveCell.EntireColumn.Select
    Selection.Insert
    
End Sub
Private Sub EliminaColumna()
    
    ActiveCell.EntireColumn.Select
    Selection.Delete
End Sub
Private Sub NuevoArchivo()
    Dim nombre, archivo As String        ' variable para guardar el nombre del archivo
    'ingreso el nombre sin extensión
    nombre = InputBox("Ingrese nombre del archivo SIN ESPACIOS Y" & Chr(13) & "Sin extencion", "Mr Flaisenber")
    'lo guardo con el nombre especicado y pongo la extención
    If Len(nombre) = 0 Then
        MsgBox "No se ingresó nombre del archivo", vbOKOnly, "Mr Flaisenber"
        nombre = InputBox("Ingrese nombre del archivo SIN ESPACIOS Y" & Chr(13) & "Sin extencion", "Mr Flaisenber")
        If Len(nombre) = 0 Then
            MsgBox "No se ingresó nombre del archivo", vbOKOnly, "Mr Flaisenber"
            nombre = InputBox("Ingrese nombre del archivo SIN ESPACIOS Y" & Chr(13) & "Sin extencion", "Mr Flaisenber")
            '               Exit Sub
        End If
        '        Exit Sub
    End If
    'formo un path con el nombre del archivo
    nombre = Application.ActiveWorkbook.Path + "\" + nombre + ".xlsx"
    'aca almaceno un resultado para comprobar si existe o no un archivo
    archivo = Dir(nombre, vbArchive)
    
    If Len(archivo) = 0 Then
        
        'No Existe
        Application.ActiveSheet.Copy        'copio la hoja activa
        ActiveWorkbook.SaveAs Filename:=Trim(nombre)
        
    Else
        'existe
        If MsgBox("Sobre escribe el archivo", vbYesNo + _
        vbExclamation, "Flaisenber") = vbYes Then
        Application.DisplayAlerts = False
        Application.ActiveSheet.Copy        'copio la hoja activa
        ActiveWorkbook.SaveAs Filename:=Trim(nombre)
        
    End If
End If

End Sub

Private Sub AjustaTexto()
    'ajusta el texto de las celdas
    ActiveSheet.Range("A7:BH7").WrapText = True
    
End Sub
Private Sub LimpiaPlantilla()
    
    Cells.Select
    Selection.Delete
    Cells.Select
    
End Sub
Sub GeneraReporte(n As String)
    
    TituloReporte
    
    'VARIABLES PARA MODIFICAR TITULOS DE COLUMNAS
    FilaInicio = 7
    TituloPedimento = 2
    TituloValorcomercialMN = 13
    TituloValorAduanaMN = 14
    TituloDescargo = 16
    TituloNota1PediEncabeza = 17
    TituloPrecioUnitario = 31
    TituloNota1PediDetalle = 35
    TituloFormaPagoIGI = 36
    TituloNota2PediDetalle = 47
    TituloFactorConversion = 55
    ColumnaNicoCorta = 27
    ColumnaNicoPega = 55
    ColumnaNicoRegresaCorta = 70
    ColumnaNicoRegresaPega = 27
    ColumnaFecEntrada = 64
    ColumnaRegimen = 65
    ColumnaTipoOpe = 66
    ColumnaIDPC = 67
    ColumnaIDIM = 68
    ColumnaIEPS = 69
    ColumnaFormaIEPS = 70
    ColumnaIVAPRV = 71
    ColumnaConsecutivo = 72
    
    'ELIMINAR COLUMNAS
    EliminaColumnaDescri = 27
    EliminaColumnaPNeto = 51
    EliminaColumnaPBruto = 51
    EliminaReferencia = 1
    EliminaConcatena = 58
    EliminaUMTValor = 58
    EliminaConsecutivo = 70
    EliminaNombreProv = 21
    EliminaValorUSD = 32
    EliminaValorComer = 32
    EliminaCalculoEB = 39
    EliminaRevision = 40
    
    'COPIA, CORTA Y PEGA COLUMNAS
    CortarPartida = 25
    CortaValorComer = 13
    Pegapartida = 51
    PegaValorComerMN = 34
    CopiaFactorMoneda = 24
    PegaFactorMoneda = 55
    
    'TITULOS DE COLUMNAS INSERTADAS
    TituloPartida = 50
    TituloRecibido = 17
    TituloNombreProv = 22
    TituloValorUSD = 33
    TituloValorComer = 34
    TituloIDEB = 47
    TituloCalculoID = 48
    TituloNota2PediEncabeza = 55
    TituloUniTarifa = 59
    TituloCantTarifa = 60
    TituloFpagoIVA = 61
    
    'INSERTAR COLUMNAS
    InsertaRecibido = 17
    InsertaNombreProv = 22
    InsertaValorUSD = 33
    InsertaIDEB = 47
    InsertaCalculoID = 48
    InsertaNota2PediEncabeza = 55
    InsertaUniTarifa = 59
    InsertaCantTarifa = 60
    InsertaFpagoIVA = 61
    FilaConcatenaInsert = 59
    ColConcatenaInsert = 59
    InsertaUMTLetra = 61
    
    'INICIALIZAN VARIABLES PARA CALCULOS Y CONCATENAR
    Filapadre = 8
    FilaReferencia = 1
    FilaFactura = 20
    FilaProducto = 25
    FilaSecuencia = 57
    FilaTipoCambio = 5
    FilaValorComer = 34
    FilaValorUSD = 33
    FilaUnidadMTInsert = 60
    FilaCantidadMTInsert = 62
    FilaNombreProv = 22
    FilaFormaPagoIVA = 63
    FilaPrevalidacion = 19
    IncrementaConsecu = 0
    
    'COPIA Y PEGA COLUMNA DE NICO
    Cells(FilaInicio, ColumnaNicoCorta).Select
    CortaColumna
    Cells(FilaInicio, ColumnaNicoPega).Select
    InsertaColumna
    
    'CAMBIA TITULOS DE COLUMNAS
    Cells(FilaInicio, TituloPedimento) = "PEDIMENTO"
    Cells(FilaInicio, TituloDescargo) = "DESCARGO X FECHA FACTURA"
    Cells(FilaInicio, TituloValorcomercialMN) = "VALOR COMERCIAL MN PEDIMENTO"
    Cells(FilaInicio, TituloValorAduanaMN) = "VALOR ADUANA MN PEDIMENTO"
    Cells(FilaInicio, TituloNota1PediEncabeza) = "NOTA INTERNA 1 PEDIMENTO (ENCABEZADO)"
    Cells(FilaInicio, TituloPrecioUnitario) = "PRECIO UNITARIO"
    Cells(FilaInicio, TituloNota1PediDetalle) = "NOTA INTERNA 1 PEDIMENTO (DETALLE)"
    Cells(FilaInicio, TituloFormaPagoIGI) = "FORMA PAGO IGI"
    Cells(FilaInicio, TituloNota2PediDetalle) = "NOTA INTERNA 2 PEDIMENTO (DETALLE)"
    'Elimina columnas
    Cells(FilaInicio, EliminaColumnaDescri).Select
    EliminaColumna
    Cells(FilaInicio, EliminaColumnaPNeto).Select
    EliminaColumna
    Cells(FilaInicio, EliminaColumnaPBruto).Select
    EliminaColumna
    'CORTA Y PEGA EN OTRA UBICACION COLUMNAS
    Cells(FilaInicio, CortarPartida).Select
    CortaColumna
    Cells(FilaInicio, Pegapartida).Select
    InsertaColumna
    Cells(FilaInicio, TituloPartida) = "SECUENCIA"
    'Inserta columnas nuevas
    Cells(FilaInicio, InsertaRecibido).Select
    InsertaColumna
    Cells(FilaInicio, TituloRecibido) = "RECIBIDO DE"
    Cells(FilaInicio, InsertaNombreProv).Select
    InsertaColumna
    Cells(FilaInicio, TituloNombreProv) = "NOMBRE PROVEEDOR"
    Cells(FilaInicio, InsertaValorUSD).Select
    InsertaColumna
    Cells(FilaInicio, TituloValorUSD) = "VALOR USD"
    'COPIA Y PEGA COLUMNAS
    Cells(FilaInicio, TituloValorcomercialMN).Select
    CopiaColumna
    Cells(FilaInicio, PegaValorComerMN).Select
    InsertaColumna
    Cells(FilaInicio, TituloValorComer) = "VALOR COMERCIAL"
    Cells(FilaInicio, CopiaFactorMoneda).Select
    CopiaColumna
    Cells(FilaInicio, PegaFactorMoneda).Select
    InsertaColumna
    Cells(FilaInicio, TituloFactorConversion) = "FACTOR DE CONVERSION"
    'INSERTA COLUMNAS SEGUNDA PARTE
    Cells(FilaInicio, InsertaIDEB).Select
    InsertaColumna
    Cells(FilaInicio, TituloIDEB) = "ID EB"
    Cells(FilaInicio, InsertaCalculoID).Select
    InsertaColumna
    Cells(FilaInicio, TituloCalculoID) = "CALCULO PARA PARTIDAS EB"
    
    Cells(FilaInicio, InsertaNota2PediEncabeza).Select
    InsertaColumna
    Cells(FilaInicio, TituloNota2PediEncabeza) = "NOTA INTERNA 2 PEDIMENTO (Encabezado)"
    Cells(FilaInicio, InsertaUniTarifa).Select
    InsertaColumna
    Cells(FilaInicio, TituloUniTarifa) = "UNIDAD TARIFA"
    Cells(FilaInicio, InsertaCantTarifa).Select
    InsertaColumna
    Cells(FilaInicio, TituloCantTarifa) = "CANTIDAD TARIFA"
    Cells(FilaInicio, InsertaFpagoIVA).Select
    InsertaColumna
    Cells(FilaInicio, TituloFpagoIVA) = "FP IVA"
    Cells(FilaInicio, ColConcatenaInsert).Select
    InsertaColumna
    Cells(FilaInicio, ColConcatenaInsert) = "Concatena"
    Cells(FilaInicio, InsertaUMTLetra).Select
    InsertaColumna
    Cells(FilaInicio, InsertaUMTLetra) = "UNIDAD DE TARIFA"
    
    '-----ULTIMAS COLUMNAS
    Cells(FilaInicio, ColumnaFecEntrada).Select
    InsertaColumna
    Cells(FilaInicio, ColumnaFecEntrada) = "FECHA ENTRADA"
    Cells(FilaInicio, ColumnaRegimen).Select
    InsertaColumna
    Cells(FilaInicio, ColumnaRegimen) = "REGIMEN"
    Cells(FilaInicio, ColumnaTipoOpe).Select
    InsertaColumna
    Cells(FilaInicio, ColumnaTipoOpe) = "TIPO OPERACION"
    Cells(FilaInicio, ColumnaIDPC).Select
    InsertaColumna
    Cells(FilaInicio, ColumnaIDPC) = "IDENTIFICADOR PC"
    Cells(FilaInicio, ColumnaIDIM).Select
    InsertaColumna
    Cells(FilaInicio, ColumnaIDIM) = "IDENTIFICADOR IM"
    Cells(FilaInicio, ColumnaIEPS).Select
    InsertaColumna
    Cells(FilaInicio, ColumnaIEPS) = "IEPS"
    Cells(FilaInicio, ColumnaFormaIEPS).Select
    InsertaColumna
    Cells(FilaInicio, ColumnaFormaIEPS) = "Forma pago IEPS"
    Cells(FilaInicio, ColumnaIVAPRV).Select
    InsertaColumna
    Cells(FilaInicio, ColumnaIVAPRV) = "IVA/PRV"
    Cells(FilaInicio, ColumnaConsecutivo).Select
    InsertaColumna
    Cells(FilaInicio, ColumnaConsecutivo) = "Consecutivo"
    
    'ORDENA COLUMAS DE REFERENCIA, SECUENCIA, FACTURA
    'Range("A7:BU400").Sort key1:=Range("A8"), Order1:=xlAscending, Header:=xlYes, Key2:=Range("Y8"), Order2:=xlAscending, Header:=xlYes, Key3:=Range("T8"), Order3:=xlAscending, Header:=xlYes
    
    'INSERTA DATOS EN COLUMNA "CONCATENAR" EN HOJA PRINCIPAL (referencia+factura+num parte+secuencia)
    u = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 8 To u Step 1
        'LLENA COLUMNA DE CONSECUTIVO
        IncrementaConsecu = IncrementaConsecu + 1
        Cells(Filapadre, ColumnaConsecutivo) = IncrementaConsecu
        
        'CONCATENACION DE REFERENCI,FACTURA, PRODUCTO, SECUENCIA, CONSECUTIVO
        Cells(Filapadre, FilaConcatenaInsert) = Cells(Filapadre, FilaReferencia) & Cells(Filapadre, FilaFactura) & Cells(Filapadre, FilaProducto) & Cells(Filapadre, FilaSecuencia)        '& Cells(Filapadre, ColumnaConsecutivo)
        'A continuación se hace se calcula y registra el Valor Dolares
        Cells(Filapadre, FilaValorUSD) = Cells(Filapadre, FilaValorComer) / Cells(Filapadre, FilaTipoCambio)
        Cells(Filapadre, FilaPrevalidacion) = "240"
        Cells(Filapadre, ColumnaTipoOpe) = "IMP"
        Cells(Filapadre, ColumnaIVAPRV) = "38"
        Filapadre = Filapadre + 1
    Next
    Filapadre = 8
    
    'RUTINA PARA TRAER DATOS DE LA HOJA "COMPLEMENTO"
    'TAMBIEN CONVIERTE LA UMT EN LETRA
    u = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 8 To u Step 1
        Cells(Filapadre, FilaUnidadMTInsert).Value = WorksheetFunction.VLookup((Cells(Filapadre, FilaConcatenaInsert).Value), Sheets("COMPLEMENTO").Range("B7:CM400"), 65, False)
        Cells(Filapadre, FilaCantidadMTInsert).Value = WorksheetFunction.VLookup((Cells(Filapadre, FilaConcatenaInsert).Value), Sheets("COMPLEMENTO").Range("B7:CM400"), 66, False)
        'Cells(Filapadre, FilaNombreProv).Value = WorksheetFunction.VLookup((Cells(Filapadre, FilaConcatenaInsert).Value), Sheets("COMPLEMENTO").Range("B7:CL250"), 88, False)
        Cells(Filapadre, FilaFormaPagoIVA).Value = WorksheetFunction.VLookup((Cells(Filapadre, FilaConcatenaInsert).Value), Sheets("COMPLEMENTO").Range("B7:CM400"), 27, False)
        Cells(Filapadre, ColumnaFecEntrada).Value = WorksheetFunction.VLookup((Cells(Filapadre, FilaConcatenaInsert).Value), Sheets("COMPLEMENTO").Range("B7:CM400"), 21, False)
        Cells(Filapadre, ColumnaFecEntrada).NumberFormat = "m/d/yyyy"
        Cells(Filapadre, ColumnaRegimen).Value = WorksheetFunction.VLookup((Cells(Filapadre, FilaConcatenaInsert).Value), Sheets("COMPLEMENTO").Range("B7:CM400"), 4, False)
        Cells(Filapadre, ColumnaIEPS).Value = WorksheetFunction.VLookup((Cells(Filapadre, FilaConcatenaInsert).Value), Sheets("COMPLEMENTO").Range("B7:CM400"), 32, False)
        
        If Cells(Filapadre, FilaUnidadMTInsert) = 1 Then
            Cells(Filapadre, InsertaUMTLetra) = "KGS"
        Else
            If Cells(Filapadre, FilaUnidadMTInsert) = 2 Then
                Cells(Filapadre, InsertaUMTLetra) = "GRS"
            Else
                If Cells(Filapadre, FilaUnidadMTInsert) = 6 Then
                    Cells(Filapadre, InsertaUMTLetra) = "PZA"
                Else
                    If Cells(Filapadre, FilaUnidadMTInsert) = 11 Then
                        Cells(Filapadre, InsertaUMTLetra) = "MIL"
                    Else
                        Cells(Filapadre, InsertaUMTLetra) = "POR CLASIFICAR"
                    End If
                End If
            End If
        End If
        
        Filapadre = Filapadre + 1
    Next
    
    'ELIMINA COLUMNAS SEGUNDA PARTE
    Cells(FilaInicio, EliminaReferencia).Select
    EliminaColumna
    Cells(FilaInicio, EliminaConcatena).Select
    EliminaColumna
    Cells(FilaInicio, EliminaUMTValor).Select
    EliminaColumna
    
    'REGRESA COLUMNA NICO A SU LUGAR DE ORIGEN
    Cells(FilaInicio, ColumnaNicoRegresaCorta).Select
    CortaColumna
    Cells(FilaInicio, ColumnaNicoRegresaPega).Select
    InsertaColumna
    
    'ELIMINA COLUMNAS DE APOYO Y NO REQUERIDAS
    Cells(FilaInicio, EliminaConsecutivo).Select
    EliminaColumna
    Cells(FilaInicio, EliminaNombreProv).Select
    EliminaColumna
    Cells(FilaInicio, EliminaValorUSD).Select
    EliminaColumna
    Cells(FilaInicio, EliminaValorComer).Select
    EliminaColumna
    'ELIMINA RANGO DE PAGO TLCAN HASTA JUSTICACION TLCUEM 6 COLUMNAS SEGUIDAS
    Range("AL:AQ").Select
    Selection.Delete
    Cells(FilaInicio, EliminaCalculoEB).Select
    EliminaColumna
    Cells(FilaInicio, EliminaRevision).Select
    EliminaColumna
    
    AjustaTexto
    'NuevoArchivo
    
End Sub

Private Sub CommandButton1_Click()
    
    GeneraReporte
    
End Sub

Private Sub CommandButton2_Click()
    LimpiaPlantilla
End Sub

