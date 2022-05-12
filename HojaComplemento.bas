Attribute VB_Name = "HojaComplemento"

' HOJA DE COMPLEMENTO

Private Sub LimpiaPlantilla2()
    
    Cells.Select
    Selection.Delete
    Cells.Select
    
End Sub

Sub Concatena(n As String)
    
    FilaTitulo = 7
    ColConcatena = 2
    FilaComienza = 8
    FilaRefe = 3
    FilaFactura = 86
    FilaParte = 87
    FilaPartida = 88
    cConsecutivo = 90
    IncrementaConsecu = 0
    
    Cells(FilaTitulo, ColConcatena).Select
    ActiveCell.EntireColumn.Select
    Selection.Insert
    Cells(FilaTitulo, ColConcatena) = "Concatena1"
    
    Cells(FilaTitulo, cConsecutivo).Select
    ActiveCell.EntireColumn.Select
    Selection.Insert
    Cells(FilaTitulo, cConsecutivo) = "Consecutivo"
    
    'ORDENA COLUMAS DE REFERENCIA, PARTIDA, FACTURA
    
    'Range("A7:CK400").Sort key1:=Range("C8"), Order1:=xlAscending, Header:=xlYes, Key2:=Range("CI8"), Order2:=xlAscending, Header:=xlYes, Key3:=Range("CH8"), Order3:=xlAscending, Header:=xlYes
    
    u = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 8 To u Step 1
        'LLENA LA COLUMNA DE CONSECUTIVO DE 1 EN 1
        IncrementaConsecu = IncrementaConsecu + 1
        Cells(FilaComienza, cConsecutivo) = IncrementaConsecu
        
        'CONCATENA POR REFERENCIA, FACTURA, NUMERO DE PARTE,PARTIDAD,CONSECUTIVO
        Cells(FilaComienza, ColConcatena) = Cells(FilaComienza, FilaRefe) & Cells(FilaComienza, FilaFactura) & Cells(FilaComienza, FilaParte) & Cells(FilaComienza, FilaPartida)        '& Cells(FilaComienza, cConsecutivo)
        FilaComienza = FilaComienza + 1
    Next
    FilaComienza = 8
    
    'CONCATENA POR REFERENCIA, FACTURA, NUMERO DE PARTE,PARTIDAD,CONSECUTIVO
    'u = Cells(Rows.Count, 1).End(xlUp).Row
    'For i = 8 To u Step 1
    '   Cells(FilaComienza, ColConcatena) = Cells(FilaComienza, FilaRefe) & Cells(FilaComienza, FilaFactura) & Cells(FilaComienza, FilaParte) & Cells(FilaComienza, FilaPartida) & Cells(FilaComienza, cConsecutivo)
    '  FilaComienza = FilaComienza + 1
    'Next
    
End Sub

Private Sub CommandButton1_Click()
    
    Concatena
    
End Sub

Private Sub CommandButton2_Click()
    
    LimpiaPlantilla2
    
End Sub
