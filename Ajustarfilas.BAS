Sub AjustarTodasLasHojasFilas()
    Dim sh As Worksheet
    Dim ultimaFila As Long
    Dim alto As Double
    
    ' Alto deseado de las filas
    alto = 20   ' <-- cambia este número al alto que quieras
    
    ' Recorre todas las hojas del libro
    For Each sh In ThisWorkbook.Worksheets
        ' Última fila con datos en la hoja
        ultimaFila = UltimaFilaUsada(sh)
        If ultimaFila < 1 Then ultimaFila = 1
        
        ' Quitar ocultamiento de filas
        sh.Rows("1:" & ultimaFila).Hidden = False
        
        ' Ajustar tamaño de filas
        sh.Rows("1:" & ultimaFila).RowHeight = alto
    Next sh
    
    MsgBox "Todas las hojas fueron ajustadas, sin filas ocultas.", vbInformation
End Sub

' ==== Función última fila usada ====
Private Function UltimaFilaUsada(ws As Worksheet) As Long
    Dim c As Range
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                          LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                          MatchCase:=False)
    If c Is Nothing Then
        UltimaFilaUsada = 0
    Else
        UltimaFilaUsada = c.Row
    End If
End Function

' =========================================
    ' git add Ajustarfilas.BAS
    ' git commit -m "Agrego comentarios en Ajustarfilas.BAS"
    ' git push



    ' cd EXCEL_VBA

