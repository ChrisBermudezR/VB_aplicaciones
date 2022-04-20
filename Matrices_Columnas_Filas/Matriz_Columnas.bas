'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Creado    : 2022-02-27
'* Autor     : ChrisBermudezR
'* Contacto   : https://chrisbermudezr.github.io/
'* Objetivo   : Transponer una matriz a una sola columna manteniendo la secuencia de la matriz por columnas
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Sub rango_columnas()

Dim rango As Variant
Dim i As Long, j As Long, k As Long
Dim col As Long

rango = Selection.Value

'Esta es la parte que permite ubicar la salida
col = Selection.Column
k = Selection.Row

'Esto recorre el rango y realiza la trasposición
For j = 1 To UBound(rango, 2)
    For i = 1 To UBound(rango, 1)
        Cells(k, col + UBound(rango, 2)).Value = rango(i, j)
        k = k + 1
    Next
Next

End Sub