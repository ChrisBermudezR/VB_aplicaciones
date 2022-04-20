'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Creado    : 2022-02-27
'* Autor     : ChrisBermudezR
'* Contacto   : https://chrisbermudezr.github.io/
'* Objetivo   : Transponer una matriz a una sola columna manteniendo la secuencia de la matriz por columnas
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Sub rango_columnas()

Dim rango As Variant
Dim x As Long, y As Long, z As Long
Dim col As Long

rango = Selection.Value

'Esta es la parte que permite ubicar la salida
col = Selection.Column
z = Selection.Row

'Esto recorre el rango y realiza la trasposición
For y = 1 To UBound(rango, 2)
    For x = 1 To UBound(rango, 1)
        Cells(z, col + UBound(rango, 2)).Value = rango(x, y)
        z = z + 1
    Next
Next

End Sub