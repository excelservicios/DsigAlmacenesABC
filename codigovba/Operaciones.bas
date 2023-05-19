Attribute VB_Name = "Operaciones"
Sub CrearTarjetas()

Dim x, y As Integer
y = 12

For x = 2 To 51
 With Sheets("Productos")
   Sheets("Tarjetas").Range("D4") = .Cells(x, 2)
   Sheets("Tarjetas").Range("B2:E10").Copy
   Range("B" & y).Select
   ActiveSheet.Paste
 
 End With
 
y = y + 10
Next



End Sub
