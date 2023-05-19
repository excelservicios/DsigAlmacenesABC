VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Movimientos"
   ClientHeight    =   7932
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7536
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnguardar_Click()
 Set cnn = New cadena
 With cnn
   .Guardar ("Movimientos")
 .rst(1).Value = Me.txtfecha.Text
 .rst(2).Value = Me.txtidproducto.Text
 .rst(3).Value = Me.txtcantidad.Text
 
 If Me.cmbmovimiento.Text = "ENTRADA" Then
 .rst(4).Value = True
 Else
 .rst(4).Value = False
 End If
 
 

 
 .rst.Update
 .rst.Requery
 .rst.Close
 
 
 End With
 
 MsgBox "Registro guardado con exito", vbInformation, "Atención"
 ListarSTock
End Sub

Private Sub cmbuscarproducto_Click()
frmlista_productos.Show vbModal
End Sub



Private Sub txtidproducto_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then


Set cnn = New cadena
 With cnn
   .ListaTabla ("select Id, Codigo, NombreProducto from Productos Where Id=" & Me.txtidproducto.Text)

   Me.txtproducto.Text = .rst(2).Value
 
 End With
End If


End Sub

Private Sub UserForm_Activate()
Me.cmbmovimiento.AddItem "ENTRADA"
Me.cmbmovimiento.AddItem "SALIDA"
Me.txtidproducto.Text = 1
ListarSTock
End Sub


Public Sub ListarSTock()
Dim FormAsControl As Object

Set cnn = New cadena
 With cnn
  Set FormAsControl = UserForm2.lista
  Set .FormAsObject = FormAsControl
  
  .TablaConsulta ("select * from vistastock")

  .FormAsObject.ColumnCount = 5
  .FormAsObject.ColumnWidths = "65;150;50;50;50"

 
 End With
End Sub

'SELECT Productos.Codigo, Productos.NombreProducto,
'iif( Movimientos.tipo=true, Movimientos.cantidad, 0) as Entrada,
'iif(Movimientos.tipo = False, Movimientos.cantidad, 0) As salida
'
'FROM Productos
' INNER JOIN Movimientos ON Productos.Id = Movimientos.idproducto;









