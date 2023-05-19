VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "PRODUCTOS"
   ClientHeight    =   8052
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6996
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btncancelar_Click()
limpiar
End Sub

Private Sub btneliminar_Click()

If MsgBox("Desea eliminar el registro? ", vbYesNo + vbInformation, "Atención") = vbYes Then
  Set cnn = New cadena

  With cnn
   .Eliminar ("Productos")
   .rst.Find "Id=" & Me.lblcodigo.Caption
   
   .rst.Delete
   .rst.Requery
   .rst.Close

  End With
 
  MsgBox "Registro eliminado con exito", vbInformation, "Atención"
  Call limpiar
End If
End Sub

Private Sub btnexportar_Click()
Set cnn = New cadena

With cnn
.ExcelTabla ("select * from vistaproductos ")

'ASIGNAR LA HOJA
 ThisWorkbook.Sheets("Productos").Range("A2:G1000").ClearContents
 ThisWorkbook.Sheets("Productos").Range("A2").Select
 ThisWorkbook.Sheets("Productos").Range("A2").CopyFromRecordset .rst
End With

End Sub

Private Sub btnguardar_Click()
 Set cnn = New cadena
 With cnn
   .Guardar ("Productos")
 .rst(1).Value = Me.txtcodigo.Text
 .rst(2).Value = Me.txtproducto.Text
 .rst(3).Value = Me.cmbunidad.Column(0)
 .rst(4).Value = Me.cmbcategoria.Column(0)
 .rst(5).Value = Me.txtventa.Text
 .rst(6).Value = Me.txtcompra.Text
 
 .rst.Update
 .rst.Requery
 .rst.Close
 
 
 End With
 
 MsgBox "Registro guardao con exito", vbInformation, "Atención"
 Call limpiar
End Sub


Public Sub limpiar()
  Me.txtcodigo.Text = Empty
  Me.txtproducto.Text = Empty
  Rem Me.cmbunidad.Text = Empty
  ' Me.cmbcategoria.Text = Empty
  Me.txtventa.Text = Empty
  Me.txtcompra.Text = Empty
  
   Me.btnguardar.Enabled = True
   Me.btneliminar.Enabled = False
   Me.btnmodificar.Enabled = False
 
 Call ListarProductos
  Me.txtcodigo.SetFocus
End Sub


Public Sub ListarCategorias()
Dim FormAsControl As Object

Set cnn = New cadena
 With cnn
  Set FormAsControl = UserForm1.cmbcategoria
  Set .FormAsObject = FormAsControl
  
  .TablaConsulta ("select * from Categorias")

  .FormAsObject.ColumnCount = 2
  .FormAsObject.ColumnWidths = "0;100"

 
 End With
End Sub

Public Sub ListarUnidades()
Dim FormAsControl As Object

Set cnn = New cadena
 With cnn
  Set FormAsControl = UserForm1.cmbunidad
  Set .FormAsObject = FormAsControl
  
  .TablaConsulta ("select * from unidades")

  .FormAsObject.ColumnCount = 2
  .FormAsObject.ColumnWidths = "0;100"

 
 End With
End Sub


Private Sub btnmodificar_Click()
 Set cnn = New cadena
 With cnn
   .Actualizar ("Productos")
   .rst.Find "id=" & Me.lblcodigo.Caption
 .rst(1).Value = Me.txtcodigo.Text
 .rst(2).Value = Me.txtproducto.Text
 .rst(3).Value = Me.cmbunidad.Column(0)
 .rst(4).Value = Me.cmbcategoria.Column(0)
 .rst(5).Value = Me.txtventa.Text
 .rst(6).Value = Me.txtcompra.Text
 
 .rst.UpdateBatch
 .rst.Requery
 .rst.Close
 
 
 End With

 MsgBox "Registro Actualizado con exito", vbInformation, "Atención"
 Call limpiar
End Sub

Function BuscarTipo(cod As Integer, tabla As String)
Set cnn = New cadena
 With cnn
   .Buscar (tabla)
   .rst.Find "id=" & cod
   BuscarTipo = .rst(1).Value
 
 End With
End Function



Private Sub lista_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 If KeyCode = 13 Then
  Set cnn = New cadena
  With cnn

  .ListaTabla ("select * from Productos where id=" & Me.lista.Column(0))
   Me.lblcodigo.Caption = .rst.Fields(0).Value
   Me.txtcodigo.Text = .rst.Fields(1).Value
   Me.txtproducto.Text = .rst.Fields(2).Value
   Me.cmbunidad.Text = BuscarTipo(.rst.Fields(3).Value, "Unidades")
   Me.cmbcategoria.Text = BuscarTipo(.rst.Fields(4).Value, "Categorias")
   Me.txtventa.Text = .rst.Fields(5).Value
   Me.txtcompra.Text = .rst.Fields(6).Value

 
 End With
 Me.btnguardar.Enabled = False
 Me.btnmodificar.Enabled = True
 Me.btneliminar.Enabled = True
 End If

End Sub

Private Sub UserForm_Initialize()
 Call ListarCategorias
 Call ListarProductos
 Call ListarUnidades
 
 Me.cmbcategoria.Text = "OTROS"
 Me.cmbunidad.Text = "OTROS"
End Sub

Public Sub ListarProductos()
Dim FormAsControl As Object

Set cnn = New cadena
 With cnn
  Set FormAsControl = UserForm1.lista
  Set .FormAsObject = FormAsControl
  
  .TablaConsulta ("select * from Productos")

  .FormAsObject.ColumnCount = 7
  .FormAsObject.ColumnWidths = "0;50;180;0;0;50;50"

 
 End With
End Sub























