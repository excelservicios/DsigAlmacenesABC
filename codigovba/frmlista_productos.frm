VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmlista_productos 
   Caption         =   "Listar Productos"
   ClientHeight    =   7476
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6996
   OleObjectBlob   =   "frmlista_productos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmlista_productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String




Private Sub lista_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  If KeyCode = 13 Then
    With UserForm2
      .txtidproducto.Text = Me.lista.Column(0)
      .txtproducto.Text = Me.lista.Column(2)
      
      Unload Me
    
    End With
  
  End If
End Sub

Private Sub txtbuscar_Change()
 If Me.optcodigo.Value = True Then
   sql = "select Id, Codigo, NombreProducto from Productos Where Codigo='" & Me.txtbuscar.Text & "'"
 
 Else
   sql = "select Id, Codigo, NombreProducto from Productos Where NombreProducto like '%" & Me.txtbuscar.Text & "%'"
 
 End If
 ListarProductos
End Sub

Private Sub UserForm_Initialize()
Me.optcodigo.Value = True
End Sub

Public Sub ListarProductos()
Dim FormAsControl As Object

Set cnn = New cadena
 With cnn
  Set FormAsControl = frmlista_productos.lista
  Set .FormAsObject = FormAsControl
  
  .TablaConsulta (sql)

  .FormAsObject.ColumnCount = 3
  .FormAsObject.ColumnWidths = "50;60;250"

 
 End With
End Sub
