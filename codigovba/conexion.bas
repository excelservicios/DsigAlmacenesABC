Attribute VB_Name = "conexion"
Public cnn As cadena

Public Sub Conectar()


 Set cn = New cadena
 cn.Conexión

' Set cn = Nothing

End Sub

Sub Registrar()
 Set cnn = New cadena
 With cnn
   .Guardar ("Productos")
 .rst(1).Value = "A003"
 .rst(2).Value = "NOMBRE DEL PRODUCTO"
 .rst(3).Value = "NIU"
 .rst(4).Value = "C2"
 .rst(5).Value = 1200
 .rst(6).Value = 700
 
 .rst.Update
 .rst.Requery
 .rst.Close
 
 
 End With
 

End Sub
