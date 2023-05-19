VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cadena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public cn As New ADODB.Connection
Public rst As ADODB.Recordset
Dim cadena As New cadena
Public FormAsControl As UserForm
Public FormAsObject As Object
Public Numerico As Integer


'DERECHOS RESERVADOS'
'FERNANDO MAMANI BLAS
'WEB: EXCELSERVICIOS.COM
'DESARROLLO WEB: DSIGPERU.NET
'EL PRINCIPIO DE LA SABIDURIA ES EL TEMOR DE JEHOVA
'----------------------------------------------------

Function Conexión()
Set cn = New ADODB.Connection

'CADENA DE CONEXION PHPMYSQL
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
With cn
  .CursorLocation = adUseClient
  'AQUI TU CADENA DE CONEXIÓN
   If .State = False Then
    .Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ThisWorkbook.Path + "/SistemaABC.accdb" + ";Persist Security Info=False;"
   End If

End With
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



    If cn.State = adStateOpen Then
        Set Conexión = cn
    End If
End Function
Function GetC()


End Function

Function ExcelTabla(Consultasql As String)
  Set rst = New ADODB.Recordset
    cadena.Conexión
    rst.Open Consultasql, cadena.cn
'    ActiveCell.CopyFromRecordset rst
'
'   ' lngcampos = rst.Fields.Count

   ' For i = 0 To lngCampos - 1
    '  Cells(ActiveCell.Row - 1, i + 1).Value = rst.Fields(i).Name
   ' Next
  
End Function

Function Guardar(tabla As String) As String
    Set rst = New ADODB.Recordset

    cadena.Conexión

    rst.Open tabla, cadena.cn, adOpenKeyset, adLockOptimistic, adCmdTable

    rst.Requery
    rst.AddNew


End Function

Function Actualizar(tabla As String) As String
 Set rst = New ADODB.Recordset

    cadena.Conexión

    rst.Open tabla, cadena.cn, adOpenKeyset, adLockOptimistic, adCmdTable
    rst.Requery


End Function

Function Eliminar(tabla As String) As String
 Set rst = New ADODB.Recordset

    cadena.Conexión

    rst.Open tabla, cadena.cn, adOpenKeyset, adLockOptimistic, adCmdTable
    rst.Requery

End Function

Function Buscar(tabla As String) As String
 Set rst = New ADODB.Recordset

 cadena.Conexión
 rst.Open tabla, cadena.cn, adOpenKeyset, adLockOptimistic, adCmdTable
 If rst.BOF Or rst.EOF Then Exit Function


End Function

Function ListaTabla(Consultasql As String)

   Set rst = New ADODB.Recordset
   cadena.Conexión

   rst.Open Consultasql, cadena.cn
  
End Function

Function Autonumerico(Consultasql As String)
   Set rst = New ADODB.Recordset
   cadena.Conexión

   rst.Open Consultasql, cadena.cn, AdOpenDinamic, adLockOptimistic
  
End Function

Function TablaConsulta(Consultasql As String)
   Set cn = Conexión
   Set rst = New ADODB.Recordset
   rst.Open Consultasql, cn
 If rst.EOF = False Or rst.BOF = False Then
    With rst
       .MoveLast
        NoOfRecords = .RecordCount
       .MoveFirst
     
        FormAsObject.ColumnCount = .Fields.Count
        FormAsObject.Column = .GetRows(NoOfRecords)
    End With
 End If
End Function


