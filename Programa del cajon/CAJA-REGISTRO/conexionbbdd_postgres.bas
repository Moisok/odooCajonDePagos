Attribute VB_Name = "conexionbbdd_postgres"
Global postgressconection As ADODB.Connection

Function postgress()
    
   ' Instancio la conexión y me conecto con la base de datos
    Set postgressconection = New ADODB.Connection
        
    ' Establecer la cadena de conexión para psqlODBC
    Dim strConn As String
    strConn = "Driver={PostgreSQL unicode};Server=192.168.0.101;Port=5432;Database=pruebas;Uid=openpg;Pwd=dev_openpgpwd;"
    
    ' Abro la conexión con la base de datos
    With postgressconection
        .Open strConn
    End With

End Function
