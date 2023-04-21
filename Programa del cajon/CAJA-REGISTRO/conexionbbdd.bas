Attribute VB_Name = "conexionbbdd"
Global MiConexión As ADODB.Connection

Global parcial_entradas As Integer
Global parcial_salidas As Integer

'FUNCIONES

'conectar con la bbdd
Function conectar(bbdd As String)
    
    ' Instancio la conexión y me conecto con la base de datos
    Set MiConexión = New ADODB.Connection
        
    ' Abro la conexión con la base de datos usando un DSN
    With MiConexión
        .Open "DSN=" & bbdd
    End With

End Function
