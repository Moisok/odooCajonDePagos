Attribute VB_Name = "conexionbbdd"
Global MiConexi�n As ADODB.Connection

Global parcial_entradas As Integer
Global parcial_salidas As Integer

'FUNCIONES

'conectar con la bbdd
Function conectar(bbdd As String)
    
    ' Instancio la conexi�n y me conecto con la base de datos
    Set MiConexi�n = New ADODB.Connection
        
    ' Abro la conexi�n con la base de datos usando un DSN
    With MiConexi�n
        .Open "DSN=" & bbdd
    End With

End Function
