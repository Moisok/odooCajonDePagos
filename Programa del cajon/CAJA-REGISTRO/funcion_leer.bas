Attribute VB_Name = "funcion_leer"
Function LeerArchivoTexto(ByVal rutaArchivo As String) As String
    Dim fso As Object
    Dim archivo As Object
    Dim contenido As String
    Dim linea As String
    Dim ejecutado As Boolean
    
    'Hay que intentarlo...
    Dim horaActual As Date
    horaActual = Now()
    
    'Hora de ejecucion
    Dim horaEjecucion As Date
    horaEjecucion = DateSerial(2023, 4, 20) + TimeSerial(12, 34, 56)
    
    ' Crear instancia del objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Verificar si el archivo existe
    If fso.FileExists(rutaArchivo) Then
        ' Abrir el archivo en modo lectura
        Set archivo = fso.OpenTextFile(rutaArchivo, 1)

        ' Leer el contenido del archivo línea por línea
        Do While Not archivo.AtEndOfStream
            linea = archivo.ReadLine
            ' Procesar la línea de alguna manera, por ejemplo, agregarla a la variable contenido
            contenido = contenido & linea & vbCrLf
        Loop

        ' Cerrar el archivo
        archivo.Close

        ' Liberar memoria
        Set archivo = Nothing
    Else
        ' El archivo no existe, mostrar mensaje de error o realizar alguna acción
        MsgBox "El archivo no existe: " & rutaArchivo, vbExclamation
    End If

    ' Liberar memoria
    Set fso = Nothing

    ' Comprobar si el archivo tiene contenido
    If Len(contenido) > 0 Then
                 
        'LeerArchivoTexto = contenido
        form_principal.Label16.Caption = "Hay un importe de: " & contenido & "€ pendiente"
        Ctotal = contenido
        'ESTO ES NUEVO
        If pagar = True Then
            form_principal.Command3.Enabled = True
            form_principal.Command3.Visible = True
        End If
        
        'form_principal.TCantidad.Text = contenido
        'form_principal.CM = "IMPORTE A COBRAR"
        'form_principal.CDatos_Click
                       
    Else
        'El archivo está vacío
        form_principal.Label16.Caption = "Sin pago pendiente, esperando..."
        form_principal.Label18.Visible = False
        form_principal.Label37.Visible = False
        form_principal.Label38.Visible = False
        form_principal.Command3.Enabled = False
        form_principal.Command3.Visible = False
        pagar = True
        'ESTO DE AQUI ES LO QUE HAY QUE METER EN UNA CONDICION QUE SOLO SE PUEDA REPETIR UNA VEZ....
        'form_principal.TCantidad.Text = 0
        'form_principal.CM = "RESET"
        'form_principal.CDatos_Click
            
        'archivo.Write ""
        'archivo.Close
        'Set archivo = Nothing

    End If
End Function
