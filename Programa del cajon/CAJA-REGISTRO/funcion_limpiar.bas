Attribute VB_Name = "funcion_limpiar"
Function LimpiarMonto(ByVal rutaArchivo As String)
    
    Dim fso As Object
    Dim archivo As Object
    
    ' Instanciamos el objeto de tipo FileSystem
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Verificamos si existe o no el archivo
    If fso.FileExists(rutaArchivo) Then
        ' Abrimos el archivo
        Set archivo = fso.OpenTextFile(rutaArchivo, 2)
        
        ' Borramos el contenido del archivo
        archivo.Write ""
        
        ' Cerramos el archivo
        archivo.Close
        
        ' Liberamos la memoria
        Set archivo = Nothing
    Else
        MsgBox "El archivo no existe: " & rutaArchivo, vbExclamation
    End If
    
    ' Liberamos memoria de nuevo
    Set fso = Nothing
End Function


