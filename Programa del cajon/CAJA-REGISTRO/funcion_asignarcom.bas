Attribute VB_Name = "funcion_asignarcom"
Private Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long

Function commnumber()
    
    Dim i As Integer
    Dim deviceName As String
    Dim targetPath As String
    Dim nullPos As Integer
    Dim Count As Integer
    Count = 0
    
    ' Enumera los puertos COM
    For i = 0 To 8
        deviceName = "COM" & i
        targetPath = String(255, 0)
        Call QueryDosDevice(deviceName, targetPath, Len(targetPath))
        nullPos = InStr(targetPath, vbNullChar)
        If nullPos > 0 Then
           ReDim Preserve arrayComm(Count) '<- Con esto redimensionamos el array
           arrayComm(Count) = Left$(targetPath, nullPos - 1) '<- Con esto almacenamos en el array
           Count = Count + 1 ' <-Con esto incrementamos el array!!!
        End If
    Next i
    
    
    'DESCOMENTAR MAS ADELANTE
    
    For i = 0 To UBound(arrayComm)
        If arrayComm(i) = "\Device\VCP0" Then
              puertoCommNumero = i
        End If
    Next i
    
    Debug.Print commnumber
  
End Function
