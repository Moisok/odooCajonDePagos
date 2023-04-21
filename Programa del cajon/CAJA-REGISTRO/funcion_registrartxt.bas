Attribute VB_Name = "funcion_registrartxt"
Function registrartexto(operaciontexto As String, Optional DATOH As Integer, Optional Datol As Integer, Optional operacion As Integer, Optional direccion As Integer)

    Dim filenumber As Integer
    
    filenumber = FreeFile()

    Open ".\logbits2.txt" For Append As #filenumber
    
    Print #filenumber, "Operacion (texto): " & operaciontexto & " | DATOH:" & DATOH & " | Dato1 " & Datol & " | Operacion:  " & operacion & " | Direccion " & _
    direccion
    
    Close filenumber

End Function
