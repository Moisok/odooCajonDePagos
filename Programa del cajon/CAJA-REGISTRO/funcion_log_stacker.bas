Attribute VB_Name = "funcion_log_stacker"
Function log_stacker(respuesta_cadena As String)

    Dim SQL As String
    Dim rs As Recordset
    
    Dim auxcadena
    
    Dim b5 As Integer
    Dim b10 As Integer
    Dim b20 As Integer
    Dim b50 As Integer
    Dim b100 As Integer
    Dim b200 As Integer
                    
    auxcadena = Split(respuesta_cadena, " ")
    
    If UBound(auxcadena) > 15 Then

        b5 = Int(auxcadena(1) & auxcadena(2) & auxcadena(3))
        b10 = Int(auxcadena(4) & auxcadena(5) & auxcadena(6))
        b20 = Int(auxcadena(7) & auxcadena(8) & auxcadena(9))
        b50 = Int(auxcadena(10) & auxcadena(11) & auxcadena(12))
        b100 = Int(auxcadena(13) & auxcadena(14) & auxcadena(15))
        b200 = Int(auxcadena(13) & auxcadena(14) & auxcadena(15))
        
        'ahi tenemos las monedas de cada valor que van al cajon
        SQL = "UPDATE log_cajon_stacker SET " & _
                "stacker_b5 = stacker_b5 + " & b5 & ", " & _
                "stacker_b10 = stacker_b10 + " & b10 & ", " & _
                "stacker_b20 = stacker_b20 + " & b20 & ", " & _
                "stacker_b50 = stacker_b50 + " & b50 & ", " & _
                "stacker_b100 = stacker_b100 + " & b100 & ", " & _
                "stacker_b200 = stacker_b200 + " & b200 & " WHERE codlog=1"
                
        MiConexión.Execute (SQL)
        
        SQL = "INSERT INTO log_cajon_stacker (stacker_b5, stacker_b10, stacker_b20, stacker_b50, stacker_b100, stacker_b200) " & _
                "VALUES (" & b5 & "," & b10 & "," & b20 & "," & b50 & "," & b100 & "," & b200 & ")"
        
        MiConexión.Execute (SQL)

                    
    End If
        
End Function

