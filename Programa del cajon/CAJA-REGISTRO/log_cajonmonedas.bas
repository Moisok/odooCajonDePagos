Attribute VB_Name = "funcion_log_cajonmonedas"
Function log_cajonmonedas(respuesta_cadena As String)

    Dim SQL As String
    Dim rs As Recordset
    
    Dim auxcadena
    
    Dim m10c As Integer
    Dim m20c As Integer
    Dim m50c As Integer
    Dim m1 As Integer
    Dim m2 As Integer
                    
    auxcadena = Split(respuesta_cadena, " ")
    
    If UBound(auxcadena) > 15 Then

        m10c = Int(auxcadena(1) & auxcadena(2) & auxcadena(3))
        m20c = Int(auxcadena(4) & auxcadena(5) & auxcadena(6))
        m50c = Int(auxcadena(7) & auxcadena(8) & auxcadena(9))
        m1 = Int(auxcadena(10) & auxcadena(11) & auxcadena(12))
        m2 = Int(auxcadena(13) & auxcadena(14) & auxcadena(15))
        
        'ahi tenemos las monedas de cada valor que van al cajon
        SQL = "UPDATE log_cajon_stacker SET " & _
                "cajon_m5c = cajon_m5c + 0, " & _
                "cajon_m10c = cajon_m10c + " & m10c & ", " & _
                "cajon_m20c = cajon_m20c + " & m20c & ", " & _
                "cajon_m50c = cajon_m50c + " & m50c & ", " & _
                "cajon_m1 = cajon_m1 + " & m1 & ", " & _
                "cajon_m2 = cajon_m2 + " & m2 & " WHERE codlog=1"
                
        MiConexión.Execute (SQL)
        
        SQL = "INSERT INTO log_cajon_stacker (cajon_m10c, cajon_m20c, cajon_m50c, cajon_m1, cajon_m2) " & _
                "VALUES (" & m10c & "," & m20c & "," & m50c & "," & m1 & "," & m2 & ")"
        
        MiConexión.Execute (SQL)

                    
    End If
        
End Function

