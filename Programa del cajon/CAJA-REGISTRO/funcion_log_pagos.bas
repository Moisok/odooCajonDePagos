Attribute VB_Name = "funcion_log_pagos"
Function log_pagos(respuesta_cadena As String)

    Dim SQL As String
    Dim rs As Recordset

    'fecha
    'cadena
    'dispositivo B o H
    'registro PE, TE, PS, TS, IN
    'importe en centimos
    'estado_dispositivo 0 1
    'direccion R=Reciclador, S=Stacker
    
    Dim auxcadena
    Dim dispositivo As String
    Dim registro As String
    Dim importe As Integer
    Dim estado_dispositivo As Integer
    Dim direccion As String
        
    dispositivo = ""
    registro = ""
    importe = 0
    estado_dispositivo = 0
            
    auxcadena = Split(respuesta_cadena, " ")
    
    If UBound(auxcadena) > 6 Then

        If auxcadena(1) = 10 Then
            dispositivo = "H"
        ElseIf auxcadena(1) = 40 Then
            dispositivo = "B"
        Else
            dispositivo = "-"
        End If
    
        Select Case auxcadena(2)
    
            Case 1: registro = "PE"
        
            Case 2: registro = "PS"
        
            Case 3: registro = "TE"
        
            Case 4: registro = "TS"
    
        End Select
    
        importe = Int(auxcadena(3) & auxcadena(4) & auxcadena(5) & auxcadena(6))
        
        If registro = "PE" Then
            parcial_entrada = parcial_entradas + importe
        End If
        
        If registro = "PS" Then
            parcial_salidas = parcial_salidas + importe
        End If

    
        estado_dispositivo = auxcadena(7)
    
        If auxcadena(8) = 1 Then
            direccion = "R"
        ElseIf auxcadena(8) = 2 Then
            direccion = "S"
        Else
            direccion = "-"
        End If
        
        form_principal.label_cobrar.Caption = form_principal.TCantidad
        form_principal.label_entradas.Caption = parcial_entradas / 100
        'form_principal.label_balance.Caption = parcial_entradas / 100 & ",00€"
        form_principal.label_salidas.Caption = parcial_salidas / 100
        'form_principal.label_devolver.Caption = parcial_salidas / 100 & ",00€"
        form_principal.label_pagado.Caption = (parcial_entradas - parcial_salidas) / 100
            
    Else

        dispositivo = "-"
        registro = "IN"
        'PUNTO DE ESCUCHA CANTIDAD 2
        importe = auxcadena(1) * 100
        'importe = auxcadena(1)
        estado_dispositivo = 0
        direccion = "-"
        pago_en_curso = True
        
    End If
        
    SQL = "INSERT INTO log_pagos (cadena, dispositivo, registro, importe, estado_dispositivo, direccion) " & _
            "VALUES ('" & respuesta_cadena & "', '" & dispositivo & "', '" & registro & "', " & importe & ", " & estado_dispositivo & ", '" & direccion & "')"
    MiConexión.Execute (SQL)
    
    form_principal.Text1.Text = SQL
    
    If dispositivo = "B" And direccion = "S" Then
    
        Select Case importe / 100
        
            Case 5: SQL = "INSERT INTO log_cajon_stacker (stacker_b5) VALUES (1)"
                    MiConexión.Execute (SQL)
                    
                    SQL = "UPDATE log_cajon_stacker SET stacker_b5=stacker_b5+1 WHERE codlog=1"
    
            Case 10: SQL = "INSERT INTO log_cajon_stacker (stacker_b10) VALUES (1)"
                    MiConexión.Execute (SQL)
                    
                    SQL = "UPDATE log_cajon_stacker SET stacker_b10=stacker_b10+1 WHERE codlog=1"
                    
            Case 20: SQL = "INSERT INTO log_cajon_stacker (stacker_b20) VALUES (1)"
                    MiConexión.Execute (SQL)
                    
                    SQL = "UPDATE log_cajon_stacker SET stacker_b20=stacker_b20+1 WHERE codlog=1"
                    
            Case 50: SQL = "INSERT INTO log_cajon_stacker (stacker_b50) VALUES (1)"
                    MiConexión.Execute (SQL)
                    
                    SQL = "UPDATE log_cajon_stacker SET stacker_b50=stacker_b50+1 WHERE codlog=1"
                    
            Case 100: SQL = "INSERT INTO log_cajon_stacker (stacker_b100) VALUES (1)"
                    MiConexión.Execute (SQL)
                    
                    SQL = "UPDATE log_cajon_stacker SET stacker_b100=stacker_b100+1 WHERE codlog=1"
                    
            Case 200: SQL = "INSERT INTO log_cajon_stacker (stacker_b200) VALUES (1)"
                    MiConexión.Execute (SQL)
                    
                    SQL = "UPDATE log_cajon_stacker SET stacker_b200=stacker_b200+1 WHERE codlog=1"
                    
        End Select
        
        MiConexión.Execute (SQL)
                    
    End If
        
End Function
