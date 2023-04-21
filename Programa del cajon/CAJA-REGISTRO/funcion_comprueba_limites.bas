Attribute VB_Name = "funcion_comprueba_limites"
Function comprueba_limites(valorbillete As Integer)

    'desde esta funcion comprobamos los limites
    'y habilitamos o deshabilitamos billetes segun lo que haya en el reciclador
    'sumamos todos los billetes que hay dentro del reciclador
    '10€ = nivel(7)
    '20€ = nivel(8)
    '50€ = nivel(9)
    '100€ = nivel(10)
    '200€ = nivel(11)
    Dim suma As Integer
    
    suma = 10 * Val(form_principal.nivel(7)) + 20 * Val(form_principal.nivel(8)) + 50 * Val(form_principal.nivel(9)) + 100 * Val(form_principal.nivel(10)) + 200 * Val(form_principal.nivel(11))
        
    Select Case valorbillete
    
        'billete de 200€
        Case 200:
            'If suma < 200 Then
                'form_principal.CM = "DESHABILITA BILL 200 ENTRADA"
            'Else
                'form_principal.CM = "HABILITA BILL 200 ENTRADA"
            'End If
        
        'billete de 100€
        Case 100:
            If suma < 100 Then
                form_principal.CM = "DESHABILITA BILL 100 ENTRADA"
            Else
                form_principal.CM = "HABILITA BILL 100 ENTRADA"
            End If
    
        'billete de 50€
        Case 50:
            If suma < 50 Then
                form_principal.CM = "DESHABILITA BILL 50 ENTRADA"
            Else
                form_principal.CM = "HABILITA BILL 50 ENTRADA"
            End If
    
        'billete de 20€
        Case 20:
            If suma < 20 And Val(form_principal.nivel(7)) = 0 Then
                form_principal.CM = "DESHABILITA BILL 20 ENTRADA"
            Else
                form_principal.CM = "HABILITA BILL 20 ENTRADA"
            End If
        
        'billete de 10€
        Case 10:
            form_principal.CM = "HABILITA BILL 10 ENTRADA"
            
    End Select
    
    form_principal.CDatos_Click
                
    While operacion_en_curso = True
                
    Wend
            
    log form_principal.TEstado
    
End Function
