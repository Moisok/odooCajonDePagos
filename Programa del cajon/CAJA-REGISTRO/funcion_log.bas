Attribute VB_Name = "funcion_log"
Function log(respuesta_cadena As String)

    Dim SQL As String
    Dim rs As Recordset
        
    'monedas
    Dim in_m5c As Integer
    Dim in_m10c As Integer
    Dim in_m20c As Integer
    Dim in_m50c As Integer
    Dim in_m1 As Integer
    Dim in_m2 As Integer
    
    Dim out_m5c As Integer
    Dim out_m10c As Integer
    Dim out_m20c As Integer
    Dim out_m50c As Integer
    Dim out_m1 As Integer
    Dim out_m2 As Integer
    
    Dim lim_m5c As Integer
    Dim lim_m10c As Integer
    Dim lim_m20c As Integer
    Dim lim_m50c As Integer
    Dim lim_m1 As Integer
    Dim lim_m2 As Integer
    
    Dim niv_m5c As Integer
    Dim niv_m10c As Integer
    Dim niv_m20c As Integer
    Dim niv_m50c As Integer
    Dim niv_m1 As Integer
    Dim niv_m2 As Integer
    
    'billetes
    Dim in_b5 As Integer
    Dim in_b10 As Integer
    Dim in_b20 As Integer
    Dim in_b50 As Integer
    Dim in_b100 As Integer
    Dim in_b200 As Integer
    
    Dim out_b5 As Integer
    Dim out_b10 As Integer
    Dim out_b20 As Integer
    Dim out_b50 As Integer
    Dim out_b100 As Integer
    Dim out_b200 As Integer

    Dim lim_b5 As Integer
    Dim lim_b10 As Integer
    Dim lim_b20 As Integer
    Dim lim_b50 As Integer
    Dim lim_b100 As Integer
    Dim lim_b200 As Integer

    Dim niv_b5 As Integer
    Dim niv_b10 As Integer
    Dim niv_b20 As Integer
    Dim niv_b50 As Integer
    Dim niv_b100 As Integer
    Dim niv_b200 As Integer
        
    Dim tot_limite As Double
    Dim tot_nivel As Double
    Dim tot_stacker As Double
    
    'leemos la cadena y hacemos split para obtener cada valor
    Dim auxcadena
    Dim in_bill As Integer
    Dim out_bill As Integer
    Dim in_moneda As Integer
    Dim out_moneda As Integer
    
    If respuesta_cadena <> respuesta_cadena_ant Then
        
        respuesta_cadena_ant = respuesta_cadena
        
        auxcadena = Split(respuesta_cadena, " ")
    
        in_bill = auxcadena(1)
        out_bill = auxcadena(2)
        in_moneda = auxcadena(3)
        out_moneda = auxcadena(4)
    
        'inhibicion billetes
        If (in_bill And 1) > 0 Then
            in_b5 = 1
        Else
            in_b5 = 0
        End If
    
        If (in_bill And 2) > 0 Then
            in_b10 = 1
        Else
            in_b10 = 0
        End If
    
        If (in_bill And 4) > 0 Then
            in_b20 = 1
        Else
            in_b20 = 0
        End If
    
        If (in_bill And 8) > 0 Then
            in_b50 = 1
        Else
            in_b50 = 0
        End If
    
        If (in_bill And 16) > 0 Then
            in_b100 = 1
        Else
            in_b100 = 0
        End If
    
        'if (in_bill And 32)>0 then
        '    in_b200 = 1
        'Else
            in_b200 = 0
        'End If
        
        'payout billetes
        If (out_bill And 1) > 0 Then
            out_b5 = 1
        Else
            out_b5 = 0
        End If
        
        If (out_bill And 2) > 0 Then
            out_b10 = 1
        Else
            out_b10 = 0
        End If
        
        If (out_bill And 4) > 0 Then
            out_b20 = 1
        Else
            out_b20 = 0
        End If
    
        If (out_bill And 8) > 0 Then
            out_b50 = 1
        Else
            out_b50 = 0
        End If
    
        If (out_bill And 16) > 0 Then
            out_b100 = 1
        Else
            out_b100 = 0
        End If
    
        'if (out_bill And 32)>0 then
        '    out_b200 = 1
        'Else
            out_b200 = 0
        'End If
    
        'inhibicion monedas
        'If (in_moneda And 1) > 0 Then
        '    in_m2c = 1
        'Else
        '    in_m2c = 0
        'End If
            
        If (in_moneda And 2) > 0 Then
            in_m5c = 1
        Else
            in_m5c = 0
        End If
    
        If (in_moneda And 4) > 0 Then
            in_m10c = 1
        Else
            in_m10c = 0
        End If
    
        If (in_moneda And 8) > 0 Then
            in_m20c = 1
        Else
            in_m20c = 0
        End If
    
        If (in_moneda And 16) > 0 Then
            in_m50c = 1
        Else
            in_m50c = 0
        End If
    
        If (in_moneda And 32) > 0 Then
            in_m1 = 1
        Else
            in_m1 = 0
        End If
        
        If (in_moneda And 64) > 0 Then
            in_m2 = 1
        Else
            in_m2 = 0
        End If
    
        'payout monedas
        'If (out_moneda And 1) > 0 Then
        '    out_m2c = 1
        'Else
        '    out_m2c = 0
        'End If
    
        If (out_moneda And 2) > 0 Then
            out_m5c = 1
        Else
            out_m5c = 0
        End If
    
        If (out_moneda And 4) > 0 Then
            out_m10c = 1
        Else
            out_m10c = 0
        End If
    
        If (out_moneda And 8) > 0 Then
            out_m20c = 1
        Else
            out_m20c = 0
        End If
        
        If (out_moneda And 16) > 0 Then
            out_m50c = 1
        Else
            out_m50c = 0
        End If
            
        If (out_moneda And 32) > 0 Then
            out_m1 = 1
        Else
            out_m1 = 0
        End If
            
        If (out_moneda And 64) > 0 Then
            out_m2 = 1
        Else
            out_m2 = 0
        End If
        
        lim_m5c = Int(auxcadena(5) & auxcadena(6) & auxcadena(7))
        lim_m10c = Int(auxcadena(8) & auxcadena(9) & auxcadena(10))
        lim_m20c = Int(auxcadena(11) & auxcadena(12) & auxcadena(13))
        lim_m50c = Int(auxcadena(14) & auxcadena(15) & auxcadena(16))
        lim_m1 = Int(auxcadena(17) & auxcadena(18) & auxcadena(19))
        lim_m2 = Int(auxcadena(20) & auxcadena(21) & auxcadena(22))
        lim_b5 = Int(auxcadena(23) & auxcadena(24) & auxcadena(25))
        lim_b10 = Int(auxcadena(26) & auxcadena(27) & auxcadena(28))
        lim_b20 = Int(auxcadena(29) & auxcadena(30) & auxcadena(31))
        lim_b50 = Int(auxcadena(32) & auxcadena(33) & auxcadena(34))
        lim_b100 = Int(auxcadena(35) & auxcadena(36) & auxcadena(37))
        'lim_b200 = Int(auxcadena(38) & auxcadena(39) & auxcadena(40))
        
        niv_m5c = Int(auxcadena(37) & auxcadena(39) & auxcadena(40))
        niv_m10c = Int(auxcadena(41) & auxcadena(42) & auxcadena(43))
        niv_m20c = Int(auxcadena(44) & auxcadena(45) & auxcadena(46))
        niv_m50c = Int(auxcadena(47) & auxcadena(48) & auxcadena(49))
        niv_m1 = Int(auxcadena(50) & auxcadena(51) & auxcadena(52))
        niv_m2 = Int(auxcadena(53) & auxcadena(54) & auxcadena(55))
        niv_b5 = Int(auxcadena(56) & auxcadena(57) & auxcadena(58))
        niv_b10 = Int(auxcadena(59) & auxcadena(60) & auxcadena(61))
        niv_b20 = Int(auxcadena(62) & auxcadena(63) & auxcadena(64))
        niv_b50 = Int(auxcadena(65) & auxcadena(66) & auxcadena(67))
        niv_b100 = Int(auxcadena(68) & auxcadena(69) & auxcadena(70))
        'niv_b200 = Int(auxcadena(74) & auxcadena(75) & auxcadena(76))
    
        'actualizamos el frame estado
        If in_m5c = 1 Then
            form_principal.in(0).FillColor = vbGreen
        Else
            form_principal.in(0).FillColor = vbRed
        End If
        
        If in_m10c = 1 Then
            form_principal.in(1).FillColor = vbGreen
        Else
            form_principal.in(1).FillColor = vbRed
        End If
    
        If in_m20c = 1 Then
            form_principal.in(2).FillColor = vbGreen
        Else
            form_principal.in(2).FillColor = vbRed
        End If
    
        If in_m50c = 1 Then
            form_principal.in(3).FillColor = vbGreen
        Else
            form_principal.in(3).FillColor = vbRed
        End If
    
        If in_m1 = 1 Then
            form_principal.in(4).FillColor = vbGreen
        Else
            form_principal.in(4).FillColor = vbRed
        End If
    
        If in_m2 = 1 Then
            form_principal.in(5).FillColor = vbGreen
        Else
            form_principal.in(5).FillColor = vbRed
        End If
    
        If in_b5 = 1 Then
            form_principal.in(6).FillColor = vbGreen
        Else
            form_principal.in(6).FillColor = vbRed
        End If
    
        If in_b10 = 1 Then
            form_principal.in(7).FillColor = vbGreen
        Else
            form_principal.in(7).FillColor = vbRed
        End If
    
        If in_b20 = 1 Then
            form_principal.in(8).FillColor = vbGreen
        Else
            form_principal.in(8).FillColor = vbRed
        End If
    
        If in_b50 = 1 Then
            form_principal.in(9).FillColor = vbGreen
        Else
            form_principal.in(9).FillColor = vbRed
        End If
    
        If in_b100 = 1 Then
            form_principal.in(10).FillColor = vbGreen
        Else
            form_principal.in(10).FillColor = vbRed
        End If
    
        If in_b200 = 1 Then
            form_principal.in(11).FillColor = vbGreen
        Else
            form_principal.in(11).FillColor = vbRed
        End If
    
        If out_m5c = 1 Then
            form_principal.out(0).FillColor = vbGreen
        Else
            form_principal.out(0).FillColor = vbRed
        End If
    
        If out_m10c = 1 Then
            form_principal.out(1).FillColor = vbGreen
        Else
            form_principal.out(1).FillColor = vbRed
        End If
        
        If out_m20c = 1 Then
            form_principal.out(2).FillColor = vbGreen
        Else
            form_principal.out(2).FillColor = vbRed
        End If
    
        If out_m50c = 1 Then
            form_principal.out(3).FillColor = vbGreen
        Else
            form_principal.out(3).FillColor = vbRed
        End If
    
        If out_m1 = 1 Then
            form_principal.out(4).FillColor = vbGreen
        Else
            form_principal.out(4).FillColor = vbRed
        End If
        
        If out_m2 = 1 Then
            form_principal.out(5).FillColor = vbGreen
        Else
            form_principal.out(5).FillColor = vbRed
        End If
    
        If out_b5 = 1 Then
            form_principal.out(6).FillColor = vbGreen
        Else
            form_principal.out(6).FillColor = vbRed
        End If
    
        If out_b10 = 1 Then
            form_principal.out(7).FillColor = vbGreen
        Else
            form_principal.out(7).FillColor = vbRed
        End If
    
        If out_b20 = 1 Then
            form_principal.out(8).FillColor = vbGreen
        Else
            form_principal.out(8).FillColor = vbRed
        End If
    
        If out_b50 = 1 Then
            form_principal.out(9).FillColor = vbGreen
        Else
            form_principal.out(9).FillColor = vbRed
        End If
    
        If out_b100 = 1 Then
            form_principal.out(10).FillColor = vbGreen
        Else
            form_principal.out(10).FillColor = vbRed
        End If
    
        If out_b200 = 1 Then
            form_principal.out(11).FillColor = vbGreen
        Else
            form_principal.out(11).FillColor = vbRed
        End If
    
        form_principal.limite(0).Caption = lim_m5c
        form_principal.limite(1).Caption = lim_m10c
        form_principal.limite(2).Caption = lim_m20c
        form_principal.limite(3).Caption = lim_m50c
        form_principal.limite(4).Caption = lim_m1
        form_principal.limite(5).Caption = lim_m2
        form_principal.limite(6).Caption = lim_b5
        form_principal.limite(7).Caption = lim_b10
        form_principal.limite(8).Caption = lim_b20
        form_principal.limite(9).Caption = lim_b50
        form_principal.limite(10).Caption = lim_b100
        form_principal.limite(11).Caption = lim_b200
        
        tot_limite = lim_m5c * 0.05 + lim_m10c * 0.1 + lim_m20c * 0.2 + lim_m50c * 0.5 + lim_m1 + lim_m2 * 2 + lim_b5 * 5 + lim_b10 * 10 + lim_b20 * 20 + lim_b50 * 50 + lim_b100 * 100 + lim_b200 * 200
                
        form_principal.nivel(0).Caption = niv_m5c
        form_principal.nivel(1).Caption = niv_m10c
        form_principal.nivel(2).Caption = niv_m20c
        form_principal.nivel(3).Caption = niv_m50c
        form_principal.nivel(4).Caption = niv_m1
        form_principal.nivel(5).Caption = niv_m2
        form_principal.nivel(6).Caption = niv_b5
        form_principal.nivel(7).Caption = niv_b10
        form_principal.nivel(8).Caption = niv_b20
        form_principal.nivel(9).Caption = niv_b50
        form_principal.nivel(10).Caption = niv_b100
        form_principal.nivel(11).Caption = niv_b200
        
        tot_nivel = niv_m5c * 0.05 + niv_m10c * 0.1 + niv_m20c * 0.2 + niv_m50c * 0.5 + niv_m1 + niv_m2 * 2 + niv_b5 * 5 + niv_b10 * 10 + niv_b20 * 20 + niv_b50 * 50 + niv_b100 * 100 + niv_b200 * 200
        
        SQL = "INSERT INTO log_ci (cadena, in_b5, in_b10, in_b20, in_b50, in_b100, in_b200, out_b5, out_b10, out_b20, out_b50, out_b100, out_b200, " & _
                                    "in_m5c, in_m10c, in_m20c, in_m50c, in_m1, in_m2, out_m5c, out_m10c, out_m20c, out_m50c, out_m1, out_m2, " & _
                                    "lim_m5c, lim_m10c, lim_m20c, lim_m50c, lim_m1, lim_m2, lim_b5, lim_b10, lim_b20, lim_b50, lim_b100, lim_b200, " & _
                                    "niv_m5c, niv_m10c, niv_m20c, niv_m50c, niv_m1, niv_m2, niv_b5, niv_b10, niv_b20, niv_b50, niv_b100, niv_b200) " & _
                                    "VALUES ('" & respuesta_cadena & "'," & in_b5 & "," & in_b10 & "," & in_b20 & "," & in_b50 & "," & in_b100 & "," & in_b200 & "," & _
                                    out_b5 & "," & out_b10 & "," & out_b20 & "," & out_b50 & "," & out_b100 & "," & out_b200 & "," & _
                                    in_m5c & "," & in_m10c & "," & in_m20c & "," & in_m50c & "," & in_m1 & "," & in_m2 & "," & _
                                    out_m5c & "," & out_m10c & "," & out_m20c & "," & out_m50c & "," & out_m1 & "," & out_m2 & "," & _
                                    lim_m5c & "," & lim_m10c & "," & lim_m20c & "," & lim_m50c & "," & lim_m1 & "," & lim_m2 & "," & _
                                    lim_b5 & "," & lim_b10 & "," & lim_b20 & "," & lim_b50 & "," & lim_b100 & "," & lim_b200 & "," & _
                                    niv_m5c & "," & niv_m10c & "," & niv_m20c & "," & niv_m50c & "," & niv_m1 & "," & niv_m2 & "," & _
                                    niv_b5 & "," & niv_b10 & "," & niv_b20 & "," & niv_b50 & "," & niv_b100 & "," & niv_b200 & ")"
                
        MiConexión.Execute (SQL)
        
        SQL = "SELECT * FROM log_cajon_stacker WHERE codlog=1"
        Set rs = MiConexión.Execute(SQL)
        
        On Error Resume Next
        
        form_principal.stacker(0).Caption = rs("cajon_m5c")
        form_principal.stacker(1).Caption = rs("cajon_m10c")
        form_principal.stacker(2).Caption = rs("cajon_m20c")
        form_principal.stacker(3).Caption = rs("cajon_m50c")
        form_principal.stacker(4).Caption = rs("cajon_m1")
        form_principal.stacker(5).Caption = rs("cajon_m2")
        form_principal.stacker(6).Caption = rs("stacker_b5")
        form_principal.stacker(7).Caption = rs("stacker_b10")
        form_principal.stacker(8).Caption = rs("stacker_b20")
        form_principal.stacker(9).Caption = rs("stacker_b50")
        form_principal.stacker(10).Caption = rs("stacker_b100")
        form_principal.stacker(11).Caption = rs("stacker_b200")
        
        tot_stacker = rs("cajon_m5c") * 0.05 + rs("cajon_m10c") * 0.1 + rs("cajon_m20c") * 0.2 + rs("cajon_m50c") * 0.5 + rs("cajon_m1") + rs("cajon_m2") * 2 + rs("stacker_b5") * 5 + rs("stacker_b10") * 10 + rs("stacker_b20") * 20 + rs("stacker_b50") * 50 + rs("stacker_b100") * 100 + rs("stacker_b200") * 200
        
    End If
    
    form_principal.total_limite.Caption = tot_limite & "€"
    form_principal.total_nivel.Caption = tot_nivel & "€"
    form_principal.total_stacker.Caption = tot_stacker & "€"
    
End Function
