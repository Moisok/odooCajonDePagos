Attribute VB_Name = "funcion_cargarbotonera"
Function cargarbotonera()

    Dim SQL As String
    Dim rs As Recordset
    Dim ruta As String
    
    Dim i As Integer
    
    'inicializamos las variables
    codart_aux = 0
    precio_aux = 0
    saldo_aux = 0
    devolver_aux = 0
    
    'inicializamos el panel total
    'form_principal.label_precio = "0,00€"
    'form_principal.label_info = "Seleccione el artículo"
    'form_principal.label_balance = "0,00€"
    'form_principal.label_devolver = "0,00€"
    
    'deshabilitamos los botones
    For i = 0 To 11
    
        'form_principal.botones_copas(i).Picture = LoadPicture("d:\visual\versiones actuales\maquina copas pepito\190121\imagenes\boton-dimmed.jpg")
        'form_principal.botones_copas(i).Picture = LoadPicture("imagenes\botonvacio.jpg")
        'form_principal.boton_descripcion(i).Visible = False
        'form_principal.boton_precio(i).Visible = False
    
    Next
        
    'cargamos los articulos por posicion
    'i = 0
   
    'SQL = "SELECT * FROM maquina_articulos WHERE posicion>0 ORDER BY posicion ASC"
    'Set rs = MiConexión.Execute(SQL)
        
    'While Not rs.EOF
        'i = rs("posicion") - 1
        'ruta = "imagenes\boton" & rs("imagen") & ".jpg"
        'form_principal.botones_copas(i).Picture = LoadPicture(ruta)
        'form_principal.boton_descripcion(i).Visible = True
        'form_principal.boton_descripcion(i).Caption = rs("descripcion")
        'form_principal.boton_precio(i).Visible = True
        'form_principal.boton_precio(i).Caption = rs("precio") & "€"
        'form_principal.botones_copas(i).Caption = rs("descripcion")
        'form_principal.botones_copas(i).Visible = True
        'i = i + 1
        'rs.MoveNext
        
    'Wend
    
End Function
