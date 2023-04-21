Attribute VB_Name = "variables"
'variables

'venta_de_copas
Global codart_aux As Integer
Global precio_aux As Integer
Global pagar As Boolean
Global Ctotal As String
  
Public Type pedido

    articulo As Integer
    
    cantidad As Integer
    
    precio As Double

End Type

Global Const lineas As Integer = 20

Global vector(lineas) As pedido

Global pos As Integer

Global global_total As Integer

Global arrayComm() As String

Global puertoCommNumero As Integer

