VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form form_ports 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSCommLib.MSComm MSComm 
      Left            =   6120
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   3960
      Width           =   4095
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "form_ports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long


Private Sub Command1_Click()
    form_ports.Visible = False
End Sub


Private Sub Form_Load()
    
    Dim i As Integer
    Dim deviceName As String
    Dim targetPath As String
    Dim nullPos As Integer
    Dim Count As Integer
    Count = 0
    
    ' Enumera los puertos COM
    For i = 1 To 8
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
    
    For i = 0 To UBound(arrayComm)
        If arrayComm(i) = "\Device\Serial1" Then
            Debug.Print "BINGO!!!!!!"
            puertoCommNumero = i
            Debug.Print puertoCommNumero
        End If
    Next i
    
    Debug.Print commnumber
    
End Sub



'Private Sub Form_Load()
   'Dim i As Integer
   'Dim MSComm As Object
   'Set MSComm = CreateObject("MSCOMMLib.MSComm")
   
   'For i = 1 To 9
      'MSComm.CommPort = i
      'On Error Resume Next
      'MSComm.PortOpen = True
      'If Err.Number = 0 Then
         'List1.AddItem "COM" & i
         'MSComm.PortOpen = False
      'End If
   'Next i
'End Sub


