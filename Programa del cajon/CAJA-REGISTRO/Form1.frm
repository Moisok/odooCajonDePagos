VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8400
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir "
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   4800
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "De momento no hay nada, estoy leyendo....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   7335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    LimpiarMonto "C:\xampp\htdocs\exportar\monto_cajon.txt"
    
    Timer1.Enabled = False
    Form1.Visible = False
    
End Sub

Private Sub Form_Load()

End Sub

Private Sub Timer1_Timer()

    LeerArchivoTexto "C:\xampp\htdocs\exportar\monto_cajon.txt"

End Sub
