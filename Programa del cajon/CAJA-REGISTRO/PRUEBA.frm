VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   3855
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "PRUEBA.frx":0000
      Top             =   360
      Width           =   7215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir "
      Height          =   1695
      Left            =   2040
      TabIndex        =   0
      Top             =   4800
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Form1.Visible = False

End Sub

Private Sub Form_Load()
    
    postgress
        
    Dim pg As String
    
    Dim rs5 As Recordset
    
    pg = "SELECT * FROM account_account"

    Set rs5 = postgressconection.Execute(pg)
    
    While Not rs5.EOF
        elid = elid & rs5("root_id") & vbCrLf
        rs5.MoveNext
    Wend
    
    
    Text1.Text = elid
    
End Sub
