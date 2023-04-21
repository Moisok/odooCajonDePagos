VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CONTROL EFECTIVO"
   ClientHeight    =   13665
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   24165
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   241.036
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   426.244
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton boton_vaciarstacker 
      Caption         =   "Vaciar Stacker"
      Height          =   420
      Left            =   120
      TabIndex        =   117
      Top             =   5400
      Width           =   1845
   End
   Begin VB.Frame framemonedas 
      Caption         =   "Monedas"
      Height          =   3255
      Left            =   13080
      TabIndex        =   86
      Top             =   9480
      Width           =   5895
      Begin VB.Label total_diferencia 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   116
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label total_despues 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   115
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label total_antes 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   114
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label diferencia 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   4320
         TabIndex        =   113
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label diferencia 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   112
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label diferencia 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   111
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label diferencia 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   110
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label diferencia 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   109
         Top             =   960
         Width           =   855
      End
      Begin VB.Label diferencia 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   108
         Top             =   600
         Width           =   855
      End
      Begin VB.Label despues 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3120
         TabIndex        =   107
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label despues 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   106
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label despues 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   105
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label despues 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   104
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label despues 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   103
         Top             =   960
         Width           =   855
      End
      Begin VB.Label despues 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   102
         Top             =   600
         Width           =   855
      End
      Begin VB.Label antes 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   101
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label antes 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   100
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label antes 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   99
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label antes 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   98
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label antes 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   97
         Top             =   960
         Width           =   855
      End
      Begin VB.Label antes 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   96
         Top             =   600
         Width           =   855
      End
      Begin VB.Line Line2 
         X1              =   5400
         X2              =   1680
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Caption         =   "DIFERENCIA"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   95
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "DESPUES"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   94
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "ANTES"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   93
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "0,05€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   92
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "0,10€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   91
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "0,20€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   90
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "0,50€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   89
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "1€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   88
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "2€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   87
         Top             =   2400
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   855
      Left            =   120
      TabIndex        =   67
      Top             =   4200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame frame_log 
      Caption         =   "Log operaciones"
      Height          =   7335
      Left            =   19080
      TabIndex        =   65
      Top             =   120
      Width           =   4935
      Begin VB.TextBox text_log_operaciones 
         Enabled         =   0   'False
         Height          =   6855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   66
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Frame frame_operacion 
      Caption         =   "Operación en curso"
      Height          =   1455
      Left            =   13080
      TabIndex        =   62
      Top             =   7920
      Width           =   5895
      Begin VB.Timer timer_label 
         Interval        =   800
         Left            =   4560
         Top             =   240
      End
      Begin VB.Label label_operacion_estado 
         Alignment       =   2  'Center
         Caption         =   "En proceso..."
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   240
         TabIndex        =   64
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label label_operacion 
         Alignment       =   2  'Center
         Caption         =   "Vaciado hopper"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   63
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame frame_entsal 
      Caption         =   "Entradas y salidas"
      Height          =   2175
      Left            =   13080
      TabIndex        =   53
      Top             =   5640
      Width           =   5895
      Begin VB.Label label_cajon 
         Alignment       =   2  'Center
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   71
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Caption         =   "CAJON"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   70
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label label_salidas 
         Alignment       =   2  'Center
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   61
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label label_entradas 
         Alignment       =   2  'Center
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   60
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label label_pagado 
         Alignment       =   2  'Center
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   59
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label label_cobrar 
         Alignment       =   2  'Center
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   58
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "PAGADO"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   57
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "SALIDAS"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   56
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "ENTRADAS"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   55
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "COBRAR"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   54
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame frame_status 
      Caption         =   "Estado"
      Height          =   5415
      Left            =   13080
      TabIndex        =   12
      Top             =   120
      Width           =   5895
      Begin VB.Label total_limite 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   85
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label total_stacker 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   84
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label total_nivel 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   83
         Top             =   5040
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   5760
         X2              =   3000
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Label stacker 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   4920
         TabIndex        =   82
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label stacker 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   4920
         TabIndex        =   81
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label stacker 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   4920
         TabIndex        =   80
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label stacker 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   4920
         TabIndex        =   79
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label stacker 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   4920
         TabIndex        =   78
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label stacker 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   4920
         TabIndex        =   77
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label stacker 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   4920
         TabIndex        =   76
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label stacker 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   75
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label stacker 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   74
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label stacker 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   73
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label stacker 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   72
         Top             =   960
         Width           =   855
      End
      Begin VB.Label stacker 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4920
         TabIndex        =   69
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "STACKER"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   68
         Top             =   240
         Width           =   855
      End
      Begin VB.Label nivel 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   3960
         TabIndex        =   52
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label nivel 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   3960
         TabIndex        =   51
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label nivel 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   3960
         TabIndex        =   50
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label nivel 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   3960
         TabIndex        =   49
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label nivel 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   3960
         TabIndex        =   48
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label nivel 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   3960
         TabIndex        =   47
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label nivel 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3960
         TabIndex        =   46
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label nivel 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   45
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label nivel 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   44
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label nivel 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   43
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label nivel 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   42
         Top             =   960
         Width           =   855
      End
      Begin VB.Label nivel 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   41
         Top             =   600
         Width           =   855
      End
      Begin VB.Label limite 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   40
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label limite 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   3000
         TabIndex        =   39
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label limite 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   3000
         TabIndex        =   38
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label limite 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   37
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label limite 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   36
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label limite 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   35
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label limite 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   34
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label limite 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   33
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label limite 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   32
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label limite 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   31
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label limite 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   30
         Top             =   960
         Width           =   855
      End
      Begin VB.Label limite 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   29
         Top             =   600
         Width           =   855
      End
      Begin VB.Shape out 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   11
         Left            =   2400
         Top             =   4560
         Width           =   255
      End
      Begin VB.Shape out 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   10
         Left            =   2400
         Top             =   4200
         Width           =   255
      End
      Begin VB.Shape out 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   9
         Left            =   2400
         Top             =   3840
         Width           =   255
      End
      Begin VB.Shape out 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   8
         Left            =   2400
         Top             =   3480
         Width           =   255
      End
      Begin VB.Shape out 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   7
         Left            =   2400
         Top             =   3120
         Width           =   255
      End
      Begin VB.Shape out 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   6
         Left            =   2400
         Top             =   2760
         Width           =   255
      End
      Begin VB.Shape out 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   5
         Left            =   2400
         Top             =   2400
         Width           =   255
      End
      Begin VB.Shape out 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   2400
         Top             =   2040
         Width           =   255
      End
      Begin VB.Shape out 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   2400
         Top             =   1680
         Width           =   255
      End
      Begin VB.Shape out 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   2400
         Top             =   1320
         Width           =   255
      End
      Begin VB.Shape out 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   2400
         Top             =   960
         Width           =   255
      End
      Begin VB.Shape out 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   2400
         Top             =   600
         Width           =   255
      End
      Begin VB.Shape in 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   11
         Left            =   1560
         Top             =   4560
         Width           =   255
      End
      Begin VB.Shape in 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   10
         Left            =   1560
         Top             =   4200
         Width           =   255
      End
      Begin VB.Shape in 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   9
         Left            =   1560
         Top             =   3840
         Width           =   255
      End
      Begin VB.Shape in 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   8
         Left            =   1560
         Top             =   3480
         Width           =   255
      End
      Begin VB.Shape in 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   7
         Left            =   1560
         Top             =   3120
         Width           =   255
      End
      Begin VB.Shape in 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   6
         Left            =   1560
         Top             =   2760
         Width           =   255
      End
      Begin VB.Shape in 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   5
         Left            =   1560
         Top             =   2400
         Width           =   255
      End
      Begin VB.Shape in 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   1560
         Top             =   2040
         Width           =   255
      End
      Begin VB.Shape in 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   1560
         Top             =   1680
         Width           =   255
      End
      Begin VB.Shape in 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   1560
         Top             =   1320
         Width           =   255
      End
      Begin VB.Shape in 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   1560
         Top             =   960
         Width           =   255
      End
      Begin VB.Shape in 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   1560
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "NIVEL"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "LIMITE"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "OUT"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2100
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "IN"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1240
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "200€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "100€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "50€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "20€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "10€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "5€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "2€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "1€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "0,50€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "0,20€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "0,10€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "0,05€"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   19080
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   7560
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox TConfi 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox TCantidad 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PARAR"
      Height          =   300
      Left            =   8880
      TabIndex        =   6
      Top             =   120
      Width           =   1485
   End
   Begin VB.TextBox TControl 
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,0 ""€"";(#.##0,0 ""€"")"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   12165
      Left            =   7800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   720
      Width           =   5055
   End
   Begin VB.TextBox TEstado 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   13080
      Width           =   12615
   End
   Begin VB.ComboBox CM 
      BackColor       =   &H80000004&
      ForeColor       =   &H00FF0000&
      Height          =   345
      ItemData        =   "CAJA-REGISTRO.frx":0000
      Left            =   4080
      List            =   "CAJA-REGISTRO.frx":0002
      TabIndex        =   3
      Text            =   "COMANDOS"
      Top             =   120
      Width           =   3855
   End
   Begin VB.Timer Tmax 
      Left            =   1440
      Top             =   0
   End
   Begin VB.CommandButton CDatos 
      Caption         =   "COMANDOS"
      Height          =   300
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1485
   End
   Begin VB.TextBox TDatos 
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,0 ""€"";(#.##0,0 ""€"")"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   12165
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   5415
   End
   Begin VB.CommandButton CClear 
      Caption         =   "CLEAR"
      Height          =   300
      Left            =   10800
      TabIndex        =   0
      Top             =   120
      Width           =   1485
   End
   Begin VB.Timer TCONSULTA 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer TRECEPCION 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "CONFIG"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "IMPORTE"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.Shape Punto 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   120
      Shape           =   3  'Circle
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents PCom As MSComm
Attribute PCom.VB_VarHelpID = -1
Private BILLETES As Integer
Private BILL As Integer
Private COMANDO As Integer
Private SubCOMANDO As Integer
Private FASE As Integer
Private SubFASE As Integer
Private TRANS As String
Private CADE_CHECK As String
Private TEn As String
Private CENTRADAS As Integer
Private ERRORES As Integer
Private CONTADOR As Integer
Private ContadorDatos As Integer
Private ContadorControl As Integer
Private INDICADOR As Byte
Private S As Byte
Private PARA As Byte
Private ACK As Integer
Private DTOS_SALIDA As String
Private DTOS_EN(25) As Integer
Private DTOS_SAL(25) As Integer
Private Variable As Integer
Private Variable1 As Integer
Private Variable2 As Integer
Private CAN_DTOS As Integer
Private CHECK As Integer
Private OPERACION As Integer
Private direccion As Integer
Private DATOL As Integer
Private DATOH As Integer
Private V_BILL_STAKER As Long
Private V_BILL_PAYOUT As Long
Private V_BILL_SCROW As Integer
Private SSE As String
Private COMPA_SSE As String

Dim cadena_ant As String
Dim cadena As String
Dim suma_control As Integer
Dim operacion_en_curso As Boolean

Dim respuesta_cadena_ant As String
Dim pago_en_curso As Boolean

Const BILLETEASTAKER = 154
Const HOST = 1
Const BILLE = 40
Const STATUS = 29
Const RUTA = 37
Const SETOPCIONPAGOS = 30
Const CANTIDADBILLETES = 42
Const ESTATUS = 29
Const INHSTAKERSCROW = 153
Const HABILITOCANALES = 231
Const INHBILL = 228
Const SI = 1
Const NO = 0

Private Sub boton_vaciarstacker_Click()

    Dim SQL As String
    
    SQL = "UPDATE log_cajon_stacker SET cajon_m5c=0, cajon_m10c=0, cajon_m20c=0, cajon_m50c=0, cajon_m1=0,cajon_m2=0, " & _
            "stacker_b5=0, stacker_b10=0, stacker_b20=0, stacker_b50=0, stacker_b100=0, stacker_b200=0 " & _
            "WHERE codlog=1"
    MiConexión.Execute (SQL)
    
    Dim i As Integer
    
    For i = 0 To 11
        
        stacker(i).Caption = 0
        
    Next
    
End Sub

Private Sub CClear_Click()

    TDatos = ""
    TControl = ""
    ContadorDatos = 0
    ContadorControl = 0
    
End Sub

Sub CDatos_Click()

    Dim CAN, A, B As Integer
    
    Dim SQL As String
    Dim rs As Recordset
    
    Select Case (CM)
    
        Case "RESET"
            
            PCom.Output = Chr(1)
        
        Case "HABILITA BILLETERO"
            
            OPERACION = 28
        
        'HABILITACION BILLETES ENTRADA
        Case "HABILITA BILL 5 ENTRADA"
                    
            OPERACION = 20
            direccion = 1
        
        Case "DESHABILITA BILL 5 ENTRADA"
            
            OPERACION = 20
            direccion = 2
        
        Case "HABILITA BILL 10 ENTRADA"
            
            OPERACION = 20
            direccion = 3
        
        Case "DESHABILITA BILL 10 ENTRADA"
            
            OPERACION = 20
            direccion = 4
        
        Case "HABILITA BILL 20 ENTRADA"
            
            OPERACION = 20
            direccion = 5
        
        Case "DESHABILITA BILL 20 ENTRADA"
            
            OPERACION = 20
            direccion = 6
        
        Case "HABILITA BILL 50 ENTRADA"
            
            OPERACION = 20
            direccion = 7
        
        Case "DESHABILITA BILL 50 ENTRADA"
            
            OPERACION = 20
            direccion = 8
        
        Case "HABILITA BILL 100 ENTRADA"
            
            OPERACION = 20
            direccion = 9
        
        Case "DESHABILITA BILL 100 ENTRADA"
            
            OPERACION = 20
            direccion = 10
        
        'HABILITACION BILLETES PAGOS
        Case "HABILITA BILL 5 PARA PAGOS"
            
            OPERACION = 21
            direccion = 1
        
        Case "DESHABILITA BILL 5 PARA PAGOS"
            
            OPERACION = 21
            direccion = 2
        
        Case "HABILITA BILL 10 PARA PAGOS"
            
            OPERACION = 21
            direccion = 3
        
        Case "DESHABILITA BILL 10 PARA PAGOS"
            
            OPERACION = 21
            direccion = 4
        
        Case "HABILITA BILL 20 PARA PAGOS"
            
            OPERACION = 21
            direccion = 5
        
        Case "DESHABILITA BILL 20 PARA PAGOS"
            
            OPERACION = 21
            direccion = 6
        
        Case "HABILITA BILL 50 PARA PAGOS"
            
            OPERACION = 21
            direccion = 7
        
        Case "DESHABILITA BILL 50 PARA PAGOS"
            
            OPERACION = 21
            direccion = 8
        
        Case "HABILITA BILL 100 PARA PAGOS"
            
            OPERACION = 21
            direccion = 9
        
        Case "DESHABILITA BILL 100 PARA PAGOS"
            
            OPERACION = 21
            direccion = 10
        
        'HABILITACION MONEDAS ENTRADA
        Case "HABILITA MONE 2 CEN ENTRADA"
            
            OPERACION = 22
            direccion = 1
        
        Case "DESHABILITA MONE 2 CEN ENTRADA"
            
            OPERACION = 22
            direccion = 2
        
        Case "HABILITA MONE 5 CEN ENTRADA"
            
            OPERACION = 22
            direccion = 3
        
        Case "DESHABILITA MONE 5 CEN ENTRADA"
            
            OPERACION = 22
            direccion = 4
        
        Case "HABILITA MONE 10 CEN ENTRADA"
            
            OPERACION = 22
            direccion = 5
        
        Case "DESHABILITA MONE 10 CEN ENTRADA"
            
            OPERACION = 22
            direccion = 6
        
        Case "HABILITA MONE 20 CEN ENTRADA"
            
            OPERACION = 22
            direccion = 7
        
        Case "DESHABILITA MONE 20 CEN ENTRADA"
            
            OPERACION = 22
            direccion = 8
        
        Case "HABILITA MONE 50 CEN ENTRADA"
            
            OPERACION = 22
            direccion = 9
        
        Case "DESHABILITA MONE 50 CEN ENTRADA"
            
            OPERACION = 22
            direccion = 10
        
        Case "HABILITA MONE 100 CEN ENTRADA"
            
            OPERACION = 22
            direccion = 11
        
        Case "DESHABILITA MONE 100 CEN ENTRADA"
            
            OPERACION = 22
            direccion = 12
        
        Case "HABILITA MONE 200 CEN ENTRADA"
            
            OPERACION = 22
            direccion = 13
        
        Case "DESHABILITA MONE 200 CEN ENTRADA"
            
            OPERACION = 22
            direccion = 14
        
        'HABILITACION MONEDA PARA PAGOS
        Case "HABILITA MONE 5 CEN PAGOS"
            
            OPERACION = 23
            direccion = 3
        
        Case "DESHABILITA MONE 5 CEN PAGOS"
            
            OPERACION = 23
            direccion = 4
        
        Case "HABILITA MONE 10 CEN PAGOS"
                
            OPERACION = 23
            direccion = 5
        
        Case "DESHABILITA MONE 10 CEN PAGOS"
            
            OPERACION = 23
            direccion = 6
          
        Case "HABILITA MONE 20 CEN PAGOS"
            
            OPERACION = 23
            direccion = 7
        
        Case "DESHABILITA MONE 20 CEN PAGOS"
            
            OPERACION = 23
            direccion = 8
          
        Case "HABILITA MONE 50 CEN PAGOS"
            
            OPERACION = 23
            direccion = 9
        
        Case "DESHABILITA MONE 50 CEN PAGOS"
            
            OPERACION = 23
            direccion = 10
        
        Case "HABILITA MONE 100 CEN PAGOS"
            
            OPERACION = 23
            direccion = 11
        
        Case "DESHABILITA MONE 100 CEN PAGOS"
            
            OPERACION = 23
            direccion = 12
        
        Case "HABILITA MONE 200 CEN PAGOS"
            
            OPERACION = 23
            direccion = 13
        
        Case "DESHABILITA MONE 200 CEN PAGOS"
            
            OPERACION = 23
            direccion = 14
        
        'NIVELES DE MONEDAS
        Case "NIVEL MONE 5 CEN"
            
            If TConfi = "" Then
                MsgBox "INSERTA CONFIGURACION", 16, "FALTAN DATOS"
                Exit Sub
            End If
            
            CAN = TConfi
            DATOH = CAN \ 256
            DATOL = CAN Mod 256
            OPERACION = 24
            direccion = 2
            
        Case "NIVEL MONE 10 CEN"
            
            If TConfi = "" Then
                MsgBox "INSERTA CONFIGURACION", 16, "FALTAN DATOS"
                Exit Sub
            End If
            
            CAN = TConfi
            DATOH = CAN \ 256
            DATOL = CAN Mod 256
            OPERACION = 24
            direccion = 3
        
        Case "NIVEL MONE 20 CEN"
            
            If TConfi = "" Then
                MsgBox "INSERTA CONFIGURACION", 16, "FALTAN DATOS"
                Exit Sub
            End If
            
            CAN = TConfi
            DATOH = CAN \ 256
            DATOL = CAN Mod 256
            OPERACION = 24
            direccion = 4
        
        Case "NIVEL MONE 50 CEN"
            
            If TConfi = "" Then
                MsgBox "INSERTA CONFIGURACION", 16, "FALTAN DATOS"
                Exit Sub
            End If
            
            CAN = TConfi
            DATOH = CAN \ 256
            DATOL = CAN Mod 256
            OPERACION = 24
            direccion = 5
        
        Case "NIVEL MONE 100 CEN"
            
            If TConfi = "" Then
                MsgBox "INSERTA CONFIGURACION", 16, "FALTAN DATOS"
                Exit Sub
            End If
            
            CAN = TConfi
            DATOH = CAN \ 256
            DATOL = CAN Mod 256
            OPERACION = 24
            direccion = 6
        
        Case "NIVEL MONE 200 CEN"
            
            If TConfi = "" Then
                MsgBox "INSERTA CONFIGURACION", 16, "FALTAN DATOS"
                Exit Sub
            End If
            
            CAN = TConfi
            DATOH = CAN \ 256
            DATOL = CAN Mod 256
            OPERACION = 24
            direccion = 7
        
        'NIVELES DE BILLETES
        Case "NIVEL BILL 5 EUROS"
            
            If TConfi = "" Then
                MsgBox "INSERTA CONFIGURACION", 16, "FALTAN DATOS"
                Exit Sub
            End If
            
            CAN = TConfi
            DATOH = CAN \ 256
            DATOL = CAN Mod 256
            OPERACION = 25
            direccion = 1
        
        Case "NIVEL BILL 10 EUROS"
            
            If TConfi = "" Then
                MsgBox "INSERTA CONFIGURACION", 16, "FALTAN DATOS"
                Exit Sub
            End If
            
            CAN = TConfi
            DATOH = CAN \ 256
            DATOL = CAN Mod 256
            OPERACION = 25
            direccion = 2
        
        Case "NIVEL BILL 20 EUROS"
            
            If TConfi = "" Then
                MsgBox "INSERTA CONFIGURACION", 16, "FALTAN DATOS"
                Exit Sub
            End If
                
            CAN = TConfi
            DATOH = CAN \ 256
            DATOL = CAN Mod 256
            OPERACION = 25
            direccion = 3
        
        Case "NIVEL BILL 50 EUROS"
            
            If TConfi = "" Then
                MsgBox "INSERTA CONFIGURACION", 16, "FALTAN DATOS"
                Exit Sub
            End If
            
            CAN = TConfi
            DATOH = CAN \ 256
            DATOL = CAN Mod 256
            OPERACION = 25
            direccion = 4
        
        Case "NIVEL BILL 100 EUROS"
            
            If TConfi = "" Then
                MsgBox "INSERTA CONFIGURACION", 16, "FALTAN DATOS"
                Exit Sub
            End If
            
            CAN = TConfi
            DATOH = CAN \ 256
            DATOL = CAN Mod 256
            OPERACION = 25
            direccion = 5
                                                                
        Case "DESHABILITA BILLETERO"
            
            OPERACION = 29
                    
        Case "HABILITA SYSTEM"
                
            OPERACION = 30
        
        Case "DESHABILITA SYSTEM"
            
            OPERACION = 31
                    
        Case "CANTIDAD BILLETES"
                
            OPERACION = 33
                    
        Case "CANTIDAD MONEDAS"
            
            OPERACION = 34
                    
        Case "CARGA BILLETES"
            
            OPERACION = 37
                    
        Case "CARGA MONEDAS"
            
            OPERACION = 38
                    
        Case "PAGA CANTIDAD EN BILLETES"
            
            If TCantidad = "" Then
                MsgBox "INSERTA CANTIDAD", 16, "FALTAN DATOS"
                Exit Sub
            End If
            
            CAN = TCantidad * 100
            DATOH = CAN \ 256
            DATOL = CAN Mod 256
            OPERACION = 40
            
        Case "PAGA CANTIDAD EN MONEDAS"
            
            If TCantidad = "" Then
                MsgBox "INSERTA CANTIDAD", 16, "FALTAN DATOS"
                Exit Sub
            End If
            
            CAN = TCantidad * 100
            DATOH = CAN \ 256
            DATOL = CAN Mod 256
            OPERACION = 41
            
        Case "VACIADO HOPPER"
            
            OPERACION = 42
                    
        Case "VACIADO PAYOUT"
            
            OPERACION = 43
        
        Case "HACER CAJA"
            
            OPERACION = 44
        
        'IMPORTE A COBRAR
        Case "IMPORTE A COBRAR"
            
            If TCantidad = "" Then
                MsgBox "INSERTA CANTIDAD", 16, "FALTAN DATOS"
                Exit Sub
            End If
            
            CAN = TCantidad * 100
            DATOH = CAN \ 256
            DATOL = CAN Mod 256
            direccion = 0
            OPERACION = 39
            
            parcial_entradas = 0
            parcial_salidas = 0
            cadena = ""
            cadena_ant = ""
            suma_control = 0
            
            label_cobrar.Caption = TCantidad
            label_entradas.Caption = parcial_entradas / 100
            label_salidas.Caption = parcial_salidas / 100
            label_pagado.Caption = (parcial_entradas - parcial_salidas) / 100
            
            'insertamos en bbdd
            log_pagos "COBRAR " & TCantidad
            
            'cargamos en monedas antes lo que hay
            SQL = "SELECT * FROM log_ci ORDER BY codlog DESC LIMIT 1"
            Set rs = MiConexión.Execute(SQL)
            
            antes(0).Caption = rs("niv_m5c")
            antes(1).Caption = rs("niv_m10c")
            antes(2).Caption = rs("niv_m20c")
            antes(3).Caption = rs("niv_m50c")
            antes(4).Caption = rs("niv_m1")
            antes(5).Caption = rs("niv_m2")
            
            total_antes.Caption = antes(0).Caption * 0.05 + antes(1).Caption * 0.1 + antes(2).Caption * 0.2 + antes(3).Caption * 0.5 + antes(4).Caption + antes(5).Caption * 2
        
    End Select
    
    Dim aux_comando As String
    
    operacion_en_curso = True
    label_operacion_estado.Caption = "En proceso..."
    label_operacion_estado.ForeColor = vbRed
    
    aux_comando = LCase(CM)
    
    Debug.Print aux_comando
    
    label_operacion.Caption = UCase(Left(aux_comando, 1)) & Right(aux_comando, Len(aux_comando) - 1)
    
    text_log_operaciones.Text = Format(Now(), "hh:mm:ss") & " " & label_operacion.Caption & " " & TCantidad & vbCrLf & text_log_operaciones.Text

End Sub

Private Sub Command1_Click()

    If PARA Then
        PARA = 0
    Else
        PARA = 1
    End If
    
End Sub

Private Sub Command2_Click()
    comprueba_limites 20
End Sub

Private Sub Form_Load()

    conectar "ci"

    Set PCom = New MSComm
    With PCom
        .RThreshold = 1
        .SThreshold = 0
        .CommPort = 15
        .Handshaking = comNone
        .Settings = "9600,N,8,1"
        .InputMode = comInputModeText
        .PortOpen = True
        .InputLen = 1
    End With
    
    operacion_en_curso = False
    pago_en_curso = False
    label_operacion.Caption = ""
    label_operacion_estado.Caption = ""
    
    ContadorDatos = 0
    ContadorControl = 0
    OPERACION = 0
    direccion = 0
    DATOL = 0
    DATOH = 0
    PARA = 0
        
    'cargamos menu en combobox
    CM.AddItem "RESET"
    CM.AddItem "HABILITA BILLETERO"
    CM.AddItem "DESHABILITA BILLETERO"
    CM.AddItem "HABILITA SYSTEM"
    CM.AddItem "DESHABILITA SYSTEM"
    CM.AddItem "IMPORTE A COBRAR"
    CM.AddItem "HACER CAJA" '44
    CM.AddItem "HABILITA BILL 5 ENTRADA"
    CM.AddItem "DESHABILITA BILL 5 ENTRADA"
    CM.AddItem "HABILITA BILL 10 ENTRADA"
    CM.AddItem "DESHABILITA BILL 10 ENTRADA"
    CM.AddItem "HABILITA BILL 20 ENTRADA"
    CM.AddItem "DESHABILITA BILL 20 ENTRADA"
    CM.AddItem "HABILITA BILL 50 ENTRADA"
    CM.AddItem "DESHABILITA BILL 50 ENTRADA"
    CM.AddItem "HABILITA BILL 100 ENTRADA"
    CM.AddItem "DESHABILITA BILL 100 ENTRADA"
    CM.AddItem "HABILITA BILL 5 PARA PAGOS"
    CM.AddItem "DESHABILITA BILL 5 PARA PAGOS"
    CM.AddItem "HABILITA BILL 10 PARA PAGOS"
    CM.AddItem "DESHABILITA BILL 10 PARA PAGOS"
    CM.AddItem "HABILITA BILL 20 PARA PAGOS"
    CM.AddItem "DESHABILITA BILL 20 PARA PAGOS"
    CM.AddItem "HABILITA BILL 50 PARA PAGOS"
    CM.AddItem "DESHABILITA BILL 50 PARA PAGOS"
    CM.AddItem "HABILITA BILL 100 PARA PAGOS"
    CM.AddItem "DESHABILITA BILL 100 PARA PAGOS"
    CM.AddItem "HABILITA MONE 5 CEN ENTRADA"
    CM.AddItem "DESHABILITA MONE 5 CEN ENTRADA"
    CM.AddItem "HABILITA MONE 10 CEN ENTRADA"
    CM.AddItem "DESHABILITA MONE 10 CEN ENTRADA"
    CM.AddItem "HABILITA MONE 20 CEN ENTRADA"
    CM.AddItem "DESHABILITA MONE 20 CEN ENTRADA"
    CM.AddItem "HABILITA MONE 50 CEN ENTRADA"
    CM.AddItem "DESHABILITA MONE 50 CEN ENTRADA"
    CM.AddItem "HABILITA MONE 100 CEN ENTRADA"
    CM.AddItem "DESHABILITA MONE 100 CEN ENTRADA"
    CM.AddItem "HABILITA MONE 200 CEN ENTRADA"
    CM.AddItem "DESHABILITA MONE 200 CEN ENTRADA"
    CM.AddItem "HABILITA MONE 5 CEN PAGOS"
    CM.AddItem "DESHABILITA MONE 5 CEN PAGOS"
    CM.AddItem "HABILITA MONE 10 CEN PAGOS"
    CM.AddItem "DESHABILITA MONE 10 CEN PAGOS"
    CM.AddItem "HABILITA MONE 20 CEN PAGOS"
    CM.AddItem "DESHABILITA MONE 20 CEN PAGOS"
    CM.AddItem "HABILITA MONE 50 CEN PAGOS"
    CM.AddItem "DESHABILITA MONE 50 CEN PAGOS"
    CM.AddItem "HABILITA MONE 100 CEN PAGOS"
    CM.AddItem "DESHABILITA MONE 100 CEN PAGOS"
    CM.AddItem "HABILITA MONE 200 CEN PAGOS"
    CM.AddItem "DESHABILITA MONE 200 CEN PAGOS"
    CM.AddItem "NIVEL MONE 5 CEN"
    CM.AddItem "NIVEL MONE 10 CEN"
    CM.AddItem "NIVEL MONE 20 CEN"
    CM.AddItem "NIVEL MONE 50 CEN"
    CM.AddItem "NIVEL MONE 100 CEN"
    CM.AddItem "NIVEL MONE 200 CEN"
    CM.AddItem "NIVEL BILL 5 EUROS"
    CM.AddItem "NIVEL BILL 10 EUROS"
    CM.AddItem "NIVEL BILL 20 EUROS"
    CM.AddItem "NIVEL BILL 50 EUROS"
    CM.AddItem "NIVEL BILL 100 EUROS"
    CM.AddItem "CANTIDAD BILLETES"
    CM.AddItem "CANTIDAD MONEDAS"
    CM.AddItem "CARGA BILLETES"
    CM.AddItem "CARGA MONEDAS"
    CM.AddItem "PAGA CANTIDAD EN BILLETES"
    CM.AddItem "PAGA CANTIDAD EN MONEDAS"
    CM.AddItem "VACIADO HOPPER"
    CM.AddItem "VACIADO PAYOUT"
    
    label_cobrar.Caption = ""
    label_pagado.Caption = ""
    label_entradas.Caption = ""
    label_salidas.Caption = ""
    label_cajon.Caption = ""
    
    'reset a la placa al iniciar
    'para inicializar la placa deshabilitamos la entrada de monedas de 5 centimos por ejemplo
    CM = "DESHABILITA MONE 10 CEN PAGOS"
    CDatos_Click
    
    Dim i As Integer
    For i = 0 To 5
    
        antes(i).Caption = ""
        despues(i).Caption = ""
        diferencia(i).Caption = ""
    
    Next
    
    total_antes.Caption = ""
    total_despues.Caption = ""
    total_diferencia.Caption = ""
           
End Sub

Private Sub CHECKSUM()

    Dim N, M As Integer
    
    CHECK = 0
    
    For N = 1 To Len(CADE_CHECK)
        CHECK = CHECK + Asc(Mid(CADE_CHECK, N, 1))
    Next
  
salidachecck:
    If CHECK = 0 Then Exit Sub

    If CHECK < 256 Then
        CHECK = 256 - CByte(CHECK)
        Exit Sub
    Else
        CHECK = CHECK - 256
        GoTo salidachecck
    End If
  
End Sub

Private Sub TCantidad_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        'Case 13
        '    KeyAscii = 0        ' Para que no "pite"
        '    SendKeys "{tab}"    ' Envía una pulsación TAB
        Case 8, 44, 48 To 57
            'Case borrar, ",", Asc("0") To Asc("9")
            ' Sólo admitir teclas consideradas numéricas
            ' El .
            ' El 8 es la tecla Backspace (borrar hacia atrás)
        Case Else
            ' No es una tecla numérica, no admitirla
            KeyAscii = 0
            Beep
    End Select

End Sub

Private Sub timer_label_Timer()

    If label_operacion_estado.Visible = True Then
        label_operacion_estado.Visible = False
    Else
        label_operacion_estado.Visible = True
    End If

End Sub

Private Sub Tmax_Timer()

    Dim N  As Integer
    
    Tmax.Enabled = False
    PCom.Output = TRANS

    INDICADOR = INDICADOR + 1
    
    If INDICADOR > 3 Then
        TEventos = "NO HAY COMUNICACION CON EL MODULO"
    Else
        Tmax.Enabled = False
        Tmax.Interval = 5000
        Tmax.Enabled = True
    End If
 
    For N = 1 To Len(TRANS)
        TSalidas = TSalidas & Asc(Mid(TRANS, N, 1)) & " "
    Next
 
    TSalidas = TSalidas & vbCrLf
    CONTADOR = CONTADOR + 1

End Sub

Private Sub Pcom_OnComm()

    Dim S As String
    
    If PCom.CommEvent <> 2 Then Exit Sub
    S = PCom.Input
    SSE = SSE & S

    TRECEPCION.Enabled = False
    TRECEPCION.Interval = 100
    TRECEPCION.Enabled = True
    
End Sub

Private Sub TRECEPCION_Timer()

    Dim N, M, L As Integer
    Dim ACK As Integer
    Dim C As Integer
    Dim cd As Integer
    Dim CM As Integer
    Dim sT As String
    
    Dim SQL As String
    Dim rs As Recordset
         
    TRECEPCION.Enabled = False
 
    'suma_control = 0
 
    sT = ""
    
    For cd = 1 To Len(SSE)
        sT = sT & Asc(Mid(SSE, cd, 1)) & " "
    Next
          
    C = Asc(Mid(SSE, 1, 1))
 
    If C = 4 Or C = 5 Then         'CHECK SUM
        
        N = 0
        
        For cd = 1 To Len(SSE) - 1
            N = N + Asc(Mid(SSE, cd, 1))
            If N > 255 Then N = N - 256
        Next
        
        If N <> Asc(Mid(SSE, Len(SSE), 1)) Then
            SSE = ""
            Exit Sub
        End If
     
     End If
 
     Select Case (C)
 
        Case 0
                    
        Case 1
            
        Case 2
            
        Case 3
        
            If PARA = 0 Then
                TDatos = TDatos & sT & vbCrLf
                ContadorDatos = ContadorDatos + 1
                If ContadorDatos > 250 Then
                    ContadorDatos = 0
                    TDatos = ""
                End If
            End If
       
        Case 4
        
            PCom.Output = Chr(0) 'ACK
            '   OPERACION = 0
            '   DIRECCION = 0
            TCONSULTA.Enabled = False
            TCONSULTA.Interval = 300
            TCONSULTA.Enabled = True
                                    
            If operacion_en_curso = True Then
            
                timer_label.Enabled = True
                
                If cadena <> sT Then
                    cadena_ant = cadena
                    cadena = sT
                End If
            
                suma_control = Val(Trim(Right(cadena_ant, 3))) + Val(Trim(Right(cadena, 3)))
                                                                
                If (Trim(Right(sT, 3)) = 44 Or Trim(Right(sT, 3)) = 14) And suma_control = 58 Then
                    'ha finalizado la operacion y hopper y reciclador estan en reposo
                    label_operacion_estado.Visible = True
                    label_operacion_estado.ForeColor = vbGreen
                    label_operacion_estado.Caption = "Finalizado con éxito"
                    
                    operacion_en_curso = False
                    timer_label.Enabled = False
                    
                    If pago_en_curso = True Then
                        
                        pago_en_curso = False
                        'comprueba_limites
                    
                    End If
                                                            
                    If antes(0).Caption <> "" Then
                       
                        'cargamos en monedas antes lo que hay
                        SQL = "SELECT * FROM log_ci ORDER BY codlog DESC LIMIT 1"
                        Set rs = MiConexión.Execute(SQL)
            
                        despues(0).Caption = rs("niv_m5c")
                        despues(1).Caption = rs("niv_m10c")
                        despues(2).Caption = rs("niv_m20c")
                        despues(3).Caption = rs("niv_m50c")
                        despues(4).Caption = rs("niv_m1")
                        despues(5).Caption = rs("niv_m2")
                        
                        total_despues.Caption = despues(0).Caption * 0.05 + despues(1).Caption * 0.1 + despues(2).Caption * 0.2 + despues(3).Caption * 0.5 + despues(4).Caption + despues(5).Caption * 2
                        
                        Dim i As Integer
                        
                        For i = 0 To 5
                            
                            diferencia(i).Caption = despues(i).Caption - antes(i).Caption
                            
                        Next
                        
                        total_diferencia.Caption = total_despues.Caption - total_antes.Caption
                                            
                    End If
                                                            
                    TCantidad.Text = ""
                    TConfi.Text = ""
                    
                    text_log_operaciones.Text = Format(Now(), "hh:mm:ss") & " En reposo " & vbCrLf & text_log_operaciones.Text
                                        
                End If
                                
            End If
                        
            If PARA = 0 Then
                TControl = TControl & sT & vbCrLf
                ContadorControl = ContadorControl + 1
                
                'analizamos si primer numero es el 39
                If Left(sT, 2) = 39 Then
                
                    log_pagos sT
                
                End If
                
                'analizamos si primer numero es el 42
                If Left(sT, 2) = 42 Then
                
                    'envia monedas al cajon
                    log_cajonmonedas sT
                
                End If
                
                If ContadorControl > 250 Then
                    ContadorControl = 0
                    TControl = ""
                End If
            
            End If
            
        Case 5
       
            'TOTAL DE PARAMETROS
            'TControl = TControl & sT & vbCrLf 'PRIMERA COMUNICACION
            TEstado = sT
            
            'insertamos en bbdd
            log TEstado
            
            PCom.Output = Chr(0) 'ACK
            DATOL = 0
            DATOH = 0
            OPERACION = 0
            direccion = 0
            TCONSULTA.Enabled = False
            TCONSULTA.Interval = 300
            TCONSULTA.Enabled = True
            
        Case 28, 29, 30, 31 ', 42, 43
            
            TControl = TControl & sT & vbCrLf
            OPERACION = 0
            PCom.Output = Chr(0) 'ACK
            TCONSULTA.Enabled = False
            TCONSULTA.Interval = 300
            TCONSULTA.Enabled = True
            
            'vaciado payout
'            If C = 43 Then
            
                'pasamos los billetes al stacker
'                log_stacker sT
            
'            End If
            
'            If C = 42 Then
            
                'pasamos monedas al cajon
'                log_cajonmonedas sT
            
'            End If
                    
        Case 33
            
            TControl = TControl & sT & vbCrLf
            OPERACION = 0
            PCom.Output = Chr(0) 'ACK
            TCONSULTA.Enabled = False
            TCONSULTA.Interval = 300
            TCONSULTA.Enabled = True
            
        Case 34
            
            TControl = TControl & sT & vbCrLf
            OPERACION = 0
            PCom.Output = Chr(0) 'ACK
            TCONSULTA.Enabled = False
            TCONSULTA.Interval = 300
            TCONSULTA.Enabled = True
       
        Case 37
            
            If PARA = 0 Then TControl = TControl & sT & vbCrLf
            OPERACION = 0
            PCom.Output = Chr(0) 'ACK
    
        Case 38
            
            If PARA = 0 Then TControl = TControl & sT & vbCrLf
            OPERACION = 0
            PCom.Output = Chr(0) 'ACK
        
        Case 39
            
            TControl = TControl & sT & vbCrLf
            
            log_pagos sT
            
            OPERACION = 0
            PCom.Output = Chr(0) 'ACK
            'TCONSULTA.Enabled = False
            'TCONSULTA.Interval = 300
            'TCONSULTA.Enabled = True
                    
        Case 44, 42, 43
            
            TControl = TControl & sT & vbCrLf
            OPERACION = 0
            PCom.Output = Chr(0) 'ACK
            'TCONSULTA.Enabled = False
            'TCONSULTA.Interval = 300
            'TCONSULTA.Enabled = True
            
            'vaciado payout
            If C = 43 Then
            
                'pasamos los billetes al stacker
                log_stacker sT
            
            End If
            
            If C = 42 Then
            
                'pasamos monedas al cajon
                log_cajonmonedas sT
            
            End If
                    
        Case 50
            
            'TOTAL DE PARAMETROS
            TControl = TControl & sT & vbCrLf
            OPERACION = 0
            PCom.Output = Chr(0) 'ACK
            TCONSULTA.Enabled = False
            TCONSULTA.Interval = 300
            TCONSULTA.Enabled = True
            
    End Select
 
    SSE = ""
    Exit Sub

    '---------------------------------------------------------------
    'CAMBIAMOS COLOR DEL PUNTO
    If Punto.BackColor = &HFF00& Then 'SI HAY RESPUESTA
        Punto.BackColor = &H80000005
    Else
        Punto.BackColor = &HFF00&
    End If
    '---------------------------------------------------------------

End Sub

Private Sub TCONSULTA_Timer()

    Dim N, M, L As Integer
    Dim S, D As String
 
    TCONSULTA.Enabled = False
    S = Chr(OPERACION) & Chr(direccion) & Chr(DATOL) & Chr(DATOH)
    M = 0
    
    For N = 1 To 4
        M = M + Asc(Mid(S, N, 1))
        L = M
        If L > 255 Then L = L - 256
    Next
 
    S = S & Chr(L)
 
    If PARA = 0 Then
        D = ""
        For N = 1 To Len(S)
            D = D & Asc(Mid(S, N, 1)) & " "
        Next
        TControl = TControl & D & "       "
    End If
    
    PCom.Output = S

End Sub
