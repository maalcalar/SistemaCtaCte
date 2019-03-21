VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFracc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fraccionamiento de Aportaciones"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   10200
   Begin VB.TextBox txtDeuda 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   6240
      MaxLength       =   8
      TabIndex        =   31
      Top             =   1620
      Width           =   930
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5295
      Left            =   240
      TabIndex        =   19
      Top             =   2040
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9340
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CRONOGRAMA DE PAGOS DE FRACCIONAMIENTO"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbE_Socio 
      Height          =   315
      ItemData        =   "frmFracc.frx":0000
      Left            =   4080
      List            =   "frmFracc.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1080
      Width           =   3255
   End
   Begin VB.ComboBox cmbTipCob 
      Height          =   315
      ItemData        =   "frmFracc.frx":0004
      Left            =   7320
      List            =   "frmFracc.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox txtNumdoc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      MaxLength       =   8
      TabIndex        =   10
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox cmbGrado 
      Height          =   315
      ItemData        =   "frmFracc.frx":0008
      Left            =   1200
      List            =   "frmFracc.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtCodSocio 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      MaxLength       =   9
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtIns 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin MSMask.MaskEdBox txtFecDsc 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecIng 
      Height          =   285
      Left            =   360
      TabIndex        =   15
      Top             =   1620
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecRenu 
      Height          =   285
      Left            =   1440
      TabIndex        =   23
      Top             =   1620
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecExclu 
      Height          =   285
      Left            =   2520
      TabIndex        =   25
      Top             =   1620
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecExpul 
      Height          =   285
      Left            =   3600
      TabIndex        =   26
      Top             =   1620
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecRein 
      Height          =   285
      Left            =   4680
      TabIndex        =   27
      Top             =   1620
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Deuda x Cobrar"
      Height          =   195
      Index           =   17
      Left            =   6060
      TabIndex        =   32
      Top             =   1440
      Width           =   1110
   End
   Begin VB.Label Label11 
      Caption         =   "Fec.Reingreso"
      Height          =   210
      Left            =   4680
      TabIndex        =   30
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Fec.Expulsión"
      Height          =   210
      Left            =   3600
      TabIndex        =   29
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "Fec.Exclusión"
      Height          =   210
      Left            =   2520
      TabIndex        =   28
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "Fec.Renuncia"
      Height          =   210
      Left            =   1440
      TabIndex        =   24
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Estado de Socio"
      Height          =   195
      Index           =   16
      Left            =   4110
      TabIndex        =   18
      Top             =   900
      Width           =   1740
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "Fecha Ing."
      Height          =   210
      Left            =   240
      TabIndex        =   16
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tipo Cobro"
      Height          =   195
      Index           =   18
      Left            =   7980
      TabIndex        =   14
      Top             =   900
      Width           =   900
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "D.N.I."
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   900
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Grado"
      Height          =   195
      Index           =   2
      Left            =   1320
      TabIndex        =   11
      Top             =   900
      Width           =   1260
   End
   Begin VB.Label Label5 
      Caption         =   "Cod.Socio"
      Height          =   195
      Left            =   2640
      TabIndex        =   8
      Top             =   420
      Width           =   975
   End
   Begin VB.Label lblCodSocio 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   600
      Width           =   6375
   End
   Begin VB.Label Label6 
      Caption         =   "Nombre de Socio"
      Height          =   195
      Left            =   4200
      TabIndex        =   6
      Top             =   420
      Width           =   1695
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Codofin"
      Height          =   195
      Left            =   1320
      TabIndex        =   5
      Top             =   420
      Width           =   1335
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha"
      Height          =   210
      Left            =   360
      TabIndex        =   1
      Top             =   420
      Width           =   855
   End
End
Attribute VB_Name = "frmFracc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   frmFracc.Left = (Screen.Width - Width) \ 2
   frmFracc.Top = 0
   

End Sub

