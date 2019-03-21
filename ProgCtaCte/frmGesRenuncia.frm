VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmGesRenuncia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Renuncias de Socios"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   16665
   Begin VB.CommandButton cmdAnular 
      Caption         =   "&Anular Renuncia"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   14400
      TabIndex        =   47
      Top             =   1080
      Width           =   1695
      Begin VB.TextBox txtSdoNew 
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   360
         MaxLength       =   8
         TabIndex        =   49
         Top             =   300
         Width           =   930
      End
      Begin VB.TextBox txtAdelanto 
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   360
         MaxLength       =   8
         TabIndex        =   48
         Top             =   780
         Width           =   930
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Saldo Actual"
         Height          =   210
         Left            =   240
         TabIndex        =   51
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Adelanto"
         Height          =   210
         Left            =   240
         TabIndex        =   50
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.CommandButton cndOtro 
      Caption         =   "&Otro Socio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      TabIndex        =   44
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Renuncia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   120
      TabIndex        =   41
      Top             =   1200
      Width           =   11535
      Begin VB.TextBox txtObserva2 
         Height          =   285
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   55
         Top             =   720
         Width           =   8895
      End
      Begin VB.TextBox txtObservac 
         Height          =   285
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   53
         Top             =   420
         Width           =   9015
      End
      Begin MSMask.MaskEdBox txtFecNew 
         Height          =   285
         Left            =   480
         TabIndex        =   42
         Top             =   420
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   2760
         TabIndex        =   54
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha Renuncia Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar Renuncia"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15240
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   6480
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
      Height          =   615
      Left            =   15240
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   2895
      Left            =   12120
      TabIndex        =   38
      Top             =   2520
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5106
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
      Caption         =   "GESTIONES DE SOCIOS"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5175
      Left            =   120
      TabIndex        =   34
      Top             =   2520
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9128
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
      Caption         =   "PAGOS DE APORTES"
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
   Begin VB.ComboBox cmbTipCob 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmGesRenuncia.frx":0000
      Left            =   8160
      List            =   "frmGesRenuncia.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   780
      Width           =   2415
   End
   Begin VB.TextBox txtCarnetPNP 
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
      Enabled         =   0   'False
      Height          =   285
      Left            =   10560
      MaxLength       =   8
      TabIndex        =   15
      Top             =   780
      Width           =   930
   End
   Begin VB.TextBox txtCarnetPIP 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   11550
      MaxLength       =   8
      TabIndex        =   14
      Top             =   780
      Width           =   930
   End
   Begin VB.TextBox txtNumDoc 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      MaxLength       =   8
      TabIndex        =   5
      Top             =   300
      Width           =   930
   End
   Begin VB.TextBox txtIns 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   4
      Top             =   300
      Width           =   330
   End
   Begin VB.TextBox txtCodigo 
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
      Left            =   3240
      MaxLength       =   8
      TabIndex        =   3
      Top             =   300
      Width           =   930
   End
   Begin VB.TextBox txtCodSocio 
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
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      MaxLength       =   8
      TabIndex        =   2
      Top             =   300
      Width           =   690
   End
   Begin VB.ComboBox cmbGrado 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmGesRenuncia.frx":0004
      Left            =   3240
      List            =   "frmGesRenuncia.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   780
      Width           =   2775
   End
   Begin VB.ComboBox cmbE_Socio 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmGesRenuncia.frx":0008
      Left            =   6000
      List            =   "frmGesRenuncia.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   780
      Width           =   2175
   End
   Begin MSMask.MaskEdBox txtFecIng 
      Height          =   285
      Left            =   12720
      TabIndex        =   20
      Top             =   315
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecRenu 
      Height          =   285
      Left            =   13800
      TabIndex        =   21
      Top             =   315
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecExclu 
      Height          =   285
      Left            =   15000
      TabIndex        =   24
      Top             =   315
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecExpul 
      Height          =   285
      Left            =   12720
      TabIndex        =   25
      Top             =   795
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecRein 
      Height          =   285
      Left            =   13800
      TabIndex        =   26
      Top             =   795
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecCondo 
      Height          =   285
      Left            =   15000
      TabIndex        =   27
      Top             =   795
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecAmnis 
      Height          =   285
      Left            =   12720
      TabIndex        =   28
      Top             =   1275
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtTope 
      Height          =   285
      Left            =   240
      TabIndex        =   35
      Top             =   330
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label13 
      Caption         =   "Total Aportes"
      Height          =   255
      Left            =   4200
      TabIndex        =   46
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5280
      TabIndex        =   45
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Mes Tope"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   37
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label lblTope 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1080
      TabIndex        =   36
      Top             =   330
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Fec.Exclusión"
      Height          =   210
      Left            =   15000
      TabIndex        =   33
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Fec.Expulsión"
      Height          =   210
      Left            =   12720
      TabIndex        =   32
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Fec.Reingreso"
      Height          =   210
      Left            =   13800
      TabIndex        =   31
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Fec.Amnistia"
      Height          =   210
      Left            =   12720
      TabIndex        =   30
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Fec.Condonac."
      Height          =   210
      Left            =   15000
      TabIndex        =   29
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Fecha Ing."
      Height          =   210
      Left            =   12600
      TabIndex        =   23
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Fec.Renuncia"
      Height          =   210
      Left            =   13800
      TabIndex        =   22
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tipo Cobro"
      Height          =   195
      Index           =   18
      Left            =   8700
      TabIndex        =   19
      Top             =   600
      Width           =   780
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Carnet PNP"
      Height          =   195
      Index           =   6
      Left            =   10560
      TabIndex        =   18
      Top             =   600
      Width           =   840
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Carnet PIP"
      Height          =   195
      Index           =   7
      Left            =   11625
      TabIndex        =   17
      Top             =   600
      Width           =   765
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "D.N.I."
      Height          =   195
      Index           =   5
      Left            =   5475
      TabIndex        =   13
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Ins"
      Height          =   195
      Index           =   4
      Left            =   4200
      TabIndex        =   12
      Top             =   120
      Width           =   210
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Codofin"
      Height          =   195
      Index           =   1
      Left            =   3315
      TabIndex        =   11
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Codigo"
      Height          =   195
      Index           =   0
      Left            =   4665
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Grado"
      Height          =   195
      Index           =   2
      Left            =   3240
      TabIndex        =   9
      Top             =   600
      Width           =   915
   End
   Begin VB.Label lblCodSocio 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   300
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de Socio"
      Height          =   195
      Left            =   6360
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Estado de Socio"
      Height          =   195
      Index           =   16
      Left            =   6270
      TabIndex        =   6
      Top             =   600
      Width           =   1170
   End
End
Attribute VB_Name = "frmGesRenuncia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Limpiar()
   txtCodigo.Text = ""
   txtIns.Text = ""
   txtCodSocio.Text = ""
   txtNumdoc.Text = ""
   lblCodSocio.Caption = ""
   txtCarnetPNP.Text = ""
   txtCarnetPIP.Text = ""
   txtFecIng.Text = "__/__/____"
   txtFecRenu.Text = "__/__/____"
   txtFecExclu.Text = "__/__/____"
   txtFecExpul.Text = "__/__/____"
   txtFecRein.Text = "__/__/____"
   txtFecCondo.Text = "__/__/____"
   txtFecAmnis.Text = "__/__/____"
   
   txtSdoNew.Text = ""
   txtAdelanto.Text = ""
   txtFecNew.Text = "__/__/____"
   
   cmbGrado.ListIndex = 0
   cmbTipCob.ListIndex = 0
   cmbE_Socio.ListIndex = 0

   Set DataGrid1.DataSource = Nothing
   Set DataGrid2.DataSource = Nothing

End Sub

Private Sub Llenar()
   Dim zz As Integer, wCod As Long, wSoc As Integer, wIns As Integer, wTot As Currency
   
   wCod = Val(txtCodigo.Text)
   
   zz = Leerado8("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(wCod) + " ")
   If zz = 0 Then
      MsgBox "Codofin Digitado No Existe", vbExclamation
      Limpiar
      Exit Sub
   End If
   wSoc = ADO8!codsocio
   wIns = ADO8!ins
   
   txtCodigo.Text = ADO8!codigo
   txtIns.Text = ADO8!ins
   txtCodSocio.Text = ADO8!codsocio
   txtNumdoc.Text = ADO8!numdoc
   txtCarnetPNP.Text = ADO8!carnetpnp
   txtCarnetPIP.Text = ADO8!carnetpip
   If IsDate(ADO8!fecing) Then
      txtFecIng.Text = Format(ADO8!fecing, "dd/mm/yyyy")
   Else
      txtFecIng.Text = "__/__/____"
   End If
   If IsDate(ADO8!fecrenu) Then
      txtFecRenu.Text = Format(ADO8!fecrenu, "dd/mm/yyyy")
   Else
      txtFecRenu.Text = "__/__/____"
   End If
   If IsDate(ADO8!fecrein) Then
      txtFecRein.Text = Format(ADO8!fecrein, "dd/mm/yyyy")
   Else
      txtFecRein.Text = "__/__/____"
   End If
   If IsDate(ADO8!fecexclu) Then
      txtFecExclu.Text = Format(ADO8!fecexclu, "dd/mm/yyyy")
   Else
      txtFecExclu.Text = "__/__/____"
   End If
   If IsDate(ADO8!fecexpul) Then
      txtFecExpul.Text = Format(ADO8!fecexpul, "dd/mm/yyyy")
   Else
      txtFecExpul.Text = "__/__/____"
   End If
   If IsDate(ADO8!feccondo) Then
      txtFecCondo.Text = Format(ADO8!feccondo, "dd/mm/yyyy")
   Else
      txtFecCondo.Text = "__/__/____"
   End If
   If IsDate(ADO8!fecamnis) Then
      txtFecAmnis.Text = Format(ADO8!fecamnis, "dd/mm/yyyy")
   Else
      txtFecAmnis.Text = "__/__/____"
   End If

   cmbGrado.ListIndex = BuscaGrado(ADO8!grado)
   cmbE_Socio.ListIndex = BuscaEsocio(ADO8!e_socio)
   cmbTipCob.ListIndex = BuscaTipCob(ADO8!tipcob)
   Set ADO8 = Nothing

   zz = Leerado7a("SELECT M.LINEA, T.NOMBRE, M.FECHA " _
                & " FROM MAESOCIO_ACCION AS M INNER JOIN MAESOCIO_TIPOACCION AS T " _
                & "   ON M.TIPO = T.TIPO " _
                & " WHERE M.CODSOCIO = " + Str(wSoc) + " " _
                & " ORDER BY M.LINEA ")
   Set DataGrid2.DataSource = ADO7a

   DataGrid2.Columns(0).Width = 500
   DataGrid2.Columns(0).Alignment = dbgCenter
   DataGrid2.Columns(0).Caption = "LIN"
    
   DataGrid2.Columns(1).Width = 2200
   DataGrid2.Columns(1).Alignment = dbgCenter
   DataGrid2.Columns(1).Caption = "ACCION"
    
   DataGrid2.Columns(2).Width = 1100
   DataGrid2.Columns(2).Alignment = dbgCenter
   DataGrid2.Columns(2).Caption = "FECHA"
   DataGrid2.Columns(2).NumberFormat = "dd/mm/yyyy"

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_PAGOSAPORTE WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Dim II As Integer, wmmm As String
   For II = 1 To 12
       wmmm = Format(II, "00")
   
       Db.BeginTrans
       Db.Execute ("INSERT INTO TMP_PAGOSAPORTE " _
       & " (CODSOCIO, CODIGO, INS, TIPCOB, SERCOB, NUMCOB, FECHA, TIPO, NOMTIPO, MONEDA, IMPORTE, GLOSA, USU) " _
       & " SELECT " _
       & "  " + Str(wSoc) + ", Z.CODIGO, Z.INS, TIPAPOR, '0" + wmmm + "', '00000'+CUOANO, " _
       & "  '" + "15/" + wmmm + "/" + "'+CUOANO, Z.TIPAPOR, " _
       & "  CASE Z.TIPAPOR " _
       & "       WHEN '1' THEN 'DIECO' " _
       & "       WHEN '2' THEN 'CAJA M/P' " _
       & "       WHEN '4' THEN 'CAJA M/P' " _
       & "  END, " _
       & "  'S', Z.IMPO" + wmmm + ", '', '" + wcodusu + "' " _
       & " FROM ZZZ_APOR_PLA AS Z INNER JOIN MAESOCIO AS M ON Z.CODIGO = M.CODIGO AND Z.INS = M.INS " _
       & " WHERE Z.CODIGO = " + Str(wCod) + " AND Z.INS = " + Str(wIns) + " AND IMPO" + wmmm + " > 0 ")
       Db.CommitTrans
   Next
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_PAGOSAPORTE " _
   & " (CODSOCIO, CODIGO, INS, TIPCOB, SERCOB, NUMCOB, FECHA, TIPO, NOMTIPO, MONEDA, IMPORTE, GLOSA, USU) " _
   & " SELECT " _
   & "  " + Str(wSoc) + ", Z.CODIGO, Z.INS, '4', Z.SERIE, RIGHT('000000000' + CAST(Z.NRO_COMP AS VARCHAR),9), " _
   & "  Z.FECHA_PAGO, '4', 'TESORERIA', LEFT(Z.MONEDA,1), Z.MONTO, Z.OBS, '" + wcodusu + "' " _
   & " FROM ZZZ_MRECIBOS AS Z INNER JOIN ZZZ_CONCEPTO AS M " _
   & "   ON Z.CONCEPTO = M.CCONCE " _
   & " WHERE Z.CODIGO = " + Str(wCod) + " AND " _
   & "          Z.INS = " + Str(wIns) + " AND " _
   & "      (Z.MARCA2 <> 'A' OR Z.MARCA2 IS NULL) AND " _
   & "      (M.aporte = 1) ")
   Db.CommitTrans
   
   wTot = 0
   zz = Leerado6a("SELECT SUM(IMPORTE) AS IMPORTE " _
                & " FROM TMP_PAGOSAPORTE " _
                & " WHERE USU = '" + wcodusu + "' ")
   If zz > 0 Then
      wTot = IIf(IsNull(ADO6a!importe), 0, ADO6a!importe)
   End If
   lblTotal.Caption = Format(wTot, "######0.00;;\ ")
   
   
   zz = Leerado6a("SELECT FECHA, TIPCOB, SERCOB, NUMCOB, NOMTIPO, MONEDA, IMPORTE, GLOSA " _
                & " FROM TMP_PAGOSAPORTE " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY FECHA ")
   Set DataGrid1.DataSource = ADO6a

   DataGrid1.Columns(0).Width = 1050
   DataGrid1.Columns(0).Alignment = dbgCenter
   DataGrid1.Columns(0).Caption = "FECHA"
   DataGrid1.Columns(0).NumberFormat = "dd/mm/yyyy"
    
   DataGrid1.Columns(1).Width = 320
   DataGrid1.Columns(1).Alignment = dbgCenter
   DataGrid1.Columns(1).Caption = "TC"
    
   DataGrid1.Columns(2).Width = 390
   DataGrid1.Columns(2).Alignment = dbgCenter
   DataGrid1.Columns(2).Caption = "SER"
    
   DataGrid1.Columns(3).Width = 1000
   DataGrid1.Columns(3).Alignment = dbgCenter
   DataGrid1.Columns(3).Caption = "DCMTO"
    
   DataGrid1.Columns(4).Width = 1600
   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Caption = "TIPO COBRO"
    
   DataGrid1.Columns(5).Width = 320
   DataGrid1.Columns(5).Alignment = dbgCenter
   DataGrid1.Columns(5).Caption = "MON"
    
   DataGrid1.Columns(6).Width = 1000
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).Caption = "IMPORTE"
   DataGrid1.Columns(6).NumberFormat = "######0.00;;\ "

   DataGrid1.Columns(7).Width = 5400
   DataGrid1.Columns(7).Alignment = dbgLeft
   DataGrid1.Columns(7).Caption = "GLOSA"
    
End Sub

Private Sub cmdAnular_Click()
   Dim wSoc As Integer, wCod As Long, wIns As Integer, wNom As String, _
       wFecNew As Date, wFecOld As Date, wMesRen As String, wLin As String
   
   If Not IsDate(txtFecNew.Text) Then
      MsgBox "Fecha de Renuncia Es Invalida", vbExclamation
      txtFecNew.Text = "__/__/____"
      txtFecNew.SetFocus
      Exit Sub
   End If
   wSoc = Val(txtCodSocio.Text)
   wCod = Val(txtCodigo.Text)
   wIns = Val(txtIns.Text)
   wFecNew = Format(txtFecNew.Text, "dd/mm/yyyy")
   wMesRen = Format(wFecNew, "yyyy/mm")
   If Day(wFecNew) > 20 Then
      If Right(wMesRen, 2) = "12" Then
         wMesRen = Format(Val(Left(wMesRen, 4)) + 1, "0000") + "/01"
      Else
         wMesRen = Left(wMesRen, 4) + "/" + Format(Val(Right(wMesRen, 2)) + 1, "00")
      End If
   End If
   wLin = ""
   wNom = Trim(lblCodSocio.Caption)
   
   zz = Leerado8("SELECT MAX(LINEA) AS LINEA " _
                & " FROM MAESOCIO_ACCION " _
                & " WHERE CODSOCIO = " + Str(wSoc) + " AND TIPO = 'REN' AND FECHA = '" + Format(wFecNew, "dd/mm/yyyy") + "' ")
   If zz > 0 Then
      wLin = ADO8!linea
   End If
   Set ADO8 = Nothing
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM MAESOCIO_ACCION " _
   & " WHERE CODSOCIO = " + Str(wSoc) + " AND LINEA = '" + wLin + "' ")
   Db.CommitTrans

   wFecOld = Format("01/01/1900", "dd/mm/yyyy")
   zz = Leerado8("SELECT * " _
                & " FROM MAESOCIO_ACCION " _
                & " WHERE CODSOCIO = " + Str(wSoc) + " AND TIPO = 'REN' " _
                & " ORDER BY FECHA DESC ")
   If zz > 0 Then
      wFecOld = Format(ADO8!fecha, "dd/mm/yyyy")
   End If
   Set ADO8 = Nothing

   If Format(wFecOld, "dd/mm/yyyy") <> Format("01/01/1900", "dd/mm/yyyy") Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECRENU = '" + Format(wFecOld, "dd/mm/yyyy") + "', E_SOCIO = E_SOCIOOLD " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECRENU = NULL, E_SOCIO = E_SOCIOOLD " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If

   Call CreateAporteAnoMes(wSoc, wanocia, wFecNew)
       
   MsgBox "Renuncia Socio " + wNom + " Anulada OK", vbExclamation
   Unload Me
End Sub

Private Sub cmdGrabar_Click()
   Dim wSoc As Integer, wCod As Long, wIns As Integer, wNom As String, _
       wFecNew As Date, wMesRen As String, wLin As String, _
       wGlo As String, wGl2 As String
   
   If Not IsDate(txtFecNew.Text) Then
      MsgBox "Fecha de Renuncia Es Invalida", vbExclamation
      txtFecNew.Text = "__/__/____"
      txtFecNew.SetFocus
      Exit Sub
   End If
   wGlo = txtObservac.Text
   wGl2 = txtObserva2.Text
   wSoc = Val(txtCodSocio.Text)
   wCod = Val(txtCodigo.Text)
   wIns = Val(txtIns.Text)
   wFecNew = Format(txtFecNew.Text, "dd/mm/yyyy")
   wMesRen = Format(wFecNew, "yyyy/mm")
   If Day(wFecNew) > 20 Then
      If Right(wMesRen, 2) = "12" Then
         wMesRen = Format(Val(Left(wMesRen, 4)) + 1, "0000") + "/01"
      Else
         wMesRen = Left(wMesRen, 4) + "/" + Format(Val(Right(wMesRen, 2)) + 1, "00")
      End If
   End If
   
   wLin = ""
   wNom = Trim(lblCodSocio.Caption)
   
   zz = Leerado8("SELECT MAX(LINEA) AS LINEA " _
                & " FROM MAESOCIO_ACCION " _
                & " WHERE CODSOCIO = " + Str(wSoc) + " ")
   If zz > 0 Then
      wLin = ADO8!linea
   End If
   Set ADO8 = Nothing
   wLin = Format(Val(wLin) + 1, "00")
    
   Db.BeginTrans
   Db.Execute ("UPDATE MAESOCIO " _
   & " SET E_SOCIOOLD = E_SOCIO " _
   & " WHERE CODSOCIO = " + Str(wSoc) + " ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE MAESOCIO " _
   & " SET FECRENU = '" + Format(wFecNew, "dd/mm/yyyy") + "', E_SOCIO = 'REN' " _
   & " WHERE CODSOCIO = " + Str(wSoc) + " ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM CTASXDET " _
   & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
   & "       CONCEPTO = '01' AND " _
   & "            MES >= '" + wMesRen + "' AND " _
   & "         TIPMOV = '1' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO MAESOCIO_ACCION " _
   & " (CODSOCIO, LINEA, TIPO, FECHA, GLOSA, GLOS2) " _
   & " VALUES " _
   & " (" + Str(wSoc) + ", '" + wLin + "', 'REN', " _
   & "  '" + Format(wFecNew, "dd/mm/yyyy") + "', " _
   & "  '" + wGlo + "', '" + wGl2 + "' ) ")
   Db.CommitTrans
   
   MsgBox "Renuncia Socio " + wNom + " Grabada OK", vbExclamation
   Unload Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cndOtro_Click()
   Limpiar
   
   txtCodigo.SetFocus
End Sub

Private Sub Form_Activate()
   frmGesRenuncia.Left = (Screen.Width - Width) \ 2
   frmGesRenuncia.Top = 0
   
   Dim a As Integer, I As Integer, wAno As String, wMes As String
   
   txtTope.Text = Left(zMesTope, 4) + "/" + Right(zMesTope, 2)
   txtTope.Enabled = False
   
   a = Leerado8("SELECT * FROM MAEGRADO ORDER BY GRADO ")
   If a > 0 Then
      ADO8.MoveFirst
      I = 0
      Do While Not ADO8.EOF
         cmbGrado.AddItem ADO8!nombre
         I = I + 1
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   
   a = Leerado8("SELECT * FROM MAEE_SOCIO ORDER BY E_SOCIO ")
   If a > 0 Then
      ADO8.MoveFirst
      I = 0
      Do While Not ADO8.EOF
         cmbE_Socio.AddItem ADO8!nombre
         I = I + 1
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   
   a = Leerado8("SELECT * FROM MAETIPCOB ORDER BY TIPCOB ")
   If a > 0 Then
      ADO8.MoveFirst
      I = 0
      Do While Not ADO8.EOF
         cmbTipCob.AddItem ADO8!nombre
         I = I + 1
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   
   txtCodigo.SetFocus
End Sub

Private Sub txtCodigo_Change()
   Dim aa As Integer
   If Val(txtCodigo.Text) <> 0 Then
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(Val(txtCodigo.Text)) + " ")
      If aa > 0 Then
         lblCodSocio.Caption = ADO8!nombre
      Else
         lblCodSocio.Caption = ""
      End If
      Set ADO8 = Nothing
   Else
      lblCodSocio.Caption = ""
   End If
End Sub

Private Sub txtCodigo_GotFocus()
   txtCodigo.SelStart = 0
   txtCodigo.SelLength = 8
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim KeyNew As Integer
   KeyNew = Asc(UCase(Chr(KeyCode)))
   
   Select Case KeyCode
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio2.Show 1
        If xseleccion <> "" Then
           txtCodigo.Text = xseleccion
        End If
   Case 65 To 90
        xlista = "1"
        xseleccion = Chr(KeyNew)
        frmSelecSocio2.Show 1
        If xseleccion <> "" Then
           txtCodigo.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   Dim aa As Integer, wSoc As Integer, wCod As Long, wIns As Integer, _
       wSdo As Currency, wAde As Currency, wE_S As String, _
       wGlo As String, wGl2 As String
   
   If KeyAscii = 13 Then
      If Len(Trim(txtCodigo.Text)) = 0 Then
         MsgBox "Codofin En Blanco", vbExclamation
         txtCodigo.Text = ""
         Exit Sub
      End If
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(Val(txtCodigo.Text)) + " ")
      If aa = 0 Then
         MsgBox "Codofin Digitado NO Existe", vbExclamation
         txtCodigo.Text = ""
         Exit Sub
      End If
      wE_S = ADO8!e_socio
      
      If ADO8!e_socio = "REN" Then
         txtFecNew.Text = Format(ADO8!fecrenu, "dd/mm/yyyy")
         txtFecNew.Enabled = False
         cmdGrabar.Enabled = False
         cmdAnular.Enabled = True
      
               
         aa = Leerado7a("SELECT * FROM MAESOCIO_ACCION " _
                    & " WHERE CODSOCIO = " + Str(ADO8!codsocio) + " AND " _
                    & "           TIPO = 'REN' AND " _
                    & "          FECHA = '" + Format(txtFecNew.Text, "dd/mm/yyyy") + "' ")
         If aa > 0 Then
            wGlo = IIf(IsNull(ADO7a!glosa), "", ADO7a!glosa)
            wGl2 = IIf(IsNull(ADO7a!glos2), "", ADO7a!glos2)
         End If
         Set ADO7a = Nothing
         txtObservac.Text = wGlo
         txtObserva2.Text = wGl2
      
      Else
         txtFecNew.Text = Format(Date, "dd/mm/yyyy")
         
         cmdGrabar.Enabled = True
         cmdAnular.Enabled = False
      End If
      If wE_S = "FAL" Or wE_S = "EXP" Or wE_S = "EXC" Or wE_S = "SEP" Or wE_S = "998" Then
         MsgBox "Socio No Esta Habil" + vbNewLine + "No Puede Renunciar", vbExclamation
         txtCodigo.Text = ""
         Exit Sub
      End If
      
      wCod = Val(txtCodigo.Text)
      lblCodSocio.Caption = ADO8!nombre
      
      Llenar
   
      wIns = Val(txtIns.Text)
      wSoc = Val(txtCodSocio.Text)
   
      wAde = 0
      wSdo = SaldoFoto(wSoc, zMesTope)
      If wSdo < 0 Then
         wAde = -wSdo
         wSdo = 0
      End If
   
      txtSdoNew.Text = Format(wSdo, "####,##0.00;;\ ")
      txtAdelanto.Text = Format(wAde, "####,##0.00;;\ ")
   
      If wE_S = "REN" Then
         cmdAnular.SetFocus
      Else
         txtFecNew.SetFocus
      End If
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtFecNew_GotFocus()
   txtFecNew.SelStart = 0
   txtFecNew.SelLength = 10
End Sub

Private Sub txtFecNew_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 40
        txtObservac.SetFocus
   End Select
End Sub

Private Sub txtFecNew_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFecNew.Text = "__/__/____" Then
         MsgBox "Fecha de Renuncia En Blanco", vbExclamation
         txtFecNew.Text = "__/__/____"
         Exit Sub
      End If
      If Not IsDate(txtFecNew.Text) Then
         MsgBox "Fecha de Renuncia Digitada Es Invalida", vbExclamation
         txtFecNew.Text = "__/__/____"
         Exit Sub
      End If
      txtObservac.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtObserva2_GotFocus()
   txtObserva2.SelStart = 0
   txtObserva2.SelLength = 50
End Sub

Private Sub txtObserva2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtObservac.SetFocus
   End Select
End Sub

Private Sub txtObserva2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdGrabar.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtObservac_GotFocus()
   txtObservac.SelStart = 0
   txtObservac.SelLength = 50
End Sub

Private Sub txtObservac_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFecNew.SetFocus
   Case 40
        txtObserva2.SetFocus
   End Select
End Sub

Private Sub txtObservac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtObserva2.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtTope_Change()
   Dim wMes As String, wAno As String
   If txtTope.Text <> "____-__" Then
      wAno = Left(txtTope.Text, 4)
      wMes = Right(txtTope.Text, 2)
               
      If wMes <> "01" And wMes <> "02" And wMes <> "03" And wMes <> "04" And wMes <> "05" And wMes <> "06" And _
         wMes <> "07" And wMes <> "08" And wMes <> "09" And wMes <> "10" And wMes <> "11" And wMes <> "12" Then
         lblTope.Caption = ""
      Else
         lblTope.Caption = Trim(funnommes(wMes)) + " " + wAno
      End If
   Else
      lblTope.Caption = ""
   End If
End Sub
