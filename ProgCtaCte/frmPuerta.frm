VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEleConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Datos del Asociado"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12540
   Icon            =   "frmPuerta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   12540
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir Estado Pagos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10440
      TabIndex        =   79
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdDIECO 
      Caption         =   "  Imprimir      Carta     DIECO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9600
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCAJMP 
      Caption         =   "Imprimir  Carta    CAJA M.P."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11040
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtPago2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   75
      Top             =   4020
      Width           =   9615
   End
   Begin VB.TextBox txtPago1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   74
      Top             =   3720
      Width           =   9615
   End
   Begin VB.OptionButton optDNI 
      Caption         =   "Consulta x DNI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   9240
      TabIndex        =   72
      Top             =   480
      Width           =   2895
   End
   Begin VB.OptionButton optCodofin 
      Caption         =   "Consulta x Codofin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   9240
      TabIndex        =   71
      Top             =   120
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.CheckBox chkVip 
      Caption         =   "Socio VIP"
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
      Height          =   255
      Left            =   5760
      TabIndex        =   68
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtObservac 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   61
      Top             =   3060
      Width           =   4935
   End
   Begin VB.TextBox txtObserva2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   60
      Top             =   3360
      Width           =   4935
   End
   Begin VB.ComboBox cmbSexo 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmPuerta.frx":030A
      Left            =   4440
      List            =   "frmPuerta.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   2100
      Width           =   1815
   End
   Begin VB.ComboBox cmbECivil 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmPuerta.frx":030E
      Left            =   2400
      List            =   "frmPuerta.frx":0310
      Style           =   2  'Dropdown List
      TabIndex        =   54
      Top             =   2100
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   120
      TabIndex        =   49
      Top             =   4440
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4471
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
      Caption         =   "RELACION DE FAMILIARES"
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
   Begin VB.Frame Frame1 
      Caption         =   "Situación de Aportes"
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
      Height          =   1455
      Left            =   7680
      TabIndex        =   44
      Top             =   2160
      Width           =   4575
      Begin VB.TextBox txtEnv541 
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
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   65
         Top             =   1080
         Width           =   930
      End
      Begin VB.TextBox txtEnv540 
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   63
         Top             =   1080
         Width           =   930
      End
      Begin VB.TextBox txtSaldo 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   48
         Top             =   240
         Width           =   930
      End
      Begin VB.TextBox txtAdelan 
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
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   45
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Asignados"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3480
         TabIndex        =   67
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Titular"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2400
         TabIndex        =   66
         Top             =   900
         Width           =   975
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Envios En Proceso"
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   645
         TabIndex        =   64
         Top             =   1140
         Width           =   1620
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sdos Cobrar"
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Adelanto"
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   2520
         TabIndex        =   46
         Top             =   300
         Width           =   765
      End
   End
   Begin VB.ComboBox cmbSituEsp 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmPuerta.frx":0312
      Left            =   9840
      List            =   "frmPuerta.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   1620
      Width           =   2055
   End
   Begin VB.ComboBox cmbSitu 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmPuerta.frx":0316
      Left            =   7800
      List            =   "frmPuerta.frx":0318
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   1620
      Width           =   2055
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
      Left            =   9960
      MaxLength       =   8
      TabIndex        =   37
      Top             =   1140
      Width           =   930
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
      Left            =   9000
      MaxLength       =   8
      TabIndex        =   36
      Top             =   1140
      Width           =   930
   End
   Begin VB.TextBox txtNumReso 
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
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   26
      Top             =   2580
      Width           =   1410
   End
   Begin VB.ComboBox cmbTipCob 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmPuerta.frx":031A
      Left            =   5160
      List            =   "frmPuerta.frx":031C
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1620
      Width           =   2655
   End
   Begin VB.ComboBox cmbE_Socio 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmPuerta.frx":031E
      Left            =   2640
      List            =   "frmPuerta.frx":0320
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1620
      Width           =   2535
   End
   Begin VB.TextBox txtNumdoc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   10920
      MaxLength       =   8
      TabIndex        =   13
      Top             =   1140
      Width           =   975
   End
   Begin VB.ComboBox cmbGrado 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmPuerta.frx":0322
      Left            =   240
      List            =   "frmPuerta.frx":0324
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1620
      Width           =   2415
   End
   Begin VB.TextBox txtCodSocio 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   9
      TabIndex        =   8
      Top             =   1140
      Width           =   975
   End
   Begin VB.TextBox txtIns 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1140
      Width           =   375
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   240
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1140
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
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
      Height          =   855
      Left            =   11040
      Picture         =   "frmPuerta.frx":0326
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOtro 
      Caption         =   "Otra Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9600
      Picture         =   "frmPuerta.frx":0768
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   1215
   End
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmPuerta.frx":0BAA
      Left            =   1080
      List            =   "frmPuerta.frx":0BAC
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   7935
   End
   Begin MSMask.MaskEdBox txtFecNac 
      Height          =   285
      Left            =   240
      TabIndex        =   20
      Top             =   2100
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecIng 
      Height          =   285
      Left            =   6360
      TabIndex        =   22
      Top             =   2100
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
      Left            =   240
      TabIndex        =   23
      Top             =   2580
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecReso 
      Height          =   285
      Left            =   1440
      TabIndex        =   27
      Top             =   2580
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
      Left            =   4080
      TabIndex        =   28
      Top             =   2580
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
      Left            =   5160
      TabIndex        =   29
      Top             =   2580
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
      Left            =   6240
      TabIndex        =   30
      Top             =   2580
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
      Left            =   1080
      TabIndex        =   50
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecMat 
      Height          =   285
      Left            =   1320
      TabIndex        =   58
      Top             =   2100
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecVip 
      Height          =   285
      Left            =   6240
      TabIndex        =   69
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin Crystal.CrystalReport Crys1 
      Left            =   120
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Reporte de Diarios"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin Crystal.CrystalReport Crys2 
      Left            =   720
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Reporte de Diarios"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label lblE_Socio 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   5040
      TabIndex        =   76
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Ultimos Cobros"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   73
      Top             =   3720
      Width           =   1050
   End
   Begin VB.Label Label9 
      Caption         =   "Fecha VIP"
      Height          =   210
      Left            =   5400
      TabIndex        =   70
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones"
      Height          =   195
      Index           =   22
      Left            =   480
      TabIndex        =   62
      Top             =   2880
      Width           =   1065
   End
   Begin VB.Label Label8 
      Caption         =   "Fecha Matrim"
      Enabled         =   0   'False
      Height          =   210
      Left            =   1320
      TabIndex        =   59
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Sexo"
      Enabled         =   0   'False
      Height          =   195
      Index           =   10
      Left            =   4680
      TabIndex        =   57
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Estado Civil"
      Enabled         =   0   'False
      Height          =   195
      Index           =   9
      Left            =   3000
      TabIndex        =   55
      Top             =   1920
      Width           =   825
   End
   Begin VB.Label lblMensaje 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   600
      TabIndex        =   53
      Top             =   7320
      Width           =   8775
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   120
      TabIndex        =   52
      Top             =   480
      Width           =   930
   End
   Begin VB.Label lblTope 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1920
      TabIndex        =   51
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Situación Especial"
      Height          =   195
      Index           =   23
      Left            =   10095
      TabIndex        =   43
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Situación Policial"
      Height          =   195
      Index           =   8
      Left            =   8160
      TabIndex        =   41
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Carnet PIP"
      Height          =   195
      Index           =   7
      Left            =   10035
      TabIndex        =   39
      Top             =   960
      Width           =   765
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Carnet PNP"
      Height          =   195
      Index           =   6
      Left            =   9000
      TabIndex        =   38
      Top             =   960
      Width           =   840
   End
   Begin VB.Label Label11 
      Caption         =   "Fec.Reingreso"
      Height          =   210
      Left            =   6240
      TabIndex        =   35
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Fec.Expulsión"
      Height          =   210
      Left            =   5160
      TabIndex        =   34
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "Fec.Exclusión"
      Height          =   210
      Left            =   4080
      TabIndex        =   33
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nro.Resol.Ingreso"
      Height          =   195
      Index           =   21
      Left            =   2640
      TabIndex        =   32
      Top             =   2400
      Width           =   1275
   End
   Begin VB.Label Label15 
      Caption         =   "Fec.Resol.Ing"
      Height          =   210
      Left            =   1440
      TabIndex        =   31
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "Fec.Renuncia"
      Height          =   210
      Left            =   240
      TabIndex        =   25
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "Fecha Ing."
      Height          =   210
      Left            =   6240
      TabIndex        =   24
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label14 
      Caption         =   "Fecha Nacim."
      Height          =   210
      Left            =   240
      TabIndex        =   21
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tipo Cobro"
      Height          =   195
      Index           =   18
      Left            =   5820
      TabIndex        =   19
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Estado de Socio"
      Height          =   195
      Index           =   16
      Left            =   2910
      TabIndex        =   17
      Top             =   1440
      Width           =   1170
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "D.N.I."
      Height          =   195
      Left            =   10920
      TabIndex        =   15
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Grado"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de Socio"
      Height          =   195
      Left            =   2640
      TabIndex        =   11
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblCodSocio 
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   1140
      Width           =   6495
   End
   Begin VB.Label Label3 
      Caption         =   "Cod.Socio"
      Height          =   195
      Left            =   1560
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Ins"
      Height          =   195
      Left            =   1200
      TabIndex        =   7
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Codofin"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmEleConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbCia_Click()
   cmbCia_KeyPress (13)
End Sub

Private Sub cmbCia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
   
   
   End If
End Sub

Private Sub cmdCAJMP_Click()
   Dim wNombre As String, wDni As String, wCodofin As String, _
       wDirec As String, wDist As String, wCorreo As String, _
       wTelefono As String, wCelular As String, wSdoPen As String, _
       wCuoIni As String, wCanMes As String, wCuoMes As String, _
       wFecDia As String, wFecMes As String, wFecAno As String, _
       zz As Integer, wSoc As Integer, wCod As Long, wIns As Integer, _
       wNomAsig1 As String, wDniAsig1 As String, wCodAsig1 As Integer, _
       wNomAsig2 As String, wDniAsig2 As String, wCodAsig2 As Integer, _
       wNomAsig3 As String, wDniAsig3 As String, wCodAsig3 As Integer, _
       wNomAsig4 As String, wDniAsig4 As String, wCodAsig4 As Integer, _
       wNomAsig5 As String, wDniAsig5 As String, wCodAsig5 As Integer, _
       wCip As String
   
   wSoc = Val(txtCodSocio.Text)
   wNombre = Trim(lblCodSocio.Caption)
   wDni = txtNumdoc.Text
   wCodofin = Trim(txtCodigo.Text) + "-" + Trim(txtIns.Text)
   wDirec = "": wDist = "": wCorreo = "": wTelefono = "": wCelular = "": wCip = ""
   wNomAsig1 = "": wDniAsig1 = "": wCodAsig1 = 0
   wNomAsig2 = "": wDniAsig2 = "": wCodAsig2 = 0
   wNomAsig3 = "": wDniAsig3 = "": wCodAsig3 = 0
   wNomAsig4 = "": wDniAsig4 = "": wCodAsig4 = 0
   wNomAsig5 = "": wDniAsig5 = "": wCodAsig5 = 0
   
   zz = Leerado8("SELECT * FROM MAEASIGNADO WHERE CODSOCIO = " + Str(wSoc) + " ")
   If zz > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
   
         Select Case ADO8!lin
         Case "01"
              wCodAsig1 = ADO8!codhijo
              wNomAsig1 = BuscaDatosSocio(wCodAsig1, 3)
              wDniAsig1 = BuscaDatosSocio(wCodAsig1, 4)
         Case "02"
              wCodAsig2 = ADO8!codhijo
              wNomAsig2 = BuscaDatosSocio(wCodAsig2, 3)
              wDniAsig2 = BuscaDatosSocio(wCodAsig2, 4)
         Case "03"
              wCodAsig3 = ADO8!codhijo
              wNomAsig3 = BuscaDatosSocio(wCodAsig3, 3)
              wDniAsig3 = BuscaDatosSocio(wCodAsig3, 4)
         Case "04"
              wCodAsig4 = ADO8!codhijo
              wNomAsig4 = BuscaDatosSocio(wCodAsig4, 3)
              wDniAsig4 = BuscaDatosSocio(wCodAsig4, 4)
         Case "05"
              wCodAsig5 = ADO8!codhijo
              wNomAsig5 = BuscaDatosSocio(wCodAsig5, 3)
              wDniAsig5 = BuscaDatosSocio(wCodAsig5, 4)
         End Select
   
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
   If zz > 0 Then
      wDirec = ADO8!direc
      wDist = ADO8!ubigeo
      If Len(Trim(ADO8!email)) > 0 Then
         wCorreo = ADO8!email
      Else
         If Len(Trim(ADO8!email2)) > 0 Then
            wCorreo = ADO8!email2
         End If
      End If
      wTelefono = Trim(IIf(IsNull(ADO8!telefono), "", Trim(ADO8!telefono)) + " " + _
                       IIf(IsNull(ADO8!telefon2), "", Trim(ADO8!telefon2)))
      wCelular = IIf(IsNull(ADO8!celular), "", Trim(ADO8!celular))
      wCip = IIf(IsNull(ADO8!carnetpnp), "", ADO8!carnetpnp)
   
   
      zz = Leerado7("SELECT * FROM MAESITU WHERE SITU = " + Str(ADO8!situ) + " ")
      If zz > 0 Then
         wSituac = ADO7!nombre
      End If
      Set ADO7 = Nothing
   
   End If
   Set ADO8 = Nothing
   
   If Len(Trim(wDist)) > 0 Then
      zz = Leerado8("SELECT * FROM MAEUBIGEO WHERE CODIGO = '" + wDist + "' ")
      If zz > 0 Then
         wDist = ADO8!nombre
      End If
      Set ADO8 = Nothing
   End If
   wFecDia = Format(Day(Date), "00")
   wFecMes = Trim(funnommes(Format(Month(Date), "00")))
   wFecAno = Right(Format(Year(Date), "0000"), 2)
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\DesctoCAJMP.RPT"
   Crys1.Formulas(0) = "NOMBRE= '" + wNombre + "' "
   Crys1.Formulas(1) = "DNI= '" + wDni + "' "
   Crys1.Formulas(2) = "CODOFIN= '" + wCodofin + "' "
   Crys1.Formulas(3) = "DIREC= '" + wDirec + "' "
   Crys1.Formulas(4) = "DIST= '" + wDist + "' "
   Crys1.Formulas(5) = "CORREO= '" + wCorreo + "' "
   Crys1.Formulas(6) = "TELEFONO= '" + wTelefono + "' "
   Crys1.Formulas(7) = "CELULAR= '" + wCelular + "' "
   Crys1.Formulas(8) = "FECDIA= '" + wFecDia + "' "
   Crys1.Formulas(9) = "FECMES= '" + wFecMes + "' "
   Crys1.Formulas(10) = "FECANO= '" + wFecAno + "' "
   Crys1.Formulas(11) = "NOMASIG1= '" + wNomAsig1 + "' "
   Crys1.Formulas(12) = "NOMASIG2= '" + wNomAsig2 + "' "
   Crys1.Formulas(13) = "NOMASIG3= '" + wNomAsig3 + "' "
   Crys1.Formulas(14) = "NOMASIG4= '" + wNomAsig4 + "' "
   Crys1.Formulas(15) = "NOMASIG5= '" + wNomAsig5 + "' "
   Crys1.Formulas(16) = "DNIASIG1= '" + wDniAsig1 + "' "
   Crys1.Formulas(17) = "DNIASIG2= '" + wDniAsig2 + "' "
   Crys1.Formulas(18) = "DNIASIG3= '" + wDniAsig3 + "' "
   Crys1.Formulas(19) = "DNIASIG4= '" + wDniAsig4 + "' "
   Crys1.Formulas(20) = "DNIASIG5= '" + wDniAsig5 + "' "
   Crys1.Formulas(21) = "CIP= '" + wCip + "' "
   Crys1.Formulas(22) = "SITUAC= '" + wSituac + "' "
   Crys1.SelectionFormula = " {MAESOCIO.CODSOCIO}=" + Str(wSoc) + " "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdDIECO_Click()
   Dim wNombre As String, wDni As String, wCodofin As String, _
       wDirec As String, wDist As String, wCorreo As String, _
       wTelefono As String, wCelular As String, wSdoPen As String, _
       wCuoIni As String, wCanMes As String, wCuoMes As String, _
       wFecDia As String, wFecMes As String, wFecAno As String, _
       zz As Integer, wSoc As Integer, wCod As Long, wIns As Integer, _
       wNomAsig1 As String, wDniAsig1 As String, wCodAsig1 As Integer, _
       wNomAsig2 As String, wDniAsig2 As String, wCodAsig2 As Integer, _
       wNomAsig3 As String, wDniAsig3 As String, wCodAsig3 As Integer, _
       wNomAsig4 As String, wDniAsig4 As String, wCodAsig4 As Integer, _
       wNomAsig5 As String, wDniAsig5 As String, wCodAsig5 As Integer, _
       wCip As String, wSituac As String
   
   wSoc = Val(txtCodSocio.Text)
   wNombre = Trim(lblCodSocio.Caption)
   wDni = txtNumdoc.Text
   wCodofin = Trim(txtCodigo.Text) + "-" + Trim(txtIns.Text)
   wDirec = "": wDist = "": wCorreo = "": wTelefono = "": wCelular = "": wCip = ""
   
   wNomAsig1 = "": wDniAsig1 = "": wCodAsig1 = 0
   wNomAsig2 = "": wDniAsig2 = "": wCodAsig2 = 0
   wNomAsig3 = "": wDniAsig3 = "": wCodAsig3 = 0
   wNomAsig4 = "": wDniAsig4 = "": wCodAsig4 = 0
   wNomAsig5 = "": wDniAsig5 = "": wCodAsig5 = 0
   
   zz = Leerado8("SELECT * FROM MAEASIGNADO WHERE CODSOCIO = " + Str(wSoc) + " ")
   If zz > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
   
         Select Case ADO8!lin
         Case "01"
              wCodAsig1 = ADO8!codhijo
              wNomAsig1 = BuscaDatosSocio(wCodAsig1, 3)
              wDniAsig1 = BuscaDatosSocio(wCodAsig1, 4)
         Case "02"
              wCodAsig2 = ADO8!codhijo
              wNomAsig2 = BuscaDatosSocio(wCodAsig2, 3)
              wDniAsig2 = BuscaDatosSocio(wCodAsig2, 4)
         Case "03"
              wCodAsig3 = ADO8!codhijo
              wNomAsig3 = BuscaDatosSocio(wCodAsig3, 3)
              wDniAsig3 = BuscaDatosSocio(wCodAsig3, 4)
         Case "04"
              wCodAsig4 = ADO8!codhijo
              wNomAsig4 = BuscaDatosSocio(wCodAsig4, 3)
              wDniAsig4 = BuscaDatosSocio(wCodAsig4, 4)
         Case "05"
              wCodAsig5 = ADO8!codhijo
              wNomAsig5 = BuscaDatosSocio(wCodAsig5, 3)
              wDniAsig5 = BuscaDatosSocio(wCodAsig5, 4)
         End Select
   
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
   If zz > 0 Then
      wDirec = ADO8!direc
      wDist = ADO8!ubigeo
      If Len(Trim(ADO8!email)) > 0 Then
         wCorreo = ADO8!email
      Else
         If Len(Trim(ADO8!email2)) > 0 Then
            wCorreo = ADO8!email2
         End If
      End If
      wTelefono = Trim(IIf(IsNull(ADO8!telefono), "", Trim(ADO8!telefono)) + " " + _
                       IIf(IsNull(ADO8!telefon2), "", Trim(ADO8!telefon2)))
      wCelular = IIf(IsNull(ADO8!celular), "", Trim(ADO8!celular))
      wCip = IIf(IsNull(ADO8!carnetpnp), "", ADO8!carnetpnp)
   
      zz = Leerado7("SELECT * FROM MAESITU WHERE SITU = " + Str(ADO8!situ) + " ")
      If zz > 0 Then
         wSituac = ADO7!nombre
      End If
      Set ADO7 = Nothing
   
   End If
   Set ADO8 = Nothing
   
   If Len(Trim(wDist)) > 0 Then
      zz = Leerado8("SELECT * FROM MAEUBIGEO WHERE CODIGO = '" + wDist + "' ")
      If zz > 0 Then
         wDist = ADO8!nombre
      End If
      Set ADO8 = Nothing
   End If
   wFecDia = Format(Day(Date), "00")
   wFecMes = Trim(funnommes(Format(Month(Date), "00")))
   wFecAno = Right(Format(Year(Date), "0000"), 2)
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\DesctoDIECO.RPT"
   Crys1.Formulas(0) = "NOMBRE= '" + wNombre + "' "
   Crys1.Formulas(1) = "DNI= '" + wDni + "' "
   Crys1.Formulas(2) = "CODOFIN= '" + wCodofin + "' "
   Crys1.Formulas(3) = "DIREC= '" + wDirec + "' "
   Crys1.Formulas(4) = "DIST= '" + wDist + "' "
   Crys1.Formulas(5) = "CORREO= '" + wCorreo + "' "
   Crys1.Formulas(6) = "TELEFONO= '" + wTelefono + "' "
   Crys1.Formulas(7) = "CELULAR= '" + wCelular + "' "
   Crys1.Formulas(8) = "FECDIA= '" + wFecDia + "' "
   Crys1.Formulas(9) = "FECMES= '" + wFecMes + "' "
   Crys1.Formulas(10) = "FECANO= '" + wFecAno + "' "
   Crys1.Formulas(11) = "NOMASIG1= '" + wNomAsig1 + "' "
   Crys1.Formulas(12) = "NOMASIG2= '" + wNomAsig2 + "' "
   Crys1.Formulas(13) = "NOMASIG3= '" + wNomAsig3 + "' "
   Crys1.Formulas(14) = "NOMASIG4= '" + wNomAsig4 + "' "
   Crys1.Formulas(15) = "NOMASIG5= '" + wNomAsig5 + "' "
   Crys1.Formulas(16) = "DNIASIG1= '" + wDniAsig1 + "' "
   Crys1.Formulas(17) = "DNIASIG2= '" + wDniAsig2 + "' "
   Crys1.Formulas(18) = "DNIASIG3= '" + wDniAsig3 + "' "
   Crys1.Formulas(19) = "DNIASIG4= '" + wDniAsig4 + "' "
   Crys1.Formulas(20) = "DNIASIG5= '" + wDniAsig5 + "' "
   Crys1.Formulas(21) = "CIP= '" + wCip + "' "
   Crys1.Formulas(22) = "SITUAC= '" + wSituac + "' "
   Crys1.SelectionFormula = " {MAESOCIO.CODSOCIO}=" + Str(wSoc) + " "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdImprimir_Click()
   Dim aa As Long, wRegAct As Long, wRegTot As Long, wSoc As Integer

   Dim wCod As Long, wIns As Long, wNom As String, wLin As Integer, _
       wRec As String, wMon As String, wImp As Currency, wFec As Date, _
       wObs As String, wNde As Currency, wnCr As Currency, wDeu As Currency, _
       wCer As Currency, wAde As Currency, wFecTope As Date, _
       wMesTope As String, wAnoTope As String, wDiaTope As String, _
       wVip As String, wCartaDieco As String, wFracSw As Boolean, wRen As Currency, _
       wSdoOld As Currency, wSdoGra As Currency, _
       wFracCargos As Currency, wFracAbonos As Currency, wFracSdoNew As Currency, wFecMax As Date, _
       wMesEnvio As String, wMesRecibe As String, _
       wImpEnvio As Currency, wImpRecibe As Currency
   
   lblMensaje.Caption = "Preparando Archivo....."
   lblMensaje.Refresh

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_ESTADO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   wMesTope = Right(txtTope.Text, 2)
   wAnoTope = Left(txtTope.Text, 4)
   wDiaTope = fundiames(wMesTope)
   wFecTope = Format(wDiaTope + "/" + wMesTope + "/" + wAnoTope, "dd/mm/yyyy")
   
   wFecMax = Format(fundiames(Format(Month(Date), "00")) + "/" + Format(Month(Date), "00") + "/" + Format(Year(Date), "0000"), "dd/mm/yyyy")
   
   wVip = ""

   aa = Leerado8("SELECT * FROM TMP_MASIVO WHERE USU = '" + wcodusu + "' ")
   If aa > 0 Then
      ADO8.MoveFirst
      wRegAct = 1
      wRegTot = aa
      Do While Not ADO8.EOF
         wSoc = ADO8!codsocio
         wCod = ADO8!codigo
         wIns = ADO8!ins
         wFracSw = False
         wRen = 0
         wFracCargos = 0: wFracAbonos = 0: wFracSdoNew = 0
   
         lblMensaje.Caption = "Socio " + Str(ADO8!codsocio) + " " + ADO8!nombre
         lblMensaje.Refresh
         
'         Db.BeginTrans
'         Db.Execute ("DELETE FROM TMP_FRACDET WHERE USU = '" + wcodusu + "'")
'         Db.CommitTrans
         
         Db.BeginTrans
         Db.Execute ("DELETE FROM TMP_ESTADO WHERE USU = '" + wcodusu + "' AND CODSOCIO = " + Str(wSoc) + " ")
         Db.CommitTrans

         aa = Leerado6a("SELECT SUM(CARGOS - ABONOS) AS DIFER " _
                    & " FROM CTASXDET " _
                    & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                    & "       CONCEPTO = '02' ")
         If aa > 0 Then
            wRen = IIf(IsNull(ADO6a!difer), 0, ADO6a!difer)
         End If
         Set ADO6a = Nothing
         
         wSdoOld = 0: wSdoGra = 0
         aa = Leerado7a("SELECT " _
                    & "  " + Str(wSoc) + ", '6', D.LINEA, " _
                    & "  D.NUMERO, D.LINEA, D.VCMTO, D.CARGOS, D.ABONOS, " _
                    & "  D.SDONEW, C.SDOPEN, D.NUMCOB, D.FECCOB, '" + wcodusu + "' " _
                    & " FROM FRACDET AS D INNER JOIN FRACCAB AS C " _
                    & "   ON D.NUMERO = C.NUMERO " _
                    & " WHERE C.CODSOCIO = " + Str(wSoc) + " " _
                    & " ORDER BY D.LINEA")
         If aa > 0 Then
            ADO7a.MoveFirst
            
            aa = Leerado6a("SELECT * FROM FRACCAB WHERE NUMERO = '" + ADO7a!numero + "'  ")
            If aa > 0 Then
               wSdoOld = ADO6a!sdopen
            End If
                    
            Do While Not ADO7a.EOF
               wSdoGra = wSdoOld - ADO7a!cargos
         
               Db.BeginTrans
               Db.Execute ("INSERT INTO TMP_ESTADO " _
               & " (CODSOCIO, TIPOREG, LINEA, " _
               & "  FRACNUMERO, FRACLINEA, FRACVCMTO, FRACCARGOS, " _
               & "  FRACABONOS, FRACSDONEW, FRACSDOOLD, FRACSDOGRA, FRACNUMCOB, FRACFECCOB, USU) " _
               & " VALUES " _
               & " (" + Str(wSoc) + ", '6', '" + ADO7a!linea + "', " _
               & "  '" + ADO7a!numero + "', '" + ADO7a!linea + "', " _
               & "  '" + Format(ADO7a!vcmto, "dd/mm/yyyy") + "', " _
               & "  " + Str(ADO7a!cargos) + ", " + Str(ADO7a!abonos) + ", " _
               & "  " + Str(ADO7a!sdonew) + ", " + Str(wSdoOld) + ", " _
               & "  " + Str(wSdoGra) + ", '" + ADO7a!numcob + "', " _
               & "  '" + Format(ADO7a!feccob, "dd/mm/yyyy") + "', '" + wcodusu + "' ) ")
               Db.CommitTrans
         
               wSdoOld = wSdoGra
               wFracCargos = wFracCargos + ADO7a!cargos
               wFracAbonos = wFracAbonos + ADO7a!abonos
               wFracSdoNew = wFracCargos - wFracAbonos
         
               ADO7a.MoveNext
            Loop
         End If
         
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_MASIVO " _
         & " SET RENOVA = " + Str(wRen) + "," _
         & "     FRACCARGOS = " + Str(wFracCargos) + ", " _
         & "     FRACABONOS = " + Str(wFracAbonos) + ", " _
         & "     FRACSDONEW = " + Str(wFracSdoNew) + " " _
         & " WHERE      USU = '" + wcodusu + "' AND " _
         & "       CODSOCIO = " + Str(wSoc) + " ")
         Db.CommitTrans
         
         aa = Leerado7("SELECT Z.* " _
                & " FROM ZZZ_MRECIBOS AS Z INNER JOIN ZZZ_CONCEPTO AS M " _
                & "   ON Z.CONCEPTO = M.CCONCE " _
                & " WHERE Z.CODIGO = " + Str(wCod) + " AND " _
                & "          Z.INS = " + Str(wIns) + " AND " _
                & "      (Z.MARCA2 <> 'A' OR Z.MARCA2 IS NULL) AND " _
                & "      (M.MARCA = 'S') " _
                & " ORDER BY Z.FECHA_PAGO, Z.SERIE, Z.NRO_COMP ")
        
'                & "      (Z.FECHA_PAGO <= '" + Format(wFecTope, "dd/mm/yyyy") + "')  " _

        If aa > 0 Then
            ADO7.MoveFirst
            wLin = 1
            Do While Not ADO7.EOF
               wRec = ADO7!serie + "-" + Format(ADO7!nro_comp, "000000")
               wMon = IIf(ADO7!moneda = "S/." Or ADO7!moneda = "S", "S", "D")
               wImp = ADO7!monto
               wFec = Format(ADO7!fecha_pago, "dd/mm/yyyy")
               wObs = Trim(IIf(IsNull(ADO7!obs), "", ADO7!obs))
   
                  Db.BeginTrans
                  Db.Execute ("INSERT INTO TMP_ESTADO " _
                  & " (CODSOCIO, LINEA, CODIGO, INS, TIPOREG, RECIBO, MONEDA, IMPORTE, FECHA, CONCEPTO, USU) " _
                  & " VALUES " _
                  & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', " + Str(wCod) + ", " + Str(wIns) + ", " _
                  & "  '2', '" + wRec + "', '" + wMon + "', " _
                  & "  " + Str(wImp) + ", '" + Format(wFec, "dd/mm/yyyy") + "', " _
                  & "  '" + GlosaLibre(wObs) + "', '" + wcodusu + "' ) ")
                  Db.CommitTrans
   
               wLin = wLin + 1
               ADO7.MoveNext
            Loop
         End If
         Set ADO7 = Nothing
   
         aa = Leerado7("SELECT * FROM ZZZ_BCORECAU " _
                & " WHERE CODIGO = " + Str(wCod) + " AND " _
                & "          INS = " + Str(wIns) + " AND " _
                & "        FECHA <= '" + Format(wFecMax, "dd/mm/yyyy") + "' " _
                & " ORDER BY FECHA, RECIBO ")
         If aa > 0 Then
            ADO7.MoveFirst
            wLin = 1
            Do While Not ADO7.EOF
               wRec = Format(ADO7!recibo, "000000")
               wMon = IIf(ADO7!moneda = "S/.", "S", "D")
               wImp = ADO7!aporte
               wFec = Format(ADO7!fecha, "dd/mm/yyyy")
               wnCr = ADO7!ncredito
               wNde = ADO7!ndebito
               wDeu = ADO7!deuda_pt2
               wCer = ADO7!dins_cer
               wAde = ADO7!adelanto
   
               Db.BeginTrans
               Db.Execute ("INSERT INTO TMP_ESTADO " _
               & " (CODSOCIO, LINEA, CODIGO, INS, TIPOREG, BCORECIBO, BCOMONEDA, BCONCREDITO, BCONDEBITO, " _
               & "  BCOAPORTE, BCOFECHA, USU) " _
               & " VALUES " _
               & " (" + Str(wSoc) + ", " + Format(wLin, "0000") + ", " + Str(wCod) + ", " + Str(wIns) + ", " _
               & "  '3', '" + wRec + "', '" + wMon + "', " _
               & "  " + Str(wnCr) + ", " + Str(wNde) + ", " + Str(wImp) + ", " _
               & "  '" + Format(wFec, "dd/mm/yyyy") + "', '" + wcodusu + "' ) ")
               Db.CommitTrans
               
               wLin = wLin + 1
               ADO7.MoveNext
            Loop
         End If
         Set ADO7 = Nothing
         
         aa = Leerado7("SELECT * FROM ZZZ_DEVOL " _
                & " WHERE CODIGO = " + Str(wCod) + " AND " _
                & "          INS = " + Str(wIns) + " AND " _
                & "        FECHA <= '" + Format(wFecMax, "dd/mm/yyyy") + "' " _
                & " ORDER BY FECHA, SERIE, NRO_COMP ")
        If aa > 0 Then
            ADO7.MoveFirst
            wLin = 1
            Do While Not ADO7.EOF
               wRec = ADO7!serie + "-" + Format(ADO7!nro_comp, "000000")
               wMon = "S"
               wImp = ADO7!importe
               wFec = Format(ADO7!fecha, "dd/mm/yyyy")
               wObs = Trim(IIf(IsNull(ADO7!glosa), "", ADO7!glosa))
   
                  Db.BeginTrans
                  Db.Execute ("INSERT INTO TMP_ESTADO " _
                  & " (CODSOCIO, LINEA, CODIGO, INS, TIPOREG, RECIBO, DEVMONEDA, DEVIMPORTE, FECHA, CONCEPTO, USU) " _
                  & " VALUES " _
                  & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', " + Str(wCod) + ", " + Str(wIns) + ", " _
                  & "  '4', '" + wRec + "', '" + wMon + "', " _
                  & "  " + Str(wImp) + ", '" + Format(wFec, "dd/mm/yyyy") + "', " _
                  & "  '" + GlosaLibre(wObs) + "', '" + wcodusu + "' ) ")
                  Db.CommitTrans
   
               wLin = wLin + 1
               ADO7.MoveNext
            Loop
         End If
         Set ADO7 = Nothing
         
         aa = Leerado7("SELECT * FROM ZZZ_APOR_PLA " _
                & " WHERE CODIGO = " + Str(wCod) + " AND " _
                & "          INS = " + Str(wIns) + " AND " _
                & "       CUOANO <= '" + wAnoTope + "'  " _
                & " ORDER BY CUOANO ")
         If aa > 0 Then
            ADO7.MoveFirst
            wLin = 1
            Do While Not ADO7.EOF
               
               If ADO7!cuoano < wAnoTope Then
               
                  Db.BeginTrans
                  Db.Execute ("INSERT INTO TMP_ESTADO " _
                  & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                  & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                  & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                  & "  TOTAL, DEUDA, USU ) " _
                  & " VALUES " _
                  & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                  & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                  & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                  & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", " + Str(ADO7!impo06) + ", " _
                  & "  " + Str(ADO7!impo07) + ", " + Str(ADO7!impo08) + ", " + Str(ADO7!impo09) + ", " _
                  & "  " + Str(ADO7!impo10) + ", " + Str(ADO7!impo11) + ", " + Str(ADO7!impo12) + ", " _
                  & "  " + Str(ADO7!totimpo) + ", " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                  Db.CommitTrans
               Else
                  If ADO7!cuoano = wAnoTope Then
                     Select Case Format(Month(wFecMax), "00")
                     Case "01"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  " + Str(ADO7!impo01) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "02"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", 0, " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "03"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "04"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03 + ADO7!impo04) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "05"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", 0, " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03 + ADO7!impo04 + ADO7!impo05) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "06"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", " + Str(ADO7!impo06) + ", " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03 + ADO7!impo04 + ADO7!impo05 + ADO7!impo06) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "07"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", " + Str(ADO7!impo06) + ", " _
                          & "  " + Str(ADO7!impo07) + ", 0, 0, " _
                          & "  0,0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03 + ADO7!impo04 + ADO7!impo05 + ADO7!impo06 + ADO7!impo07) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "08"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", " + Str(ADO7!impo06) + ", " _
                          & "  " + Str(ADO7!impo07) + ", " + Str(ADO7!impo08) + ", 0, " _
                          & "  0,0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03 + ADO7!impo04 + ADO7!impo05 + ADO7!impo06 + ADO7!impo07 + ADO7!impo08) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "09"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", " + Str(ADO7!impo06) + ", " _
                          & "  " + Str(ADO7!impo07) + ", " + Str(ADO7!impo08) + ", " + Str(ADO7!impo09) + ", " _
                          & "  0,0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03 + ADO7!impo04 + ADO7!impo05 + ADO7!impo06 + ADO7!impo07 + ADO7!impo08 + ADO7!impo09) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "10"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", " + Str(ADO7!impo06) + ", " _
                          & "  " + Str(ADO7!impo07) + ", " + Str(ADO7!impo08) + ", " + Str(ADO7!impo09) + ", " _
                          & "  " + Str(ADO7!impo10) + ",0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03 + ADO7!impo04 + ADO7!impo05 + ADO7!impo06 + ADO7!impo07 + ADO7!impo08 + ADO7!impo09 + ADO7!impo10) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "11"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", " + Str(ADO7!impo06) + ", " _
                          & "  " + Str(ADO7!impo07) + ", " + Str(ADO7!impo08) + ", " + Str(ADO7!impo09) + ", " _
                          & "  " + Str(ADO7!impo10) + ", " + Str(ADO7!impo11) + ", 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03 + ADO7!impo04 + ADO7!impo05 + ADO7!impo06 + ADO7!impo07 + ADO7!impo08 + ADO7!impo09 + ADO7!impo10 + ADO7!impo11) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "12"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", " + Str(ADO7!impo06) + ", " _
                          & "  " + Str(ADO7!impo07) + ", " + Str(ADO7!impo08) + ", " + Str(ADO7!impo09) + ", " _
                          & "  " + Str(ADO7!impo10) + ", " + Str(ADO7!impo11) + ", " + Str(ADO7!impo12) + ", " _
                          & "  " + Str(ADO7!totimpo) + ", " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     End Select
                  
                  End If
               End If
               
               wLin = wLin + 1
               ADO7.MoveNext
            Loop
         End If
         Set ADO7 = Nothing
   
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_ESTADO " _
         & " SET TOTAL = IMP01 + IMP02 + IMP03 + IMP04 + IMP05 + IMP06 + " _
         & "             IMP07 + IMP08 + IMP09 + IMP10 + IMP11 + IMP12 " _
         & " WHERE USU = '" + wcodusu + "' ")
         Db.CommitTrans
   
         ADO8.MoveNext
      Loop
   End If

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ESTADO " _
   & " SET NOMCOB = 'DIECO 1' " _
   & " WHERE (TIPCOB = '1') AND " _
   & "       (USU = '" + wcodusu + "') ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ESTADO " _
   & " SET NOMCOB = 'DIECO 2' " _
   & " WHERE (TIPCOB = '2') AND " _
   & "       (USU = '" + wcodusu + "') ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ESTADO " _
   & " SET NOMCOB = 'CAJA MP' " _
   & " WHERE (TIPCOB = '4') AND " _
   & "       (USU = '" + wcodusu + "') ")
   Db.CommitTrans
   
   aa = Leerado8("SELECT * FROM TMP_MASIVO WHERE USU = '" + wcodusu + "' ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wSoc = ADO8!codsocio
         wCod = ADO8!codigo
         wIns = ADO8!ins

         aa = Leerado7("SELECT * FROM TMP_ESTADO WHERE USU = '" + wcodusu + "' AND CODSOCIO = " + Str(wSoc) + " ")
         If aa = 0 Then
            Db.BeginTrans
            Db.Execute ("INSERT INTO TMP_ESTADO " _
            & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
            & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
            & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
            & "  TOTAL, DEUDA, USU ) " _
            & " VALUES " _
            & " (" + Str(wSoc) + ", '0001', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
            & "  '',  '', '', " _
            & "  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '" + wcodusu + "') ")
            Db.CommitTrans
         End If

         ADO8.MoveNext
      Loop
   End If

   wVip = "": wCartaDieco = ""
   aa = Leerado8a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
   If aa > 0 Then
      wVip = IIf(ADO8a!vip = True, "SOCIO VIP", "")
      wCartaDieco = IIf(ADO8a!cartadieco = True, "ASOCIADO SIN CARTA AUTORIZACION DIECO", "")
   End If

   wMesEnvio = BuscaUltimoDiecoCajMP(wSoc, 1)
   wImpEnvio = Val(BuscaUltimoDiecoCajMP(wSoc, 2))
   wMesRecibe = BuscaUltimoDiecoCajMP(wSoc, 3)
   wImpRecibe = Val(BuscaUltimoDiecoCajMP(wSoc, 4))
   
   lblMensaje.Caption = ""
   lblMensaje.Refresh

   Crys2.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys2.ReportFileName = xraiz + "ReportCtaCte\EstadoCtaMasivo.RPT"
   Crys2.Formulas(0) = "SOCIOVIP= '" + wVip + "' "
   Crys2.Formulas(1) = "CARTADIECO= '" + wCartaDieco + "' "
   Crys2.Formulas(2) = "MESCIERRE= '" + wAnoTope + "-" + wMesTope + "' "
   Crys2.Formulas(3) = "MESENVIO= '" + wMesEnvio + "' "
   Crys2.Formulas(4) = "MONTOENVIO= '" + Format(wImpEnvio, "###0.00") + "' "
   Crys2.Formulas(5) = "MESRECIBE= '" + wMesRecibe + "' "
   Crys2.Formulas(6) = "MONTORECIBE= '" + Format(wImpRecibe, "###0.00") + "' "
   Crys2.SelectionFormula = " {TMP_ESTADO.USU}='" + wcodusu + "' "
   Crys2.WindowState = crptMaximized
   Crys2.Action = 1
End Sub

Private Sub cmdOtro_Click()
   Limpiar

   optCodofin.Value = True
   txtCodigo.Enabled = True
   txtNumdoc.Enabled = False

   txtCodigo.SetFocus
End Sub

Private Sub cmdSalir_Click()
   Db.Close
   End
End Sub

Private Sub Form_Activate()
   frmEleConsulta.Left = (Screen.Width - Width) \ 2
   frmEleConsulta.Top = 0
   
   Dim a As Integer, wcia As String
   
   cmbCia.Clear
   a = LeeradoMaster3("SELECT * FROM COMPANIAS ORDER BY CODIGOCIA ")
   If a > 0 Then
      ADOMaster3.MoveFirst
      Do While Not ADOMaster3.EOF
         cmbCia.AddItem ADOMaster3!codigocia + " " + Trim(ADOMaster3!NombreCia)
         
         ADOMaster3.MoveNext
      Loop
   End If
   Set ADOMaster3 = Nothing
   
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
   
   a = Leerado8("SELECT * FROM MAESITU ORDER BY SITU ")
   If a > 0 Then
      ADO8.MoveFirst
      I = 0
      Do While Not ADO8.EOF
         cmbSitu.AddItem ADO8!nombre
         I = I + 1
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   
   a = Leerado8("SELECT * FROM MAESITUESP ORDER BY SITUESP ")
   If a > 0 Then
      ADO8.MoveFirst
      I = 0
      Do While Not ADO8.EOF
         cmbSituEsp.AddItem ADO8!nombre
         I = I + 1
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   
   a = Leerado8("SELECT * FROM MAEECIVIL ORDER BY ECIVIL ")
   If a > 0 Then
      ADO8.MoveFirst
      I = 0
      Do While Not ADO8.EOF
         cmbECivil.AddItem Trim(ADO8!nombre)
         I = I + 1
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   
   a = Leerado8("SELECT * FROM MAESEXO ORDER BY SEXO ")
   If a > 0 Then
      ADO8.MoveFirst
      I = 0
      Do While Not ADO8.EOF
         cmbSexo.AddItem ADO8!nombre
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
   
   cmbCia.ListIndex = 0
   cmbCia.Enabled = False
   
   
   Dim zDiaFin As Date, zDiaHoy As Date, _
       wmescia As String, wanocia As String
   
   wmescia = Format(Month(Date), "00")
   wanocia = Format(Year(Date), "0000")
   
   zDiaFin = fundiames(wmescia) + "/" + wmescia + "/" + wanocia
   zDiaHoy = Format(Date, "dd/mm/yyyy")
   
   
   If Format(zDiaHoy, "dd/mm/yyyy") < Format(zDiaFin, "dd/mm/yyyy") Then
      If wmescia > "01" Then
         zMesTope = wanocia + Format(Val(wmescia) - 1, "00")
      Else
         zMesTope = Format(Val(wanocia) - 1, "0000") + "12"
      End If
   Else
      zMesTope = wanocia + wmescia
   End If
   
'   wanocia = Format(Year(Date), "0000")
'   wmescia = Format(Month(Date), "00")
   
'   zMesTope = wanocia + wmescia
   
   txtTope.Text = Left(zMesTope, 4) + "/" + Right(zMesTope, 2)
   txtTope.Enabled = False
   
   txtCodigo.Enabled = True
   txtNumdoc.Enabled = False
   txtCodigo.SetFocus
End Sub

Private Sub optCodofin_Click()
   Limpiar
   
   txtNumdoc.Enabled = False
   txtCodigo.Enabled = True
   txtCodigo.SetFocus
End Sub

Private Sub optDNI_Click()
   Limpiar
   
   txtNumdoc.Enabled = True
   txtCodigo.Enabled = False
   txtNumdoc.SetFocus
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
   Dim aa As Integer
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
      
      llenadet
      LlenaCab
      LlenaPagos
      
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(Val(txtCodigo.Text)) + " ")
      If aa > 0 Then
         txtSaldo.Text = Format(ADO8!deuda_pt2, "#####0.00;;\ ")
         txtAdelan.Text = Format(ADO8!adelanto, "#####0.00;;\ ")
         txtEnv540.Text = Format(ADO8!env_540, "#####0.00;;\ ")
         txtEnv541.Text = Format(ADO8!env_541, "#####0.00;;\ ")
         Select Case ADO8!tipcob
         Case "01"
              cmdDIECO.Enabled = True
              cmdCAJMP.Enabled = False
         Case "02"
              cmdDIECO.Enabled = False
              cmdCAJMP.Enabled = True
         Case "03"
              cmdDIECO.Enabled = False
              cmdCAJMP.Enabled = False
         End Select
      End If
      Set ADO8 = Nothing
      
      aa = Leerado8("SELECT * FROM TMP_PADRON2018 WHERE CODIGO = " + Str(Val(txtCodigo.Text)) + " ")
      If aa > 0 Then
         lblMensaje.Caption = "ASOCIADO EN PADRON DE VOTANTES"
         lblMensaje.ForeColor = RGB(255, 0, 0)
      Else
         lblMensaje.Caption = "NO ESTA EN PADRON DE VOTANTES"
         lblMensaje.ForeColor = RGB(0, 0, 255)
      End If
      Set ADO8 = Nothing
      
      cmdImprimir.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub Limpiar()
   txtCodSocio.Text = ""
   lblCodSocio.Caption = ""
   txtCodigo.Text = ""
   txtIns.Text = ""
   txtNumdoc.Text = ""
   txtCarnetPNP.Text = ""
   txtCarnetPIP.Text = ""
   txtFecNac.Text = "__/__/____"
   txtFecMat.Text = "__/__/____"
   txtFecIng.Text = "__/__/____"
   txtFecRenu.Text = "__/__/____"
   txtFecReso.Text = "__/__/____"
   txtFecExpul.Text = "__/__/____"
   txtFecExclu.Text = "__/__/____"
   txtFecRein.Text = "__/__/____"

   lblE_Socio.Caption = ""

   cmbGrado.ListIndex = 0
   cmbSitu.ListIndex = 0
   cmbSituEsp.ListIndex = 0
   cmbECivil.ListIndex = 0
   cmbSexo.ListIndex = 0
   cmbE_Socio.ListIndex = 0
   cmbTipCob.ListIndex = 0

   chkVip.Value = vbUnchecked
   txtFecVip.Text = "__/__/____"
   lblMensaje.Caption = ""
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_FAMILIA WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Set DataGrid1.DataSource = Nothing
End Sub

Private Sub llenadet()
   Dim aa As Integer, wCod As Long, wSoc As Integer, wDni As String, wNomE_S As String
   wCod = Val(txtCodigo.Text)
   wDni = txtNumdoc.Text
   wSoc = 0

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_FAMILIA WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   If optCodofin.Value = True Then
      aa = Leerado7a("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(wCod) + " ")
   Else
      aa = Leerado7a("SELECT * FROM MAESOCIO WHERE NUMDOC = '" + wDni + "' ")
   End If
   If aa > 0 Then
      wSoc = ADO7a!codsocio
      txtCodSocio.Text = ADO7a!codsocio
      txtIns.Text = ADO7a!ins
      txtNumdoc.Text = ADO7a!numdoc
      lblCodSocio.Caption = ADO7a!nombre
      txtCodigo.Text = ADO7a!codigo
      txtCarnetPNP.Text = ADO7a!carnetpnp
      txtCarnetPIP.Text = ADO7a!carnetpip
      txtNumReso.Text = IIf(IsNull(ADO7a!nreso_ing), "", ADO7a!nreso_ing)
      txtObservac.Text = IIf(IsNull(ADO7a!observac), "", ADO7a!observac)
      txtObserva2.Text = IIf(IsNull(ADO7a!observa2), "", ADO7a!observa2)
      If IsDate(ADO7a!fecnac) Then
         txtFecNac.Text = Format(ADO7a!fecnac, "dd/mm/yyyy")
      Else
         txtFecNac.Text = "__/__/____"
      End If
      If IsDate(ADO7a!fecing) Then
         txtFecIng.Text = Format(ADO7a!fecing, "dd/mm/yyyy")
      Else
         txtFecIng.Text = "__/__/____"
      End If
      If IsDate(ADO7a!fecmat) Then
         txtFecMat.Text = Format(ADO7a!fecmat, "dd/mm/yyyy")
      Else
         txtFecMat.Text = "__/__/____"
      End If
      If IsDate(ADO7a!fecrenu) Then
         txtFecRenu.Text = Format(ADO7a!fecrenu, "dd/mm/yyyy")
      Else
         txtFecRenu.Text = "__/__/____"
      End If
      If IsDate(ADO7a!freso_ing) Then
         txtFecReso.Text = Format(ADO7a!freso_ing, "dd/mm/yyyy")
      Else
         txtFecReso.Text = "__/__/____"
      End If
      If IsDate(ADO7a!fecexpul) Then
         txtFecExpul.Text = Format(ADO7a!fecexpul, "dd/mm/yyyy")
      Else
         txtFecExpul.Text = "__/__/____"
      End If
      If IsDate(ADO7a!fecexclu) Then
         txtFecExclu.Text = Format(ADO7a!fecexclu, "dd/mm/yyyy")
      Else
         txtFecExclu.Text = "__/__/____"
      End If
      If IsDate(ADO7a!fecrein) Then
         txtFecRein.Text = Format(ADO7a!fecrein, "dd/mm/yyyy")
      Else
         txtFecRein.Text = "__/__/____"
      End If
      If IsDate(ADO7a!fecvip) Then
         txtFecVip.Text = Format(ADO7a!fecvip, "dd/mm/yyyy")
      Else
         txtFecVip.Text = "__/__/____"
      End If
   
      cmbGrado.ListIndex = BuscaGrado(ADO7a!grado)
      cmbSitu.ListIndex = BuscaSitu(ADO7a!situ)
      cmbSituEsp.ListIndex = BuscaSituEsp(ADO7a!situesp)
      cmbECivil.ListIndex = BuscaECivil(ADO7a!ecivil)
      cmbSexo.ListIndex = BuscaSexo(ADO7a!sexo)
      cmbE_Socio.ListIndex = BuscaEsocio(ADO7a!e_socio)
      cmbTipCob.ListIndex = BuscaTipCob(ADO7a!tipcob)
   
      lblE_Socio.Caption = BuscaEsocio(ADO7a!e_socio)
   
      wNomE_S = ""
      aa = Leerado8("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + ADO7a!e_socio + "' ")
      If aa > 0 Then
         wNomE_S = ADO8!nombre
      End If
      Set ADO8 = Nothing
      
      lblE_Socio.Caption = wNomE_S
      
      If ADO7a!vip = True Then
         chkVip.Value = vbChecked
      Else
         chkVip.Value = vbUnchecked
      End If
   
      Db.BeginTrans
      Db.Execute ("INSERT INTO TMP_FAMILIA " _
      & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMTIPOPARIENTE, NOMBRE, FECNAC, USU) " _
      & " SELECT " _
      & "  F.CODSOCIO, F.TIPOPARIENTE, F.LIN, F.NUMDOC, " _
      & "  T.NOMBRE, F.NOMBRE, F.FECNAC, '" + wcodusu + "' " _
      & " FROM MAEFAMILIA AS F LEFT JOIN MAETIPOPARIENTE AS T " _
      & "   ON F.TIPOPARIENTE = T.TIPOPARIENTE " _
      & " WHERE F.CODSOCIO = " + Str(wSoc) + "  ")
      Db.CommitTrans
   
   End If

   aa = Leerado2("SELECT NOMTIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC " _
                 & " FROM TMP_FAMILIA " _
                 & " WHERE USU = '" + wcodusu + "' AND " _
                 & "       CODSOCIO = " + Str(wSoc) + " ")
   Set DataGrid1.DataSource = ADO2

   DataGrid1.Columns(0).Width = 1500  ' NOMTIPOPARIENTE
   DataGrid1.Columns(0).Alignment = dbgLeft
   DataGrid1.Columns(0).Caption = "TIPO"
       
   DataGrid1.Columns(1).Width = 500    ' LIN
   DataGrid1.Columns(1).Alignment = dbgCenter
   DataGrid1.Columns(1).Caption = "LIN"
       
   DataGrid1.Columns(2).Width = 950   ' NUMDOC
   DataGrid1.Columns(2).Alignment = dbgCenter
   DataGrid1.Columns(2).Caption = "D.N.I."
       
   DataGrid1.Columns(3).Width = 4800  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE"
    
   DataGrid1.Columns(4).Width = 1150  ' FECNAC
   DataGrid1.Columns(4).Alignment = dbgCenter
   DataGrid1.Columns(4).Caption = "FEC.NAC."
   DataGrid1.Columns(4).NumberFormat = "dd/mm/yyyy"
   
End Sub

Private Sub LlenaCab()
   Dim aa As Integer, wRegAct As Integer, wRegTot As Integer, wSoc As Integer, wCod As Long, _
       wApo As Currency, wcob As Currency, wDif As Currency, wMesUno As String, wMesDos As String, _
       wSdo As Currency, wAde As Currency, wEnv540 As Currency, wEnv541 As Currency, _
       wMesTope As String
   wCod = Val(txtCodigo.Text)
   wMesTope = Left(txtTope.Text, 4) + Right(txtTope.Text, 2)
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_MASIVO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_MASIVO " _
   & " (CODSOCIO, CODIGO, INS, NUMDOC, NOMBRE, E_SOCIO, DEUDA_PT2, ADELANTO, " _
   & "  FECING, FECREIN, FECBAJ, FECRENO, FECRENU, FECEXCLU, FECEXPUL, " _
   & "  TOTAPO, TOTCOB, DESDE, HASTA, USU) " _
   & " SELECT " _
   & "  CODSOCIO, CODIGO, INS, NUMDOC, NOMBRE, E_SOCIO, DEUDA_PT2, ADELANTO, " _
   & "  FECING, FECREIN, FECBAJ, FECRENO, FECRENU, FECEXCLU, FECEXPUL, " _
   & "  0, 0, '', '', '" + wcodusu + "' " _
   & " FROM MAESOCIO " _
   & " WHERE CODIGO = " + Str(wCod) + " ")
   Db.CommitTrans

   aa = Leerado2("SELECT * " _
                & " FROM TMP_MASIVO " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY E_SOCIO, NOMBRE ")
   If aa > 0 Then
      ADO2.MoveFirst
      wRegAct = 1
      wRegTot = aa
      Do While Not ADO2.EOF
         DoEvents
         lblMensaje.Caption = Trim(Format(wRegAct, "####0")) + " / " + _
                              Trim(Format(wRegTot, "####0")) + _
                              " Socio " + Trim(ADO2!nombre)
         lblMensaje.Refresh
         
         wCod = ADO2!codigo
         wSoc = ADO2!codsocio
         wApo = 0: wSdo = 0: wAde = 0
         wMesUno = ""
         wMesDos = ""
         wcob = 0
         wDif = wApo - wcob
         
         wSdo = SaldoFoto(wSoc, wMesTope)
         If wSdo < 0 Then
            wAde = -wSdo
            wSdo = 0
         End If
         
         wEnv540 = EnvioDiecoCMP(wSoc, wMesTope, 1)
         wEnv541 = EnvioDiecoCMP(wSoc, wMesTope, 2)
         
         Db.BeginTrans
         Db.Execute ("UPDATE MAESOCIO " _
         & " SET DEUDA_PT2 = " + Str(wSdo) + ", " _
         & "      ADELANTO = " + Str(wAde) + ", ENV_540 = " + Str(wEnv540) + ", ENV_541 = " + Str(wEnv541) + " " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " ")
         Db.CommitTrans
   
         Db.BeginTrans
         Db.Execute ("UPDATE ZZZ_MAESTRO " _
         & " SET DEUDA_PT2 = " + Str(wSdo) + ", " _
         & "      ADELANTO = " + Str(wAde) + ", ENV_540 = " + Str(wEnv540) + ", ENV_541 = " + Str(wEnv541) + " " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " ")
         Db.CommitTrans
   
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_MASIVO " _
         & " SET TOTAPO = " + Str(wApo) + ", " _
         & "     TOTCOB = " + Str(wcob) + ", " _
         & "      DIFER = " + Str(wDif) + ", " _
         & "      DESDE = '" + wMesUno + "', " _
         & "      HASTA = '" + wMesDos + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "            USU = '" + wcodusu + "' ")
         Db.CommitTrans
   
         wRegAct = wRegAct + 1
         ADO2.MoveNext
      Loop
   End If

   DoEvents
   lblMensaje.Caption = ""
   lblMensaje.Refresh
End Sub

Private Sub txtNumdoc_GotFocus()
   txtNumdoc.SelStart = 0
   txtNumdoc.SelLength = Len(Trim(txtNumdoc.Text))
End Sub

Private Sub txtNumdoc_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtNumdoc.Text)) = 0 Then
         MsgBox "DNI En Blanco", vbExclamation
         txtNumdoc.Text = ""
         Exit Sub
      End If
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE NUMDOC = '" + txtNumdoc.Text + "' ")
      If aa = 0 Then
         MsgBox "DNI Digitado No Existe", vbExclamation
         txtNumdoc.Text = ""
         Exit Sub
      End If
         
      llenadet
      LlenaCab
      LlenaPagos
      
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(Val(txtCodigo.Text)) + " ")
      If aa > 0 Then
         txtSaldo.Text = Format(ADO8!deuda_pt2, "#####0.00;;\ ")
         txtAdelan.Text = Format(ADO8!adelanto, "#####0.00;;\ ")
         txtEnv540.Text = Format(ADO8!env_540, "#####0.00;;\ ")
         txtEnv541.Text = Format(ADO8!env_541, "#####0.00;;\ ")
         If ADO8!cartadieco = True Then
            lblCartaDieco.Caption = "Asociado Sin Carta Autorizacion DIECO"
         Else
            lblCartaDieco.Caption = ""
         End If
      End If
      
      cmdOtro.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
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

Private Sub txtTope_GotFocus()
   txtTope.SelStart = 0
   txtTope.SelLength = 10
End Sub

Private Sub txtTope_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 40
        txtCodigo.SetFocus
   End Select
End Sub

Private Sub txtTope_KeyPress(KeyAscii As Integer)
   Dim wMes As String, wAno As String
   If KeyAscii = 13 Then
      If txtTope.Text = "____/__" Then
         MsgBox "Mes Tope En Blanco", vbExclamation
         txtTope.Text = "____/__"
         Exit Sub
      End If
      wAno = Left(txtTope.Text, 4)
      wMes = Right(txtTope.Text, 2)
      If wMes <> "01" And wMes <> "02" And wMes <> "03" And wMes <> "04" And _
         wMes <> "05" And wMes <> "06" And wMes <> "07" And wMes <> "08" And _
         wMes <> "09" And wMes <> "10" And wMes <> "11" And wMes <> "12" Then
         MsgBox "Mes Digitado Es Errado", vbExclamation
         txtTope.Text = "____/__"
         Exit Sub
      End If
      If wAno < "2017" And wAno > "2030" Then
         MsgBox "Año Digitado Es Errado", vbExclamation
         txtTope.Text = "____/__"
         Exit Sub
      End If
      txtCodigo.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii), 1) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub LlenaPagos()
   Dim wSoc As Integer, wCod As Long, wIns As Integer, _
       wPag1 As String, wPag2 As String
   wSoc = Val(txtCodSocio.Text)
   wCod = Val(txtCodigo.Text)
   wIns = Val(txtIns.Text)
   
   txtPago1.Text = ""
   txtPago2.Text = ""
   
   wPag1 = "": wPag2 = ""
   aa = Leerado8("SELECT * FROM DIECOCAB " _
                & " WHERE CODSOCIO = " + Str(wSoc) + " OR " _
                & "       CODASIG1 = " + Str(wSoc) + " OR " _
                & "       CODASIG2 = " + Str(wSoc) + " OR " _
                & "       CODASIG3 = " + Str(wSoc) + " OR " _
                & "       CODASIG4 = " + Str(wSoc) + " OR " _
                & "       CODASIG5 = " + Str(wSoc) + " " _
                & " ORDER BY MES DESC ")
   If aa > 0 Then
      Select Case True
      Case ADO8!codsocio = wSoc
           If ADO8!dscsocio > 0 Then
              txtPago1.Text = "DIECO MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscsocio, "###00.00")
           End If
           ADO8.MoveNext
           If Not ADO8.EOF Then
              If ADO8!dscsocio > 0 Then
                 txtPago2.Text = "DIECO MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscsocio, "###00.00")
              End If
           End If
      Case ADO8!codasig1 = wSoc
           If ADO8!dscasig1 > 0 Then
              txtPago1.Text = "DIECO MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig1, "###00.00")
           End If
           ADO8.MoveNext
           If Not ADO8.EOF Then
              If ADO8!dscasig1 > 0 Then
                 txtPago2.Text = "DIECO MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig1, "###00.00")
              End If
           End If
      Case ADO8!codasig2 = wSoc
           If ADO8!dscasig2 > 0 Then
              txtPago1.Text = "DIECO MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig2, "###00.00")
           End If
           ADO8.MoveNext
           If Not ADO8.EOF Then
              If ADO8!dscasig2 > 0 Then
                 txtPago2.Text = "DIECO MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig2, "###00.00")
              End If
           End If
      Case ADO8!codasig3 = wSoc
           If ADO8!dscasig3 > 0 Then
              txtPago1.Text = "DIECO MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig3, "###00.00")
           End If
           ADO8.MoveNext
           If Not ADO8.EOF Then
              If ADO8!dscasig3 > 0 Then
                 txtPago2.Text = "DIECO MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig3, "###00.00")
              End If
           End If
      Case ADO8!codasig4 = wSoc
           If ADO8!dscasig4 > 0 Then
              txtPago1.Text = "DIECO MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig4, "###00.00")
           End If
           ADO8.MoveNext
           If Not ADO8.EOF Then
              If ADO8!dscasig4 > 0 Then
                 txtPago2.Text = "DIECO MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig4, "###00.00")
              End If
           End If
      Case ADO8!codasig5 = wSoc
           If ADO8!dscasig5 > 0 Then
              txtPago1.Text = "DIECO MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig5, "###00.00")
           End If
           ADO8.MoveNext
           If Not ADO8.EOF Then
              If ADO8!dscasig5 > 0 Then
                 txtPago2.Text = "DIECO MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig5, "###00.00")
              End If
           End If
      End Select
   End If
   Set ADO8 = Nothing
    
   aa = Leerado8("SELECT * FROM CAJMPCAB " _
                & " WHERE CODSOCIO = " + Str(wSoc) + " OR " _
                & "       CODASIG1 = " + Str(wSoc) + " OR " _
                & "       CODASIG2 = " + Str(wSoc) + " OR " _
                & "       CODASIG3 = " + Str(wSoc) + " OR " _
                & "       CODASIG4 = " + Str(wSoc) + " OR " _
                & "       CODASIG5 = " + Str(wSoc) + " " _
                & " ORDER BY MES DESC ")
   If aa > 0 Then
      Select Case True
      Case ADO8!codsocio = wSoc
           If ADO8!dscsocio > 0 Then
              txtPago1.Text = "CAJA MP MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscsocio, "###00.00")
           End If
           ADO8.MoveNext
           If Not ADO8.EOF Then
              If ADO8!dscsocio > 0 Then
                 txtPago2.Text = "CAJA MP MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscsocio, "###00.00")
              End If
           End If
      Case ADO8!codasig1 = wSoc
           If ADO8!dscasig1 > 0 Then
              txtPago1.Text = "CAJA MP MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig1, "###00.00")
           End If
           ADO8.MoveNext
           If Not ADO8.EOF Then
              If ADO8!dscasig1 > 0 Then
                 txtPago2.Text = "CAJA MP MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig1, "###00.00")
              End If
           End If
      Case ADO8!codasig2 = wSoc
           If ADO8!dscasig2 > 0 Then
              txtPago1.Text = "CAJA MP MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig2, "###00.00")
           End If
           ADO8.MoveNext
           If Not ADO8.EOF Then
              If ADO8!dscasig2 > 0 Then
                 txtPago2.Text = "CAJA MP MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig2, "###00.00")
              End If
           End If
      Case ADO8!codasig3 = wSoc
           If ADO8!dscasig3 > 0 Then
              txtPago1.Text = "CAJA MP MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig3, "###00.00")
           End If
           ADO8.MoveNext
           If Not ADO8.EOF Then
              If ADO8!dscasig3 > 0 Then
                 txtPago2.Text = "CAJA MP MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig3, "###00.00")
              End If
           End If
      Case ADO8!codasig4 = wSoc
           If ADO8!dscasig4 > 0 Then
              txtPago1.Text = "CAJA MP MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig4, "###00.00")
           End If
           ADO8.MoveNext
           If Not ADO8.EOF Then
              If ADO8!dscasig4 > 0 Then
                 txtPago2.Text = "CAJA MP MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig4, "###00.00")
              End If
           End If
      Case ADO8!codasig5 = wSoc
           If ADO8!dscasig5 > 0 Then
              txtPago1.Text = "CAJA MP MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig5, "###00.00")
           End If
           ADO8.MoveNext
           If Not ADO8.EOF Then
              If ADO8!dscasig5 > 0 Then
                 txtPago2.Text = "CAJA MP MES " + Format(ADO8!mes, "0000-00") + "  S/." + Format(ADO8!dscasig5, "###00.00")
              End If
           End If
      End Select
   End If
   Set ADO8 = Nothing
   
   aa = Leerado8("SELECT Z.* " _
                & " FROM ZZZ_MRECIBOS AS Z INNER JOIN ZZZ_CONCEPTO AS M " _
                & "   ON Z.CONCEPTO = M.CCONCE " _
                & " WHERE Z.CODIGO = " + Str(wCod) + " AND " _
                & "          Z.INS = " + Str(wIns) + " AND " _
                & "      (Z.MARCA2 <> 'A' OR Z.MARCA2 IS NULL) AND " _
                & "      (M.MARCA = 'S') " _
                & " ORDER BY FECHA_PAGO DESC")
   If aa > 0 Then
      If ADO8!monto > 0 Then
         txtPago1.Text = Trim(txtPago1.Text) + _
              " TESORERIA " + Format(ADO8!fecha_pago, "dd/mm/yyyy") + "  " + _
              IIf(ADO8!moneda = "S/.", "S/.", "US$") + _
              Format(ADO8!monto, "###0.00") + " " + Trim(ADO8!obs)
      End If
      ADO8.MoveNext
      If Not ADO8.EOF Then
         If ADO8!monto > 0 Then
            txtPago2.Text = Trim(txtPago2.Text) + _
                 " TESORERIA " + Format(ADO8!fecha_pago, "dd/mm/yyyy") + "  " + _
                 IIf(ADO8!moneda = "S/.", "S/.", "US$") + _
                 Format(ADO8!monto, "###0.00") + " " + Trim(ADO8!obs)
         End If
      End If
   End If
   Set ADO8 = Nothing

End Sub

