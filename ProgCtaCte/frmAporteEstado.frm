VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAporteEstado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estado de Pagos de Asociados"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   12765
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6840
      MaxLength       =   8
      TabIndex        =   80
      Top             =   1725
      Width           =   975
   End
   Begin VB.OptionButton optDni 
      Caption         =   "Consulta x DNI"
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
      Height          =   375
      Left            =   9120
      TabIndex        =   79
      Top             =   840
      Width           =   2175
   End
   Begin VB.OptionButton optCodofin 
      Caption         =   "Consulta x Codofin"
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
      Height          =   375
      Left            =   9120
      TabIndex        =   78
      Top             =   480
      Value           =   -1  'True
      Width           =   2535
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
      Left            =   480
      TabIndex        =   74
      Top             =   4680
      Width           =   1455
   End
   Begin VB.ComboBox cmbTipCob 
      Height          =   315
      ItemData        =   "frmAporteEstado.frx":0000
      Left            =   3360
      List            =   "frmAporteEstado.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   68
      Top             =   4155
      Width           =   2655
   End
   Begin VB.ComboBox cmbGrado 
      Height          =   315
      ItemData        =   "frmAporteEstado.frx":0004
      Left            =   9240
      List            =   "frmAporteEstado.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   66
      Top             =   1725
      Width           =   2895
   End
   Begin VB.TextBox txtTomo 
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
      Left            =   2400
      MaxLength       =   8
      TabIndex        =   51
      Top             =   4155
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
      Height          =   285
      Left            =   7200
      MaxLength       =   10
      TabIndex        =   50
      Top             =   4155
      Width           =   1410
   End
   Begin VB.TextBox txtDirec 
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   39
      Top             =   2685
      Width           =   6015
   End
   Begin VB.TextBox txtUbi1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6120
      MaxLength       =   2
      TabIndex        =   38
      Text            =   " "
      Top             =   2685
      Width           =   375
   End
   Begin VB.TextBox txtTelefono 
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
      Left            =   120
      MaxLength       =   20
      TabIndex        =   37
      Top             =   3165
      Width           =   2610
   End
   Begin VB.TextBox txtCelular 
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
      Left            =   5400
      MaxLength       =   10
      TabIndex        =   36
      Top             =   3165
      Width           =   1410
   End
   Begin VB.TextBox txteMail 
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   35
      Top             =   3660
      Width           =   3975
   End
   Begin VB.TextBox txtRefer 
      Height          =   285
      Left            =   6840
      MaxLength       =   50
      TabIndex        =   34
      Top             =   3165
      Width           =   5295
   End
   Begin VB.ComboBox cmbE_Socio 
      Height          =   315
      ItemData        =   "frmAporteEstado.frx":0008
      Left            =   8040
      List            =   "frmAporteEstado.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   3660
      Width           =   3255
   End
   Begin VB.TextBox txtEMail2 
      Height          =   285
      Left            =   4080
      MaxLength       =   50
      TabIndex        =   32
      Top             =   3660
      Width           =   3975
   End
   Begin VB.TextBox txtTelefon2 
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
      Left            =   2760
      MaxLength       =   20
      TabIndex        =   31
      Top             =   3165
      Width           =   2610
   End
   Begin VB.TextBox txtUbi2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6480
      MaxLength       =   2
      TabIndex        =   30
      Text            =   " "
      Top             =   2685
      Width           =   375
   End
   Begin VB.TextBox txtUbi3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6840
      MaxLength       =   2
      TabIndex        =   29
      Text            =   " "
      Top             =   2685
      Width           =   375
   End
   Begin VB.ComboBox cmbECivil 
      Height          =   315
      ItemData        =   "frmAporteEstado.frx":000C
      Left            =   8280
      List            =   "frmAporteEstado.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   2205
      Width           =   2055
   End
   Begin VB.ComboBox cmbSexo 
      Height          =   315
      ItemData        =   "frmAporteEstado.frx":0010
      Left            =   10320
      List            =   "frmAporteEstado.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   2205
      Width           =   1815
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
      Height          =   285
      Left            =   120
      MaxLength       =   8
      TabIndex        =   16
      Top             =   2205
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
      Height          =   285
      Left            =   1080
      MaxLength       =   8
      TabIndex        =   15
      Top             =   2205
      Width           =   930
   End
   Begin VB.ComboBox cmbSitu 
      Height          =   315
      ItemData        =   "frmAporteEstado.frx":0014
      Left            =   2040
      List            =   "frmAporteEstado.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2205
      Width           =   2055
   End
   Begin VB.ComboBox cmbSituEsp 
      Height          =   315
      ItemData        =   "frmAporteEstado.frx":0018
      Left            =   4080
      List            =   "frmAporteEstado.frx":001A
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2205
      Width           =   2055
   End
   Begin VB.CommandButton cndOtro 
      Caption         =   "&Otra Consulta"
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
      Left            =   6360
      TabIndex        =   12
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar"
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
      Left            =   8520
      TabIndex        =   11
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   9840
      TabIndex        =   10
      Top             =   5880
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
      Left            =   11160
      TabIndex        =   9
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox txtCodSocio 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      MaxLength       =   9
      TabIndex        =   2
      Top             =   1725
      Width           =   975
   End
   Begin VB.TextBox txtIns 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7800
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1725
      Width           =   375
   End
   Begin VB.TextBox txtNumdoc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8160
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1725
      Width           =   975
   End
   Begin MSMask.MaskEdBox txtFecNac 
      Height          =   285
      Left            =   6120
      TabIndex        =   21
      Top             =   2205
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecMat 
      Height          =   285
      Left            =   7200
      TabIndex        =   25
      Top             =   2205
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
      Left            =   120
      TabIndex        =   52
      Top             =   4155
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
      Left            =   1200
      TabIndex        =   53
      Top             =   4155
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecReso 
      Height          =   285
      Left            =   6000
      TabIndex        =   54
      Top             =   4155
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
      Left            =   8640
      TabIndex        =   55
      Top             =   4155
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
      Left            =   9720
      TabIndex        =   56
      Top             =   4155
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
      Left            =   10800
      TabIndex        =   57
      Top             =   4155
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin Crystal.CrystalReport Crys2 
      Left            =   12240
      Top             =   960
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
   Begin MSMask.MaskEdBox txtTope 
      Height          =   285
      Left            =   480
      TabIndex        =   71
      Top             =   405
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecVip 
      Height          =   285
      Left            =   960
      TabIndex        =   75
      Top             =   5040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblCartaDieco 
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
      Left            =   2280
      TabIndex        =   77
      Top             =   5040
      Width           =   10215
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha VIP"
      Height          =   210
      Left            =   120
      TabIndex        =   76
      Top             =   5040
      Width           =   855
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
      Left            =   1320
      TabIndex        =   73
      Top             =   405
      Width           =   1935
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
      Left            =   240
      TabIndex        =   72
      Top             =   195
      Width           =   1170
   End
   Begin VB.Label lblMensaje 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   70
      Top             =   5760
      Width           =   5655
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tipo Cobro"
      Height          =   195
      Index           =   18
      Left            =   4020
      TabIndex        =   69
      Top             =   3975
      Width           =   780
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Grado"
      Height          =   195
      Index           =   2
      Left            =   9240
      TabIndex        =   67
      Top             =   1545
      Width           =   915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tomo Legajo"
      Height          =   195
      Index           =   17
      Left            =   2400
      TabIndex        =   65
      Top             =   3975
      Width           =   930
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "Fecha Ing."
      Height          =   210
      Left            =   0
      TabIndex        =   64
      Top             =   3975
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "Fec.Renuncia"
      Height          =   210
      Left            =   1200
      TabIndex        =   63
      Top             =   3975
      Width           =   1095
   End
   Begin VB.Label Label15 
      Caption         =   "Fec.Resol.Ing"
      Height          =   210
      Left            =   6000
      TabIndex        =   62
      Top             =   3975
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nro.Resol.Ingreso"
      Height          =   195
      Index           =   21
      Left            =   7200
      TabIndex        =   61
      Top             =   3975
      Width           =   1275
   End
   Begin VB.Label Label13 
      Caption         =   "Fec.Exclusión"
      Height          =   210
      Left            =   8640
      TabIndex        =   60
      Top             =   3975
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Fec.Expulsión"
      Height          =   210
      Left            =   9720
      TabIndex        =   59
      Top             =   3975
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Fec.Reingreso"
      Height          =   210
      Left            =   10800
      TabIndex        =   58
      Top             =   3975
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Dirección"
      Height          =   195
      Index           =   11
      Left            =   600
      TabIndex        =   49
      Top             =   2505
      Width           =   675
   End
   Begin VB.Label Label9 
      Caption         =   "Ubicación Geográfica"
      Height          =   195
      Left            =   6240
      TabIndex        =   48
      Top             =   2505
      Width           =   2535
   End
   Begin VB.Label lblUbigeo 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7200
      TabIndex        =   47
      Top             =   2685
      Width           =   4935
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Teléfonos"
      Height          =   195
      Index           =   12
      Left            =   255
      TabIndex        =   46
      Top             =   2985
      Width           =   705
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Celular 2"
      Height          =   195
      Index           =   13
      Left            =   5505
      TabIndex        =   45
      Top             =   2985
      Width           =   615
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Correo Electrónico"
      Height          =   195
      Index           =   14
      Left            =   360
      TabIndex        =   44
      Top             =   3480
      Width           =   1305
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Referencia"
      Height          =   195
      Index           =   15
      Left            =   6840
      TabIndex        =   43
      Top             =   2985
      Width           =   780
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Estado de Socio"
      Height          =   195
      Index           =   16
      Left            =   8310
      TabIndex        =   42
      Top             =   3480
      Width           =   1170
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Correo Electrónico 2"
      Height          =   195
      Index           =   19
      Left            =   4320
      TabIndex        =   41
      Top             =   3480
      Width           =   1440
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Celular"
      Height          =   195
      Index           =   20
      Left            =   3120
      TabIndex        =   40
      Top             =   2985
      Width           =   480
   End
   Begin VB.Label Label8 
      Caption         =   "Fecha Matrim"
      Height          =   210
      Left            =   7200
      TabIndex        =   28
      Top             =   2025
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Estado Civil"
      Height          =   195
      Index           =   9
      Left            =   8880
      TabIndex        =   27
      Top             =   2025
      Width           =   825
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Sexo"
      Height          =   195
      Index           =   10
      Left            =   10560
      TabIndex        =   26
      Top             =   2025
      Width           =   360
   End
   Begin VB.Label Label14 
      Caption         =   "Fecha Nacim."
      Height          =   210
      Left            =   6120
      TabIndex        =   22
      Top             =   2025
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Carnet PNP"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   2025
      Width           =   840
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Carnet PIP"
      Height          =   195
      Index           =   7
      Left            =   1155
      TabIndex        =   19
      Top             =   2025
      Width           =   765
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Situación Policial"
      Height          =   195
      Index           =   8
      Left            =   2400
      TabIndex        =   18
      Top             =   2025
      Width           =   1200
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Situación Especial"
      Height          =   195
      Index           =   23
      Left            =   4335
      TabIndex        =   17
      Top             =   2025
      Width           =   1305
   End
   Begin VB.Label Label3 
      Caption         =   "Cod.Socio"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1545
      Width           =   975
   End
   Begin VB.Label lblCodSocio 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   1725
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de Socio"
      Height          =   195
      Left            =   1200
      TabIndex        =   6
      Top             =   1545
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Codofin"
      Height          =   195
      Left            =   6840
      TabIndex        =   5
      Top             =   1545
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Ins"
      Height          =   195
      Left            =   7800
      TabIndex        =   4
      Top             =   1545
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "D.N.I."
      Height          =   195
      Left            =   8160
      TabIndex        =   3
      Top             =   1545
      Width           =   975
   End
End
Attribute VB_Name = "frmAporteEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
   txtDirec.Text = ""
   txtUbi1.Text = ""
   txtUbi2.Text = ""
   txtUbi3.Text = ""
   txtTelefono.Text = ""
   txtTelefon2.Text = ""
   txtCelular.Text = ""
   txtRefer.Text = ""
   txteMail.Text = ""
   txtEMail2.Text = ""
   txtTomo.Text = ""
   txtFecIng.Text = "__/__/____"
   txtFecRenu.Text = "__/__/____"
   txtFecReso.Text = "__/__/____"
   txtFecExpul.Text = "__/__/____"
   txtFecExclu.Text = "__/__/____"
   txtFecRein.Text = "__/__/____"

   lblCartaDieco.Caption = ""

   cmbGrado.ListIndex = 0
   cmbSitu.ListIndex = 0
   cmbSituEsp.ListIndex = 0
   cmbECivil.ListIndex = 0
   cmbSexo.ListIndex = 0
   cmbE_Socio.ListIndex = 0
   cmbTipCob.ListIndex = 0
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
  '                   Select Case Format(Month(wFecMax), "00")
                     Select Case wMesTope
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

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cndOtro_Click()
   Limpiar

   optCodofin.Value = True
   
   txtNumdoc.Enabled = False
   txtCodigo.Enabled = True

   txtCodigo.SetFocus
End Sub

Private Sub Form_Activate()
   frmAporteEstado.Left = (Screen.Width - Width) \ 2
   frmAporteEstado.Top = 0
   
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
   
   Call Limpiar
   
   txtCodigo.SetFocus
End Sub

Private Sub LlenaCab()
   Dim aa As Integer, wRegAct As Integer, wRegTot As Integer, wSoc As Integer, _
       wApo As Currency, wcob As Currency, wDif As Currency, wMesUno As String, wMesDos As String, _
       wSdo As Currency, wAde As Currency, wEnv540 As Currency, wEnv541 As Currency, _
       wMesTope As String
   wSoc = Val(txtCodSocio.Text)
   wMesTope = Left(txtTope.Text, 4) + Right(txtTope.Text, 2)
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_MASIVO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_MASIVO " _
   & " (CODSOCIO, CODIGO, INS, NUMDOC, NOMBRE, E_SOCIO, DEUDA_PT2, ADELANTO, " _
   & "  FECING, FECREIN, FECBAJ, FECRENO, FECRENU, FECEXCLU, FECEXPUL, FECAMNI, " _
   & "  NRESO_ING, FRESO_ING, TOTAPO, TOTCOB, DESDE, HASTA, USU) " _
   & " SELECT " _
   & "  CODSOCIO, CODIGO, INS, NUMDOC, NOMBRE, E_SOCIO, DEUDA_PT2, ADELANTO, " _
   & "  FECING, FECREIN, FECBAJ, FECRENO, FECRENU, FECEXCLU, FECEXPUL, FECAMNI, " _
   & "  NRESO_ING, FRESO_ING, 0, 0, '', '', '" + wcodusu + "' " _
   & " FROM MAESOCIO " _
   & " WHERE CODSOCIO = " + Str(wSoc) + " ")
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
         & "      ADELANTO = " + Str(wAde) + ", " _
         & "      ENV_540 = " + Str(wEnv540) + ", ENV_541 = " + Str(wEnv541) + " " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " ")
         Db.CommitTrans
   
         Db.BeginTrans
         Db.Execute ("UPDATE ZZZ_MAESTRO " _
         & " SET DEUDA_PT2 = " + Str(wSdo) + ", " _
         & "      ADELANTO = " + Str(wAde) + ", " _
         & "      ENV_540 = " + Str(wEnv540) + ", ENV_541 = " + Str(wEnv541) + " " _
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

Private Sub llenadet()
   Dim aa As Integer, wCod As Long, wDni As String
   wCod = Val(txtCodigo.Text)
   wDni = txtNumdoc.Text

   If optCodofin.Value = True Then
      If Len(Trim(txtIns.Text)) > 0 Then
         aa = Leerado7a("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(wCod) + " AND INS = " + Str(txtIns.Text) + " ")
      Else
         aa = Leerado7a("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(wCod) + " ")
      End If
   Else
      aa = Leerado7a("SELECT * FROM MAESOCIO WHERE NUMDOC = '" + wDni + "' ")
   End If
   If aa > 0 Then
      txtCodSocio.Text = ADO7a!codsocio
      txtCodigo.Text = ADO7a!codigo
      txtIns.Text = ADO7a!ins
      txtNumdoc.Text = ADO7a!numdoc
      txtNumdoc.Text = ADO7a!numdoc
      txtCarnetPNP.Text = ADO7a!carnetpnp
      txtCarnetPIP.Text = ADO7a!carnetpip
      txtDirec.Text = ADO7a!direc
      txtUbi1.Text = Mid(ADO7a!ubigeo, 1, 2)
      txtUbi2.Text = Mid(ADO7a!ubigeo, 3, 2)
      txtUbi3.Text = Mid(ADO7a!ubigeo, 5, 2)
      txtTelefono.Text = ADO7a!telefono
      txtTelefon2.Text = ADO7a!telefon2
      txtCelular.Text = ADO7a!celular
      txteMail.Text = ADO7a!email
      txtEMail2.Text = ADO7a!email2
      txtRefer.Text = ADO7a!refer
      txtTomo.Text = ADO7a!tomo
      txtNumReso.Text = ADO7a!nreso_ing
      
      If ADO7a!cartadieco = True Then
         lblCartaDieco.Caption = "ASOCIADO SIN CARTA AUTORIZACION DIECO"
      Else
         lblCartaDieco.Caption = ""
      End If
      
      If IsDate(ADO7a!fecnac) Then
         txtFecNac.Text = Format(ADO7a!fecnac, "dd/mm/yyyy")
      Else
         txtFecNac.Text = "__/__/____"
      End If
      If IsDate(ADO7a!fecmat) Then
         txtFecMat.Text = Format(ADO7a!fecmat, "dd/mm/yyyy")
      Else
         txtFecMat.Text = "__/__/____"
      End If
      If IsDate(ADO7a!fecing) Then
         txtFecIng.Text = Format(ADO7a!fecing, "dd/mm/yyyy")
      Else
         txtFecIng.Text = "__/__/____"
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
   
      If ADO7a!vip = True Then
         chkVip.Value = vbChecked
      Else
         chkVip.Value = vbUnchecked
      End If
   
      cmbGrado.ListIndex = BuscaGrado(ADO7a!grado)
      cmbSitu.ListIndex = BuscaSitu(ADO7a!situ)
      cmbSituEsp.ListIndex = BuscaSituEsp(ADO7a!situesp)
      cmbECivil.ListIndex = BuscaECivil(ADO7a!ecivil)
      cmbSexo.ListIndex = BuscaSexo(ADO7a!sexo)
      cmbE_Socio.ListIndex = BuscaEsocio(ADO7a!e_socio)
      cmbTipCob.ListIndex = BuscaTipCob(ADO7a!tipcob)
   End If

End Sub

Private Sub optCodofin_Click()
   Limpiar
   
   txtNumdoc.Enabled = False
   txtCodigo.Enabled = True
   txtCodigo.SetFocus
End Sub

Private Sub optDNI_Click()
   Limpiar
   
   txtCodigo.Enabled = False
   txtNumdoc.Enabled = True
   txtNumdoc.SetFocus
End Sub

Private Sub txtCodigo_Change()
'   Dim aa As Integer
'   If Val(txtCodigo.Text) <> 0 Then
'      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(Val(txtCodigo.Text)) + " ")
'      If aa > 0 Then
'         lblCodSocio.Caption = ADO8!nombre
'      Else
'         lblCodSocio.Caption = ""
'      End If
'      Set ADO8 = Nothing
'   Else
'      lblCodSocio.Caption = ""
'   End If
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
           txtIns.Text = xselecIns
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
      If Len(Trim(txtIns.Text)) <> 0 Then
         aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(Val(txtCodigo.Text)) + " AND INS = " + Str(Val(txtIns.Text)) + " ")
      Else
         aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(Val(txtCodigo.Text)) + " ")
      End If
      If aa = 0 Then
         MsgBox "Codofin Digitado NO Existe", vbExclamation
         txtCodigo.Text = ""
         Exit Sub
      End If
      lblCodSocio.Caption = ADO8!nombre
      
      llenadet
      LlenaCab
   
      cmdImprimir.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtCodSocio_Change()
   Dim aa As Integer
   If Val(txtCodSocio.Text) <> 0 Then
      If Len(Trim(txtIns.Text)) > 0 Then
         aa = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " AND INS = " + Str(Val(txtIns.Text)) + " ")
      Else
         aa = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
      End If
      If aa > 0 Then
         lblCodSocio.Caption = ADO6a!nombre
      Else
         lblCodSocio.Caption = ""
         Limpiar
      End If
      Set ADO6a = Nothing
   Else
      lblCodSocio.Caption = ""
   End If
End Sub

Private Sub txtCodSocio_GotFocus()
   txtCodSocio.SelStart = 0
   If Len(Trim(txtCodSocio.Text)) > 0 Then
      txtCodSocio.SelLength = Len(Trim(txtCodSocio.Text))
   Else
      txtCodSocio.SelLength = 8
   End If
End Sub

Private Sub txtCodSocio_KeyDown(KeyCode As Integer, Shift As Integer)
   
   Dim KeyNew As Integer
   KeyNew = Asc(UCase(Chr(KeyCode)))
   
   Select Case KeyCode
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtCodSocio.Text = xseleccion
        End If
   Case 65 To 90
        xlista = "1"
        xseleccion = Chr(KeyNew)
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtCodSocio.Text = xseleccion
        End If
          
   End Select
End Sub

Private Sub txtCodSocio_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtCodSocio.Text)) = 0 Then
         MsgBox "Codigo Socio En Blanco", vbExclamation
         txtCodSocio.Text = ""
         Exit Sub
      End If
      If Len(Trim(txtIns.Text)) > 0 Then
         aa = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " AND INS = " + Str(Val(txtIns.Text)) + " ")
      Else
         aa = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
      End If
      If aa = 0 Then
         MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
         txtCodSocio.Text = ""
         Exit Sub
      End If
      lblCodSocio.Caption = ADO8!nombre
      
      llenadet
      LlenaCab
      
      cmdImprimir.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
'         KeyAscii = 0
'      End If
   End If
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
      
      cmdImprimir.SetFocus
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
        cmbE_Socio.SetFocus
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
      txtCodSocio.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii), 1) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub


