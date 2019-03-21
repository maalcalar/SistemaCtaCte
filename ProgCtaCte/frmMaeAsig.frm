VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMaeAsig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Familiares Asignados Por Socio"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   12075
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtObservac5 
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
      Left            =   7200
      MaxLength       =   20
      TabIndex        =   65
      Top             =   2860
      Width           =   2010
   End
   Begin VB.TextBox txtObservac4 
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
      Left            =   7200
      MaxLength       =   20
      TabIndex        =   64
      Top             =   2580
      Width           =   2010
   End
   Begin VB.TextBox txtObservac3 
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
      Left            =   7200
      MaxLength       =   20
      TabIndex        =   63
      Top             =   2300
      Width           =   2010
   End
   Begin VB.TextBox txtObservac2 
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
      Left            =   7200
      MaxLength       =   20
      TabIndex        =   62
      Top             =   2020
      Width           =   2010
   End
   Begin VB.TextBox txtObservac1 
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
      Left            =   7200
      MaxLength       =   20
      TabIndex        =   60
      Top             =   1740
      Width           =   2010
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
      Left            =   8640
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   3480
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
      Left            =   9960
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtEstado5 
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
      Left            =   9240
      MaxLength       =   1
      TabIndex        =   55
      Top             =   2860
      Width           =   330
   End
   Begin VB.TextBox txtEstado4 
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
      Left            =   9240
      MaxLength       =   1
      TabIndex        =   53
      Top             =   2580
      Width           =   330
   End
   Begin VB.TextBox txtEstado3 
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
      Left            =   9240
      MaxLength       =   1
      TabIndex        =   51
      Top             =   2300
      Width           =   330
   End
   Begin VB.TextBox txtEstado2 
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
      Left            =   9240
      MaxLength       =   1
      TabIndex        =   49
      Top             =   2020
      Width           =   330
   End
   Begin VB.TextBox txtEstado1 
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
      Left            =   9240
      MaxLength       =   1
      TabIndex        =   47
      Top             =   1740
      Width           =   330
   End
   Begin VB.ComboBox cmbE_Socio 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmMaeAsig.frx":0000
      Left            =   6120
      List            =   "frmMaeAsig.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   900
      Width           =   3255
   End
   Begin VB.ComboBox cmbTipCob 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmMaeAsig.frx":0004
      Left            =   6120
      List            =   "frmMaeAsig.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   420
      Width           =   3255
   End
   Begin VB.TextBox txtNombre 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   34
      Top             =   900
      Width           =   6015
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
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      MaxLength       =   8
      TabIndex        =   33
      Top             =   420
      Width           =   690
   End
   Begin VB.ComboBox cmbGrado 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmMaeAsig.frx":0008
      Left            =   3120
      List            =   "frmMaeAsig.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   420
      Width           =   3015
   End
   Begin VB.TextBox txtCodofin 
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
      Left            =   840
      MaxLength       =   8
      TabIndex        =   31
      Top             =   420
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
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   30
      Top             =   420
      Width           =   330
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
      Left            =   2160
      MaxLength       =   8
      TabIndex        =   29
      Top             =   420
      Width           =   930
   End
   Begin VB.TextBox txtSocio5 
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
      Left            =   600
      MaxLength       =   8
      TabIndex        =   21
      Top             =   2860
      Width           =   690
   End
   Begin VB.TextBox txtSocio4 
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
      Left            =   600
      MaxLength       =   8
      TabIndex        =   17
      Top             =   2580
      Width           =   690
   End
   Begin VB.TextBox txtSocio3 
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
      Left            =   600
      MaxLength       =   8
      TabIndex        =   13
      Top             =   2300
      Width           =   690
   End
   Begin VB.TextBox txtSocio1 
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
      Left            =   600
      MaxLength       =   8
      TabIndex        =   1
      Top             =   1740
      Width           =   690
   End
   Begin VB.TextBox txtSocio2 
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
      Left            =   600
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2020
      Width           =   690
   End
   Begin MSMask.MaskEdBox txtFecTop1 
      Height          =   285
      Left            =   10680
      TabIndex        =   48
      Top             =   1740
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecTop2 
      Height          =   285
      Left            =   10680
      TabIndex        =   50
      Top             =   2025
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecTop3 
      Height          =   285
      Left            =   10680
      TabIndex        =   52
      Top             =   2295
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecTop4 
      Height          =   285
      Left            =   10680
      TabIndex        =   54
      Top             =   2580
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecTop5 
      Height          =   285
      Left            =   10680
      TabIndex        =   56
      Top             =   2865
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblEstado5 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9600
      TabIndex        =   70
      Top             =   2860
      Width           =   1095
   End
   Begin VB.Label lblEstado4 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9600
      TabIndex        =   69
      Top             =   2580
      Width           =   1095
   End
   Begin VB.Label lblEstado3 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9600
      TabIndex        =   68
      Top             =   2300
      Width           =   1095
   End
   Begin VB.Label lblEstado2 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9600
      TabIndex        =   67
      Top             =   2020
      Width           =   1095
   End
   Begin VB.Label lblEstado1 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9600
      TabIndex        =   66
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Observaciones"
      Height          =   195
      Index           =   8
      Left            =   7110
      TabIndex        =   61
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Familiares Asignados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   59
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fecha Final"
      Height          =   195
      Index           =   7
      Left            =   10800
      TabIndex        =   46
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      Height          =   195
      Index           =   6
      Left            =   9360
      TabIndex        =   45
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Estado de Socio"
      Height          =   195
      Index           =   16
      Left            =   6390
      TabIndex        =   44
      Top             =   720
      Width           =   1170
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tipo Cobro"
      Height          =   195
      Index           =   18
      Left            =   6660
      TabIndex        =   42
      Top             =   240
      Width           =   780
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Codigo"
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   40
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Apellidos y Nombres de Asociado"
      Height          =   195
      Index           =   3
      Left            =   600
      TabIndex        =   39
      Top             =   720
      Width           =   3915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Grado"
      Height          =   195
      Index           =   2
      Left            =   3120
      TabIndex        =   38
      Top             =   240
      Width           =   915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Codofin"
      Height          =   195
      Index           =   1
      Left            =   915
      TabIndex        =   37
      Top             =   240
      Width           =   540
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Ins"
      Height          =   195
      Index           =   4
      Left            =   1800
      TabIndex        =   36
      Top             =   240
      Width           =   210
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "D.N.I."
      Height          =   195
      Index           =   5
      Left            =   2355
      TabIndex        =   35
      Top             =   240
      Width           =   420
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "5.-"
      Height          =   195
      Index           =   32
      Left            =   360
      TabIndex        =   28
      Top             =   2860
      Width           =   180
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "4.-"
      Height          =   195
      Index           =   31
      Left            =   360
      TabIndex        =   27
      Top             =   2580
      Width           =   180
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "3.-"
      Height          =   195
      Index           =   30
      Left            =   360
      TabIndex        =   26
      Top             =   2300
      Width           =   180
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "2.-"
      Height          =   195
      Index           =   24
      Left            =   360
      TabIndex        =   25
      Top             =   2020
      Width           =   180
   End
   Begin VB.Label lblSocio5 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2760
      TabIndex        =   24
      Top             =   2860
      Width           =   4335
   End
   Begin VB.Label lblIns5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2280
      TabIndex        =   23
      Top             =   2860
      Width           =   375
   End
   Begin VB.Label lblCodigo5 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1320
      TabIndex        =   22
      Top             =   2860
      Width           =   855
   End
   Begin VB.Label lblSocio4 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2760
      TabIndex        =   20
      Top             =   2580
      Width           =   4335
   End
   Begin VB.Label lblIns4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2280
      TabIndex        =   19
      Top             =   2580
      Width           =   375
   End
   Begin VB.Label lblCodigo4 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1320
      TabIndex        =   18
      Top             =   2580
      Width           =   855
   End
   Begin VB.Label lblSocio3 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   2300
      Width           =   4335
   End
   Begin VB.Label lblIns3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   2300
      Width           =   375
   End
   Begin VB.Label lblCodigo3 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   2300
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Ins"
      Height          =   195
      Index           =   29
      Left            =   2280
      TabIndex        =   12
      Top             =   1560
      Width           =   210
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Codofin"
      Height          =   195
      Index           =   28
      Left            =   1395
      TabIndex        =   11
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Codigo"
      Height          =   195
      Index           =   27
      Left            =   705
      TabIndex        =   10
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "1.-"
      Height          =   195
      Index           =   26
      Left            =   360
      TabIndex        =   9
      Top             =   1740
      Width           =   180
   End
   Begin VB.Label lblSocio1 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   1740
      Width           =   4335
   End
   Begin VB.Label lblIns1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   1740
      Width           =   375
   End
   Begin VB.Label lblCodigo1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   1740
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nombre del Asociado"
      Height          =   195
      Index           =   25
      Left            =   3240
      TabIndex        =   5
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label lblSocio2 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   2020
      Width           =   4335
   End
   Begin VB.Label lblIns2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   2020
      Width           =   375
   End
   Begin VB.Label lblCodigo2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   2020
      Width           =   855
   End
End
Attribute VB_Name = "frmMaeAsig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGrabar_Click()


   MsgBox "Cambios Grabados OK", vbExclamation
   Unload Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Form_Initialize
   
   cmdGrabar.Enabled = False
   cmdSalir.SetFocus
End Sub

Private Sub Form_Initialize()
   frmMaeAsig.Left = (Screen.Width - Width) \ 2
   frmMaeAsig.Top = 0
   
   Dim a As Integer, I As Integer
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

   txtCodigo.Text = zSocio

   LlenaCab
End Sub

Private Sub LlenaCab()
   Dim yy As Integer
   
   Dim wSoc1 As Integer, wSoc2 As Integer, wSoc3 As Integer, wSoc4 As Integer, wSoc5 As Integer, _
       wCod1 As Long, wCod2 As Long, wCod3 As Long, wCod4 As Long, wCod5 As Long, _
       wIns1 As Integer, wIns2 As Integer, wIns3 As Integer, wIns4 As Integer, wIns5 As Integer, _
       wNom1 As String, wNom2 As String, wNom3 As String, wNom4 As String, wNom5 As String, _
       wObs1 As String, wObs2 As String, wObs3 As String, wObs4 As String, wObs5 As String, _
       wEst1 As String, wEst2 As String, wEst3 As String, wEst4 As String, wEst5 As String, _
       wFec1 As Date, wFec2 As Date, wFec3 As Date, wFec4 As Date, wFec5 As Date
   
   wSoc1 = 0: wSoc2 = 0: wSoc3 = 0: wSoc4 = 0: wSoc5 = 0
   wCod1 = 0: wCod2 = 0: wCod3 = 0: wCod4 = 0: wCod5 = 0
   wIns1 = 0: wIns2 = 0: wIns3 = 0: wIns4 = 0: wIns5 = 0
   wNom1 = "": wNom2 = "": wNom3 = "": wNom4 = "": wNom5 = ""
   wObs1 = "": wObs2 = "": wObs3 = "": wObs4 = "": wObs5 = ""
   wEst1 = "": wEst2 = "": wEst3 = "": wEst4 = "": wEst5 = ""
   wFec1 = Format("01/01/1900", "dd/mm/yyyy"): wFec2 = Format("01/01/1900", "dd/mm/yyyy"): wFec3 = Format("01/01/1900", "dd/mm/yyyy"): wFec4 = Format("01/01/1900", "dd/mm/yyyy"): wFec5 = Format("01/01/1900", "dd/mm/yyyy")
   
   yy = Leerado5a("SELECT * FROM MAEASIGNADO " _
                & " WHERE CODSOCIO = " + Str(zSocio) + " AND " _
                & "            LIN = '01' ")
   If yy > 0 Then
      wSoc1 = ADO5a!codhijo
      wEst1 = ADO5a!estado
      wObs1 = ADO5a!observ
      wCod1 = BuscaDatosSocio(wSoc1, 1)
      wIns1 = BuscaDatosSocio(wSoc1, 2)
      wNom1 = BuscaDatosSocio(wSoc1, 3)
      If IsDate(ADO5a!fectop) Then
         wFec1 = Format(ADO5a!fectop, "dd/mm/yyyy")
      Else
         wFec1 = Format("01/01/1900", "dd/mm/yyyy")
      End If
   End If

   yy = Leerado5a("SELECT * FROM MAEASIGNADO " _
                & " WHERE CODSOCIO = " + Str(zSocio) + " AND " _
                & "            LIN = '02' ")
   If yy > 0 Then
      wSoc2 = ADO5a!codhijo
      wEst2 = ADO5a!estado
      wObs2 = ADO5a!observ
      wCod2 = BuscaDatosSocio(wSoc2, 1)
      wIns2 = BuscaDatosSocio(wSoc2, 2)
      wNom2 = BuscaDatosSocio(wSoc2, 3)
      If IsDate(ADO5a!fectop) Then
         wFec2 = Format(ADO5a!fectop, "dd/mm/yyyy")
      Else
         wFec2 = Format("01/01/1900", "dd/mm/yyyy")
      End If
   End If

   yy = Leerado5a("SELECT * FROM MAEASIGNADO " _
                & " WHERE CODSOCIO = " + Str(zSocio) + " AND " _
                & "            LIN = '03' ")
   If yy > 0 Then
      wSoc3 = ADO5a!codhijo
      wEst3 = ADO5a!estado
      wObs3 = ADO5a!observ
      wCod3 = BuscaDatosSocio(wSoc3, 1)
      wIns3 = BuscaDatosSocio(wSoc3, 2)
      wNom3 = BuscaDatosSocio(wSoc3, 3)
      If IsDate(ADO5a!fectop) Then
         wFec3 = Format(ADO5a!fectop, "dd/mm/yyyy")
      Else
         wFec3 = Format("01/01/1900", "dd/mm/yyyy")
      End If
   End If

   yy = Leerado5a("SELECT * FROM MAEASIGNADO " _
                & " WHERE CODSOCIO = " + Str(zSocio) + " AND " _
                & "            LIN = '04' ")
   If yy > 0 Then
      wSoc4 = ADO5a!codhijo
      wEst4 = ADO5a!estado
      wObs4 = ADO5a!observ
      wCod4 = BuscaDatosSocio(wSoc4, 1)
      wIns4 = BuscaDatosSocio(wSoc4, 2)
      wNom4 = BuscaDatosSocio(wSoc4, 3)
      If IsDate(ADO5a!fectop) Then
         wFec4 = Format(ADO5a!fectop, "dd/mm/yyyy")
      Else
         wFec4 = Format("01/01/1900", "dd/mm/yyyy")
      End If
   End If

   yy = Leerado5a("SELECT * FROM MAEASIGNADO " _
                & " WHERE CODSOCIO = " + Str(zSocio) + " AND " _
                & "            LIN = '05' ")
   If yy > 0 Then
      wSoc5 = ADO5a!codhijo
      wEst5 = ADO5a!estado
      wObs5 = ADO5a!observ
      wCod5 = BuscaDatosSocio(wSoc5, 1)
      wIns5 = BuscaDatosSocio(wSoc5, 2)
      wNom5 = BuscaDatosSocio(wSoc5, 3)
      If IsDate(ADO5a!fectop) Then
         wFec5 = Format(ADO5a!fectop, "dd/mm/yyyy")
      Else
         wFec5 = Format("01/01/1900", "dd/mm/yyyy")
      End If
   End If

   txtSocio1.Text = wSoc1
   lblCodigo1.Caption = Format(wCod1, "#######0;;\ ")
   lblIns1.Caption = Format(wIns1, "0;;\ ")
   lblSocio1.Caption = wNom1
   txtObservac1.Text = wObs1
   txtEstado1.Text = wEst1
   If wFec1 = Format("01/01/1900", "dd/mm/yyyy") Then
      txtFecTop1.Text = "__/__/____"
   Else
      txtFecTop1.Text = Format(wFec1, "dd/mm/yyyy")
   End If

   txtSocio2.Text = wSoc2
   lblCodigo2.Caption = Format(wCod2, "#######0;;\ ")
   lblIns2.Caption = Format(wIns2, "0;;\ ")
   lblSocio2.Caption = wNom2
   txtObservac2.Text = wObs2
   txtEstado2.Text = wEst2
   If wFec2 = Format("01/01/1900", "dd/mm/yyyy") Then
      txtFecTop2.Text = "__/__/____"
   Else
      txtFecTop2.Text = Format(wFec2, "dd/mm/yyyy")
   End If

   txtSocio3.Text = wSoc3
   lblCodigo3.Caption = Format(wCod3, "#######0;;\ ")
   lblIns3.Caption = Format(wIns3, "0;;\ ")
   lblSocio3.Caption = wNom3
   txtObservac3.Text = wObs3
   txtEstado3.Text = wEst3
   If wFec3 = Format("01/01/1900", "dd/mm/yyyy") Then
      txtFecTop3.Text = "__/__/____"
   Else
      txtFecTop3.Text = Format(wFec3, "dd/mm/yyyy")
   End If

   txtSocio4.Text = wSoc4
   lblCodigo4.Caption = Format(wCod4, "#######0;;\ ")
   lblIns4.Caption = Format(wIns4, "0;;\ ")
   lblSocio4.Caption = wNom4
   txtObservac4.Text = wObs4
   txtEstado4.Text = wEst4
   If wFec4 = Format("01/01/1900", "dd/mm/yyyy") Then
      txtFecTop4.Text = "__/__/____"
   Else
      txtFecTop4.Text = Format(wFec4, "dd/mm/yyyy")
   End If

   txtSocio5.Text = wSoc5
   lblCodigo5.Caption = Format(wCod5, "#######0;;\ ")
   lblIns5.Caption = Format(wIns5, "0;;\ ")
   lblSocio5.Caption = wNom5
   txtObservac5.Text = wObs5
   txtEstado5.Text = wEst5
   If wFec5 = Format("01/01/1900", "dd/mm/yyyy") Then
      txtFecTop5.Text = "__/__/____"
   Else
      txtFecTop5.Text = Format(wFec5, "dd/mm/yyyy")
   End If

End Sub

Private Sub txtCodigo_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodigo.Text)) + " ")
   If aa > 0 Then
   
      txtIns.Text = ADO6a!ins
      txtNumDoc.Text = ADO6a!numdoc
      txtNombre.Text = ADO6a!nombre
      
'      txtE_socio.Text = ADO6a!e_socio
'      txtGrado.Text = ADO6a!grado
'      txtTipCob.Text = ADO6a!tipcob
   
      cmbGrado.ListIndex = BuscaGrado(ADO6a!grado)
      cmbE_Socio.ListIndex = BuscaEsocio(ADO6a!e_socio)
      cmbTipCob.ListIndex = BuscaTipCob(ADO6a!tipcob)
   
   End If
   Set ADO6a = Nothing
End Sub

Private Sub txtEstado1_Change()
   Select Case txtEstado1.Text
   Case "H"
        lblEstado1.Caption = "Habilitado"
   Case "D"
        lblEstado1.Caption = "Deshabilitado"
   Case Else
        lblEstado1.Caption = ""
   End Select
End Sub

Private Sub txtEstado1_GotFocus()
   txtEstado1.SelStart = 0
   txtEstado1.SelLength = 1
End Sub

Private Sub txtEstado1_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtObservac1.SetFocus
   Case 40
        txtFecTop1.SetFocus
   End Select
End Sub

Private Sub txtEstado1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txtEstado1.Text)) = 0 Then
         If Len(Trim(txtSocio1.Text)) <> 0 Then
            MsgBox "Estado Debe Ser D o H", vbExclamation
            txtEstado1.Text = "D"
            Exit Sub
         End If
      Else
         If Len(Trim(txtSocio1.Text)) = 0 Then
            MsgBox "Socio Esta En Blanco", vbExclamation
            txtEstado1.Text = ""
            Exit Sub
         End If
      End If
      txtFecTop1.SetFocus
   Else
      If InStr(1, "DH" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtEstado2_Change()
   Select Case txtEstado2.Text
   Case "H"
        lblEstado2.Caption = "Habilitado"
   Case "D"
        lblEstado2.Caption = "Deshabilitado"
   Case Else
        lblEstado2.Caption = ""
   End Select
End Sub

Private Sub txtEstado2_GotFocus()
   txtEstado2.SelStart = 0
   txtEstado2.SelLength = 1
End Sub

Private Sub txtEstado2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtObservac2.SetFocus
   Case 40
        txtFecTop2.SetFocus
   End Select
End Sub

Private Sub txtEstado3_Change()
   Select Case txtEstado3.Text
   Case "H"
        lblEstado3.Caption = "Habilitado"
   Case "D"
        lblEstado3.Caption = "Deshabilitado"
   Case Else
        lblEstado3.Caption = ""
   End Select
End Sub

Private Sub txtEstado3_GotFocus()
   txtEstado3.SelStart = 0
   txtEstado3.SelLength = 1
End Sub

Private Sub txtEstado3_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtObservac3.SetFocus
   Case 40
        txtFecTop3.SetFocus
   End Select
End Sub

Private Sub txtEstado4_Change()
   Select Case txtEstado4.Text
   Case "H"
        lblEstado4.Caption = "Habilitado"
   Case "D"
        lblEstado4.Caption = "Deshabilitado"
   Case Else
        lblEstado4.Caption = ""
   End Select
End Sub

Private Sub txtEstado4_GotFocus()
   txtEstado4.SelStart = 0
   txtEstado4.SelLength = 1
End Sub

Private Sub txtEstado4_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtObservac4.SetFocus
   Case 40
        txtFecTop4.SetFocus
   End Select
End Sub

Private Sub txtEstado5_Change()
   Select Case txtEstado5.Text
   Case "H"
        lblEstado5.Caption = "Habilitado"
   Case "D"
        lblEstado5.Caption = "Deshabilitado"
   Case Else
        lblEstado5.Caption = ""
   End Select
End Sub

Private Sub txtEstado5_GotFocus()
   txtEstado5.SelStart = 0
   txtEstado5.SelLength = 1
End Sub

Private Sub txtEstado5_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtObservac5.SetFocus
   Case 40
        txtFecTop5.SetFocus
   End Select
End Sub

Private Sub txtFecTop1_GotFocus()
   txtFecTop1.SelStart = 0
   txtFecTop1.SelLength = 10
End Sub

Private Sub txtFecTop1_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtEstado1.SetFocus
   Case 40
        txtSocio2.SetFocus
   End Select
End Sub

Private Sub txtFecTop1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
   Else
   End If

End Sub

Private Sub txtObservac1_GotFocus()
   txtObservac1.SelStart = 0
   txtObservac1.SelLength = Len(Trim(txtObservac1.Text))
End Sub

Private Sub txtObservac1_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtSocio1.SetFocus
   Case 40
        txtEstado1.SetFocus
   End Select
End Sub

Private Sub txtObservac1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtEstado1.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtObservac2_GotFocus()
   txtObservac2.SelStart = 0
   txtObservac2.SelLength = Len(Trim(txtObservac2.Text))
End Sub

Private Sub txtObservac2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtSocio2.SetFocus
   Case 40
        txtEstado2.SetFocus
   End Select
End Sub

Private Sub txtObservac2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtEstado2.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtObservac3_GotFocus()
   txtObservac3.SelStart = 0
   txtObservac3.SelLength = Len(Trim(txtObservac3.Text))
End Sub

Private Sub txtObservac3_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtSocio3.SetFocus
   Case 40
        txtEstado3.SetFocus
   End Select
End Sub

Private Sub txtObservac3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtEstado3.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtObservac4_GotFocus()
   txtObservac4.SelStart = 0
   txtObservac4.SelLength = Len(Trim(txtObservac4.Text))
End Sub

Private Sub txtObservac4_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtSocio4.SetFocus
   Case 40
        txtEstado4.SetFocus
   End Select
End Sub

Private Sub txtObservac4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtEstado4.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtObservac5_GotFocus()
   txtObservac5.SelStart = 0
   txtObservac5.SelLength = Len(Trim(txtObservac5.Text))
End Sub

Private Sub txtObservac5_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtSocio5.SetFocus
   Case 40
        txtEstado5.SetFocus
   End Select
End Sub

Private Sub txtObservac5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtEstado5.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtSocio1_GotFocus()
   txtSocio1.SelStart = 0
   txtSocio1.SelLength = Len(Trim(txtSocio1.Text))
End Sub

Private Sub txtSocio1_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
   
   Case 40
        txtObservac1.SetFocus
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtSocio1.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtSocio1_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtSocio1.Text)) <> 0 Then
         aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio1.Text)) + " ")
         If aa = 0 Then
            MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
            txtSocio1.Text = ""
            Exit Sub
         End If
         lblCodigo1.Caption = ADO8!codigo
         lblIns1.Caption = ADO8!ins
         lblSocio1.Caption = ADO8!nombre
      Else
         lblCodigo1.Caption = ""
         lblIns1.Caption = ""
         lblSocio1.Caption = ""
         
         txtEstado1.Text = ""
         txtObservac1.Text = ""
         txtFecTop1.Text = "__/__/____"
      End If
      
      txtObservac1.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtSocio2_GotFocus()
   txtSocio2.SelStart = 0
   txtSocio2.SelLength = Len(Trim(txtSocio2.Text))
End Sub

Private Sub txtSocio2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFecTop1.SetFocus
   Case 40
        txtObservac2.SetFocus
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtSocio2.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtSocio2_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtSocio2.Text)) <> 0 Then
         aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio2.Text)) + " ")
         If aa = 0 Then
            MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
            txtSocio2.Text = ""
            Exit Sub
         End If
      
         lblCodigo2.Caption = ADO8!codigo
         lblIns2.Caption = ADO8!ins
         lblSocio2.Caption = ADO8!nombre
      Else
         lblCodigo2.Caption = ""
         lblIns2.Caption = ""
         lblSocio2.Caption = ""
         
         txtEstado2.Text = ""
         txtObservac2.Text = ""
         txtFecTop2.Text = "__/__/____"
      End If
      txtObservac2.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtSocio3_GotFocus()
   txtSocio3.SelStart = 0
   txtSocio3.SelLength = Len(Trim(txtSocio3.Text))
End Sub

Private Sub txtSocio3_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFecTop2.SetFocus
   Case 40
        txtObservac3.SetFocus
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtSocio3.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtSocio3_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtSocio3.Text)) <> 0 Then
         aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio3.Text)) + " ")
         If aa = 0 Then
            MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
            txtSocio3.Text = ""
            Exit Sub
         End If
      
         lblCodigo3.Caption = ADO8!codigo
         lblIns3.Caption = ADO8!ins
         lblSocio3.Caption = ADO8!nombre
      Else
         lblCodigo3.Caption = ""
         lblIns3.Caption = ""
         lblSocio3.Caption = ""
         
         txtEstado3.Text = ""
         txtObservac3.Text = ""
         txtFecTop3.Text = "__/__/____"
      End If
      txtObservac3.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtSocio4_GotFocus()
   txtSocio4.SelStart = 0
   txtSocio4.SelLength = Len(Trim(txtSocio4.Text))
End Sub

Private Sub txtSocio4_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFecTop3.SetFocus
   Case 40
        txtObservac4.SetFocus
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtSocio4.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtSocio4_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtSocio4.Text)) <> 0 Then
         aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio4.Text)) + " ")
         If aa = 0 Then
            MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
            txtSocio4.Text = ""
            Exit Sub
         End If
      
         lblCodigo4.Caption = ADO8!codigo
         lblIns4.Caption = ADO8!ins
         lblSocio4.Caption = ADO8!nombre
      Else
         lblCodigo4.Caption = ""
         lblIns4.Caption = ""
         lblSocio4.Caption = ""
         
         txtEstado4.Text = ""
         txtObservac4.Text = ""
         txtFecTop4.Text = "__/__/____"
      End If
      txtObservac4.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtSocio5_GotFocus()
   txtSocio5.SelStart = 0
   txtSocio5.SelLength = Len(Trim(txtSocio5.Text))
End Sub

Private Sub txtSocio5_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFecTop4.SetFocus
   Case 40
        txtObservac5.SetFocus
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtSocio5.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtSocio5_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtSocio5.Text)) <> 0 Then
         aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio5.Text)) + " ")
         If aa = 0 Then
            MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
            txtSocio5.Text = ""
            Exit Sub
         End If
      
         lblCodigo5.Caption = ADO8!codigo
         lblIns5.Caption = ADO8!ins
         lblSocio5.Caption = ADO8!nombre
      Else
         lblCodigo5.Caption = ""
         lblIns5.Caption = ""
         lblSocio5.Caption = ""
         
         txtEstado5.Text = ""
         txtObservac5.Text = ""
         txtFecTop5.Text = "__/__/____"
      End If
      txtObservac5.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub
