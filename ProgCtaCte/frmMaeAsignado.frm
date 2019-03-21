VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMaeAsignado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Hijos"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   14760
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
      Left            =   10440
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7920
      Width           =   975
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtro x Nombre"
      Height          =   615
      Left            =   1680
      TabIndex        =   19
      Top             =   7680
      Width           =   8175
      Begin VB.TextBox txtFiltrar 
         Height          =   285
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   22
         Top             =   240
         Width           =   4935
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Filtrar x Nombre"
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Mostrar Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame fraMantenimiento 
      Caption         =   "Mantenimiento"
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
      Height          =   1815
      Left            =   12000
      TabIndex        =   13
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1320
         Width           =   975
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
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdDeshacer 
         Caption         =   "&Deshacer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame fraDesplaza 
      Caption         =   "Desplazamiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   12000
      TabIndex        =   8
      Top             =   2520
      Width           =   2655
      Begin VB.CommandButton cmdMover 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdMover 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdMover 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdMover 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1920
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Consultas"
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
      Height          =   735
      Left            =   12000
      TabIndex        =   6
      Top             =   1800
      Width           =   2655
      Begin VB.CommandButton cmdExporta 
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
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraDetalle 
      Caption         =   "Detalle del Registro"
      ForeColor       =   &H00C00000&
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   11775
      Begin VB.ComboBox cmbE_Socio 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMaeAsignado.frx":0000
         Left            =   6360
         List            =   "frmMaeAsignado.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   900
         Width           =   3255
      End
      Begin VB.ComboBox cmbGrado 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMaeAsignado.frx":0004
         Left            =   120
         List            =   "frmMaeAsignado.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   900
         Width           =   3015
      End
      Begin VB.ComboBox cmbTipCob 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMaeAsignado.frx":0008
         Left            =   3120
         List            =   "frmMaeAsignado.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   900
         Width           =   3255
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
         Left            =   2400
         MaxLength       =   8
         TabIndex        =   78
         Top             =   420
         Width           =   930
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
         Height          =   285
         Left            =   345
         MaxLength       =   8
         TabIndex        =   44
         Top             =   2145
         Width           =   930
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
         Height          =   285
         Left            =   360
         MaxLength       =   8
         TabIndex        =   43
         Top             =   1860
         Width           =   930
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
         Height          =   285
         Left            =   345
         MaxLength       =   8
         TabIndex        =   42
         Top             =   2415
         Width           =   930
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
         Height          =   285
         Left            =   345
         MaxLength       =   8
         TabIndex        =   41
         Top             =   2700
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
         Height          =   285
         Left            =   345
         MaxLength       =   8
         TabIndex        =   40
         Top             =   2985
         Width           =   930
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
         Height          =   285
         Left            =   8640
         MaxLength       =   1
         TabIndex        =   39
         Top             =   1860
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
         Height          =   285
         Left            =   8640
         MaxLength       =   1
         TabIndex        =   38
         Top             =   2145
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
         Height          =   285
         Left            =   8640
         MaxLength       =   1
         TabIndex        =   37
         Top             =   2415
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
         Height          =   285
         Left            =   8640
         MaxLength       =   1
         TabIndex        =   36
         Top             =   2700
         Width           =   330
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
         Height          =   285
         Left            =   8640
         MaxLength       =   1
         TabIndex        =   35
         Top             =   2985
         Width           =   330
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
         Height          =   285
         Left            =   6600
         MaxLength       =   20
         TabIndex        =   34
         Top             =   1860
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
         Height          =   285
         Left            =   6600
         MaxLength       =   20
         TabIndex        =   33
         Top             =   2145
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
         Height          =   285
         Left            =   6600
         MaxLength       =   20
         TabIndex        =   32
         Top             =   2415
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
         Height          =   285
         Left            =   6600
         MaxLength       =   20
         TabIndex        =   31
         Top             =   2700
         Width           =   2010
      End
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
         Height          =   285
         Left            =   6600
         MaxLength       =   20
         TabIndex        =   30
         Top             =   2985
         Width           =   2010
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
         Left            =   1035
         MaxLength       =   1
         TabIndex        =   26
         Top             =   420
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
         Left            =   120
         MaxLength       =   8
         TabIndex        =   25
         Top             =   420
         Width           =   930
      End
      Begin VB.TextBox txtCodSocio 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   2
         Top             =   420
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtFecTop1 
         Height          =   285
         Left            =   10680
         TabIndex        =   45
         Top             =   1860
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecTop2 
         Height          =   285
         Left            =   10680
         TabIndex        =   46
         Top             =   2145
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecTop3 
         Height          =   285
         Left            =   10680
         TabIndex        =   47
         Top             =   2415
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecTop4 
         Height          =   285
         Left            =   10680
         TabIndex        =   48
         Top             =   2700
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecTop5 
         Height          =   285
         Left            =   10680
         TabIndex        =   49
         Top             =   2985
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblCodigo5 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1560
         TabIndex        =   89
         Top             =   2985
         Width           =   735
      End
      Begin VB.Label lblCodigo4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1560
         TabIndex        =   88
         Top             =   2700
         Width           =   735
      End
      Begin VB.Label lblCodigo3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1560
         TabIndex        =   87
         Top             =   2415
         Width           =   735
      End
      Begin VB.Label lblCodigo2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1560
         TabIndex        =   86
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Estado de Socio"
         Height          =   195
         Index           =   16
         Left            =   6630
         TabIndex        =   85
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Grado"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   83
         Top             =   720
         Width           =   915
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cobro"
         Height          =   195
         Index           =   18
         Left            =   3660
         TabIndex        =   82
         Top             =   720
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "D.N.I."
         Height          =   195
         Index           =   5
         Left            =   2595
         TabIndex        =   79
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lblIns2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1320
         TabIndex        =   77
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label lblSocio2 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2280
         TabIndex        =   76
         Top             =   2145
         Width           =   4335
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre del Asociado"
         Height          =   195
         Index           =   25
         Left            =   2760
         TabIndex        =   75
         Top             =   1680
         Width           =   1515
      End
      Begin VB.Label lblCodigo1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1560
         TabIndex        =   74
         Top             =   1860
         Width           =   735
      End
      Begin VB.Label lblIns1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1320
         TabIndex        =   73
         Top             =   1860
         Width           =   255
      End
      Begin VB.Label lblSocio1 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2280
         TabIndex        =   72
         Top             =   1860
         Width           =   4335
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1.-"
         Height          =   195
         Index           =   26
         Left            =   105
         TabIndex        =   71
         Top             =   1860
         Width           =   180
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         Height          =   195
         Index           =   27
         Left            =   1680
         TabIndex        =   70
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Codofin"
         Height          =   195
         Index           =   28
         Left            =   360
         TabIndex        =   69
         Top             =   1680
         Width           =   660
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Ins"
         Height          =   195
         Index           =   29
         Left            =   1200
         TabIndex        =   68
         Top             =   1680
         Width           =   330
      End
      Begin VB.Label lblIns3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1320
         TabIndex        =   67
         Top             =   2415
         Width           =   255
      End
      Begin VB.Label lblSocio3 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2280
         TabIndex        =   66
         Top             =   2415
         Width           =   4335
      End
      Begin VB.Label lblIns4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1320
         TabIndex        =   65
         Top             =   2700
         Width           =   255
      End
      Begin VB.Label lblSocio4 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2280
         TabIndex        =   64
         Top             =   2700
         Width           =   4335
      End
      Begin VB.Label lblIns5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1320
         TabIndex        =   63
         Top             =   2985
         Width           =   255
      End
      Begin VB.Label lblSocio5 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2280
         TabIndex        =   62
         Top             =   2985
         Width           =   4335
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "2.-"
         Height          =   195
         Index           =   24
         Left            =   105
         TabIndex        =   61
         Top             =   2145
         Width           =   180
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "3.-"
         Height          =   195
         Index           =   30
         Left            =   105
         TabIndex        =   60
         Top             =   2415
         Width           =   180
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "4.-"
         Height          =   195
         Index           =   31
         Left            =   105
         TabIndex        =   59
         Top             =   2700
         Width           =   180
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "5.-"
         Height          =   195
         Index           =   32
         Left            =   105
         TabIndex        =   58
         Top             =   2985
         Width           =   180
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Index           =   6
         Left            =   8760
         TabIndex        =   57
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Index           =   7
         Left            =   10800
         TabIndex        =   56
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   8
         Left            =   6510
         TabIndex        =   55
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Label lblEstado1 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   9000
         TabIndex        =   54
         Top             =   1860
         Width           =   1695
      End
      Begin VB.Label lblEstado2 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   9000
         TabIndex        =   53
         Top             =   2145
         Width           =   1695
      End
      Begin VB.Label lblEstado3 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   9000
         TabIndex        =   52
         Top             =   2415
         Width           =   1695
      End
      Begin VB.Label lblEstado4 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   9000
         TabIndex        =   51
         Top             =   2700
         Width           =   1695
      End
      Begin VB.Label lblEstado5 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   9000
         TabIndex        =   50
         Top             =   2985
         Width           =   1695
      End
      Begin VB.Label Label2 
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
         TabIndex        =   29
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ins"
         Height          =   195
         Index           =   4
         Left            =   1035
         TabIndex        =   28
         Top             =   240
         Width           =   210
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codofin"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   27
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "Cod.Socio"
         Height          =   195
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblCodSocio 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3360
         TabIndex        =   4
         Top             =   420
         Width           =   6255
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre de Socio"
         Height          =   195
         Left            =   3720
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3855
      Left            =   960
      TabIndex        =   0
      Top             =   3720
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "RELACION DE FAMILIARES  ASIGNADOS"
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
      Left            =   1680
      TabIndex        =   24
      Top             =   8400
      Width           =   8055
   End
End
Attribute VB_Name = "frmMaeAsignado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Limpiar()

   txtCodSocio.Text = ""
   txtCodigo.Text = ""
   txtIns.Text = ""
   txtNumdoc.Text = ""
   lblCodSocio.Caption = ""
   
   cmbGrado.ListIndex = 0
   cmbTipCob.ListIndex = 0
   cmbE_Socio.ListIndex = 0

   txtSocio1.Text = ""
   lblCodigo1.Caption = ""
   lblIns1.Caption = ""
   lblSocio1.Caption = ""
   txtObservac1.Text = ""
   txtEstado1.Text = ""
   lblEstado1.Caption = ""
   txtFecTop1.Text = "__/__/____"
   
   txtSocio2.Text = ""
   lblCodigo2.Caption = ""
   lblIns2.Caption = ""
   lblSocio2.Caption = ""
   txtObservac2.Text = ""
   txtEstado2.Text = ""
   lblEstado2.Caption = ""
   txtFecTop2.Text = "__/__/____"
   
   txtSocio3.Text = ""
   lblCodigo3.Caption = ""
   lblIns3.Caption = ""
   lblSocio3.Caption = ""
   txtObservac3.Text = ""
   txtEstado3.Text = ""
   lblEstado3.Caption = ""
   txtFecTop3.Text = "__/__/____"
   
   txtSocio4.Text = ""
   lblCodigo4.Caption = ""
   lblIns4.Caption = ""
   lblSocio4.Caption = ""
   txtObservac4.Text = ""
   txtEstado4.Text = ""
   lblEstado4.Caption = ""
   txtFecTop4.Text = "__/__/____"
   
   txtSocio5.Text = ""
   lblCodigo5.Caption = ""
   lblIns5.Caption = ""
   lblSocio5.Caption = ""
   txtObservac5.Text = ""
   txtEstado5.Text = ""
   lblEstado5.Caption = ""
   txtFecTop5.Text = "__/__/____"
End Sub

Sub refrescar()
   If ADO1.BOF Then Exit Sub
   If ADO1.EOF Then Exit Sub
   
   Dim wSoc As Integer
   
   wSoc = ADO1!codsocio
   
   txtCodSocio.Text = ADO1!codsocio
   txtCodigo.Text = ADO1!codigo
   txtIns.Text = ADO1!ins
   txtNumdoc.Text = ADO1!numdoc
   lblCodSocio.Caption = ADO1!nombre
   
   cmbGrado.ListIndex = BuscaGrado(ADO1!grado)
   cmbE_Socio.ListIndex = BuscaEsocio(ADO1!e_socio)
   cmbTipCob.ListIndex = BuscaTipCob(ADO1!tipcob)

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_MAEASIGNADODET WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_MAEASIGNADODET " _
   & " (CODSOCIO, LIN, SOCHIJO, CODHIJO, INSHIJO, NOMHIJO, ESTADO, OBSERV, FECTOP, USU) " _
   & " SELECT " _
   & "  M.CODSOCIO, M.LIN, M.CODHIJO, S.CODIGO, S.INS, S.NOMBRE, M.ESTADO, M.OBSERV, " _
   & "  M.FECTOP, '" + wcodusu + "' " _
   & " FROM MAEASIGNADO AS M INNER JOIN MAESOCIO AS S " _
   & "   ON M.CODHIJO = S.CODSOCIO " _
   & " WHERE M.CODSOCIO = " + Str(wSoc) + " ")
   Db.CommitTrans

   aa = Leerado3("SELECT * FROM TMP_MAEASIGNADODET " _
                & " WHERE USU = '" + wcodusu + "' AND CODSOCIO = " + Str(wSoc) + "  " _
                & " ORDER BY LIN ")
   If aa > 0 Then
      ADO3.MoveFirst
      Do While Not ADO3.EOF
   
         Select Case ADO3!lin
         Case "01"
              txtSocio1.Text = ADO3!codhijo
              lblCodigo1.Caption = ADO3!socHijo
              lblIns1.Caption = ADO3!InsHijo
              lblSocio1.Caption = ADO3!nomhijo
              txtObservac1.Text = ADO3!observ
              txtEstado1.Text = ADO3!estado
              If IsDate(ADO3!fectop) Then
                 txtFecTop1.Text = ADO3!fectop
              Else
                 txtFecTop1.Text = "__/__/____"
              End If
         Case "02"
              txtSocio2.Text = ADO3!codhijo
              lblCodigo2.Caption = ADO3!socHijo
              lblIns2.Caption = ADO3!InsHijo
              lblSocio2.Caption = ADO3!nomhijo
              txtObservac2.Text = ADO3!observ
              txtEstado2.Text = ADO3!estado
              If IsDate(ADO3!fectop) Then
                 txtFecTop2.Text = ADO3!fectop
              Else
                 txtFecTop2.Text = "__/__/____"
              End If
         Case "03"
              txtSocio3.Text = ADO3!codhijo
              lblCodigo3.Caption = ADO3!socHijo
              lblIns3.Caption = ADO3!InsHijo
              lblSocio3.Caption = ADO3!nomhijo
              txtObservac3.Text = ADO3!observ
              txtEstado3.Text = ADO3!estado
              If IsDate(ADO3!fectop) Then
                 txtFecTop3.Text = ADO3!fectop
              Else
                 txtFecTop3.Text = "__/__/____"
              End If
         Case "04"
              txtSocio4.Text = ADO3!codhijo
              lblCodigo4.Caption = ADO3!socHijo
              lblIns4.Caption = ADO3!InsHijo
              lblSocio4.Caption = ADO3!nomhijo
              txtObservac4.Text = ADO3!observ
              txtEstado4.Text = ADO3!estado
              If IsDate(ADO3!fectop) Then
                 txtFecTop4.Text = ADO3!fectop
              Else
                 txtFecTop4.Text = "__/__/____"
              End If
         Case "05"
              txtSocio5.Text = ADO3!codhijo
              lblCodigo5.Caption = ADO3!socHijo
              lblIns5.Caption = ADO3!InsHijo
              lblSocio5.Caption = ADO3!nomhijo
              txtObservac5.Text = ADO3!observ
              txtEstado5.Text = ADO3!estado
              If IsDate(ADO3!fectop) Then
                 txtFecTop5.Text = ADO3!fectop
              Else
                 txtFecTop5.Text = "__/__/____"
              End If
         End Select
   
         ADO3.MoveNext
      Loop
   End If
End Sub

Sub grabar()
   On Error GoTo err
   
   Dim aa As Integer, wSoc As Integer, wCod As Long, wIns As Integer, wNumDoc As String, wNombre As String, _
       wTipCob As String, wGrado As Integer, wEsocio, _
       wSocHijo1 As Integer, wCodHijo1 As Long, wInsHijo1 As Integer, wNomHijo1 As String, wObserv1 As String, wEstado1 As String, wFecTop1 As Date, _
       wSocHijo2 As Integer, wCodHijo2 As Long, wInsHijo2 As Integer, wNomHijo2 As String, wObserv2 As String, wEstado2 As String, wFecTop2 As Date, _
       wSocHijo3 As Integer, wCodHijo3 As Long, wInsHijo3 As Integer, wNomHijo3 As String, wObserv3 As String, wEstado3 As String, wFecTop3 As Date, _
       wSocHijo4 As Integer, wCodHijo4 As Long, wInsHijo4 As Integer, wNomHijo4 As String, wObserv4 As String, wEstado4 As String, wFecTop4 As Date, _
       wSocHijo5 As Integer, wCodHijo5 As Long, wInsHijo5 As Integer, wNomHijo5 As String, wObserv5 As String, wEstado5 As String, wFecTop5 As Date
         
   wSoc = Val(txtCodSocio.Text)
   wCod = Val(txtCodigo.Text)
   wIns = Val(txtIns.Text)
   wNumDoc = txtNumdoc.Text
   wNombre = Trim(lblCodSocio.Caption)
   wGrado = BuscaCodGrado(cmbGrado.List(cmbGrado.ListIndex))
   wEsocio = BuscaCodEsocio(cmbE_Socio.List(cmbE_Socio.ListIndex))
   wTipCob = BuscaCodTipCob(cmbTipCob.List(cmbTipCob.ListIndex))
   
   If Len(Trim(wNombre)) = 0 Then
      MsgBox "Nombre En Blanco", vbExclamation
      Exit Sub
   End If
   wSocHijo1 = 0: wCodHijo1 = 0: wInsHijo1 = 0: wNomHijo1 = "": wObserv1 = "": wEstado1 = "": wFecTop1 = Format("01/01/1900", "dd/mm/yyyy")
   wSocHijo2 = 0: wCodHijo2 = 0: wInsHijo2 = 0: wNomHijo2 = "": wObserv2 = "": wEstado2 = "": wFecTop2 = Format("01/01/1900", "dd/mm/yyyy")
   wSocHijo3 = 0: wCodHijo3 = 0: wInsHijo3 = 0: wNomHijo3 = "": wObserv3 = "": wEstado3 = "": wFecTop3 = Format("01/01/1900", "dd/mm/yyyy")
   wSocHijo4 = 0: wCodHijo4 = 0: wInsHijo4 = 0: wNomHijo4 = "": wObserv4 = "": wEstado4 = "": wFecTop4 = Format("01/01/1900", "dd/mm/yyyy")
   wSocHijo5 = 0: wCodHijo5 = 0: wInsHijo5 = 0: wNomHijo5 = "": wObserv5 = "": wEstado5 = "": wFecTop5 = Format("01/01/1900", "dd/mm/yyyy")
   
   wCodHijo1 = Val(txtSocio1.Text)
   wSocHijo1 = Val(lblCodigo1.Caption)
   wInsHijo1 = Val(lblIns1.Caption)
   wNomHijo1 = Trim(lblSocio1.Caption)
   wObserv1 = txtObservac1.Text
   wEstado1 = txtEstado1.Text
   If IsDate(txtFecTop1.Text) Then
      wFecTop1 = Format(txtFecTop1.Text, "dd/mm/yyyy")
   End If
   
   wCodHijo2 = Val(txtSocio2.Text)
   wSocHijo2 = Val(lblCodigo2.Caption)
   wInsHijo2 = Val(lblIns2.Caption)
   wNomHijo2 = Trim(lblSocio2.Caption)
   wObserv2 = txtObservac2.Text
   wEstado2 = txtEstado2.Text
   If IsDate(txtFecTop2.Text) Then
      wFecTop2 = Format(txtFecTop2.Text, "dd/mm/yyyy")
   End If
   
   wCodHijo3 = Val(txtSocio3.Text)
   wSocHijo3 = Val(lblCodigo3.Caption)
   wInsHijo3 = Val(lblIns3.Caption)
   wNomHijo3 = Trim(lblSocio3.Caption)
   wObserv3 = txtObservac3.Text
   wEstado3 = txtEstado3.Text
   If IsDate(txtFecTop3.Text) Then
      wFecTop3 = Format(txtFecTop3.Text, "dd/mm/yyyy")
   End If
   
   wCodHijo4 = Val(txtSocio4.Text)
   wSocHijo4 = Val(lblCodigo4.Caption)
   wInsHijo4 = Val(lblIns4.Caption)
   wNomHijo4 = Trim(lblSocio4.Caption)
   wObserv4 = txtObservac4.Text
   wEstado4 = txtEstado4.Text
   If IsDate(txtFecTop4.Text) Then
      wFecTop4 = Format(txtFecTop4.Text, "dd/mm/yyyy")
   End If
   
   wCodHijo5 = Val(txtSocio5.Text)
   wSocHijo5 = Val(lblCodigo5.Caption)
   wInsHijo5 = Val(lblIns5.Caption)
   wNomHijo5 = Trim(lblSocio5.Caption)
   wObserv5 = txtObservac5.Text
   wEstado5 = txtEstado5.Text
   If IsDate(txtFecTop5.Text) Then
      wFecTop5 = Format(txtFecTop5.Text, "dd/mm/yyyy")
   End If
   
   If wSocHijo1 <> 0 Or wSocHijo2 <> 0 Or wSocHijo3 <> 0 Or _
      wSocHijo4 <> 0 Or wSocHijo5 <> 0 Then
   
      Call CreaDet(wSoc, "01", wSocHijo1, wCodHijo1, wInsHijo1, wNomHijo1, wObserv1, wEstado1, wFecTop1)
      Call CreaDet(wSoc, "02", wSocHijo2, wCodHijo2, wInsHijo2, wNomHijo2, wObserv2, wEstado2, wFecTop2)
      Call CreaDet(wSoc, "03", wSocHijo3, wCodHijo3, wInsHijo3, wNomHijo3, wObserv3, wEstado3, wFecTop3)
      Call CreaDet(wSoc, "04", wSocHijo4, wCodHijo4, wInsHijo4, wNomHijo4, wObserv4, wEstado4, wFecTop4)
      Call CreaDet(wSoc, "05", wSocHijo5, wCodHijo5, wInsHijo5, wNomHijo5, wObserv5, wEstado5, wFecTop5)
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_MAEASIGNADOCAB " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_MAEASIGNADODET " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEASIGNADO " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If
   
   ADO1.Requery
   LlenaCab1
   ADO1.Find "CODIGO = " + Str(Val(wCod)) + " "
   ADO1.Find "   INS = " + Str(Val(wIns)) + " "
   MsgBox "Asignados del Socio " + wNombre + vbNewLine + vbNewLine + _
          "            Grabados OK", vbExclamation
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub CreaDet(zSoc As Integer, zLinH As String, zSocH As Integer, zCodH As Long, zInsH As Integer, zNomH As String, zObsH As String, zEstH As String, zFecH As Date)

   Dim zz As Integer
   
   If zSocH <> 0 Then
      aa = Leerado8("SELECT * FROM TMP_MAEASIGNADODET " _
                 & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
                 & "            LIN = '" + zLinH + "' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO TMP_MAEASIGNADODET " _
         & " (CODSOCIO, LIN, SOCHIJO, CODHIJO, INSHIJO, NOMHIJO, ESTADO, OBSERV, USU) " _
         & " VALUES " _
         & " (" + Str(zSoc) + ", '" + zLinH + "', " _
         & "  " + Str(zCodH) + ", " + Str(zSocH) + ", " + Str(zInsH) + ", " _
         & "  '" + zNomH + "', '" + zEstH + "', '" + zObsH + "', '" + wcodusu + "' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_MAEASIGNADODET " _
         & " SET CODHIJO = " + Str(zSocH) + ", " _
         & "     SOCHIJO = " + Str(zCodH) + ", " _
         & "     INSHIJO = " + Str(zInsH) + ", " _
         & "     NOMHIJO = '" + zNomH + "', " _
         & "      ESTADO = '" + zEstH + "', " _
         & "      OBSERV = '" + zObsH + "' " _
         & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
         & "            LIN = '" + zLinH + "' AND " _
         & "            USU = '" + wcodusu + "' ")
         Db.CommitTrans
      End If
      
      aa = Leerado8("SELECT * FROM MAEASIGNADO " _
                 & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
                 & "            LIN = '" + zLinH + "' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEASIGNADO " _
         & " (CODSOCIO, LIN, CODHIJO, ESTADO, OBSERV) " _
         & " VALUES " _
         & " (" + Str(zSoc) + ", '" + zLinH + "', " _
         & "  " + Str(zSocH) + ", " _
         & "  '" + zEstH + "', '" + zObsH + "' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEASIGNADO " _
         & " SET CODHIJO = " + Str(zSocH) + ", " _
         & "      ESTADO = '" + zEstH + "', " _
         & "      OBSERV = '" + zObsH + "' " _
         & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
         & "            LIN = '" + zLinH + "' ")
         Db.CommitTrans
      End If
      
      If Format(zFecH, "dd/mm/yyyy") > Format("01/01/1900", "dd/mm/yyyy") Then
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_MAEASIGNADODET " _
         & " SET FECTOP = '" + Format(zFecH, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
         & "            LIN = '" + zLinH + "' AND " _
         & "            USU = '" + wcodusu + "' ")
         Db.CommitTrans
      
         Db.BeginTrans
         Db.Execute ("UPDATE MAEASIGNADO " _
         & " SET FECTOP = '" + Format(zFecH, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
         & "            LIN = '" + zLinH + "' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_MAEASIGNADODET " _
         & " SET FECTOP = NULL " _
         & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
         & "            LIN = '" + zLinH + "' AND " _
         & "            USU = '" + wcodusu + "' ")
         Db.CommitTrans
      
         Db.BeginTrans
         Db.Execute ("UPDATE MAEASIGNADO " _
         & " SET FECTOP = NULL " _
         & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
         & "            LIN = '" + zLinH + "' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_MAEASIGNADODET " _
      & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
      & "            LIN = '" + zLinH + "' AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
End Sub

Private Sub editar(estado As Boolean)
   txtCodSocio.Enabled = estado
   txtCodigo.Enabled = estado
   txtIns.Enabled = estado
   txtNumdoc.Enabled = estado
   
   txtSocio1.Enabled = estado
   txtObservac1.Enabled = estado
   txtEstado1.Enabled = estado
   txtFecTop1.Enabled = estado
   
   txtSocio2.Enabled = estado
   txtObservac2.Enabled = estado
   txtEstado2.Enabled = estado
   txtFecTop2.Enabled = estado
   
   txtSocio3.Enabled = estado
   txtObservac3.Enabled = estado
   txtEstado3.Enabled = estado
   txtFecTop3.Enabled = estado
   
   txtSocio4.Enabled = estado
   txtObservac4.Enabled = estado
   txtEstado4.Enabled = estado
   txtFecTop4.Enabled = estado
   
   txtSocio5.Enabled = estado
   txtObservac5.Enabled = estado
   txtEstado5.Enabled = estado
   txtFecTop5.Enabled = estado
   
   cmdNuevo.Visible = Not estado
   cmdModificar.Visible = Not estado
   cmdEliminar.Visible = Not estado
   
   DataGrid1.Enabled = Not estado
   fraDesplaza.Enabled = Not estado
   fraFiltro.Enabled = Not estado
   
   cmdGrabar.Visible = estado
   cmdDeshacer.Visible = estado
   cmdExporta.Visible = Not estado
   cmdSalir.Visible = Not estado
End Sub

Private Sub cmdDeshacer_Click()
   MsgBox "Los Cambios Efectuados Se Perderán", vbExclamation
   ACCION = 0
   
   editar (False)
   
   Limpiar
   refrescar
End Sub

Private Sub cmdExporta_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(12) As String, wreg As Integer, wTot As Integer
   Dim wNom As String, wCod As Long, wIns As Integer, wCodHijo As Long, wInsHijo As Integer, _
       wFec As Date, wSoc As Integer
   Heading(0) = "SOCIO"
   Heading(1) = "CODOFIN"
   Heading(2) = "INS"
   Heading(3) = "NOMBRE"
   Heading(4) = "LIN"
   Heading(5) = "SOCIO HIJO"
   Heading(6) = "CODIGO HIJO"
   Heading(7) = "INS"
   Heading(8) = "NOMBRE HIJO"
   Heading(9) = "ESTADO"
   Heading(10) = "OBSERVAC"
   Heading(11) = "FECHA TOPE"
   Heading(12) = "TIPO COBRO"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 13)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 13)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "MAESTRO DE ASIGNADOS"
        For I = 1 To 13 Step 1
            .Cells(3, I) = Heading(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 9
        objExcel.Columns("B").ColumnWidth = 10
        objExcel.Columns("C").ColumnWidth = 4
        objExcel.Columns("D").ColumnWidth = 50
        objExcel.Columns("E").ColumnWidth = 5
        objExcel.Columns("F").ColumnWidth = 9
        objExcel.Columns("G").ColumnWidth = 10
        objExcel.Columns("H").ColumnWidth = 4
        objExcel.Columns("I").ColumnWidth = 50
        objExcel.Columns("J").ColumnWidth = 8
        objExcel.Columns("K").ColumnWidth = 20
        objExcel.Columns("L").ColumnWidth = 11
        objExcel.Columns("M").ColumnWidth = 20
   End With
   
   Db.BeginTrans
   Db.Execute ("delete from tmp_maeasignadodet where usu = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_MAEASIGNADODET " _
   & " (CODSOCIO, LIN, SOCHIJO, CODHIJO, INSHIJO, NOMHIJO, OBSERV, ESTADO, FECTOP, USU) " _
   & " SELECT " _
   & "  M.CODSOCIO, M.LIN, M.CODHIJO, S.CODIGO, S.INS, S.NOMBRE, M.OBSERV, " _
   & "  M.ESTADO, M.FECTOP, '" + wcodusu + "' " _
   & " FROM MAEASIGNADO AS M INNER JOIN MAESOCIO AS S " _
   & "   ON M.CODHIJO = S.CODSOCIO ")
   Db.CommitTrans
   
   aa = Leerado7("SELECT C.CODSOCIO, C.CODIGO, C.INS, C.NOMBRE, D.LIN, D.SOCHIJO, D.CODHIJO, D.INSHIJO, D.NOMHIJO, D.ESTADO, D.OBSERV, D.FECTOP, M.TIPCOB " _
                & " FROM TMP_MAEASIGNADODET AS D INNER JOIN TMP_MAEASIGNADOCAB AS C ON D.USU = C.USU AND D.CODSOCIO = C.CODSOCIO " _
                & "                              INNER JOIN MAESOCIO           AS M ON C.CODSOCIO = M.CODSOCIO " _
                & " where c.usu = '" + wcodusu + "' " _
                & " ORDER BY C.NOMBRE, D.LIN ")
   If aa > 0 Then
      wTot = aa
      V = 4
      H = 1
      wreg = 1
      wSoc = ADO7!codsocio
      Do While Not ADO7.EOF
         DoEvents
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wreg, "####0") + " / " + Format(wTot, "####0")
         lblMensaje.Refresh
         
         If ADO7!codsocio <> wSoc Then
            V = V + 1
            wSoc = ADO7!codsocio
         End If
         
         objExcel.Cells(V, H + 0) = ADO7!codsocio
         objExcel.Cells(V, H + 1) = ADO7!codigo
         objExcel.Cells(V, H + 2) = ADO7!ins
         objExcel.Cells(V, H + 3) = ADO7!nombre
         objExcel.Cells(V, H + 4) = ADO7!lin
         objExcel.Cells(V, H + 5) = ADO7!socHijo
         objExcel.Cells(V, H + 6) = ADO7!codhijo
         objExcel.Cells(V, H + 7) = ADO7!InsHijo
         objExcel.Cells(V, H + 8) = ADO7!nomhijo
         objExcel.Cells(V, H + 9) = ADO7!estado
         objExcel.Cells(V, H + 10) = ADO7!observ
         If IsDate(ADO7!fectop) Then
            wFec = ADO7!fectop
            objExcel.Cells(V, H + 11) = wFec
         End If
         Select Case ADO7!tipcob
         Case "01"
              objExcel.Cells(V, H + 12) = "DIECO"
         Case "02"
              objExcel.Cells(V, H + 12) = "CAJA MILITAR POLICIAL"
         Case "03"
              objExcel.Cells(V, H + 12) = "TESORERIA AOPIP"
         End Select
         
         wreg = wreg + 1
         V = V + 1
         ADO7.MoveNext
      Loop
      Set ADO7 = Nothing
      objExcel.Visible = True
      Set objExcel = Nothing
   End If
   lblMensaje.Caption = ""
   lblMensaje.Refresh
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdGrabar_Click()
   On Error GoTo err
   Dim wSoc As Integer
   If ACCION = 1 Then
      wSoc = Val(txtCodSocio.Text)
      If Leerado7("SELECT * FROM TMP_MAEASIGNADO WHERE CODSOCIO = " + wSoc + " ") > 0 Then
         MsgBox "Codigo Socio Ya Existe", vbExclamation
         Limpiar
         txtCodSocio.SetFocus
         Exit Sub
      End If
   End If
   grabar
   editar False
   ACCION = 0
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdEliminar_Click()
   On Error GoTo err
   
   Dim wSoc As Integer, wNew As Integer, aa As Integer
   wSoc = ADO1!codsocio
   wNew = 0
   ADO1.MoveNext
   If Not ADO1.EOF Then
      wNew = ADO1!codsocio
   Else
      ADO1.MovePrevious
      ADO1.MovePrevious
      If ADO1.BOF Then
         wNew = 0
      Else
         wNew = ADO1!codsocio
      End If
   End If
   
   If MsgBox("¿Esta seguro de borrar Asignados de Socio " + Str(wSoc) + "?", vbYesNo + vbDefaultButton2 + vbQuestion, "Advertencia") = vbYes Then
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_MAEASIGNADOCAB WHERE CODSOCIO = " + Str(wSoc) + " AND USU = '" + wcodusu + "' ")
      Db.CommitTrans
      
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_MAEASIGNADODET WHERE CODSOCIO = " + Str(wSoc) + " AND USU = '" + wcodusu + "' ")
      Db.CommitTrans
      
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEASIGNADO WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
      
      ADO1.Requery
      LlenaCab
      LlenaCab1
      Limpiar
      refrescar
      
      If wNew <> 0 Then
         ADO1.Find "CODSOCIO=" + Str(wNew) + ""
      End If
      MsgBox "Asignados de Socio " + Str(wSoc) + " " + wNom + vbNewLine + _
             "Eliminado OK", vbExclamation
   End If
   ACCION = 0
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdModificar_Click()
   ACCION = 2
   editar True
   refrescar
   txtCodSocio.Enabled = False
   
   txtSocio1.SetFocus
End Sub

Private Sub cmdNuevo_Click()
   ACCION = 1
   editar True
   Limpiar
   
   txtCodigo.SetFocus
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   Select Case ColIndex
   Case 0
        ADO1.Sort = "CODSOCIO"
   Case 1
        ADO1.Sort = "CODIGO"
   Case 3
        ADO1.Sort = "NOMBRE"
   End Select
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If ACCION = 0 Then
      Limpiar
      refrescar
   End If
End Sub

Private Sub Form_Activate()
   Form_Initialize
End Sub

Private Sub Form_Initialize()
   frmMaeAsignado.Left = (Screen.Width - Width) \ 2
   frmMaeAsignado.Top = 0
   
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

   LlenaCab
   LlenaCab1
   Limpiar
   refrescar
   editar False
   
   DataGrid1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_MAEASIGNADOCAB WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_MAEASIGNADODET WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_MAEASIGNADOCAB WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_MAEASIGNADODET WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_MAEASIGNADOCAB " _
   & " (CODSOCIO, CODIGO, INS, NOMBRE, NUMDOC, E_SOCIO, GRADO, TIPCOB, CANT, USU) " _
   & " SELECT " _
   & "  M.CODSOCIO, S.CODIGO, S.INS, S.NOMBRE, S.NUMDOC, S.E_SOCIO, S.GRADO, " _
   & "  S.TIPCOB, COUNT(M.CODSOCIO), '" + wcodusu + "' " _
   & " FROM MAEASIGNADO AS M INNER JOIN MAESOCIO AS S " _
   & "   ON M.CODSOCIO = S.CODSOCIO " _
   & " GROUP BY M.CODSOCIO, S.CODIGO, S.INS, S.NOMBRE, S.NUMDOC, " _
   & "          S.E_SOCIO, S.GRADO, S.TIPCOB ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_MAEASIGNADODET " _
   & " (CODSOCIO, LIN, SOCHIJO, CODHIJO, INSHIJO, NOMHIJO, OBSERV, ESTADO, FECTOP, USU) " _
   & " SELECT " _
   & "  M.CODSOCIO, M.LIN, M.CODHIJO, S.CODIGO, S.INS, S.NOMBRE, M.OBSERV, " _
   & "  M.ESTADO, M.FECTOP, '" + wcodusu + "' " _
   & " FROM MAEASIGNADO AS M INNER JOIN MAESOCIO AS S " _
   & "   ON M.CODHIJO = S.CODSOCIO ")
   Db.CommitTrans

   Dim aa As Integer

   aa = Leerado("SELECT CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, CANT, NUMDOC, GRADO, TIPCOB, USU " _
                & " FROM TMP_MAEASIGNADOCAB " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO1
End Sub
   
Private Sub LlenaCab1()
   
   DataGrid1.Columns(0).Width = 800
   DataGrid1.Columns(0).Alignment = dbgRight
   DataGrid1.Columns(0).Caption = "SOCIO"
    
   DataGrid1.Columns(1).Width = 1000
   DataGrid1.Columns(1).Alignment = dbgRight
   DataGrid1.Columns(1).Caption = "CODIGO"
    
   DataGrid1.Columns(2).Width = 570
   DataGrid1.Columns(2).Alignment = dbgCenter
   DataGrid1.Columns(2).Caption = "INS"
    
   DataGrid1.Columns(3).Width = 5500
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE"
    
   DataGrid1.Columns(4).Width = 700
   DataGrid1.Columns(4).Alignment = dbgCenter
   DataGrid1.Columns(4).Caption = "ESTADO"
    
   DataGrid1.Columns(5).Width = 500
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).Caption = "CANT"
    
   DataGrid1.Columns(6).Visible = False
   DataGrid1.Columns(7).Visible = False
   DataGrid1.Columns(8).Visible = False
   DataGrid1.Columns(9).Visible = False
End Sub

Private Sub optFiltro_Click()
   If optTodos.Value = True Then
      txtFiltrar.Text = ""
      txtFiltrar.Enabled = False
      DataGrid1.SetFocus
   Else
      txtFiltrar.Enabled = True
      optFiltro.Value = True
      txtFiltrar.SetFocus
   End If
End Sub

Private Sub optFiltro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      optTodos_Click
   End If
End Sub

Private Sub optTodos_Click()
   If optTodos.Value = True Then
      txtFiltrar.Text = ""
      txtFiltrar.Enabled = False
      LlenaCab
      LlenaCab1
      Limpiar
      refrescar
   Else
      txtFiltrar.Enabled = True
      optFiltro.Value = True
   End If
End Sub

Private Sub optTodos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If optTodos.Value = True Then
         txtFiltrar.Text = ""
         txtFiltrar.Enabled = False
         ADO1.Filter = ""
         Set DataGrid1.DataSource = ADO1
         DataGrid1.SetFocus
      Else
         txtFiltrar.Enabled = True
         optFiltro.Value = True
         txtFiltrar.SetFocus
      End If
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
      lblCodSocio.Caption = ADO8!nombre
      
      txtCodSocio.Text = ADO8!codsocio
      txtIns.Text = ADO8!ins
      txtNumdoc.Text = ADO8!numdoc
      lblCodSocio.Caption = ADO8!nombre
   
      cmbGrado.ListIndex = BuscaGrado(ADO8!grado)
      cmbE_Socio.ListIndex = BuscaEsocio(ADO8!e_socio)
      cmbTipCob.ListIndex = BuscaTipCob(ADO8!tipcob)

      txtSocio1.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtEstado1_Change()
   Dim zz As Integer
   zz = Leerado8("SELECT * FROM MAEESTADOASIGNADO WHERE CODIGO = '" + txtEstado1.Text + "' ")
   If zz > 0 Then
      lblEstado1.Caption = ADO8!nombre
   Else
      lblEstado1.Caption = ""
   End If
   Set ADO8 = Nothing
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
   Case 116
        xlista = "9"
        xseleccion = ""
        frmSeleccion.Show 1
        If xseleccion <> "" Then
           txtEstado1.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtEstado1_KeyPress(KeyAscii As Integer)
   Dim zz As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtEstado1.Text)) = 0 Then
         If Len(Trim(txtSocio1.Text)) <> 0 Then
            MsgBox "Estado Asignado Esta En Blanco", vbExclamation
            txtEstado1.Text = "H"
            Exit Sub
         End If
      Else
         zz = Leerado8("SELECT * FROM MAEESTADOASIGNADO WHERE CODIGO = '" + txtEstado1.Text + "' ")
         If zz = 0 Then
            MsgBox "Estado del Socio Es Invalido", vbExclamation
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
   Dim zz As Integer
   zz = Leerado8("SELECT * FROM MAEESTADOASIGNADO WHERE CODIGO = '" + txtEstado2.Text + "' ")
   If zz > 0 Then
      lblEstado2.Caption = ADO8!nombre
   Else
      lblEstado2.Caption = ""
   End If
   Set ADO8 = Nothing
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
   Case 116
        xlista = "9"
        xseleccion = ""
        frmSeleccion.Show 1
        If xseleccion <> "" Then
           txtEstado2.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtEstado2_KeyPress(KeyAscii As Integer)
   Dim zz As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtEstado2.Text)) = 0 Then
         If Len(Trim(txtSocio2.Text)) <> 0 Then
            MsgBox "Estado Asignado Esta En Blanco", vbExclamation
            txtEstado2.Text = "H"
            Exit Sub
         End If
      Else
         zz = Leerado8("SELECT * FROM MAEESTADOASIGNADO WHERE CODIGO = '" + txtEstado2.Text + "' ")
         If zz = 0 Then
            MsgBox "Estado del Socio Es Invalido", vbExclamation
            txtEstado2.Text = ""
            Exit Sub
         End If
      End If
      txtFecTop2.SetFocus
   Else
      If InStr(1, "DH" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtEstado3_Change()
   Dim zz As Integer
   zz = Leerado8("SELECT * FROM MAEESTADOASIGNADO WHERE CODIGO = '" + txtEstado3.Text + "' ")
   If zz > 0 Then
      lblEstado3.Caption = ADO8!nombre
   Else
      lblEstado3.Caption = ""
   End If
   Set ADO8 = Nothing
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
   Case 116
        xlista = "9"
        xseleccion = ""
        frmSeleccion.Show 1
        If xseleccion <> "" Then
           txtEstado3.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtEstado3_KeyPress(KeyAscii As Integer)
   Dim zz As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtEstado3.Text)) = 0 Then
         If Len(Trim(txtSocio3.Text)) <> 0 Then
            MsgBox "Estado Asignado Esta En Blanco", vbExclamation
            txtEstado3.Text = "H"
            Exit Sub
         End If
      Else
         zz = Leerado8("SELECT * FROM MAEESTADOASIGNADO WHERE CODIGO = '" + txtEstado3.Text + "' ")
         If zz = 0 Then
            MsgBox "Estado del Socio Es Invalido", vbExclamation
            txtEstado3.Text = ""
            Exit Sub
         End If
      End If
      txtFecTop3.SetFocus
   Else
      If InStr(1, "DH" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtEstado4_Change()
   Dim zz As Integer
   zz = Leerado8("SELECT * FROM MAEESTADOASIGNADO WHERE CODIGO = '" + txtEstado4.Text + "' ")
   If zz > 0 Then
      lblEstado4.Caption = ADO8!nombre
   Else
      lblEstado4.Caption = ""
   End If
   Set ADO8 = Nothing
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
   Case 116
        xlista = "9"
        xseleccion = ""
        frmSeleccion.Show 1
        If xseleccion <> "" Then
           txtEstado4.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtEstado4_KeyPress(KeyAscii As Integer)
   Dim zz As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtEstado4.Text)) = 0 Then
         If Len(Trim(txtSocio4.Text)) <> 0 Then
            MsgBox "Estado Asignado Esta En Blanco", vbExclamation
            txtEstado4.Text = "H"
            Exit Sub
         End If
      Else
         zz = Leerado8("SELECT * FROM MAEESTADOASIGNADO WHERE CODIGO = '" + txtEstado4.Text + "' ")
         If zz = 0 Then
            MsgBox "Estado del Socio Es Invalido", vbExclamation
            txtEstado4.Text = ""
            Exit Sub
         End If
      End If
      txtFecTop4.SetFocus
   Else
      If InStr(1, "DH" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtEstado5_Change()
   Dim zz As Integer
   zz = Leerado8("SELECT * FROM MAEESTADOASIGNADO WHERE CODIGO = '" + txtEstado5.Text + "' ")
   If zz > 0 Then
      lblEstado5.Caption = ADO8!nombre
   Else
      lblEstado5.Caption = ""
   End If
   Set ADO8 = Nothing
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
   Case 116
        xlista = "9"
        xseleccion = ""
        frmSeleccion.Show 1
        If xseleccion <> "" Then
           txtEstado5.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtEstado5_KeyPress(KeyAscii As Integer)
   Dim zz As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtEstado5.Text)) = 0 Then
         If Len(Trim(txtSocio5.Text)) <> 0 Then
            MsgBox "Estado Asignado Esta En Blanco", vbExclamation
            txtEstado5.Text = "H"
            Exit Sub
         End If
      Else
         zz = Leerado8("SELECT * FROM MAEESTADOASIGNADO WHERE CODIGO = '" + txtEstado5.Text + "' ")
         If zz = 0 Then
            MsgBox "Estado del Socio Es Invalido", vbExclamation
            txtEstado5.Text = ""
            Exit Sub
         End If
      End If
      txtFecTop5.SetFocus
   Else
      If InStr(1, "DH" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
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
      If txtFecTop1.Text <> "__/__/____" Then
         If Not IsDate(txtFecTop1.Text) Then
            MsgBox "Fecha Tope Es Invalida", vbExclamation
            txtFecTop1.Text = "__/__/____"
            Exit Sub
         End If
      End If
      txtSocio2.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFecTop2_GotFocus()
   txtFecTop2.SelStart = 0
   txtFecTop2.SelLength = 10
End Sub

Private Sub txtFecTop2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtEstado2.SetFocus
   Case 40
        txtSocio3.SetFocus
   End Select
End Sub

Private Sub txtFecTop2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFecTop2.Text <> "__/__/____" Then
         If Not IsDate(txtFecTop2.Text) Then
            MsgBox "Fecha Tope Es Invalida", vbExclamation
            txtFecTop2.Text = "__/__/____"
            Exit Sub
         End If
      End If
      txtSocio3.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFecTop3_GotFocus()
   txtFecTop3.SelStart = 0
   txtFecTop3.SelLength = 10
End Sub

Private Sub txtFecTop3_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtEstado3.SetFocus
   Case 40
        txtSocio4.SetFocus
   End Select
End Sub

Private Sub txtFecTop3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFecTop3.Text <> "__/__/____" Then
         If Not IsDate(txtFecTop3.Text) Then
            MsgBox "Fecha Tope Es Invalida", vbExclamation
            txtFecTop3.Text = "__/__/____"
            Exit Sub
         End If
      End If
      txtSocio4.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFecTop4_GotFocus()
   txtFecTop4.SelStart = 0
   txtFecTop4.SelLength = 10
End Sub

Private Sub txtFecTop4_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtEstado4.SetFocus
   Case 40
        txtSocio5.SetFocus
   End Select
End Sub

Private Sub txtFecTop4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFecTop4.Text <> "__/__/____" Then
         If Not IsDate(txtFecTop4.Text) Then
            MsgBox "Fecha Tope Es Invalida", vbExclamation
            txtFecTop4.Text = "__/__/____"
            Exit Sub
         End If
      End If
      txtSocio5.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFecTop5_GotFocus()
   txtFecTop5.SelStart = 0
   txtFecTop5.SelLength = 10
End Sub

Private Sub txtFecTop5_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtEstado5.SetFocus
   Case 40
        cmdGrabar.SetFocus
   End Select
End Sub

Private Sub txtFecTop5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFecTop5.Text <> "__/__/____" Then
         If Not IsDate(txtFecTop5.Text) Then
            MsgBox "Fecha Tope Es Invalida", vbExclamation
            txtFecTop5.Text = "__/__/____"
            Exit Sub
         End If
      End If
      cmdGrabar.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFiltrar_GotFocus()
   txtFiltrar.SelStart = 0
   txtFiltrar.SelLength = Len(Trim(txtFiltrar.Text))
End Sub

Private Sub txtFiltrar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFiltrar.Text = "" Then
         MsgBox "Filtro En Blanco", vbExclamation
         Exit Sub
      End If
      ADO1.Filter = "NOMBRE LIKE '%" & Trim(txtFiltrar) & "%' "
      refrescar
      DataGrid1.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
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
   Dim KeyNew As Integer
   KeyNew = Asc(UCase(Chr(KeyCode)))
   
   Select Case KeyCode
   Case 38
   
   Case 40
        txtObservac1.SetFocus
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio2.Show 1
        If xseleccion <> "" Then
           lblCodigo1.Caption = xselecSocio
           txtSocio1.Text = xseleccion
        End If
   Case 65 To 90
        xlista = "1"
        xseleccion = Chr(KeyNew)
        frmSelecSocio2.Show 1
        If xseleccion <> "" Then
           lblCodigo1.Caption = xselecSocio
           txtSocio1.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtSocio1_KeyPress(KeyAscii As Integer)
   Dim aa As Integer, wSoc As Integer, wPad As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtSocio1.Text)) <> 0 Then
         aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(lblCodigo1.Caption)) + " ")
         If aa = 0 Then
            MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
            txtSocio1.Text = ""
            Exit Sub
         End If
         If ADO8!e_socio <> "HIJ" Then
            MsgBox "Socio Asignado NO Es Hijo", vbExclamation
            txtSocio1.Text = ""
            Exit Sub
         End If
         txtSocio1.Text = ADO8!codigo
         lblCodigo1.Caption = ADO8!codsocio
         lblIns1.Caption = ADO8!ins
         lblSocio1.Caption = Trim(ADO8!nombre)
         txtEstado1.Text = "H"
         
         wPad = Val(txtCodSocio.Text)
         wSoc = Val(lblCodigo1.Caption)
         
         aa = Leerado8("SELECT * FROM TMP_MAEASIGNADODET WHERE SOCHIJO = " + Str(wSoc) + " AND (LIN <> '01' OR CODSOCIO <> " + Str(wPad) + ")  ")
         If aa > 0 Then
            MsgBox "Hijo " + lblSocio1 + " Ya Tiene Asignación", vbExclamation
            txtSocio1.Text = ""
            lblCodigo1.Caption = ""
            lblIns1.Caption = ""
            lblSocio1.Caption = ""
            txtEstado1.Text = ""
            Exit Sub
         End If
         
         aa = Leerado8("SELECT * FROM MAEASIGNADO WHERE CODHIJO = " + Str(wSoc) + " AND (LIN <> '01' OR CODSOCIO <> " + Str(wPad) + ")  ")
         If aa > 0 Then
            MsgBox "Hijo " + lblSocio1 + " Ya Tiene Asignación", vbExclamation
            txtSocio1.Text = ""
            lblCodigo1.Caption = ""
            lblIns1.Caption = ""
            lblSocio1.Caption = ""
            txtEstado1.Text = ""
            Exit Sub
         End If
         
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
   Dim KeyNew As Integer
   KeyNew = Asc(UCase(Chr(KeyCode)))
   
   Select Case KeyCode
   Case 38
        txtFecTop1.SetFocus
   Case 40
        txtObservac2.SetFocus
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio2.Show 1
        If xseleccion <> "" Then
           lblCodigo2.Caption = xselecSocio
           txtSocio2.Text = xseleccion
        End If
   Case 65 To 90
        xlista = "1"
        xseleccion = Chr(KeyNew)
        frmSelecSocio2.Show 1
        If xseleccion <> "" Then
           lblCodigo2.Caption = xselecSocio
           txtSocio2.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtSocio2_KeyPress(KeyAscii As Integer)
   Dim aa As Integer, wSoc As Integer, wPad As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtSocio2.Text)) <> 0 Then
         aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(lblCodigo2.Caption)) + " ")
         If aa = 0 Then
            MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
            txtSocio2.Text = ""
            Exit Sub
         End If
         If ADO8!e_socio <> "HIJ" Then
            MsgBox "Socio Asignado NO Es Hijo", vbExclamation
            txtSocio2.Text = ""
            Exit Sub
         End If
         txtSocio2.Text = ADO8!codigo
         lblCodigo2.Caption = ADO8!codsocio
         lblIns2.Caption = ADO8!ins
         lblSocio2.Caption = Trim(ADO8!nombre)
         txtEstado2.Text = "H"
         
         wPad = Val(txtCodSocio.Text)
         wSoc = Val(lblCodigo2.Caption)
         
         If Val(txtSocio2.Text) = Val(txtSocio1.Text) Or _
            Val(txtSocio2.Text) = Val(txtSocio3.Text) Or _
            Val(txtSocio2.Text) = Val(txtSocio4.Text) Or _
            Val(txtSocio2.Text) = Val(txtSocio5.Text) Then
            MsgBox "Hijo " + lblSocio2 + " Ya Tiene Asignación", vbExclamation
            txtSocio2.Text = ""
            lblCodigo2.Caption = ""
            lblIns2.Caption = ""
            lblSocio2.Caption = ""
            txtEstado2.Text = ""
            Exit Sub
         End If
         
         aa = Leerado8("SELECT * FROM MAEASIGNADO WHERE CODHIJO = " + Str(wSoc) + " AND (LIN <> '02' OR CODSOCIO <> " + Str(wPad) + ")  ")
         If aa > 0 Then
            MsgBox "Hijo " + lblSocio2 + " Ya Tiene Asignación", vbExclamation
            txtSocio2.Text = ""
            lblCodigo2.Caption = ""
            lblIns2.Caption = ""
            lblSocio2.Caption = ""
            txtEstado2.Text = ""
            Exit Sub
         End If
         
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
   Dim KeyNew As Integer
   KeyNew = Asc(UCase(Chr(KeyCode)))
   
   Select Case KeyCode
   Case 38
        txtFecTop2.SetFocus
   Case 40
        txtObservac3.SetFocus
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio2.Show 1
        If xseleccion <> "" Then
           lblCodigo3.Caption = xselecSocio
           txtSocio3.Text = xseleccion
        End If
   Case 65 To 90
        xlista = "1"
        xseleccion = Chr(KeyNew)
        frmSelecSocio2.Show 1
        If xseleccion <> "" Then
           lblCodigo3.Caption = xselecSocio
           txtSocio3.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtSocio3_KeyPress(KeyAscii As Integer)
   Dim aa As Integer, wSoc As Integer, wPad As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtSocio3.Text)) <> 0 Then
         aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(lblCodigo3.Caption)) + " ")
         If aa = 0 Then
            MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
            txtSocio3.Text = ""
            Exit Sub
         End If
         If ADO8!e_socio <> "HIJ" Then
            MsgBox "Socio Asignado NO Es Hijo", vbExclamation
            txtSocio3.Text = ""
            Exit Sub
         End If
         txtSocio3.Text = ADO8!codigo
         lblCodigo3.Caption = ADO8!codsocio
         lblIns3.Caption = ADO8!ins
         lblSocio3.Caption = Trim(ADO8!nombre)
         txtEstado3.Text = "H"
         
         wPad = Val(txtCodSocio.Text)
         wSoc = Val(lblCodigo3.Caption)
         
         If Val(txtSocio3.Text) = Val(txtSocio1.Text) Or _
            Val(txtSocio3.Text) = Val(txtSocio2.Text) Or _
            Val(txtSocio3.Text) = Val(txtSocio4.Text) Or _
            Val(txtSocio3.Text) = Val(txtSocio5.Text) Then
            MsgBox "Hijo " + lblSocio3 + " Ya Tiene Asignación", vbExclamation
            txtSocio3.Text = ""
            lblCodigo3.Caption = ""
            lblIns3.Caption = ""
            lblSocio3.Caption = ""
            txtEstado3.Text = ""
            Exit Sub
         End If
         
         aa = Leerado8("SELECT * FROM MAEASIGNADO WHERE CODHIJO = " + Str(wSoc) + " AND (LIN <> '03' OR CODSOCIO <> " + Str(wPad) + ")  ")
         If aa > 0 Then
            MsgBox "Hijo " + lblSocio3 + " Ya Tiene Asignación", vbExclamation
            txtSocio3.Text = ""
            lblCodigo3.Caption = ""
            lblIns3.Caption = ""
            lblSocio3.Caption = ""
            txtEstado3.Text = ""
            Exit Sub
         End If
         
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
   Dim KeyNew As Integer
   KeyNew = Asc(UCase(Chr(KeyCode)))
   
   Select Case KeyCode
   Case 38
        txtFecTop3.SetFocus
   Case 40
        txtObservac4.SetFocus
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio2.Show 1
        If xseleccion <> "" Then
           lblCodigo4.Caption = xselecSocio
           txtSocio4.Text = xseleccion
        End If
   Case 65 To 90
        xlista = "1"
        xseleccion = Chr(KeyNew)
        frmSelecSocio2.Show 1
        If xseleccion <> "" Then
           lblCodigo4.Caption = xselecSocio
           txtSocio4.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtSocio4_KeyPress(KeyAscii As Integer)
   Dim aa As Integer, wSoc As Integer, wPad As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtSocio4.Text)) <> 0 Then
         aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(lblCodigo4.Caption)) + " ")
         If aa = 0 Then
            MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
            txtSocio4.Text = ""
            Exit Sub
         End If
         If ADO8!e_socio <> "HIJ" Then
            MsgBox "Socio Asignado NO Es Hijo", vbExclamation
            txtSocio4.Text = ""
            Exit Sub
         End If
         txtSocio4.Text = ADO8!codigo
         lblCodigo4.Caption = ADO8!codsocio
         lblIns4.Caption = ADO8!ins
         lblSocio4.Caption = Trim(ADO8!nombre)
         txtEstado4.Text = "H"
         
         wPad = Val(txtCodSocio.Text)
         wSoc = Val(lblCodigo4.Caption)
         
         If Val(txtSocio4.Text) = Val(txtSocio1.Text) Or _
            Val(txtSocio4.Text) = Val(txtSocio2.Text) Or _
            Val(txtSocio4.Text) = Val(txtSocio3.Text) Or _
            Val(txtSocio4.Text) = Val(txtSocio5.Text) Then
            MsgBox "Hijo " + lblSocio4 + " Ya Tiene Asignación", vbExclamation
            txtSocio4.Text = ""
            lblCodigo4.Caption = ""
            lblIns4.Caption = ""
            lblSocio4.Caption = ""
            txtEstado4.Text = ""
            Exit Sub
         End If
         
         aa = Leerado8("SELECT * FROM MAEASIGNADO WHERE CODHIJO = " + Str(wSoc) + " AND (LIN <> '04' OR CODSOCIO <> " + Str(wPad) + ")  ")
         If aa > 0 Then
            MsgBox "Hijo " + lblSocio4 + " Ya Tiene Asignación", vbExclamation
            txtSocio4.Text = ""
            lblCodigo4.Caption = ""
            lblIns4.Caption = ""
            lblSocio4.Caption = ""
            txtEstado4.Text = ""
            Exit Sub
         End If
         
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
   Dim KeyNew As Integer
   KeyNew = Asc(UCase(Chr(KeyCode)))
   
   Select Case KeyCode
   Case 38
        txtFecTop4.SetFocus
   Case 40
        txtObservac5.SetFocus
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio2.Show 1
        If xseleccion <> "" Then
           lblCodigo5.Caption = xselecSocio
           txtSocio5.Text = xseleccion
        End If
   Case 65 To 90
        xlista = "1"
        xseleccion = Chr(KeyNew)
        frmSelecSocio2.Show 1
        If xseleccion <> "" Then
           lblCodigo5.Caption = xselecSocio
           txtSocio5.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtSocio5_KeyPress(KeyAscii As Integer)
   Dim aa As Integer, wSoc As Integer, wPad As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtSocio5.Text)) <> 0 Then
         aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(lblCodigo5.Caption)) + " ")
         If aa = 0 Then
            MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
            txtSocio5.Text = ""
            Exit Sub
         End If
         If ADO8!e_socio <> "HIJ" Then
            MsgBox "Socio Asignado NO Es Hijo", vbExclamation
            txtSocio5.Text = ""
            Exit Sub
         End If
         txtSocio5.Text = ADO8!codigo
         lblCodigo5.Caption = ADO8!codsocio
         lblIns5.Caption = ADO8!ins
         lblSocio5.Caption = Trim(ADO8!nombre)
         txtEstado5.Text = "H"
         
         wPad = Val(txtCodSocio.Text)
         wSoc = Val(lblCodigo5.Caption)
         
         If Val(txtSocio5.Text) = Val(txtSocio1.Text) Or _
            Val(txtSocio5.Text) = Val(txtSocio2.Text) Or _
            Val(txtSocio5.Text) = Val(txtSocio3.Text) Or _
            Val(txtSocio5.Text) = Val(txtSocio4.Text) Then
            MsgBox "Hijo " + lblSocio5 + " Ya Tiene Asignación", vbExclamation
            txtSocio5.Text = ""
            lblCodigo5.Caption = ""
            lblIns5.Caption = ""
            lblSocio5.Caption = ""
            txtEstado5.Text = ""
            Exit Sub
         End If
         
         aa = Leerado8("SELECT * FROM MAEASIGNADO WHERE CODHIJO = " + Str(wSoc) + " AND (LIN <> '05' OR CODSOCIO <> " + Str(wPad) + ")  ")
         If aa > 0 Then
            MsgBox "Hijo " + lblSocio5 + " Ya Tiene Asignación", vbExclamation
            txtSocio5.Text = ""
            lblCodigo5.Caption = ""
            lblIns5.Caption = ""
            lblSocio5.Caption = ""
            txtEstado5.Text = ""
            Exit Sub
         End If
         
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

