VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMaeAsignadoOLD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Hijos"
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10695
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
      Left            =   11400
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
      Left            =   11400
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
      Left            =   11400
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
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   11175
      Begin VB.ComboBox cmbE_Socio 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMaeAsignadoOLD.frx":0000
         Left            =   6360
         List            =   "frmMaeAsignadoOLD.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   900
         Width           =   3255
      End
      Begin VB.ComboBox cmbGrado 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMaeAsignadoOLD.frx":0004
         Left            =   120
         List            =   "frmMaeAsignadoOLD.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   900
         Width           =   3015
      End
      Begin VB.ComboBox cmbTipCob 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMaeAsignadoOLD.frx":0008
         Left            =   3120
         List            =   "frmMaeAsignadoOLD.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   84
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
         TabIndex        =   82
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
         Height          =   285
         Left            =   345
         MaxLength       =   8
         TabIndex        =   43
         Top             =   1860
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
         Height          =   285
         Left            =   345
         MaxLength       =   8
         TabIndex        =   42
         Top             =   2415
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
         Height          =   285
         Left            =   345
         MaxLength       =   8
         TabIndex        =   41
         Top             =   2700
         Width           =   690
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
         Width           =   690
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
         Left            =   10080
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
         Left            =   10080
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
         Left            =   10080
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
         Left            =   10080
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
         Left            =   10080
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
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Estado de Socio"
         Height          =   195
         Index           =   16
         Left            =   6630
         TabIndex        =   89
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
         TabIndex        =   87
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
         TabIndex        =   86
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
         TabIndex        =   83
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lblCodigo2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   81
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblIns2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1920
         TabIndex        =   80
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblSocio2 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2280
         TabIndex        =   79
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
         TabIndex        =   78
         Top             =   1680
         Width           =   1515
      End
      Begin VB.Label lblCodigo1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   77
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label lblIns1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1920
         TabIndex        =   76
         Top             =   1860
         Width           =   375
      End
      Begin VB.Label lblSocio1 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2280
         TabIndex        =   75
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
         TabIndex        =   74
         Top             =   1860
         Width           =   180
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         Height          =   195
         Index           =   27
         Left            =   450
         TabIndex        =   73
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codofin"
         Height          =   195
         Index           =   28
         Left            =   1155
         TabIndex        =   72
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ins"
         Height          =   195
         Index           =   29
         Left            =   1920
         TabIndex        =   71
         Top             =   1680
         Width           =   210
      End
      Begin VB.Label lblCodigo3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   70
         Top             =   2415
         Width           =   855
      End
      Begin VB.Label lblIns3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1920
         TabIndex        =   69
         Top             =   2415
         Width           =   375
      End
      Begin VB.Label lblSocio3 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2280
         TabIndex        =   68
         Top             =   2415
         Width           =   4335
      End
      Begin VB.Label lblCodigo4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   67
         Top             =   2700
         Width           =   855
      End
      Begin VB.Label lblIns4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1920
         TabIndex        =   66
         Top             =   2700
         Width           =   375
      End
      Begin VB.Label lblSocio4 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2280
         TabIndex        =   65
         Top             =   2700
         Width           =   4335
      End
      Begin VB.Label lblCodigo5 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   64
         Top             =   2985
         Width           =   855
      End
      Begin VB.Label lblIns5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1920
         TabIndex        =   63
         Top             =   2985
         Width           =   375
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
         Left            =   10200
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
         Width           =   1095
      End
      Begin VB.Label lblEstado2 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   9000
         TabIndex        =   53
         Top             =   2145
         Width           =   1095
      End
      Begin VB.Label lblEstado3 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   9000
         TabIndex        =   52
         Top             =   2415
         Width           =   1095
      End
      Begin VB.Label lblEstado4 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   9000
         TabIndex        =   51
         Top             =   2700
         Width           =   1095
      End
      Begin VB.Label lblEstado5 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   9000
         TabIndex        =   50
         Top             =   2985
         Width           =   1095
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
      Height          =   3735
      Left            =   960
      TabIndex        =   0
      Top             =   3840
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6588
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
Attribute VB_Name = "frmMaeAsignadoOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Limpiar()

   txtCodSocio.Text = ""
   txtCodigo.Text = ""
   txtIns.Text = ""
   txtNumDoc.Text = ""
   lblCodSocio.Caption = ""
   cmbGrado.ListIndex = 0
   cmbTipCob.ListIndex = 0
   cmbE_Socio.ListIndex = 0
End Sub

Sub refrescar()
   If ADO1.BOF Then Exit Sub
   If ADO1.EOF Then Exit Sub
   txtCodigo.Text = ADO1!grado
   txtNombre.Text = IIf(IsNull(ADO1!nombre), "", ADO1!nombre)
   cmbGrupo.ListIndex = ADO1!gradogrupo - 1
End Sub

Private Sub cmdSalir_Click()
   Unload Me
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
   Db.Execute ("INSERT INTO TMP_MAEASIGNADOCAB " _
   & " (CODSOCIO, CODIGO, INS, NOMBRE, NUMDOC, E_SOCIO, GRADO, TIPCOB, CANT, USU) " _
   & " SELECT " _
   & "  M.CODSOCIO, S.CODIGO, S.INS, S.NOMBRE, S.NUMDOC, S.E_SOCIO, S.GRADO, " _
   & "  S.TIPCOB, COUNT(M.CODSOCIO), '" + wcodusu + "' " _
   & " FROM MAEASIGNADO AS M INNER JOIN MAESOCIO AS S " _
   & "   ON M.CODSOCIO = S.CODSOCIO ")
   Db.CommitTrans

   Dim aa As Integer

   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, CANT, NUMDOC, GRADO, TIPCOB, USU " _
                & " FROM TMP_MAEASIGNADOCAB " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO2

   DataGrid1.Columns(0).Width = 700
   DataGrid1.Columns(0).Alignment = dbgRight
   DataGrid1.Columns(0).Caption = "SOCIO"
    
   DataGrid1.Columns(1).Width = 900
   DataGrid1.Columns(1).Alignment = dbgRight
   DataGrid1.Columns(1).Caption = "CODIGO"
    
   DataGrid1.Columns(2).Width = 370
   DataGrid1.Columns(2).Alignment = dbgCenter
   DataGrid1.Columns(2).Caption = "INS"
    
   DataGrid1.Columns(3).Width = 3500
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE"
    
   DataGrid1.Columns(4).Width = 500
   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Caption = "ESTADO"
    
   DataGrid1.Columns(5).Width = 400
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).Caption = "CANT"
    
   DataGrid1.Columns(6).Visible = False
   DataGrid1.Columns(7).Visible = False
   DataGrid1.Columns(8).Visible = False
   DataGrid1.Columns(9).Visible = False
End Sub

Private Sub LlenaDet()
   If ADO2.EOF Or ADO2.BOF Then Exit Sub
   
   Dim aa As Integer
   
   txtCodSocio.Text = ADO2!codsocio
   txtCodigo.Text = ADO2!codigo
   txtIns.Text = ADO2!ins
   txtNumDoc.Text = ADO2!numdoc
   lblCodSocio.Caption = ADO2!nombre
   
   cmbGrado.ListIndex = BuscaGrado(ADO2!grado)
   cmbE_Socio.ListIndex = BuscaEsocio(ADO2!e_socio)
   cmbTipCob.ListIndex = BuscaTipCob(ADO2!tipcob)

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_MAEASIGNADODET WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_MAEASIGNADODET " _
   & " (CODSOCIO, LIN, SOCHIJO, CODHIJO, INSHIJO, NOMHIJO, ESTADO, OBSERV, FECTOP, USU) " _
   & " SELECT " _
   & "  M.CODSOCIO, M.LIN, M.SOCHIJO, S.CODIGO, S.INS, S.NOMBRE, M.ESTADO, M.OBSERV, " _
   & "  M.FECTOP, '" + wcodusu + "' " _
   & " FROM MAEASIGNADO AS M INNER JOIN MAESOCIO AS S " _
   & "   ON M.SOCHIJO = S.CODSOCIO " _
   & " WHERE S.CODSOCIO = '" + txtCodSocio.Text + "' ")
   Db.CommitTrans

   aa = Leerado3("SELECT * FROM TMP_MAEASIGNADODET " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY LIN ")
   If aa > 0 Then
      ADO3.MoveFirst
      Do While Not ADO3.EOF
   
         Select Case ADO3!lin
         Case "01"
              txtSocio1.Text = ADO2!sochijo
              lblCodigo1.Caption = ADO2!codhijo
              lblIns1.Caption = ADO2!ins
              lblSocio1.Caption = ADO2!nombre
              txtObservac1.Text = ADO2!observ
              txtEstado1.Text = ADO2!estado
              If IsDate(ADO2!fectop) Then
                 txtFecTop1.Text = ADO2!fectop
              Else
                 txtFecTop1.Text = "__/__/____"
              End If
         Case "02"
              txtSocio2.Text = ADO2!sochijo
              lblCodigo2.Caption = ADO2!codhijo
              lblIns2.Caption = ADO2!ins
              lblSocio2.Caption = ADO2!nombre
              txtObservac2.Text = ADO2!observ
              txtEstado2.Text = ADO2!estado
              If IsDate(ADO2!fectop) Then
                 txtFecTop2.Text = ADO2!fectop
              Else
                 txtFecTop2.Text = "__/__/____"
              End If
         Case "03"
              txtSocio3.Text = ADO2!sochijo
              lblCodigo3.Caption = ADO2!codhijo
              lblIns3.Caption = ADO2!ins
              lblSocio3.Caption = ADO2!nombre
              txtObservac3.Text = ADO2!observ
              txtEstado3.Text = ADO2!estado
              If IsDate(ADO2!fectop) Then
                 txtFecTop3.Text = ADO2!fectop
              Else
                 txtFecTop3.Text = "__/__/____"
              End If
         Case "04"
              txtSocio4.Text = ADO2!sochijo
              lblCodigo4.Caption = ADO2!codhijo
              lblIns4.Caption = ADO2!ins
              lblSocio4.Caption = ADO2!nombre
              txtObservac4.Text = ADO2!observ
              txtEstado4.Text = ADO2!estado
              If IsDate(ADO2!fectop) Then
                 txtFecTop4.Text = ADO2!fectop
              Else
                 txtFecTop4.Text = "__/__/____"
              End If
         Case "05"
              txtSocio5.Text = ADO2!sochijo
              lblCodigo5.Caption = ADO2!codhijo
              lblIns5.Caption = ADO2!ins
              lblSocio5.Caption = ADO2!nombre
              txtObservac5.Text = ADO2!observ
              txtEstado5.Text = ADO2!estado
              If IsDate(ADO2!fectop) Then
                 txtFecTop5.Text = ADO2!fectop
              Else
                 txtFecTop5.Text = "__/__/____"
              End If
         End Select
   
         ADO3.MoveNext
      Loop
   End If
  



End Sub
