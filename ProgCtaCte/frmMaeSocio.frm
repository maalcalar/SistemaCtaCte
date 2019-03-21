VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMaeSocio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Socios"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   15870
   Begin VB.CommandButton cmdAsignado 
      Caption         =   "Ver Asignados"
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
      Left            =   14640
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdFamiliar 
      Caption         =   "Ver &Familiares"
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
      Left            =   13440
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Caption         =   "Filtro x Nombre Pariente"
      Height          =   615
      Left            =   120
      TabIndex        =   99
      Top             =   8240
      Width           =   8175
      Begin VB.OptionButton optTodosPariente 
         Caption         =   "Mostrar Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optFiltroPariente 
         Caption         =   "Filtrar x Nombre"
         Height          =   255
         Left            =   1560
         TabIndex        =   101
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtFiltrarPariente 
         Height          =   285
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   100
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Filtro x DNI"
      Height          =   615
      Left            =   8400
      TabIndex        =   93
      Top             =   8240
      Width           =   4095
      Begin VB.TextBox txtFiltrarDni 
         Height          =   285
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   96
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optFiltroDni 
         Caption         =   "Filtrar x DNI"
         Height          =   255
         Left            =   1560
         TabIndex        =   95
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optTodosDni 
         Caption         =   "Mostrar Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Filtro x Codofin"
      Height          =   615
      Left            =   8400
      TabIndex        =   89
      Top             =   7640
      Width           =   4095
      Begin VB.OptionButton optTodosCodoFin 
         Caption         =   "Mostrar Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optFiltroCodoFin 
         Caption         =   "Filtrar x Codofin"
         Height          =   255
         Left            =   1560
         TabIndex        =   91
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtFiltrarCodofin 
         Height          =   285
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   90
         Top             =   240
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
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   13440
      TabIndex        =   20
      Top             =   2320
      Width           =   2295
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
         Left            =   1560
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
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
         Left            =   1080
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
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
         Left            =   600
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
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
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
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
      Left            =   13440
      TabIndex        =   18
      Top             =   1600
      Width           =   2295
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
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
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
      ForeColor       =   &H00C00000&
      Height          =   1600
      Left            =   13440
      TabIndex        =   12
      Top             =   0
      Width           =   2295
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
         Left            =   1200
         TabIndex        =   17
         Top             =   240
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
         Left            =   1200
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   680
         Width           =   975
      End
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
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
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   680
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1120
         Width           =   975
      End
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtro x Nombre"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   7640
      Width           =   8175
      Begin VB.TextBox txtFiltrar 
         Height          =   285
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   11
         Top             =   240
         Width           =   5055
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Filtrar x Nombre"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Mostrar Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame fraDetalles 
      Caption         =   "Detalles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3720
      Left            =   -120
      TabIndex        =   2
      Top             =   0
      Width           =   13455
      Begin VB.TextBox txtMesFall 
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
         Left            =   9720
         MaxLength       =   6
         TabIndex        =   127
         Top             =   2660
         Width           =   690
      End
      Begin VB.TextBox txtPromocion 
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
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   120
         Top             =   2660
         Width           =   1050
      End
      Begin VB.TextBox txtAnoDirec 
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
         Left            =   3930
         MaxLength       =   10
         TabIndex        =   118
         Top             =   2660
         Width           =   1050
      End
      Begin VB.TextBox txtDirectivo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   360
         MaxLength       =   3
         TabIndex        =   115
         Text            =   " "
         Top             =   2660
         Width           =   495
      End
      Begin VB.TextBox txtNumResoR 
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
         Left            =   9480
         MaxLength       =   10
         TabIndex        =   113
         Top             =   2220
         Width           =   1050
      End
      Begin VB.CheckBox chkCartaDieco 
         Caption         =   "Falta Carta Autorización"
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
         Left            =   5520
         TabIndex        =   108
         Top             =   3400
         Width           =   2775
      End
      Begin VB.CheckBox chkVip 
         Caption         =   "Socio VIP"
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
         Left            =   5640
         TabIndex        =   105
         Top             =   3100
         Width           =   1215
      End
      Begin VB.CheckBox chkFamiliar 
         Caption         =   "Tiene Familiares Registrados"
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
         Left            =   8640
         TabIndex        =   104
         Top             =   3100
         Width           =   2775
      End
      Begin VB.CheckBox chkAsignado 
         Caption         =   "Tiene Asignados Para Dscto"
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
         Left            =   8640
         TabIndex        =   103
         Top             =   3400
         Width           =   2775
      End
      Begin VB.TextBox txtUbi3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   98
         Text            =   " "
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtUbi2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   97
         Text            =   " "
         Top             =   1320
         Width           =   375
      End
      Begin VB.ComboBox cmbSituEsp 
         Height          =   315
         ItemData        =   "frmMaeSocio.frx":0000
         Left            =   3120
         List            =   "frmMaeSocio.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtObserva2 
         Height          =   285
         Left            =   360
         MaxLength       =   50
         TabIndex        =   86
         Top             =   3400
         Width           =   5055
      End
      Begin VB.TextBox txtObservac 
         Height          =   285
         Left            =   360
         MaxLength       =   50
         TabIndex        =   84
         Top             =   3100
         Width           =   5055
      End
      Begin VB.TextBox txtNumReso 
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
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   76
         Top             =   2220
         Width           =   1050
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
         Left            =   360
         MaxLength       =   20
         TabIndex        =   72
         Top             =   1760
         Width           =   2130
      End
      Begin VB.TextBox txtEMail2 
         Height          =   285
         Left            =   9120
         MaxLength       =   50
         TabIndex        =   70
         Top             =   1760
         Width           =   2055
      End
      Begin VB.ComboBox cmbTipCob 
         Height          =   315
         ItemData        =   "frmMaeSocio.frx":0004
         Left            =   11040
         List            =   "frmMaeSocio.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   840
         Width           =   1935
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
         Left            =   2640
         MaxLength       =   8
         TabIndex        =   62
         Top             =   2220
         Width           =   930
      End
      Begin VB.ComboBox cmbE_Socio 
         Height          =   315
         ItemData        =   "frmMaeSocio.frx":0008
         Left            =   11160
         List            =   "frmMaeSocio.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   1760
         Width           =   2175
      End
      Begin VB.TextBox txtRefer 
         Height          =   285
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   58
         Top             =   1760
         Width           =   3255
      End
      Begin VB.TextBox txteMail 
         Height          =   285
         Left            =   7200
         MaxLength       =   50
         TabIndex        =   56
         Top             =   1760
         Width           =   1935
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
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   54
         Top             =   1760
         Width           =   1410
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
         Left            =   10560
         MaxLength       =   20
         TabIndex        =   52
         Top             =   1320
         Width           =   2610
      End
      Begin VB.TextBox txtUbi1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   49
         Text            =   " "
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtDirec 
         Height          =   285
         Left            =   360
         MaxLength       =   50
         TabIndex        =   47
         Top             =   1320
         Width           =   5535
      End
      Begin VB.ComboBox cmbSexo 
         Height          =   315
         ItemData        =   "frmMaeSocio.frx":000C
         Left            =   9240
         List            =   "frmMaeSocio.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cmbECivil 
         Height          =   315
         ItemData        =   "frmMaeSocio.frx":0010
         Left            =   7560
         List            =   "frmMaeSocio.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox cmbSitu 
         Height          =   315
         ItemData        =   "frmMaeSocio.frx":0014
         Left            =   960
         List            =   "frmMaeSocio.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   840
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
         Height          =   285
         Left            =   12150
         MaxLength       =   10
         TabIndex        =   35
         Top             =   380
         Width           =   1170
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
         Left            =   11280
         MaxLength       =   10
         TabIndex        =   33
         Top             =   380
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
         Height          =   285
         Left            =   2400
         MaxLength       =   8
         TabIndex        =   31
         Top             =   380
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
         Height          =   285
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   28
         Top             =   380
         Width           =   330
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
         Height          =   285
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   27
         Top             =   380
         Width           =   930
      End
      Begin VB.ComboBox cmbGrado 
         Height          =   315
         ItemData        =   "frmMaeSocio.frx":0018
         Left            =   9000
         List            =   "frmMaeSocio.frx":001A
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   380
         Width           =   2295
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
         Left            =   360
         MaxLength       =   8
         TabIndex        =   4
         Top             =   380
         Width           =   690
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   3
         Top             =   380
         Width           =   5655
      End
      Begin MSMask.MaskEdBox txtFecNac 
         Height          =   285
         Left            =   5280
         TabIndex        =   39
         Top             =   840
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
         Left            =   6480
         TabIndex        =   41
         Top             =   840
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
         TabIndex        =   66
         Top             =   2220
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
         Left            =   1440
         TabIndex        =   68
         Top             =   2220
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
         Left            =   3600
         TabIndex        =   74
         Top             =   2220
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
         Left            =   5880
         TabIndex        =   78
         Top             =   2220
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
         Left            =   7080
         TabIndex        =   80
         Top             =   2220
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
         Left            =   8280
         TabIndex        =   82
         Top             =   2220
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
         Left            =   7080
         TabIndex        =   106
         Top             =   3100
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecCondo 
         Height          =   285
         Left            =   10680
         TabIndex        =   109
         Top             =   2220
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
         Left            =   11880
         TabIndex        =   111
         Top             =   2220
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecFall 
         Height          =   285
         Left            =   10440
         TabIndex        =   126
         Top             =   2660
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
         Caption         =   "Fallecimiento"
         Height          =   195
         Index           =   27
         Left            =   10200
         TabIndex        =   125
         Top             =   2480
         Width           =   915
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Left            =   11760
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblPromocion 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6120
         TabIndex        =   122
         Top             =   2660
         Width           =   3495
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Promoción"
         Height          =   195
         Index           =   26
         Left            =   5070
         TabIndex        =   121
         Top             =   2480
         Width           =   990
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Año Directivo"
         Height          =   195
         Index           =   25
         Left            =   3990
         TabIndex        =   119
         Top             =   2480
         Width           =   960
      End
      Begin VB.Label Label13 
         Caption         =   "Directivos??"
         Height          =   195
         Left            =   360
         TabIndex        =   117
         Top             =   2480
         Width           =   2535
      End
      Begin VB.Label lblDirectivo 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   840
         TabIndex        =   116
         Top             =   2660
         Width           =   3015
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nro.Resol.ReIng."
         Height          =   195
         Index           =   24
         Left            =   9360
         TabIndex        =   114
         Top             =   2040
         Width           =   1230
      End
      Begin VB.Label Label11 
         Caption         =   "Fec.Condonac."
         Height          =   210
         Left            =   10680
         TabIndex        =   112
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Fec.Amnistia"
         Height          =   210
         Left            =   11880
         TabIndex        =   110
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha VIP"
         Height          =   210
         Left            =   7200
         TabIndex        =   107
         Top             =   2920
         Width           =   855
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Situación Especial"
         Height          =   195
         Index           =   23
         Left            =   3375
         TabIndex        =   88
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   22
         Left            =   600
         TabIndex        =   85
         Top             =   2920
         Width           =   1065
      End
      Begin VB.Label Label7 
         Caption         =   "Fec.Reingreso"
         Height          =   210
         Left            =   8280
         TabIndex        =   83
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Fec.Expulsión"
         Height          =   210
         Left            =   7080
         TabIndex        =   81
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Fec.Exclusión"
         Height          =   210
         Left            =   5880
         TabIndex        =   79
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nro.Resol.Ing."
         Height          =   195
         Index           =   21
         Left            =   4815
         TabIndex        =   77
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Label Label4 
         Caption         =   "Fec.Resol.Ing"
         Height          =   210
         Left            =   3600
         TabIndex        =   75
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Celular"
         Height          =   195
         Index           =   20
         Left            =   720
         TabIndex        =   73
         Top             =   1580
         Width           =   480
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Correo Electrónico 2"
         Height          =   195
         Index           =   19
         Left            =   9360
         TabIndex        =   71
         Top             =   1580
         Width           =   1440
      End
      Begin VB.Label Label3 
         Caption         =   "Fec.Renuncia"
         Height          =   210
         Left            =   1440
         TabIndex        =   69
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Fecha Ing."
         Height          =   210
         Left            =   240
         TabIndex        =   67
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cobro"
         Height          =   195
         Index           =   18
         Left            =   11580
         TabIndex        =   65
         Top             =   660
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tomo Legajo"
         Height          =   195
         Index           =   17
         Left            =   2640
         TabIndex        =   63
         Top             =   2040
         Width           =   930
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Estado de Socio"
         Height          =   195
         Index           =   16
         Left            =   11430
         TabIndex        =   61
         Top             =   1580
         Width           =   1170
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Referencia"
         Height          =   195
         Index           =   15
         Left            =   3960
         TabIndex        =   59
         Top             =   1580
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Correo Electrónico"
         Height          =   195
         Index           =   14
         Left            =   7440
         TabIndex        =   57
         Top             =   1580
         Width           =   1305
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Celular 2"
         Height          =   195
         Index           =   13
         Left            =   2625
         TabIndex        =   55
         Top             =   1580
         Width           =   615
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Teléfonos"
         Height          =   195
         Index           =   12
         Left            =   10695
         TabIndex        =   53
         Top             =   1140
         Width           =   705
      End
      Begin VB.Label lblUbigeo 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   7080
         TabIndex        =   51
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label9 
         Caption         =   "Ubicación Geográfica"
         Height          =   195
         Left            =   6120
         TabIndex        =   50
         Top             =   1140
         Width           =   2535
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   195
         Index           =   11
         Left            =   840
         TabIndex        =   48
         Top             =   1140
         Width           =   675
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sexo"
         Height          =   195
         Index           =   10
         Left            =   9840
         TabIndex        =   46
         Top             =   660
         Width           =   360
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Estado Civil"
         Height          =   195
         Index           =   9
         Left            =   8040
         TabIndex        =   44
         Top             =   660
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Matrim"
         Height          =   210
         Left            =   6480
         TabIndex        =   42
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Fecha Nacim."
         Height          =   210
         Left            =   5280
         TabIndex        =   40
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Situación Policial"
         Height          =   195
         Index           =   8
         Left            =   1320
         TabIndex        =   38
         Top             =   660
         Width           =   1200
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Carnet PIP"
         Height          =   195
         Index           =   7
         Left            =   12360
         TabIndex        =   36
         Top             =   200
         Width           =   765
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Carnet PNP"
         Height          =   195
         Index           =   6
         Left            =   11160
         TabIndex        =   34
         Top             =   200
         Width           =   840
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "D.N.I."
         Height          =   195
         Index           =   5
         Left            =   2595
         TabIndex        =   32
         Top             =   200
         Width           =   420
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ins"
         Height          =   195
         Index           =   4
         Left            =   2040
         TabIndex        =   30
         Top             =   200
         Width           =   210
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codofin"
         Height          =   195
         Index           =   1
         Left            =   1155
         TabIndex        =   29
         Top             =   200
         Width           =   540
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Grado"
         Height          =   195
         Index           =   2
         Left            =   9120
         TabIndex        =   26
         Top             =   200
         Width           =   915
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Apellidos y Nombres de Asociado"
         Height          =   195
         Index           =   3
         Left            =   3840
         TabIndex        =   6
         Top             =   200
         Width           =   3915
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   465
         TabIndex        =   5
         Top             =   200
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
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
      Left            =   14160
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   8280
      Width           =   1095
   End
   Begin Crystal.CrystalReport Crys1 
      Left            =   12720
      Top             =   7800
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3840
      Left            =   0
      TabIndex        =   7
      Top             =   3800
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   6773
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "TABLA DE SOCIOS"
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
            LCID            =   3082
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
            LCID            =   3082
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
      Left            =   1800
      TabIndex        =   1
      Top             =   3825
      Width           =   9375
   End
End
Attribute VB_Name = "frmMaeSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACCION As Byte
Dim marca As Variant, wcia As String
    
Sub Limpiar()
   txtCodigo.Text = ""
   txtCodofin.Text = ""
   txtIns.Text = ""
   txtNumdoc.Text = ""
   txtNombre.Text = ""
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
   txteMail.Text = ""
   txtEMail2.Text = ""
   txtRefer.Text = ""
   txtTomo.Text = ""
   txtFecIng.Text = "__/__/____"
   txtFecRenu.Text = "__/__/____"
   txtNumReso.Text = ""
   txtNumResoR.Text = ""
   txtFecReso.Text = "__/__/____"
   txtFecExclu.Text = "__/__/____"
   txtFecExpul.Text = "__/__/____"
   txtFecRein.Text = "__/__/____"
   txtFecAmnis.Text = "__/__/____"
   txtFecCondo.Text = "__/__/____"
   txtFecFall.Text = "__/__/____"
   txtMesFall.Text = ""
   
   txtDirectivo.Text = ""
   txtAnoDirec.Text = ""
   txtPromocion.Text = ""
   txtObservac.Text = ""
   txtObserva2.Text = ""
   
   cmbGrado.ListIndex = 0
   cmbECivil.ListIndex = 0
   cmbSitu.ListIndex = 0
   cmbSituEsp.ListIndex = 0
   cmbSexo.ListIndex = 0
   cmbE_Socio.ListIndex = 0
   cmbTipCob.ListIndex = 0

   chkCartaDieco.Value = vbUnchecked
   chkVip.Value = vbUnchecked
   txtFecVip.Text = "__/__/____"

   chkAsignado.Value = vbUnchecked
   chkFamiliar.Value = vbUnchecked
End Sub

Sub refrescar()
   If ADO1.BOF Then Exit Sub
   If ADO1.EOF Then Exit Sub
   
   txtCodigo.Text = ADO1!codsocio
   txtCodofin.Text = ADO1!codigo
   txtIns.Text = ADO1!ins
   txtNumdoc.Text = ADO1!numdoc
   txtNombre.Text = IIf(IsNull(ADO1!nombre), "", ADO1!nombre)
   txtCarnetPNP.Text = IIf(IsNull(ADO1!carnetpnp), "", ADO1!carnetpnp)
   txtCarnetPIP.Text = IIf(IsNull(ADO1!carnetpip), "", ADO1!carnetpip)
   txtDirec.Text = IIf(IsNull(ADO1!direc), "", ADO1!direc)
   txtUbi1.Text = IIf(IsNull(ADO1!ubigeo), "", Mid(ADO1!ubigeo, 1, 2))
   txtUbi2.Text = IIf(IsNull(ADO1!ubigeo), "", Mid(ADO1!ubigeo, 3, 2))
   txtUbi3.Text = IIf(IsNull(ADO1!ubigeo), "", Mid(ADO1!ubigeo, 5, 2))
   txtTelefono.Text = IIf(IsNull(ADO1!telefono), "", ADO1!telefono)
   txtTelefon2.Text = IIf(IsNull(ADO1!telefon2), "", ADO1!telefon2)
   txtCelular.Text = IIf(IsNull(ADO1!celular), "", ADO1!celular)
   txteMail.Text = IIf(IsNull(ADO1!email), "", ADO1!email)
   txtEMail2.Text = IIf(IsNull(ADO1!email2), "", ADO1!email2)
   txtRefer.Text = IIf(IsNull(ADO1!refer), "", ADO1!refer)
   txtTomo.Text = IIf(IsNull(ADO1!tomo), "", ADO1!tomo)
   
   If IsDate(ADO1!fecnac) Then
      txtFecNac.Text = Format(ADO1!fecnac, "dd/mm/yyyy")
   Else
      txtFecNac.Text = "__/__/____"
   End If
   If IsDate(ADO1!fecmat) Then
      txtFecMat.Text = Format(ADO1!fecmat, "dd/mm/yyyy")
   Else
      txtFecMat.Text = "__/__/____"
   End If
   If IsDate(ADO1!fecing) Then
      txtFecIng.Text = Format(ADO1!fecing, "dd/mm/yyyy")
   Else
      txtFecIng.Text = "__/__/____"
   End If
   If IsDate(ADO1!fecrenu) Then
      txtFecRenu.Text = Format(ADO1!fecrenu, "dd/mm/yyyy")
   Else
      txtFecRenu.Text = "__/__/____"
   End If
   If IsDate(ADO1!freso_ing) Then
      txtFecReso.Text = Format(ADO1!freso_ing, "dd/mm/yyyy")
   Else
      txtFecReso.Text = "__/__/____"
   End If
   If IsDate(ADO1!fecrein) Then
      txtFecRein.Text = Format(ADO1!fecrein, "dd/mm/yyyy")
   Else
      txtFecRein.Text = "__/__/____"
   End If
   If IsDate(ADO1!fecexpul) Then
      txtFecExpul.Text = Format(ADO1!fecexpul, "dd/mm/yyyy")
   Else
      txtFecExpul.Text = "__/__/____"
   End If
   If IsDate(ADO1!fecexclu) Then
      txtFecExclu.Text = Format(ADO1!fecexclu, "dd/mm/yyyy")
   Else
      txtFecExclu.Text = "__/__/____"
   End If
   If IsDate(ADO1!fecrein) Then
      txtFecRein.Text = Format(ADO1!fecrein, "dd/mm/yyyy")
   Else
      txtFecRein.Text = "__/__/____"
   End If
   If IsDate(ADO1!fecvip) Then
      txtFecVip.Text = Format(ADO1!fecvip, "dd/mm/yyyy")
   Else
      txtFecVip.Text = "__/__/____"
   End If
   If IsDate(ADO1!fecamnis) Then
      txtFecAmnis.Text = Format(ADO1!fecamnis, "dd/mm/yyyy")
   Else
      txtFecAmnis.Text = "__/__/____"
   End If
   If IsDate(ADO1!feccondo) Then
      txtFecCondo.Text = Format(ADO1!feccondo, "dd/mm/yyyy")
   Else
      txtFecCondo.Text = "__/__/____"
   End If
   If IsDate(ADO1!fecfall) Then
      txtFecFall.Text = Format(ADO1!fecfall, "dd/mm/yyyy")
   Else
      txtFecFall.Text = "__/__/____"
   End If
   txtMesFall.Text = IIf(IsNull(ADO1!mesfall), "", ADO1!mesfall)
   
   txtNumReso.Text = IIf(IsNull(ADO1!nreso_ing), "", ADO1!nreso_ing)
   txtNumResoR.Text = IIf(IsNull(ADO1!nreso_reing), "", ADO1!nreso_reing)
   txtObservac.Text = IIf(IsNull(ADO1!observac), "", ADO1!observac)
   txtObserva2.Text = IIf(IsNull(ADO1!observa2), "", ADO1!observa2)
   
   txtDirectivo.Text = IIf(IsNull(ADO1!directivo), "", ADO1!directivo)
   txtAnoDirec.Text = IIf(IsNull(ADO1!anodirec), "", ADO1!anodirec)
   txtPromocion.Text = IIf(IsNull(ADO1!promocion), "", ADO1!promocion)
   
   cmbGrado.ListIndex = BuscaGrado(ADO1!grado)
   cmbSitu.ListIndex = BuscaSitu(ADO1!situ)
   cmbSituEsp.ListIndex = BuscaSituEsp(ADO1!situesp)
   cmbECivil.ListIndex = BuscaECivil(ADO1!ecivil)
   cmbSexo.ListIndex = BuscaSexo(ADO1!sexo)
   cmbE_Socio.ListIndex = BuscaEsocio(ADO1!e_socio)
   cmbTipCob.ListIndex = BuscaTipCob(ADO1!tipcob)
   
   If ADO1!asignado = True Then
      cmdAsignado.Enabled = True
      chkAsignado.Value = vbChecked
   Else
      cmdAsignado.Enabled = False
      chkAsignado.Value = vbUnchecked
   End If
   
   If ADO1!familiar = True Then
      cmdFamiliar.Enabled = True
      chkFamiliar.Value = vbChecked
   Else
      cmdFamiliar.Enabled = False
      chkFamiliar.Value = vbUnchecked
   End If
   
   If ADO1!vip = True Then
      chkVip.Value = vbChecked
   Else
      chkVip.Value = vbUnchecked
      txtFecVip.Text = "__/__/____"
   End If

   If ADO1!cartadieco = True Then
      chkCartaDieco.Value = vbChecked
   Else
      chkCartaDieco.Value = vbUnchecked
   End If
End Sub

Sub grabar()
   On Error GoTo err
   
   Dim aa As Integer, _
       wSoc As Integer, wCod As Long, wIns As Integer, wNom As String, wGrado As Integer, _
       wCarPNP As Long, wCarPIP As String, wDir As String, wUbigeo As String, _
       wTelefo As String, wTelef2 As String, wCelula As String, weMail As String, wEMail2 As String, wRefer As String, _
       wTomo As Integer, wSitua As Integer, wSituEsp As Integer, wSexo As String, wEcivi As String, wESoci As String, _
       wFecNac As Date, wFecMat As Date, wTipCob As String, wFecIng As Date, _
       wFecRenu As Date, wNumDoc As String, wFecReso As Date, wNumReso As String, wNumResoR As String, _
       wNomGra As String, wNomSitu As String, wNomEso As String, wNomCob As String, _
       wFecRein As Date, wFecExclu As Date, wFecExpul As Date, wFecVip As Date, _
       wObservac As String, wObserva2 As String, wFecAmnis As Date, wFecCondo As Date, wESoc2 As String, _
       wswActiva As Boolean, wMon As String, wApo As Currency, wDirectivo As String, wAnoDirec As String, wPromocion As String, _
       wMesIni As String, wMesFecIng As String, wAnoFecIng As String, wMesFall As String, wFecFall As Date
   
   wSoc = Val(txtCodigo.Text)
   wCod = Val(txtCodofin.Text)
   wIns = Val(txtIns.Text)
   wNom = txtNombre.Text
   wNumDoc = txtNumdoc.Text
   wswActiva = False
   wESoc2 = ""
   wMon = ""
   wApo = 0
   
   If Len(Trim(wNumDoc)) = 0 Then
      MsgBox "DNI Del Asociado En Blanco", vbExclamation
      txtNumdoc.Text = ""
      Exit Sub
   End If
   If Len(Trim(wNom)) = 0 Then
      MsgBox "Nombre Del Asociado En Blanco", vbExclamation
      txtNombre.Text = ""
      Exit Sub
   End If
   If Len(Trim(wCod)) = 0 Or Len(Trim(wIns)) = 0 Then
      MsgBox "Codofin En Blanco", vbExclamation
      txtCodigo.Text = ""
      Exit Sub
   End If
   
' Valida Nombre
   aa = Leerado8("SELECT * FROM MAESOCIO WHERE NOMBRE = '" + wNom + "' ")
   If aa > 0 Then
      If ADO8!codsocio <> wSoc Or ADO8!codigo <> wCod Or ADO8!ins <> wIns Then
         MsgBox "Nombre de Asociado Ya Existe", vbExclamation
         Exit Sub
      End If
   End If

' Valida DNI
   aa = Leerado8("SELECT * FROM MAESOCIO WHERE NUMDOC = '" + wNumDoc + "' ")
   If aa > 0 Then
      If ADO8!codsocio <> wSoc Or ADO8!codigo <> wCod Or ADO8!ins <> wIns Then
         MsgBox "DNI de Asociado Ya Existe", vbExclamation
         Exit Sub
      End If
   End If
' Valida CODOFIN
   aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(wCod) + " AND INS = " + Str(wIns) + " ")
   If aa > 0 Then
      If ADO8!codsocio <> wSoc Then
         MsgBox "CODOFIN de Asociado Ya Existe", vbExclamation
         Exit Sub
      End If
   End If
   If Not IsDate(txtFecIng.Text) Then
      MsgBox "Fecha de Ingreso Es Invalida", vbExclamation
      txtFecIng.SetFocus
      Exit Sub
   End If
   
   If IsDate(txtFecIng.Text) Then
      wFecIng = Format(txtFecIng.Text, "dd/mm/yyyy")
      wMesFecIng = Format(Month(wFecIng), "00")
      wAnoFecIng = Format(Year(wFecIng), "0000")
   End If
   wCarPNP = Val(txtCarnetPNP.Text)
   wCarPIP = txtCarnetPIP.Text
   wDir = txtDirec.Text
   wUbigeo = txtUbi1.Text + txtUbi2.Text + txtUbi3.Text
   wRefer = txtRefer.Text
   wTelefo = txtTelefono.Text
   wTelef2 = txtTelefon2.Text
   wCelula = txtCelular.Text
   weMail = txteMail.Text
   wEMail2 = txteMail.Text
   wRefer = txtRefer.Text
   wTomo = Val(txtTomo.Text)
   wNomGra = cmbGrado.Text
   wNomSitu = cmbSitu.Text
   wNomEso = cmbE_Socio.Text
   wNomCob = cmbTipCob.Text
   wNumReso = txtNumReso.Text
   wNumResoR = txtNumResoR.Text
   wObservac = txtObservac.Text
   wObserva2 = txtObserva2.Text
   wDirectivo = txtDirectivo.Text
   wAnoDirec = txtAnoDirec.Text
   wPromocion = txtPromocion.Text
   wMesFall = txtMesFall.Text
   
   If IsDate(txtFecFall.Text) Then
      wFecFall = Format(txtFecFall.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFecNac.Text) Then
      wFecNac = Format(txtFecNac.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFecMat.Text) Then
      wFecMat = Format(txtFecMat.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFecRenu.Text) Then
      wFecRenu = Format(txtFecRenu.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFecReso.Text) Then
      wFecReso = Format(txtFecReso.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFecRein.Text) Then
      wFecRein = Format(txtFecRein.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFecExpul.Text) Then
      wFecExpul = Format(txtFecExpul.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFecExclu.Text) Then
      wFecExclu = Format(txtFecExclu.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFecVip.Text) Then
      wFecVip = Format(txtFecVip.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFecAmnis.Text) Then
      wFecAmnis = Format(txtFecAmnis.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFecCondo.Text) Then
      wFecCondo = Format(txtFecCondo.Text, "dd/mm/yyyy")
   End If
   
   wGrado = BuscaCodGrado(cmbGrado.List(cmbGrado.ListIndex))
   wSitua = BuscaCodSitu(cmbSitu.List(cmbSitu.ListIndex))
   wSituEsp = BuscaCodSituEsp(cmbSituEsp.List(cmbSituEsp.ListIndex))
   wEcivi = BuscaCodECivil(cmbECivil.List(cmbECivil.ListIndex))
   wSexo = BuscaCodSexo(cmbSexo.List(cmbSexo.ListIndex))
   wESoci = BuscaCodEsocio(cmbE_Socio.List(cmbE_Socio.ListIndex))
   wTipCob = BuscaCodTipCob(cmbTipCob.List(cmbTipCob.ListIndex))
   
   aa = Leerado8("SELECT * FROM MAESOCIO " _
                & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " ")
   If aa > 0 Then
      wESoc2 = ADO8!e_socio
   End If
   Set ADO8 = Nothing
   
   
   Dim wSdoOld As Currency, wAdelan As Currency
   
   wSdoOld = 0: wAdelan = 0
   
   If (wESoc2 = "EXC" Or wESoc2 = "EXP" Or wESoc2 = "FAL" Or wESoc2 = "REN" Or wESoc2 = "SEP") And _
      (wESoci = "TIT" Or wESoci = "HIJ" Or wESoci = "NIE" Or wESoci = "HER" Or wESoci = "CIV" Or _
       wESoci = "CI2" Or wESoci = "TRA" Or wESoci = "VIU" Or wESoci = "ADH") Then
       
       aa = Leerado8("SELECT * FROM ZZZ_MAESTRO_INICIAL " _
                    & " WHERE CODIGO = " + Str(wCod) + " AND " _
                    & "          INS = " + Str(wIns) + " ")
       If aa > 0 Then
          wSdoOld = IIf(IsNull(ADO8!deuda_pt2), 0, ADO8!deuda_pt2)
          wAdelan = IIf(IsNull(ADO8!adelanto), 0, ADO8!adelanto)
       End If
       Set ADO8 = Nothing
       
       If MsgBox("Desea Activar El Saldo Que Habia Al 21 Oct 2017 " + vbNewLine + _
                 IIf(wSdoOld > 0, "Saldo Anterior->" + Format(wSdoOld, "#####0.00"), _
                                  "Adelanto->" + Format(wAdelan, "#####0.00")) + " ??", _
                 vbQuestion + vbYesNo, "Deudas Anteriores") = vbYes Then
          wswActiva = True
      End If
   End If
   
   If Len(Trim(wCod)) = 0 Then
      MsgBox "Codigo En Blanco", vbExclamation
      Exit Sub
   End If
   
   If Len(Trim(wNom)) = 0 Then
      MsgBox "Nombre En Blanco", vbExclamation
      Exit Sub
   End If
   
   aa = Leerado8("SELECT * FROM MAESOCIO " _
                & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " ")
   If aa = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO MAESOCIO " _
      & " (CODSOCIO, CODIGO, INS, NUMDOC, CARNETPNP, CARNETPIP, NOMBRE, GRADO, " _
      & "  SITU, SITUESP, SEXO, ECIVIL, DIREC, UBIGEO, TELEFONO, TELEFON2, CELULAR, EMAIL, EMAIL2, REFER, " _
      & "  E_SOCIO, TOMO, TIPCOB, NRESO_ING, OBSERVAC, OBSERVA2, NRESO_REING, DIRECTIVO, ANODIREC, PROMOCION, " _
      & "  MESFALL ) " _
      & " VALUES " _
      & " (" + Str(Val(wSoc)) + ", " + Str(wCod) + ", " + Str(wIns) + ", '" + wNumDoc + "', " + Str(wCarPNP) + ", " _
      & "  '" + wCarPIP + "', '" + wNom + "', " + Str(wGrado) + ", " + Str(wSitua) + ", " + Str(wSituEsp) + ", " _
      & "  '" + wSexo + "', '" + wEcivi + "', '" + wDir + "', '" + wUbigeo + "', " _
      & "  '" + wTelefo + "', '" + wTelef2 + "', '" + wCelula + "', '" + weMail + "', " _
      & "  '" + wEMail2 + "', '" + wRefer + "', " _
      & "  '" + wESoci + "', " + Str(wTomo) + ", '" + wTipCob + "', '" + wNumReso + "', " _
      & "  '" + wObservac + "', '" + wObserva2 + "', '" + wNumResoR + "', '" + wDirectivo + "', " _
      & "  '" + wAnoDirec + "', '" + wPromocion + "', '" + wMesFall + "' ) ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET NOMBRE = '" + wNom + "', CODIGO = " + Str(wCod) + ", INS = " + Str(wIns) + ", NUMDOC = '" + wNumDoc + "', " _
      & "     CARNETPNP = " + Str(wCarPNP) + ", CARNETPIP = '" + wCarPIP + "', " _
      & "     GRADO = " + Str(wGrado) + ", SITU = " + Str(wSitua) + ", SITUESP = " + Str(wSituEsp) + ", SEXO = '" + wSexo + "', " _
      & "     ECIVIL = '" + wEcivi + "', DIREC = '" + wDir + "', UBIGEO = '" + wUbigeo + "', " _
      & "     TELEFONO = '" + wTelefo + "', TELEFON2 = '" + wTelef2 + "', " _
      & "     CELULAR = '" + wCelula + "', EMAIL = '" + weMail + "', " _
      & "     EMAIL2 = '" + wEMail2 + "', " _
      & "     REFER = '" + wRefer + "', E_SOCIO = '" + wESoci + "', TOMO = " + Str(wTomo) + ", " _
      & "     TIPCOB = '" + wTipCob + "', NRESO_ING = '" + wNumReso + "', " _
      & "     OBSERVAC = '" + wObservac + "', OBSERVA2 = '" + wObserva2 + "', NRESO_REING = '" + wNumResoR + "', " _
      & "     DIRECTIVO = '" + wDirectivo + "', ANODIREC = '" + wAnoDirec + "', PROMOCION = '" + wPromocion + "', " _
      & "     MESFALL = '" + wMesFall + "' " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " ")
      Db.CommitTrans
   End If
   Set ADO8 = Nothing
   
   If IsDate(txtFecNac.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECNAC = '" + Format(wFecNac, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECNAC = NULL " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecMat.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECMAT = '" + Format(wFecMat, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECMAT = NULL " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecIng.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECING = '" + Format(wFecIng, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECING = NULL " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecRenu.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECRENU = '" + Format(wFecRenu, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECRENU = NULL " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecReso.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FRESO_ING = '" + Format(wFecReso, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FRESO_ING = NULL " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecRein.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECREIN = '" + Format(wFecRein, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECREIN = NULL " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecExpul.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECEXPUL = '" + Format(wFecExpul, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECEXPUL = NULL " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecExclu.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECEXCLU = '" + Format(wFecExclu, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECEXCLU = NULL " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecVip.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECVIP = '" + Format(wFecVip, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECVIP = NULL " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecAmnis.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECAMNIS = '" + Format(wFecAmnis, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECAMNIS = NULL " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecCondo.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECCONDO = '" + Format(wFecCondo, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECCONDO = NULL " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecFall.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECFALL = '" + Format(wFecFall, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FECFALL = NULL " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If
   
   If chkAsignado.Value = vbChecked Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET ASIGNADO = 1 " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET ASIGNADO = 0 " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If
   
   If chkFamiliar.Value = vbChecked Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FAMILIAR = 1 " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET FAMILIAR = 0 " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If
   
   If chkVip.Value = vbChecked Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET VIP = 1 " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET VIP = 0 " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If
   
   If chkCartaDieco.Value = vbChecked Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET CARTADIECO = 1 " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET CARTADIECO = 0 " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   End If
   
   aa = Leerado8("SELECT * FROM TMP_MAESOCIO " _
                & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
                & "       USU = '" + wcodusu + "' ")
   If aa = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO TMP_MAESOCIO " _
      & " (CODSOCIO, CODIGO, INS, NUMDOC, CARNETPNP, CARNETPIP, NOMBRE, GRADO, " _
      & "  SITU, SITUESP, SEXO, ECIVIL, DIREC, UBIGEO, TELEFONO, TELEFON2, CELULAR, EMAIL, EMAIL2, REFER, " _
      & "  E_SOCIO, TOMO, TIPCOB, NOMGRA, NOMSITU, NOMESO, NOMCOB, NRESO_ING, " _
      & "  OBSERVAC, OBSERVA2, NRESO_REING, DIRECTIVO, ANODIREC, PROMOCION, MESFALL, USU ) " _
      & " VALUES " _
      & " (" + Str(Val(wSoc)) + ", " + Str(wCod) + ", " + Str(wIns) + ", '" + wNumDoc + "', " + Str(wCarPNP) + ", " _
      & "  '" + wCarPIP + "', '" + wNom + "', " + Str(wGrado) + ", " + Str(wSitua) + ", " + Str(wSituEsp) + ", " _
      & "  '" + wSexo + "', '" + wEcivi + "', '" + wDir + "', '" + wUbigeo + "', " _
      & "  '" + wTelefo + "', '" + wTelef2 + "', '" + wCelula + "', '" + weMail + "', " _
      & "  '" + wEMail2 + "', '" + wRefer + "', " _
      & "  '" + wESoci + "', " + Str(wTomo) + ", '" + wTipCob + "', " _
      & "  '" + wNomGra + "', '" + wNomSitu + "', '" + wNomEso + "', '" + wNomCob + "', " _
      & "  '" + wNumReso + "', '" + wObservac + "', '" + wObserva2 + "', '" + wNumResoR + "', " _
      & "  '" + wDirectivo + "', '" + wAnoDirec + "', '" + wPromocion + "', '" + wMesFall + "', '" + wcodusu + "' ) ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET NOMBRE = '" + wNom + "', CODIGO = " + Str(wCod) + ", INS = " + Str(wIns) + ", NUMDOC = '" + wNumDoc + "', " _
      & "     CARNETPNP = " + Str(wCarPNP) + ", CARNETPIP = '" + wCarPIP + "', " _
      & "     GRADO = " + Str(wGrado) + ", SITU = " + Str(wSitua) + ", SITUESP = " + Str(wSituEsp) + ", SEXO = '" + wSexo + "', " _
      & "     ECIVIL = '" + wEcivi + "', DIREC = '" + wDir + "', UBIGEO = '" + wUbigeo + "', " _
      & "     TELEFONO = '" + wTelefo + "', TELEFON2 = '" + wTelef2 + "', CELULAR = '" + wCelula + "', EMAIL = '" + weMail + "', " _
      & "     EMAIL2 = '" + wEMail2 + "', " _
      & "     REFER = '" + wRefer + "', E_SOCIO = '" + wESoci + "', TOMO = " + Str(wTomo) + ", " _
      & "     TIPCOB = '" + wTipCob + "', " _
      & "     NOMGRA = '" + wNomGra + "', NOMSITU = '" + wNomSitu + "', " _
      & "     NOMESO = '" + wNomEso + "', NOMCOB  = '" + wNomCob + "', NRESO_ING = '" + wNumReso + "', " _
      & "     OBSERVAC = '" + wObservac + "', OBSERVA2 = '" + wObserva2 + "', NRESO_REING = '" + wNumResoR + "', " _
      & "     DIRECTIVO = '" + wDirectivo + "', ANODIREC = '" + wAnoDirec + "', PROMOCION = '" + wPromocion + "', " _
      & "     MESFALL = '" + wMesFall + "' " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   Set ADO8 = Nothing
   
   If IsDate(txtFecNac.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECNAC = '" + Format(wFecNac, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECNAC = NULL " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecMat.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECMAT = '" + Format(wFecMat, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECMAT = NULL " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecIng.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECING = '" + Format(wFecIng, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECING = NULL " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecRenu.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECRENU = '" + Format(wFecRenu, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECRENU = NULL " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecReso.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FRESO_ING = '" + Format(wFecReso, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FRESO_ING = NULL " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecRein.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECREIN = '" + Format(wFecRein, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECREIN = NULL " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecExpul.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECEXPUL = '" + Format(wFecExpul, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECEXPUL = NULL " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecExclu.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECEXCLU = '" + Format(wFecExclu, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECEXCLU = NULL " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecVip.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECVIP = '" + Format(wFecVip, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECVIP = NULL " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecAmnis.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECAMNIS = '" + Format(wFecAmnis, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECAMNIS = NULL " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecCondo.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECCONDO = '" + Format(wFecCondo, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECCONDO = NULL " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFecFall.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECFALL = '" + Format(wFecFall, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FECFALL = NULL " _
      & " WHERE CODSOCIO = " + Str(Val(wSoc)) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   If chkAsignado.Value = vbChecked Then
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET ASIGNADO = 1 " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET ASIGNADO = 0 " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   If chkFamiliar.Value = vbChecked Then
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FAMILIAR = 1 " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET FAMILIAR = 0 " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   If chkVip.Value = vbChecked Then
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET VIP = 1 " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET VIP = 0 " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   If chkCartaDieco.Value = vbChecked Then
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET CARTADIECO = 1 " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAESOCIO " _
      & " SET CARTADIECO = 0 " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   If (wESoci = "TIT" Or wESoci = "HIJ" Or wESoci = "NIE" Or wESoci = "HER" Or wESoci = "CIV" Or _
       wESoci = "CI1" Or wESoci = "CI2" Or wESoci = "TRA" Or wESoci = "VIU" Or wESoci = "ADH") Then
   
       Call CreateAporteAnoMes(wSoc, wAnoFecIng, wFecIng)
   End If
   
   If wswActiva = True Then
      aa = Leerado8("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + wESoci + "' ")
      If aa > 0 Then
         wMon = ADO8!moneda
         wApo = ADO8!aporte
      End If
      Set ADO8 = Nothing
      
      aa = Leerado8("SELECT * FROM CTASXCAB " _
                & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                & "            MES = '2017/09' AND " _
                & "       CONCEPTO = '01' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO CTASXCAB " _
         & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, " _
         & "  CARGOS, ABONOS, SDONEW ) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", '2017/09', '01', '" + wESoci + "', '" + wMon + "', " _
         & "  " + Str(wSdoOld) + ", " + Str(wAdelan) + ", 0)  ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE CTASXCAB " _
         & " SET CARGOS = " + Str(wSdoOld) + ", " _
         & "     ABONOS = " + Str(wAdelan) + " " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "            MES = '2017/09' AND " _
         & "       CONCEPTO = '01' ")
         Db.CommitTrans
      End If
      Set ADO8 = Nothing
   
      aa = Leerado8("SELECT * FROM CTASXDET " _
                & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                & "            MES = '2017/09' AND " _
                & "       CONCEPTO = '01' AND " _
                & "         TIPCOB = '00' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO CTASXDET " _
         & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
         & "  TIPCAM, DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW, OBS ) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", '2017/09', '01', '00', '', '', '', '1', '01/09/2017', " _
         & "  0, 0, 0, 0, " + Str(wSdoOld) + ", " + Str(wAdelan) + ", 0, '')  ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE CTASXDET " _
         & " SET CARGOS = " + Str(wSdoOld) + ", " _
         & "     ABONOS = " + Str(wAdelan) + " " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "            MES = '2017/09' AND " _
         & "       CONCEPTO = '01' AND " _
         & "         TIPCOB = '00' ")
         Db.CommitTrans
      End If
   
      Call ActualizaSaldos(wSoc, "2017/09", "01")
   
   End If
   
   Call abrirEVENTO
   
   aa = LeeradoEvento1("SELECT * FROM MAESOCIO WHERE CODIGO = '" + Format(wCod, "00000000") + "' AND INS = '" + Format(wIns, "0") + "' ")
   If aa = 0 Then
      DbEvento.BeginTrans
      DbEvento.Execute ("INSERT INTO MAESOCIO " _
      & " (INS, CODIGO, NOMBRE, GRADO, DNI, DIREC, DISTRITO, UBIGEO, TELFS, REFERE, EMAIL, " _
      & "  ESPOSA, MADRE, PADRE, HIJO01, HIJO02, HIJO03, HIJO04, HIJO05, HIJO06, HIJO07, " _
      & "  RESOL, E_SOCIO ) " _
      & " VALUES " _
      & " ('" + Format(wIns, "0") + "', '" + Format(wCod, "00000000") + "', '" + Left(wNom, 50) + "', " + Str(wGrado) + ", " _
      & "  '" + wNumDoc + "', '" + wDir + "', '', '" + wUbigeo + "', '" + wTelefo + "', '" + wRefer + "', " _
      & "  '" + weMail + "', '', '', '', '', '', '', '', '', '', '', '" + wNumReso + "', '" + wESoci + "' ) ")
      DbEvento.CommitTrans
   Else
      DbEvento.BeginTrans
      DbEvento.Execute ("UPDATE MAESOCIO " _
      & " SET NOMBRE = '" + wNom + "', DNI = '" + wNumDoc + "', DIREC = '" + Trim(wDir) + "', " _
      & "     DISTRITO = '', UBIGEO = '" + Trim(wUbigeo) + "', TELFS = '" + wTelefo + "', REFERE = '" + Trim(wRefer) + "', " _
      & "     EMAIL = '" + Trim(weMail) + "', E_SOCIO = '" + Trim(wESoci) + "' " _
      & " WHERE CODIGO = '" + Format(wCod, "00000000") + "' AND " _
      & "          INS = '" + Format(wIns, "0") + "'  ")
      DbEvento.CommitTrans
   End If
   
   DbEvento.Close
   
   ADO1.Requery
   LlenaCab1
   ADO1.Find "CODSOCIO = " + Str(Val(wCod)) + " "
   MsgBox "Socio " + Str(wCod) + " " + wNom + vbNewLine + _
          "Grabado OK", vbExclamation
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub editar(estado As Boolean)
   FraDetalles.Enabled = estado
   
   cmdNuevo.Visible = Not estado
   cmdModificar.Visible = Not estado
   cmdEliminar.Visible = Not estado
   
   DataGrid1.Enabled = Not estado
   fraDesplaza.Enabled = Not estado
   fraFiltro.Enabled = Not estado
   
   cmdGrabar.Visible = estado
   cmdDeshacer.Visible = estado
   cmdExporta.Visible = Not estado
   cmdCerrar.Visible = Not estado
End Sub

Private Sub cmbGrupo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdGrabar.SetFocus
   End If
End Sub

Private Sub chkAsignado_Click()
   cmdAsignado.Enabled = IIf(chkAsignado.Value = vbChecked, True, False)
End Sub

Private Sub chkCartaDieco_Click()
   If chkFamiliar.Enabled Then
      chkFamiliar.SetFocus
   End If
End Sub

Private Sub chkFamiliar_Click()
   cmdFamiliar.Enabled = IIf(chkFamiliar.Value = vbChecked, True, False)
End Sub

Private Sub chkVip_Click()
   If txtFecVip.Enabled = True Then
      txtFecVip.SetFocus
   Else
      If chkCartaDieco.Enabled = True Then
         chkCartaDieco.SetFocus
      End If
   End If
End Sub

Private Sub cmbE_Socio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtTomo.SetFocus
   End If
End Sub

Private Sub cmbECivil_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmbSexo.SetFocus
   End If
End Sub

Private Sub cmbGrado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtCarnetPNP.SetFocus
   End If
End Sub

Private Sub cmbSexo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmbTipCob.SetFocus
   End If
End Sub

Private Sub cmbSitu_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmbSituEsp.SetFocus
   End If
End Sub

Private Sub cmbSituEsp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFecNac.SetFocus
   End If
End Sub

Private Sub cmbTipCob_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtDirec.SetFocus
   End If
End Sub

Private Sub cmdAsignado_Click()
   zSocio = txtCodigo.Text

   frmMaeAsig.Show
End Sub

Private Sub cmdCerrar_Click()
   Unload Me
End Sub

Private Sub cmdDeshacer_Click()
   MsgBox "Los Cambios Efectuados Se Perderán", vbExclamation
   ACCION = 0
   
   editar (False)
   
   Limpiar
   refrescar
End Sub

Private Sub cmdEliminar_Click()
   On Error GoTo err
   
   Dim wcon As Integer, wNom As String, wNew As Integer, aa As Integer
   wcon = ADO1!codsocio
   wNom = Trim(ADO1!nombre)
   wNew = 0
   ADO1.MoveNext
   If Not ADO1.EOF Then
      wNew = ADO1!codsocio
   Else
      ADO1.MovePrevious
      ADO1.MovePrevious
      If ADO1.BOF Then
         wNew = ""
      Else
         wNew = ADO1!codsocio
      End If
   End If
   
   If MsgBox("¿Esta seguro de borrar Codigo " + Format(wcon, "#######0") + "?", vbYesNo + vbDefaultButton2 + vbQuestion, "Advertencia") = vbYes Then
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAESOCIO " _
      & " WHERE CODSOCIO = " + Str(wcon) + " ")
      Db.CommitTrans
      
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_MAESOCIO " _
      & " WHERE CODSOCIO = " + Str(wcon) + " AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
      
      ADO1.Requery
      LlenaCab
      LlenaCab1
      Limpiar
      refrescar
      
      MsgBox "Socio " + Str(wcon) + " " + wNom + vbNewLine + _
             "Eliminado OK", vbExclamation
      
      If wNew <> 0 Then
         ADO1.Find "CODSOCIO=" + Str(Val(wNew)) + ""
      End If
   End If
   ACCION = 0
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdExporta_Click()
   On Error GoTo err
   
   Dim aa As Long, I As Integer, Heading(16) As String, wreg As Long, wTot As Long
   Dim wNomCob As String, wFecIng As Date, wNomGra As String, wNomUbi As String
   Heading(0) = "SOCIO"
   Heading(1) = "CODIGO"
   Heading(2) = "INS"
   Heading(3) = "NOMBRE"
   Heading(4) = "ESTADO"
   Heading(5) = "GRADO"
   Heading(6) = "D.N.I."
   Heading(7) = "FEC.ING"
   Heading(8) = "TIP.COB"
   Heading(9) = "DIRECCION"
   Heading(10) = "REFER"
   Heading(11) = "UBIGEO"
   Heading(12) = "TELEFONO"
   Heading(13) = "TELEFONO2"
   Heading(14) = "CELULAR"
   Heading(15) = "EMAIL"
   Heading(16) = "EMAIL2"
   
   aa = Leerado3("SELECT * FROM TMP_MAESOCIO WHERE USU = '" + wcodusu + "' ORDER BY NOMBRE ")
   If aa > 0 Then
      wTot = aa
      Set objExcel = New Excel.Application
      objExcel.SheetsInNewWorkbook = 1
      objExcel.Workbooks.Add
      With objExcel
           .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
           .Range(.Cells(3, 1), .Cells(3, 17)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 17)).Font.Bold = True
           .Cells(1, 1) = wnomcia
           .Cells(2, 1) = "MAESTRO DE SOCIOS"
           For I = 1 To 17 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           objExcel.Columns("A").ColumnWidth = 9
           objExcel.Columns("B").ColumnWidth = 10
           objExcel.Columns("C").ColumnWidth = 4
           objExcel.Columns("D").ColumnWidth = 50
           objExcel.Columns("E").ColumnWidth = 6
           objExcel.Columns("F").ColumnWidth = 15
           objExcel.Columns("G").ColumnWidth = 11
           objExcel.Columns("H").ColumnWidth = 11
           objExcel.Columns("I").ColumnWidth = 20
           objExcel.Columns("J").ColumnWidth = 50
           objExcel.Columns("K").ColumnWidth = 50
           objExcel.Columns("L").ColumnWidth = 20
           objExcel.Columns("M").ColumnWidth = 20
           objExcel.Columns("N").ColumnWidth = 20
           objExcel.Columns("O").ColumnWidth = 20
           objExcel.Columns("P").ColumnWidth = 25
           objExcel.Columns("Q").ColumnWidth = 25
      End With
      V = 4
      H = 1
      wreg = 1
      Do While Not ADO3.EOF
         DoEvents
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wreg, "####0") + " / " + Format(wTot, "####0")
         lblMensaje.Refresh
         
         wNomCob = ""
         aa = Leerado7("SELECT * FROM MAETIPCOB " _
                    & " WHERE TIPCOB = '" + ADO3!tipcob + "' ")
         If aa > 0 Then
            wNomCob = ADO7!nombre
         End If
         Set ADO7 = Nothing
         
         wNomGra = ""
         aa = Leerado7("SELECT * FROM MAEGRADO " _
                    & " WHERE GRADO = " + Str(ADO3!grado) + " ")
         If aa > 0 Then
            wNomGra = ADO7!nombre
         End If
         Set ADO7 = Nothing
         
         wNomUbi = ""
         aa = Leerado7("SELECT * FROM MAEUBIGEO " _
                    & " WHERE CODIGO = '" + ADO3!ubigeo + "' ")
         If aa > 0 Then
            wNomUbi = ADO7!nombre
         End If
         Set ADO7 = Nothing
         
         objExcel.Cells(V, H + 0) = ADO3!codsocio
         objExcel.Cells(V, H + 1) = ADO3!codigo
         objExcel.Cells(V, H + 2) = ADO3!ins
         objExcel.Cells(V, H + 3) = IIf(IsNull(ADO3!nombre), "", ADO3!nombre)
         objExcel.Cells(V, H + 4) = ADO3!e_socio
         objExcel.Cells(V, H + 5) = wNomGra
         objExcel.Cells(V, H + 6) = ADO3!numdoc
         If IsDate(ADO3!fecing) Then
            wFecIng = ADO3!fecing
            objExcel.Cells(V, H + 7) = wFecIng
         End If
         objExcel.Cells(V, H + 8) = wNomCob
         objExcel.Cells(V, H + 9) = ADO3!direc
         objExcel.Cells(V, H + 10) = ADO3!refer
         objExcel.Cells(V, H + 11) = wNomUbi
         objExcel.Cells(V, H + 12) = ADO3!telefono
         objExcel.Cells(V, H + 13) = ADO3!telefon2
         objExcel.Cells(V, H + 14) = ADO3!celular
         objExcel.Cells(V, H + 15) = ADO3!email
         objExcel.Cells(V, H + 16) = ADO3!email2
         
         wreg = wreg + 1
         V = V + 1
         ADO3.MoveNext
      Loop
      Set ADO3 = Nothing
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

Private Sub cmdFamiliar_Click()
   zSocio = txtCodigo.Text

   frmMaeFamilia.Show
End Sub

Private Sub cmdGrabar_Click()
   On Error GoTo err
   Dim wSoc As Integer, wCod As Long, wIns As Integer, wDni As String, wNom As String
   If ACCION = 1 Then
      wSoc = Val(txtCodigo.Text)
      wCod = Val(txtCodofin.Text)
      wIns = Val(txtIns.Text)
      wDni = txtNumdoc.Text
      wNom = txtNombre.Text
      
      If wSoc = 0 Then
         MsgBox "Codigo Socio En Blanco", vbExclamation
         txtCodigo.Text = ""
         Exit Sub
      End If
      If wCod = 0 Then
         MsgBox "Codofin En Blanco", vbExclamation
         txtCodofin.Text = ""
         Exit Sub
      End If
      If wIns = 0 Then
         MsgBox "Codofin En Blanco", vbExclamation
         txtIns.Text = ""
         Exit Sub
      End If
      If Len(Trim(wNom)) = 0 Then
         MsgBox "Nombre Del Asociado En Blanco", vbExclamation
         txtNombre.Text = ""
         Exit Sub
      End If
      If Len(Trim(wDni)) = 0 Then
         MsgBox "DNI Del Asociado En Blanco", vbExclamation
         txtNumdoc.Text = ""
         Exit Sub
      End If
      
      If Leerado2("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ") > 0 Then
         MsgBox "Codigo de Socio Ya Existe", vbExclamation
         Limpiar
         txtCodigo.SetFocus
         Exit Sub
      End If
      If Leerado2("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(wCod) + " AND INS = " + Str(wIns) + " ") > 0 Then
         MsgBox "Codofin Ya Existe", vbExclamation
         Limpiar
         txtCodigo.SetFocus
         Exit Sub
      End If
      If Leerado2("SELECT * FROM MAESOCIO WHERE NUMDOC = '" + wDni + "' ") > 0 Then
         MsgBox "DNI de Socio Ya Existe", vbExclamation
         Limpiar
         txtCodigo.SetFocus
         Exit Sub
      End If
      If Leerado2("SELECT * FROM MAESOCIO WHERE NOMBRE = '" + wNom + "' ") > 0 Then
         MsgBox "Nombre de Socio Ya Existe", vbExclamation
         Limpiar
         txtCodigo.SetFocus
         Exit Sub
      End If
   
   End If
   grabar
   ACCION = 0
   editar False
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdModificar_Click()
   ACCION = 2
   editar True
   refrescar
   txtCodigo.Enabled = False
   txtCodofin.Enabled = False
   txtIns.Enabled = False
   txtNumdoc.SetFocus
End Sub

Private Sub cmdMover_Click(Index As Integer)
    With ADO1
    If .BOF And .EOF Then
       Exit Sub
    End If
    Select Case Index
    Case 0
        .MoveFirst
    Case 1
        .MovePrevious
        If .BOF Then .MoveFirst
    Case 2
        .MoveNext
        If .EOF Then .MoveLast
    Case 3
        .MoveLast
    End Select
    End With
    refrescar
End Sub

Private Sub cmdNuevo_Click()
   Dim wNew As String, aa As Integer
   
   wNew = 0
   aa = Leerado8("SELECT MAX(CODSOCIO) AS CODSOCIO FROM MAESOCIO ")
   If aa > 0 Then
      wNew = IIf(IsNull(ADO8!codsocio), 0, ADO8!codsocio)
   End If
   Set ADO8 = Nothing
   
   wNew = wNew + 1
   
   ACCION = 1
   editar True
   Limpiar
   
   txtFecIng.Text = Format(Date, "dd/mm/yyyy")
   cmbE_Socio.ListIndex = 9
   cmbTipCob.ListIndex = 3
   
   txtCodigo.Text = wNew
   txtCodigo.Enabled = False
   
   txtCodofin.SetFocus
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   Select Case ColIndex
   Case 0
        ADO1.Sort = "GRADO"
   Case 1
        ADO1.Sort = "NOMBRE"
   End Select
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   On Error GoTo err

   Dim wreg As Integer, zReg As String, file As String
   
   If ACCION = 0 Then
      Limpiar
      refrescar
   
   
      zReg = Trim(Format(ADO1!codsocio, "00000"))
   
      file = "P:\fotos\" + zReg + "jpg.jpg"
      
      If Len(Dir$(file)) Then
         Image1.Picture = LoadPicture(file)
      Else
         Image1.Picture = LoadPicture("P:\fotos\SinFoto.jpg")
      End If
      Image1.Refresh
   End If
  
   Exit Sub
err:
   Resume Next
End Sub

Private Sub Form_Activate()
   frmMaeSocio.Left = (Screen.Width - Width) \ 2
   frmMaeSocio.Top = 0
   optTodos.Value = True
   
   cmdEliminar.Enabled = False
   
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
   
   LlenaCab
   LlenaCab1
   Limpiar
   refrescar
   editar False
   Call DataGrid1_RowColChange(0, 0)
   
   DataGrid1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set ADO1 = Nothing
End Sub

Private Sub LlenaCab()
   Dim a As Integer
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_MAESOCIO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_MAESOCIO " _
   & " (CODSOCIO, CODIGO, INS, NUMDOC, CARNETPNP, CARNETPIP, NOMBRE, GRADO, SITU, " _
   & "  SITUESP, FECNAC, FECMAT, SEXO, ECIVIL, DIREC, UBIGEO, TELEFONO, TELEFON2, " _
   & "  CELULAR, EMAIL, EMAIL2, REFER, E_SOCIO, TOMO, TIPCOB, FECING, FECRENU, " _
   & "  NRESO_ING, FRESO_ING, FECREIN, OBSERVAC, OBSERVA2, " _
   & "  ASIGNADO, FAMILIAR, VIP, FECVIP, CARTADIECO, FECAMNIS, FECCONDO, " _
   & "  DIRECTIVO, ANODIREC, PROMOCION, NRESO_REING, MESFALL, FECFALL, USU ) " _
   & " SELECT " _
   & "  CODSOCIO, CODIGO, INS, NUMDOC, CARNETPNP, CARNETPIP, NOMBRE, GRADO, SITU, " _
   & "  SITUESP, FECNAC, FECMAT, SEXO, ECIVIL, DIREC, UBIGEO, TELEFONO, TELEFON2, " _
   & "  CELULAR, EMAIL, EMAIL2, REFER, E_SOCIO, TOMO, TIPCOB, FECING, FECRENU, " _
   & "  NRESO_ING, FRESO_ING, FECREIN, OBSERVAC, OBSERVA2, " _
   & "  ASIGNADO, FAMILIAR, VIP, FECVIP, CARTADIECO, FECAMNIS, FECCONDO, " _
   & "  DIRECTIVO, ANODIREC, PROMOCION, NRESO_REING, MESFALL, FECFALL, '" + wcodusu + "' " _
   & " FROM MAESOCIO ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_MAESOCIO " _
   & " SET NOMGRA = M.NOMBRE " _
   & " FROM TMP_MAESOCIO AS T INNER JOIN MAEGRADO AS M " _
   & "   ON T.GRADO = M.GRADO " _
   & " WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_MAESOCIO " _
   & " SET NOMSITU = M.NOMBRE " _
   & " FROM TMP_MAESOCIO AS T INNER JOIN MAESITU AS M " _
   & "   ON T.SITU = M.SITU " _
   & " WHERE USU = '" + wcodusu + "' ")
5   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_MAESOCIO " _
   & " SET NOMSITUESP = M.NOMBRE " _
   & " FROM TMP_MAESOCIO AS T INNER JOIN MAESITUESP AS M " _
   & "   ON T.SITUESP = M.SITUESP " _
   & " WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_MAESOCIO " _
   & " SET NOMESO = M.NOMBRE " _
   & " FROM TMP_MAESOCIO AS T INNER JOIN MAEE_SOCIO AS M " _
   & "   ON T.E_SOCIO = M.E_SOCIO " _
   & " WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_MAESOCIO " _
   & " SET NOMCOB = M.NOMBRE " _
   & " FROM TMP_MAESOCIO AS T INNER JOIN MAETIPCOB AS M " _
   & "   ON T.TIPCOB = M.TIPCOB " _
   & " WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   txtFiltrar.Text = ""
   txtFiltrarCodofin.Text = ""
   
   a = Leerado("SELECT CODSOCIO, CODIGO, INS, NUMDOC, NOMBRE, NOMGRA, NOMSITU, SEXO, ECIVIL, " _
                & "    NOMCOB, NOMESO, ASIGNADO, FAMILIAR, " _
                & "    CARNETPNP, CARNETPIP, FECNAC, FECMAT, DIREC, UBIGEO, TELEFONO, TELEFON2, CELULAR, " _
                & "    EMAIL, EMAIL2, REFER, TOMO, TIPCOB, SITU, GRADO, E_SOCIO, FECING, FECRENU, " _
                & "    NRESO_ING, FRESO_ING, FECEXCLU, FECEXPUL, FECREIN, OBSERVAC, OBSERVA2, SITUESP, " _
                & "    NOMSITUESP, VIP, FECVIP, CARTADIECO, FECAMNIS, FECCONDO, DIRECTIVO, " _
                & "    ANODIREC, PROMOCION, NRESO_REING, MESFALL, FECFALL, USU  " _
                & " FROM TMP_MAESOCIO " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY CODSOCIO ")
   Set DataGrid1.DataSource = ADO1
End Sub

Private Sub LlenaCab1()
   DataGrid1.Columns(0).Width = 550
   DataGrid1.Columns(0).Alignment = dbgRight
   DataGrid1.Columns(0).Caption = "SOCIO"
    
   DataGrid1.Columns(1).Width = 800
   DataGrid1.Columns(1).Alignment = dbgRight
   DataGrid1.Columns(1).Caption = "CODOFIN"
    
   DataGrid1.Columns(2).Width = 400
   DataGrid1.Columns(2).Alignment = dbgCenter
   DataGrid1.Columns(2).Caption = "INS"
    
   DataGrid1.Columns(3).Width = 900
   DataGrid1.Columns(3).Alignment = dbgCenter
   DataGrid1.Columns(3).Caption = "D.N.I."
    
   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Width = 4200
   DataGrid1.Columns(4).Caption = "NOMBRE"

   DataGrid1.Columns(5).Width = 1200
   DataGrid1.Columns(5).Alignment = dbgCenter
   DataGrid1.Columns(5).Caption = "GRADO"

   DataGrid1.Columns(6).Width = 1200
   DataGrid1.Columns(6).Alignment = dbgCenter
   DataGrid1.Columns(6).Caption = "SITUAC"

   DataGrid1.Columns(7).Width = 500
   DataGrid1.Columns(7).Alignment = dbgCenter
   DataGrid1.Columns(7).Caption = "SEXO"

   DataGrid1.Columns(8).Width = 500
   DataGrid1.Columns(8).Alignment = dbgCenter
   DataGrid1.Columns(8).Caption = "E.CIVIL"

   DataGrid1.Columns(9).Width = 1800
   DataGrid1.Columns(9).Alignment = dbgCenter
   DataGrid1.Columns(9).Caption = "TIP.COB"

   DataGrid1.Columns(10).Width = 1400
   DataGrid1.Columns(10).Alignment = dbgCenter
   DataGrid1.Columns(10).Caption = "E_SOCIO"

   DataGrid1.Columns(11).Width = 600
   DataGrid1.Columns(11).Alignment = dbgCenter
   DataGrid1.Columns(11).Caption = "ASIG"

   DataGrid1.Columns(12).Width = 600
   DataGrid1.Columns(12).Alignment = dbgCenter
   DataGrid1.Columns(12).Caption = "FAMIL"

   DataGrid1.Columns(13).Visible = False
   DataGrid1.Columns(14).Visible = False
   DataGrid1.Columns(15).Visible = False
   DataGrid1.Columns(16).Visible = False
   DataGrid1.Columns(17).Visible = False
   DataGrid1.Columns(18).Visible = False
   DataGrid1.Columns(19).Visible = False
   DataGrid1.Columns(20).Visible = False
   DataGrid1.Columns(21).Visible = False
   DataGrid1.Columns(22).Visible = False
   DataGrid1.Columns(23).Visible = False
   DataGrid1.Columns(24).Visible = False
   DataGrid1.Columns(25).Visible = False
   DataGrid1.Columns(26).Visible = False
   DataGrid1.Columns(27).Visible = False
   DataGrid1.Columns(28).Visible = False
   DataGrid1.Columns(29).Visible = False
   DataGrid1.Columns(30).Visible = False
   DataGrid1.Columns(31).Visible = False
   DataGrid1.Columns(32).Visible = False
   DataGrid1.Columns(33).Visible = False
   DataGrid1.Columns(34).Visible = False
   DataGrid1.Columns(35).Visible = False
   DataGrid1.Columns(36).Visible = False
   DataGrid1.Columns(37).Visible = False
   DataGrid1.Columns(38).Visible = False
   DataGrid1.Columns(39).Visible = False
   DataGrid1.Columns(40).Visible = False
   DataGrid1.Columns(41).Visible = False
   DataGrid1.Columns(42).Visible = False
   DataGrid1.Columns(43).Visible = False
   DataGrid1.Columns(44).Visible = False
   DataGrid1.Columns(45).Visible = False
   DataGrid1.Columns(46).Visible = False
   DataGrid1.Columns(47).Visible = False
   DataGrid1.Columns(48).Visible = False
   DataGrid1.Columns(49).Visible = False
   DataGrid1.Columns(50).Visible = False
'   DataGrid1.Columns(51).Visible = False
'   DataGrid1.Columns(52).Visible = False
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

Private Sub optFiltrocodofin_Click()
   If optTodosCodoFin.Value = True Then
      txtFiltrarCodofin.Text = ""
      txtFiltrarCodofin.Enabled = False
      DataGrid1.SetFocus
   Else
      txtFiltrarCodofin.Enabled = True
      optFiltroCodoFin.Value = True
      txtFiltrarCodofin.SetFocus
   End If
End Sub

Private Sub optFiltroCodofin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      optTodosCodofin_Click
   End If
End Sub

Private Sub optFiltroDni_Click()
   If optTodosDni.Value = True Then
      txtFiltrarDni.Text = ""
      txtFiltrarDni.Enabled = False
      DataGrid1.SetFocus
   Else
      txtFiltrarDni.Enabled = True
      optFiltroDni.Value = True
      txtFiltrarDni.SetFocus
   End If
End Sub

Private Sub optFiltroDni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      optTodosDni_Click
   End If
End Sub

Private Sub optFiltroPariente_Click()
   If optTodosPariente.Value = True Then
      txtFiltrarPariente.Text = ""
      txtFiltrarPariente.Enabled = False
      DataGrid1.SetFocus
   Else
      txtFiltrarPariente.Enabled = True
      optFiltroPariente.Value = True
      txtFiltrarPariente.SetFocus
   End If
End Sub

Private Sub optFiltroPariente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      optTodosPariente_Click
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

Private Sub optTodosCodofin_Click()
   If optTodosCodoFin.Value = True Then
      txtFiltrarCodofin.Text = ""
      txtFiltrarCodofin.Enabled = False
      LlenaCab
      LlenaCab1
      Limpiar
      refrescar
   Else
      txtFiltrarCodofin.Enabled = True
      optFiltroCodoFin.Value = True
   End If
End Sub

Private Sub optTodosCodofin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If optTodosCodoFin.Value = True Then
         txtFiltrarCodofin.Text = ""
         txtFiltrarCodofin.Enabled = False
         ADO1.Filter = ""
         Set DataGrid1.DataSource = ADO1
         DataGrid1.SetFocus
      Else
         txtFiltrarCodofin.Enabled = True
         optFiltroCodoFin.Value = True
         txtFiltrarCodofin.SetFocus
      End If
   End If
End Sub

Private Sub optTodosDni_Click()
   If optTodosDni.Value = True Then
      txtFiltrarDni.Text = ""
      txtFiltrarDni.Enabled = False
      LlenaCab
      LlenaCab1
      Limpiar
      refrescar
   Else
      txtFiltrarDni.Enabled = True
      optFiltroDni.Value = True
   End If
End Sub

Private Sub optTodosDni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If optTodosDni.Value = True Then
         txtFiltrarDni.Text = ""
         txtFiltrarDni.Enabled = False
         ADO1.Filter = ""
         Set DataGrid1.DataSource = ADO1
         DataGrid1.SetFocus
      Else
         txtFiltrarDni.Enabled = True
         optFiltroDni.Value = True
         txtFiltrarDni.SetFocus
      End If
   End If
End Sub

Private Sub optTodosPariente_Click()
   If optTodosPariente.Value = True Then
      txtFiltrarPariente.Text = ""
      txtFiltrarPariente.Enabled = False
      LlenaCab
      LlenaCab1
      Limpiar
      refrescar
   Else
      txtFiltrarPariente.Enabled = True
      optFiltroPariente.Value = True
   End If
End Sub

Private Sub optTodosPariente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If optTodosPariente.Value = True Then
         txtFiltrarPariente.Text = ""
         txtFiltrarPariente.Enabled = False
         ADO1.Filter = ""
         Set DataGrid1.DataSource = ADO1
         DataGrid1.SetFocus
      Else
         txtFiltrarPariente.Enabled = True
         optFiltroPariente.Value = True
         txtFiltrarPariente.SetFocus
      End If
   End If
End Sub

Private Sub txtAnoDirec_GotFocus()
   txtAnoDirec.SelStart = 0
   txtAnoDirec.SelLength = 10
End Sub

Private Sub txtAnoDirec_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDirectivo.SetFocus
   Case 40
        txtPromocion.SetFocus
   End Select
End Sub

Private Sub txtAnoDirec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtPromocion.SetFocus
   Else
      If InStr(1, "0123456789-" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtCarnetPIP_GotFocus()
   txtCarnetPIP.SelStart = 0
   txtCarnetPIP.SelLength = 8
End Sub

Private Sub txtCarnetPIP_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtCarnetPNP.SetFocus
   Case 40
        cmbSitu.SetFocus
   End Select
End Sub

Private Sub txtCarnetPIP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmbSitu.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtCarnetPNP_GotFocus()
   txtCarnetPNP.SelStart = 0
   txtCarnetPNP.SelLength = 8
End Sub

Private Sub txtCarnetPNP_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        cmbGrado.SetFocus
   Case 40
        txtCarnetPIP.SetFocus
   End Select
End Sub

Private Sub txtCarnetPNP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtCarnetPIP.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtCelular_GotFocus()
   txtCelular.SelStart = 0
   txtCelular.SelLength = 10
End Sub

Private Sub txtCelular_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtTelefon2.SetFocus
   Case 40
        txtRefer.SetFocus
   End Select
End Sub

Private Sub txtCelular_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtRefer.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtCodofin_GotFocus()
   txtCodofin.SelStart = 0
   txtCodofin.SelLength = 8
End Sub

Private Sub txtCodofin_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 40
        txtIns.SetFocus
   End Select
End Sub

Private Sub txtCodofin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txtCodofin.Text)) = 0 Then
         MsgBox "Codofin En Blanco", vbExclamation
         txtCodofin.Text = ""
         Exit Sub
      End If
      
      txtIns.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDirec_GotFocus()
   txtDirec.SelStart = 0
   txtDirec.SelLength = 50
End Sub

Private Sub txtDirec_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        cmbSexo.SetFocus
   Case 40
        txtUbi1.SetFocus
   End Select
End Sub

Private Sub txtDirec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtUbi1.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtDirectivo_Change()
   Dim aa As Integer
   If Len(Trim(txtDirectivo.Text)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEDIRECTIVO WHERE DIRECTIVO = '" + txtDirectivo.Text + "' ")
      If aa > 0 Then
         lblDirectivo.Caption = ADO8!nombre
      Else
         lblDirectivo.Caption = ""
      End If
      Set ADO8 = Nothing
   Else
      lblDirectivo.Caption = ""
   End If
End Sub

Private Sub txtDirectivo_GotFocus()
   txtDirectivo.SelStart = 0
   txtDirectivo.SelLength = 3
End Sub

Private Sub txtDirectivo_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        If txtFecAmnis.Enabled = True Then
           txtFecAmnis.SetFocus
        End If
   Case 40
        txtAnoDirec.SetFocus
   Case 116
        xlista = "DV"
        xseleccion = ""
        frmSeleccion.Show 1
        If xseleccion <> "" Then
           txtDirectivo.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtDirectivo_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtDirectivo.Text)) > 0 Then
         aa = Leerado8("SELECT * FROM MAEDIRECTIVO WHERE DIRECTIVO = '" + txtDirectivo.Text + "' ")
         If aa = 0 Then
            MsgBox "Directivo Digitado NO existe", vbExclamation
            txtDirectivo.Text = ""
            Exit Sub
         End If
      End If
      txtAnoDirec.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txteMail_GotFocus()
   txteMail.SelStart = 0
   txteMail.SelLength = 50
End Sub

Private Sub txteMail_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtRefer.SetFocus
   Case 40
        txtEMail2.SetFocus
   End Select
End Sub

Private Sub txteMail_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtEMail2.SetFocus
   Else
      KeyAscii = Asc(LCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtEMail2_GotFocus()
   txtEMail2.SelStart = 0
   txtEMail2.SelLength = 50
End Sub

Private Sub txtEMail2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txteMail.SetFocus
   Case 40
        cmbE_Socio.SetFocus
   End Select
End Sub

Private Sub txtEMail2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmbE_Socio.SetFocus
   Else
      KeyAscii = Asc(LCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtFecAmnis_GotFocus()
   txtFecAmnis.SelStart = 0
   txtFecAmnis.SelLength = 10
End Sub

Private Sub txtFecAmnis_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFecCondo.SetFocus
   Case 40
        txtDirectivo.SetFocus
   End Select
End Sub

Private Sub txtFecAmnis_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFecAmnis.Text <> "__/__/____" Then
         If Not IsDate(txtFecAmnis.Text) Then
            MsgBox "Fecha Amnistia Digitada Es Invalida", vbExclamation
            txtFecAmnis.Text = "__/__/____"
            txtFecAmnis.SetFocus
            Exit Sub
         End If
      End If
      txtDirectivo.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFecCondo_GotFocus()
   txtFecCondo.SelStart = 0
   txtFecCondo.SelLength = 10
End Sub

Private Sub txtFecCondo_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFecRein.SetFocus
   Case 40
        txtFecAmnis.SetFocus
   End Select
End Sub

Private Sub txtFecCondo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFecCondo.Text <> "__/__/____" Then
         If Not IsDate(txtFecCondo.Text) Then
            MsgBox "Fecha Condonación Digitada Es Invalida", vbExclamation
            txtFecCondo.Text = "__/__/____"
            txtFecCondo.SetFocus
            Exit Sub
         End If
      End If
      txtFecAmnis.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFecExclu_GotFocus()
   txtFecExclu.SelStart = 0
   txtFecExclu.SelLength = 10
End Sub

Private Sub txtFecExclu_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtNumReso.SetFocus
   Case 40
        txtFecExpul.SetFocus
   End Select
End Sub

Private Sub txtFecExclu_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFecExclu.Text <> "__/__/____" Then
         If Not IsDate(txtFecExclu.Text) Then
            MsgBox "Fecha Resolución Digitada Es Invalida", vbExclamation
            txtFecExclu.Text = "__/__/____"
            txtFecExclu.SetFocus
            Exit Sub
         End If
      End If
      txtFecExpul.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFecExpul_GotFocus()
   txtFecExpul.SelStart = 0
   txtFecExpul.SelLength = 10
End Sub

Private Sub txtFecExpul_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFecExclu.SetFocus
   Case 40
        txtFecRein.SetFocus
   End Select
End Sub

Private Sub txtFecExpul_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsDate(txtFecExpul.Text) And txtFecExpul.Text <> "__/__/____" Then
         MsgBox "Fecha Exclusión Es Invalida", vbExclamation
         txtFecExpul.Text = "__/__/____"
         Exit Sub
      End If
      txtFecRein.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFecFall_GotFocus()
   txtFecFall.SelStart = 0
   txtFecFall.SelLength = 10
End Sub

Private Sub txtFecFall_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtMesFall.SetFocus
   Case 40
        txtObservac.SetFocus
   End Select
End Sub

Private Sub txtFecFall_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtObservac.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFecIng_GotFocus()
   txtFecIng.SelStart = 0
   txtFecIng.SelLength = 10
End Sub

Private Sub txtFecIng_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        cmbE_Socio.SetFocus
   Case 40
        txtTomo.SetFocus
   End Select
End Sub

Private Sub txtFecIng_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsDate(txtFecIng.Text) And txtFecIng.Text <> "__/__/____" Then
         MsgBox "Fecha Ingreso Es Invalida", vbExclamation
         txtFecIng.Text = "__/__/____"
         Exit Sub
      End If
      txtTomo.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFecMat_GotFocus()
   txtFecMat.SelStart = 0
   txtFecMat.SelLength = 10
End Sub

Private Sub txtFecMat_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFecNac.SetFocus
   Case 40
        cmbECivil.SetFocus
   End Select
End Sub

Private Sub txtFecMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsDate(txtFecMat.Text) And txtFecMat.Text <> "__/__/____" Then
         MsgBox "Fecha Digitada Es Invalida", vbExclamation
         txtFecMat.Text = "__/__/____"
         Exit Sub
      End If
      cmbECivil.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFecNac_GotFocus()
   txtFecNac.SelStart = 0
   txtFecNac.SelLength = 10
End Sub

Private Sub txtFecNac_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        cmbSituEsp.SetFocus
   Case 40
        txtFecMat.SetFocus
   End Select
End Sub

Private Sub txtFecNac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsDate(txtFecNac.Text) And txtFecNac.Text <> "__/__/____" Then
         MsgBox "Fecha Digitada Es Invalida", vbExclamation
         txtFecNac.Text = "__/__/____"
         Exit Sub
      End If
      txtFecMat.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFecRein_GotFocus()
   txtFecRein.SelStart = 0
   txtFecRein.SelLength = 10
End Sub

Private Sub txtFecRein_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFecExpul.SetFocus
   Case 40
        txtFecCondo.SetFocus
   End Select
End Sub

Private Sub txtFecRein_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsDate(txtFecRein.Text) And txtFecRein.Text <> "__/__/____" Then
         MsgBox "Fecha Digitada Es Invalida", vbExclamation
         txtFecRein.Text = "__/__/____"
         Exit Sub
      End If
      txtFecCondo.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFecRenu_GotFocus()
   txtFecRenu.SelStart = 0
   txtFecRenu.SelLength = 10
End Sub

Private Sub txtFecRenu_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFecIng.SetFocus
   Case 40
        txtTomo.SetFocus
   End Select
End Sub

Private Sub txtFecRenu_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsDate(txtFecRenu.Text) And txtFecRenu.Text <> "__/__/____" Then
         MsgBox "Fecha Digitada Es Invalida", vbExclamation
         txtFecRenu.Text = "__/__/____"
         Exit Sub
      End If
      txtTomo.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFecReso_GotFocus()
   txtFecReso.SelStart = 0
   txtFecReso.SelLength = 10
End Sub

Private Sub txtFecReso_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtTomo.SetFocus
   Case 40
        txtNumReso.SetFocus
   End Select
End Sub

Private Sub txtFecReso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFecReso.Text <> "__/__/____" Then
         If Not IsDate(txtFecReso.Text) Then
            MsgBox "Fecha Resolución Digitada Es Invalida", vbExclamation
            txtFecReso.Text = "__/__/____"
            txtFecReso.SetFocus
            Exit Sub
         End If
      End If
      txtNumReso.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFecVip_GotFocus()
   txtFecVip.SelStart = 0
   txtFecVip.SelLength = 10
End Sub

Private Sub txtFecVip_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      chkCartaDieco.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFiltrar_Change()
   Dim a As Long
   a = Leerado("SELECT CODSOCIO, CODIGO, INS, NUMDOC, NOMBRE, NOMGRA, NOMSITU, SEXO, ECIVIL, " _
                & "    NOMCOB, NOMESO, ASIGNADO, FAMILIAR, " _
                & "    CARNETPNP, CARNETPIP, FECNAC, FECMAT, DIREC, UBIGEO, TELEFONO, TELEFON2, CELULAR, " _
                & "    EMAIL, EMAIL2, REFER, TOMO, TIPCOB, SITU, GRADO, E_SOCIO, FECING, FECRENU, " _
                & "    NRESO_ING, FRESO_ING, FECEXCLU, FECEXPUL, FECREIN, OBSERVAC, OBSERVA2, SITUESP, " _
                & "    NOMSITUESP, VIP, FECVIP, CARTADIECO, FECAMNIS, FECCONDO, DIRECTIVO, " _
                & "    ANODIREC, PROMOCION, NRESO_REING, MESFALL, FECFALL, USU  " _
                & " FROM TMP_MAESOCIO " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "      NOMBRE LIKE '%" + Trim(txtFiltrar.Text) + "%' " _
                & " ORDER BY CODSOCIO ")
   Set DataGrid1.DataSource = ADO1

   LlenaCab1
   Limpiar
   refrescar
End Sub

Private Sub txtFiltrar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtFiltrarCodofin_Change()
   Dim a As Long
   
   a = Leerado("SELECT CODSOCIO, CODIGO, INS, NUMDOC, NOMBRE, NOMGRA, NOMSITU, SEXO, ECIVIL, " _
                & "    NOMCOB, NOMESO, ASIGNADO, FAMILIAR, " _
                & "    CARNETPNP, CARNETPIP, FECNAC, FECMAT, DIREC, UBIGEO, TELEFONO, TELEFON2, CELULAR, " _
                & "    EMAIL, EMAIL2, REFER, TOMO, TIPCOB, SITU, GRADO, E_SOCIO, FECING, FECRENU, " _
                & "    NRESO_ING, FRESO_ING, FECEXCLU, FECEXPUL, FECREIN, OBSERVAC, OBSERVA2, SITUESP, " _
                & "    NOMSITUESP, VIP, FECVIP, CARTADIECO, FECAMNIS, FECCONDO, DIRECTIVO, " _
                & "    ANODIREC, PROMOCION, NRESO_REING, MESFALL, FECFALL, USU  " _
                & " FROM TMP_MAESOCIO " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "      CODIGO LIKE '%" + Trim(txtFiltrarCodofin.Text) + "%' " _
                & " ORDER BY CODSOCIO ")
   Set DataGrid1.DataSource = ADO1

   LlenaCab1
   Limpiar
   refrescar
End Sub

Private Sub txtFiltrarDni_Change()
   Dim a As Long
   a = Leerado("SELECT CODSOCIO, CODIGO, INS, NUMDOC, NOMBRE, NOMGRA, NOMSITU, SEXO, ECIVIL, " _
                & "    NOMCOB, NOMESO, ASIGNADO, FAMILIAR, " _
                & "    CARNETPNP, CARNETPIP, FECNAC, FECMAT, DIREC, UBIGEO, TELEFONO, TELEFON2, CELULAR, " _
                & "    EMAIL, EMAIL2, REFER, TOMO, TIPCOB, SITU, GRADO, E_SOCIO, FECING, FECRENU, " _
                & "    NRESO_ING, FRESO_ING, FECEXCLU, FECEXPUL, FECREIN, OBSERVAC, OBSERVA2, SITUESP, " _
                & "    NOMSITUESP, VIP, FECVIP, CARTADIECO, FECAMNIS, FECCONDO, DIRECTIVO, " _
                & "    ANODIREC, PROMOCION, NRESO_REING, MESFALL, FECFALL, USU  " _
                & " FROM TMP_MAESOCIO " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "      NUMDOC LIKE '%" + Trim(txtFiltrarDni.Text) + "%' " _
                & " ORDER BY CODSOCIO ")
   Set DataGrid1.DataSource = ADO1

   LlenaCab1
   Limpiar
   refrescar
End Sub

Private Sub txtFiltrarPariente_Change()
   Dim a As Long, wNom As String, wNomSocio As String, _
       wSoc As Integer, wTip As String, wLin As String
   
   wNom = Trim(txtFiltrarPariente.Text)
   wSoc = 0: wTip = "": wLin = "": wNomSocio = ""
   a = Leerado8("SELECT * FROM MAEFAMILIA WHERE NOMBRE LIKE '%" + wNom + "%' ")
   If a > 0 Then
      wSoc = ADO8!codsocio
      wTip = ADO8!tipopariente
      wLin = ADO8!lin
   End If
   Set ADO8 = Nothing
   
   a = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
   If a > 0 Then
      wNomSocio = ADO8!nombre
   End If
   Set ADO8 = Nothing
   
   a = Leerado("SELECT CODSOCIO, CODIGO, INS, NUMDOC, NOMBRE, NOMGRA, NOMSITU, SEXO, ECIVIL, " _
                & "    NOMCOB, NOMESO, ASIGNADO, FAMILIAR, " _
                & "    CARNETPNP, CARNETPIP, FECNAC, FECMAT, DIREC, UBIGEO, TELEFONO, TELEFON2, CELULAR, " _
                & "    EMAIL, EMAIL2, REFER, TOMO, TIPCOB, SITU, GRADO, E_SOCIO, FECING, FECRENU, " _
                & "    NRESO_ING, FRESO_ING, FECEXCLU, FECEXPUL, FECREIN, OBSERVAC, OBSERVA2, SITUESP, " _
                & "    NOMSITUESP, VIP, FECVIP, CARTADIECO, FECAMNIS, FECCONDO, DIRECTIVO, " _
                & "    ANODIREC, PROMOCION, NRESO_REING, MESFALL, FECFALL, USU  " _
                & " FROM TMP_MAESOCIO " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "      NOMBRE LIKE '%" + Trim(wNomSocio) + "%' " _
                & " ORDER BY CODSOCIO ")
   Set DataGrid1.DataSource = ADO1

   LlenaCab1
   Limpiar
   refrescar
End Sub

Private Sub txtFiltrarPariente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtIns_GotFocus()
   txtIns.SelStart = 0
   txtIns.SelLength = 1
End Sub

Private Sub txtIns_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtCodofin.SetFocus
   Case 40
        txtNumdoc.SetFocus
   End Select
End Sub

Private Sub txtIns_KeyPress(KeyAscii As Integer)
   Dim aa As Integer, wCod As Long, wIns As Integer, wSoc As Integer, wNom As String
   If KeyAscii = 13 Then
      If Len(Trim(txtIns.Text)) = 0 Then
         MsgBox "Institución En Blanco", vbExclamation
         txtIns.Text = ""
         Exit Sub
      End If
      wSoc = Val(txtCodigo.Text)
      wCod = Val(txtCodofin.Text)
      wIns = Val(txtIns.Text)
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(wCod) + " AND INS = " + Str(wIns) + " AND CODSOCIO <> " + Str(wSoc) + " ")
      If aa > 0 Then
         wNom = Trim(ADO8!nombre)
         MsgBox "Codofin Ya Existe En Socio" + vbNewLine + wNom, vbExclamation
         txtCodofin.Text = ""
         txtIns.Text = ""
         txtCodofin.SetFocus
         Exit Sub
      End If
      Set ADO8 = Nothing
      
      txtNumdoc.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtMesFall_GotFocus()
   txtMesFall.SelStart = 0
   txtMesFall.SelLength = 6
End Sub

Private Sub txtMesFall_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtPromocion.SetFocus
   Case 40
        txtFecFall.SetFocus
   End Select
End Sub

Private Sub txtMesFall_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFecFall.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtNombre_GotFocus()
    txtNombre.SelStart = 0
    txtNombre.SelLength = 50
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
   Dim aa As Integer, wSoc As Integer, wCod As Long, wIns As Integer, wNom As String
   
   If KeyAscii = 13 Then
      If Len(Trim(txtNombre)) = 0 Then
         MsgBox "Nombre En Blanco", vbExclamation
         Exit Sub
      End If
      wSoc = Val(txtCodigo.Text)
      wCod = Val(txtCodofin.Text)
      wIns = Val(txtIns.Text)
      wNom = Trim(txtNombre.Text)
      aa = Leerado8("SELECT * FROM MAESOCIO " _
                    & " WHERE NOMBRE LIKE '" + wNom + "%' AND " _
                    & "       (CODSOCIO <> " + Str(wSoc) + " OR " _
                    & "        CODIGO <> " + Str(wCod) + " AND " _
                    & "        INS <> " + Str(wIns) + " ) ")
      If aa > 0 Then
         MsgBox "Nombre Digitado Ya existe", vbExclamation
         txtNombre.Text = ""
         Exit Sub
      End If
      
      cmbGrado.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtNumdoc_GotFocus()
   txtNumdoc.SelStart = 0
   txtNumdoc.SelLength = 8
End Sub

Private Sub txtNumDoc_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtIns.SetFocus
   Case 40
        txtNombre.SetFocus
   End Select
End Sub

Private Sub txtNumdoc_KeyPress(KeyAscii As Integer)
   Dim aa As Integer, wSoc As Integer, wDni As String, wNom As String
   
   If KeyAscii = 13 Then
      If Len(Trim(txtNumdoc.Text)) = 0 Then
         MsgBox "DNI Esta En Blanco", vbExclamation
         txtNumdoc.Text = ""
         Exit Sub
      End If
      wSoc = Val(txtCodigo.Text)
      wDni = Trim(txtNumdoc.Text)
      
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO <> " + Str(wSoc) + " AND NUMDOC = '" + wDni + "' ")
      If aa > 0 Then
         wNom = ADO8!nombre
         MsgBox "DNI Ya Existe en Socio " + vbNewLine + wNom, vbExclamation
         txtNumdoc.Text = ""
         Exit Sub
      End If
      txtNombre.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtNumReso_GotFocus()
   txtNumReso.SelStart = 0
   txtNumReso.SelLength = 10
End Sub

Private Sub txtNumReso_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFecReso.SetFocus
   Case 40
        txtNumResoR.SetFocus
   End Select
End Sub

Private Sub txtNumReso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtDirectivo.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtNumResoR_GotFocus()
   txtNumResoR.SelStart = 0
   txtNumResoR.SelLength = 10
End Sub

Private Sub txtNumResoR_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtNumReso.SetFocus
   Case 40
        txtDirectivo.SetFocus
   End Select
End Sub

Private Sub txtNumResoR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtDirectivo.SetFocus
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
        txtFecFall.SetFocus
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

Private Sub txtPromocion_Change()
   Dim aa As Integer
   If Len(Trim(txtPromocion.Text)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEPROMOCION WHERE PROMOCION = '" + txtPromocion + "' ")
      If aa > 0 Then
         lblPromocion.Caption = ADO8!nombre
      Else
         lblPromocion.Caption = ""
      End If
      Set ADO8 = Nothing
   Else
      lblPromocion.Caption = ""
   End If
End Sub

Private Sub txtPromocion_GotFocus()
   txtPromocion.SelStart = 0
   txtPromocion.SelLength = 10
End Sub

Private Sub txtPromocion_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtAnoDirec.SetFocus
   Case 40
        txtMesFall.SetFocus
   Case 116
        xlista = "PR"
        xseleccion = ""
        frmSeleccion.Show 1
        If xseleccion <> "" Then
           txtPromocion.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtPromocion_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtPromocion.Text)) > 0 Then
         aa = Leerado8("SELECT * FROM MAEPROMOCION WHERE PROMOCION = '" + txtPromocion.Text + "' ")
         If aa = 0 Then
            MsgBox "Promocion Digitada No Existe", vbExclamation
            txtPromocion.Text = ""
            Exit Sub
         End If
      End If
      txtMesFall.SetFocus
   Else
      If InStr(1, "0123456789-" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtRefer_GotFocus()
   txtRefer.SelStart = 0
   txtRefer.SelLength = 50
End Sub

Private Sub txtRefer_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtCelular.SetFocus
   Case 40
        txteMail.SetFocus
   End Select
End Sub

Private Sub txtRefer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txteMail.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtTelefon2_GotFocus()
   txtTelefon2.SelStart = 0
   txtTelefon2.SelLength = 20
End Sub

Private Sub txtTelefon2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtTelefono.SetFocus
   Case 40
        txtCelular.SetFocus
   End Select
End Sub

Private Sub txtTelefon2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtCelular.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtTelefono_GotFocus()
   txtTelefono.SelStart = 0
   txtTelefono.SelLength = 20
End Sub

Private Sub txtTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtUbi3.SetFocus
   Case 40
        txtTelefon2.SetFocus
   End Select
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtTelefon2.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtTomo_GotFocus()
   txtTomo.SelStart = 0
   txtTomo.SelLength = 8
End Sub

Private Sub txtTomo_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFecIng.SetFocus
   Case 40
        txtFecReso.SetFocus
   End Select
End Sub

Private Sub txtTomo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFecReso.Enabled = True Then
         txtFecReso.SetFocus
      End If
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtUbi1_GotFocus()
   txtUbi1.SelStart = 0
   txtUbi1.SelLength = 2
End Sub

Private Sub txtUbi1_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDirec.SetFocus
   Case 40
        txtUbi2.SetFocus
   Case 116
        xlista = "U1"
        xseleccion = ""
        frmSeleccion.Show 1
        If xseleccion <> "" Then
           txtUbi1.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtUbi1_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtUbi1.Text)) > 0 Then
         aa = Leerado8("SELECT * FROM MAEUBIGEO WHERE CODIGO = '" + txtUbi1.Text + "0000' ")
         If aa = 0 Then
            MsgBox "Departamento Digitado No Existe", vbExclamation
            txtUbi1.Text = ""
            Exit Sub
         End If
      End If
      txtUbi2.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtUbi2_GotFocus()
   txtUbi2.SelStart = 0
   txtUbi2.SelLength = 2
End Sub

Private Sub txtUbi2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtUbi1.SetFocus
   Case 40
        txtUbi3.SetFocus
   Case 116
        xlista = "U2"
        xseleccion = txtUbi1.Text
        frmSeleccion.Show 1
        If xseleccion <> "" Then
           txtUbi2.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtUbi2_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtUbi2.Text)) > 0 Then
         aa = Leerado8("SELECT * FROM MAEUBIGEO WHERE CODIGO = '" + txtUbi1.Text + txtUbi2.Text + "00' ")
         If aa = 0 Then
            MsgBox "Provincia Digitada No Existe", vbExclamation
            txtUbi2.Text = ""
            Exit Sub
         End If
      End If
      txtUbi3.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtUbi3_Change()
   Dim aa As Integer
   aa = Leerado8("SELECT * FROM MAEUBIGEO WHERE CODIGO = '" + txtUbi1.Text + txtUbi2.Text + txtUbi3.Text + "' ")
   If aa > 0 Then
      lblUbigeo.Caption = ADO8!nombre
   Else
      lblUbigeo.Caption = ""
   End If
   Set ADO8 = Nothing
End Sub

Private Sub txtUbi3_GotFocus()
   txtUbi3.SelStart = 0
   txtUbi3.SelLength = 2
End Sub

Private Sub txtUbi3_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtUbi2.SetFocus
   Case 40
        txtTelefono.SetFocus
   Case 116
        xlista = "U3"
        xseleccion = txtUbi1.Text + txtUbi2.Text
        frmSeleccion.Show 1
        If xseleccion <> "" Then
           txtUbi3.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtUbi3_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtUbi2.Text)) > 0 Then
         aa = Leerado8("SELECT * FROM MAEUBIGEO WHERE CODIGO = '" + txtUbi1.Text + txtUbi2.Text + txtUbi3.Text + "' ")
         If aa = 0 Then
            MsgBox "Distrito Digitado No Existe", vbExclamation
            txtUbi3.Text = ""
            Exit Sub
         End If
         lblUbigeo.Caption = ADO8!nombre
      End If
      txtTelefono.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub CreateAporteAnoMes(zSoc As Integer, zAno As String, zFecIng As Date)
   On Error GoTo err
   
   Dim aa As Integer, zCod As Long, zIns As Integer, zMesIni As String, _
       zDia As Integer, zMes As Integer, zmmm As String, zFec As Date, _
       zApo As Currency, zMon As String, zE_s As String
   zDia = Day(zFecIng)
   zMes = Month(zFecIng)
   zAno = Year(zFecIng)
   
   aa = Leerado8a("SELECT S.CODSOCIO, S.CODIGO, S.INS, S.E_SOCIO, E.MONEDA, E.APORTE " _
                & " FROM MAESOCIO AS S INNER JOIN MAEE_SOCIO AS E " _
                & "   ON S.E_SOCIO = E.E_SOCIO " _
                & " WHERE CODSOCIO = " + Str(zSoc) + " ")
   If aa > 0 Then
      zE_s = ADO8a!e_socio
      zMon = ADO8a!moneda
      zApo = ADO8a!aporte
   End If
   Set ADO8a = Nothing
   
   If zDia >= 20 Then
      If zMes = 12 Then
         zMesIni = Format(zAno + 1, "00") + "/" + "/01"
         zMes = 1
         zAno = zAno + 1
      Else
         zMesIni = Format(zAno, "0000") + "/" + Format(zMes + 1, "00")
         zMes = zMes + 1
      End If
   Else
      zMesIni = Format(zAno, "0000") + "/" + Format(zMes, "00")
   End If
   
   Dim II As Integer
   
   For II = zMes To 12
       zmmm = Format(II, "00")
       zFec = Format("01/" + zmmm + "/" + Format(zAno, "0000"), "dd/mm/yyyy")
          
       aa = Leerado6a("SELECT * FROM CTASXCAB " _
                  & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
                  & "            MES = '" + Format(zAno, "0000") + "/" + zmmm + "' AND " _
                  & "       CONCEPTO = '01' ")
       If aa = 0 Then
          Db.BeginTrans
          Db.Execute ("INSERT INTO CTASXCAB " _
          & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, CARGOS, ABONOS, SDONEW ) " _
          & " VALUES " _
          & " (" + Str(zSoc) + ", '" + Format(zAno, "0000") + "/" + zmmm + "', '01', '" + zE_s + "', '" + zMon + "', " _
          & "  " + Str(zApo) + ", 0, " + Str(zApo) + " ) ")
          Db.CommitTrans
       End If
                
       aa = Leerado6a("SELECT * FROM CTASXDET " _
                  & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
                  & "            MES = '" + Format(zAno, "0000") + "/" + zmmm + "' AND " _
                  & "       CONCEPTO = '01' ")
       If aa = 0 Then
          Db.BeginTrans
          Db.Execute ("INSERT INTO CTASXDET " _
          & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
          & "  DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW) " _
          & " VALUES " _
          & " (" + Str(zSoc) + ", '" + Format(zAno, "0000") + "/" + zmmm + "', '01', '00', '', '', '', '1', " _
          & "  '" + Format(zFec, "dd/mm/yyyy") + "', 0, 0, 0, " + Str(zApo) + ", " _
          & "  0, " + Str(zApo) + " ) ")
          Db.CommitTrans
       End If
   Next
   Exit Sub
err:
   MsgBox err.Description
   Resume Next
End Sub
