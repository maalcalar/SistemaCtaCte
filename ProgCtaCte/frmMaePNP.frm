VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMaePNP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de PNP No Asociados"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   12915
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtro x Nombre"
      Height          =   615
      Left            =   360
      TabIndex        =   76
      Top             =   7320
      Width           =   8295
      Begin VB.OptionButton optTodos 
         Caption         =   "Mostrar Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Filtrar x Nombre"
         Height          =   255
         Left            =   1560
         TabIndex        =   78
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtFiltrar 
         Height          =   285
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   77
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Filtro x Nombre"
      Height          =   615
      Left            =   360
      TabIndex        =   72
      Top             =   7920
      Width           =   4095
      Begin VB.TextBox txtFiltrarCodofin 
         Height          =   285
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   75
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optFiltroCodoFin 
         Caption         =   "Filtrar x Codofin"
         Height          =   255
         Left            =   1560
         TabIndex        =   74
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optTodosCodoFin 
         Caption         =   "Mostrar Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Filtro x Nombre"
      Height          =   615
      Left            =   4560
      TabIndex        =   68
      Top             =   7920
      Width           =   4095
      Begin VB.OptionButton optTodosDni 
         Caption         =   "Mostrar Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optFiltroDni 
         Caption         =   "Filtrar x DNI"
         Height          =   255
         Left            =   1560
         TabIndex        =   70
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtFiltrarDni 
         Height          =   285
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   69
         Top             =   240
         Width           =   975
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
      Left            =   10920
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Frame FraDetalles 
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
      Height          =   3735
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   10335
      Begin VB.ComboBox cmbTipCob 
         Height          =   315
         ItemData        =   "frmMaePNP.frx":0000
         Left            =   8160
         List            =   "frmMaePNP.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   900
         Width           =   2055
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
         Left            =   600
         MaxLength       =   8
         TabIndex        =   60
         Top             =   3240
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
         Left            =   600
         MaxLength       =   8
         TabIndex        =   55
         Top             =   2960
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
         Left            =   600
         MaxLength       =   8
         TabIndex        =   50
         Top             =   2680
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
         Height          =   285
         Left            =   600
         MaxLength       =   8
         TabIndex        =   45
         Top             =   2400
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
         Left            =   600
         MaxLength       =   8
         TabIndex        =   36
         Top             =   2120
         Width           =   690
      End
      Begin VB.TextBox txtObservac 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   34
         Top             =   1380
         Width           =   8295
      End
      Begin VB.TextBox txtObserva2 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   33
         Top             =   1680
         Width           =   8295
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
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   28
         Top             =   900
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
         Left            =   3120
         MaxLength       =   8
         TabIndex        =   27
         Top             =   900
         Width           =   930
      End
      Begin VB.ComboBox cmbSitu 
         Height          =   315
         ItemData        =   "frmMaePNP.frx":0004
         Left            =   4080
         List            =   "frmMaePNP.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   900
         Width           =   2055
      End
      Begin VB.ComboBox cmbSituEsp 
         Height          =   315
         ItemData        =   "frmMaePNP.frx":0008
         Left            =   6120
         List            =   "frmMaePNP.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   900
         Width           =   2055
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   19
         Top             =   420
         Width           =   5655
      End
      Begin VB.ComboBox cmbGrado 
         Height          =   315
         ItemData        =   "frmMaePNP.frx":000C
         Left            =   120
         List            =   "frmMaePNP.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   900
         Width           =   2055
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
         Left            =   120
         MaxLength       =   8
         TabIndex        =   17
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
         Height          =   285
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   16
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
         Height          =   285
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   15
         Top             =   420
         Width           =   930
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cobro"
         Height          =   195
         Index           =   18
         Left            =   8700
         TabIndex        =   67
         Top             =   720
         Width           =   780
      End
      Begin VB.Label lblCodigo5 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1320
         TabIndex        =   64
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lblIns5 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2280
         TabIndex        =   63
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label lblSocio5 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2760
         TabIndex        =   62
         Top             =   3240
         Width           =   5655
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "5.-"
         Height          =   195
         Index           =   16
         Left            =   360
         TabIndex        =   61
         Top             =   3240
         Width           =   180
      End
      Begin VB.Label lblCodigo4 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1320
         TabIndex        =   59
         Top             =   2960
         Width           =   855
      End
      Begin VB.Label lblIns4 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2280
         TabIndex        =   58
         Top             =   2960
         Width           =   375
      End
      Begin VB.Label lblSocio4 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2760
         TabIndex        =   57
         Top             =   2960
         Width           =   5655
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "4.-"
         Height          =   195
         Index           =   15
         Left            =   360
         TabIndex        =   56
         Top             =   2960
         Width           =   180
      End
      Begin VB.Label lblCodigo3 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1320
         TabIndex        =   54
         Top             =   2680
         Width           =   855
      End
      Begin VB.Label lblIns3 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2280
         TabIndex        =   53
         Top             =   2680
         Width           =   375
      End
      Begin VB.Label lblSocio3 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2760
         TabIndex        =   52
         Top             =   2680
         Width           =   5655
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "3.-"
         Height          =   195
         Index           =   14
         Left            =   360
         TabIndex        =   51
         Top             =   2680
         Width           =   180
      End
      Begin VB.Label lblCodigo2 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1320
         TabIndex        =   49
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblIns2 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2280
         TabIndex        =   48
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label lblSocio2 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2760
         TabIndex        =   47
         Top             =   2400
         Width           =   5655
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "2.-"
         Height          =   195
         Index           =   13
         Left            =   360
         TabIndex        =   46
         Top             =   2400
         Width           =   180
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre del Asociado"
         Height          =   195
         Index           =   12
         Left            =   3240
         TabIndex        =   44
         Top             =   1935
         Width           =   1515
      End
      Begin VB.Label lblCodigo1 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1320
         TabIndex        =   43
         Top             =   2120
         Width           =   855
      End
      Begin VB.Label lblIns1 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2280
         TabIndex        =   42
         Top             =   2120
         Width           =   375
      End
      Begin VB.Label lblSocio1 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2760
         TabIndex        =   41
         Top             =   2115
         Width           =   5655
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1.-"
         Height          =   195
         Index           =   11
         Left            =   360
         TabIndex        =   40
         Top             =   2120
         Width           =   180
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         Height          =   195
         Index           =   10
         Left            =   705
         TabIndex        =   39
         Top             =   1940
         Width           =   495
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codofin"
         Height          =   195
         Index           =   9
         Left            =   1395
         TabIndex        =   38
         Top             =   1940
         Width           =   540
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ins"
         Height          =   195
         Index           =   0
         Left            =   2280
         TabIndex        =   37
         Top             =   1940
         Width           =   210
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   22
         Left            =   360
         TabIndex        =   35
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Carnet PNP"
         Height          =   195
         Index           =   6
         Left            =   2160
         TabIndex        =   32
         Top             =   720
         Width           =   840
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Carnet PIP"
         Height          =   195
         Index           =   7
         Left            =   3195
         TabIndex        =   31
         Top             =   720
         Width           =   765
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Situación Policial"
         Height          =   195
         Index           =   8
         Left            =   4440
         TabIndex        =   30
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Situación Especial"
         Height          =   195
         Index           =   23
         Left            =   6375
         TabIndex        =   29
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Apellidos y Nombres "
         Height          =   195
         Index           =   3
         Left            =   2880
         TabIndex        =   24
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Grado"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   915
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codofin"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   22
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ins"
         Height          =   195
         Index           =   4
         Left            =   1080
         TabIndex        =   21
         Top             =   240
         Width           =   210
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "D.N.I."
         Height          =   195
         Index           =   5
         Left            =   1635
         TabIndex        =   20
         Top             =   240
         Width           =   420
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
      Left            =   10560
      TabIndex        =   12
      Top             =   1800
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
         TabIndex        =   13
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
      Left            =   10560
      TabIndex        =   7
      Top             =   2520
      Width           =   2295
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
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
         Left            =   1560
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
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
      Height          =   1815
      Left            =   10560
      TabIndex        =   1
      Top             =   0
      Width           =   2295
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1320
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
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   840
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
         TabIndex        =   4
         TabStop         =   0   'False
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
         Left            =   1200
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
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
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   360
      TabIndex        =   0
      Top             =   3840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5953
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
      Caption         =   "Tabla de PNP No Asociados"
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
End
Attribute VB_Name = "frmMaePNP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Limpiar()
   txtCodofin.Text = ""
   txtIns.Text = ""
   txtNumdoc.Text = ""
   txtNombre.Text = ""
   txtCarnetPNP.Text = ""
   txtCarnetPIP.Text = ""
   txtObservac.Text = ""
   txtObserva2.Text = ""
   
   txtSocio1.Text = ""
   txtSocio2.Text = ""
   txtSocio3.Text = ""
   txtSocio4.Text = ""
   txtSocio5.Text = ""
   
   lblCodigo1.Caption = ""
   lblCodigo2.Caption = ""
   lblCodigo3.Caption = ""
   lblCodigo4.Caption = ""
   lblCodigo5.Caption = ""
   
   lblIns1.Caption = ""
   lblIns2.Caption = ""
   lblIns3.Caption = ""
   lblIns4.Caption = ""
   lblIns5.Caption = ""
   
   lblSocio1.Caption = ""
   lblSocio2.Caption = ""
   lblSocio3.Caption = ""
   lblSocio4.Caption = ""
   lblSocio5.Caption = ""
   
   cmbGrado.ListIndex = 0
   cmbSitu.ListIndex = 0
   cmbSituEsp.ListIndex = 0
   cmbTipCob.ListIndex = 0
End Sub

Sub refrescar()
   If ADO1.BOF Then Exit Sub
   If ADO1.EOF Then Exit Sub
   txtCodofin.Text = ADO1!codigo
   txtIns.Text = ADO1!ins
   txtNumdoc.Text = ADO1!numdoc
   txtNombre.Text = IIf(IsNull(ADO1!nombre), "", ADO1!nombre)
   txtCarnetPNP.Text = IIf(IsNull(ADO1!carnetpnp), "", ADO1!carnetpnp)
   txtCarnetPIP.Text = IIf(IsNull(ADO1!carnetpip), "", ADO1!carnetpip)
   
   txtObservac.Text = IIf(IsNull(ADO1!observac), "", ADO1!observac)
   txtObserva2.Text = IIf(IsNull(ADO1!observa2), "", ADO1!observa2)
   
   txtSocio1.Text = IIf(IsNull(ADO1!codsocio1), 0, Format(ADO1!codsocio1, "#####0;;\ "))
   txtSocio2.Text = IIf(IsNull(ADO1!codsocio2), 0, Format(ADO1!codsocio2, "#####0;;\ "))
   txtSocio3.Text = IIf(IsNull(ADO1!codsocio3), 0, Format(ADO1!codsocio3, "#####0;;\ "))
   txtSocio4.Text = IIf(IsNull(ADO1!codsocio4), 0, Format(ADO1!codsocio4, "#####0;;\ "))
   txtSocio5.Text = IIf(IsNull(ADO1!codsocio5), 0, Format(ADO1!codsocio5, "#####0;;\ "))
   
   cmbGrado.ListIndex = BuscaGrado(ADO1!grado)
   cmbSitu.ListIndex = BuscaSitu(ADO1!situ)
   cmbSituEsp.ListIndex = BuscaSituEsp(ADO1!situesp)
   cmbTipCob.ListIndex = BuscaTipCob(ADO1!tipcob)
End Sub

Sub grabar()
   On Error GoTo err
   
   Dim aa As Integer, _
       wCod As Long, wIns As Integer, wNom As String, wGrado As Integer, _
       wCarPNP As Long, wCarPIP As String, _
       wSitua As Integer, wSituEsp As Integer, _
       wTipCob As String, wNumDoc As String, _
       wNomGra As String, wNomSitu As String, wNomEso As String, wNomCob As String, _
       wObservac As String, wObserva2 As String, _
       wSocio1 As Integer, wSocio2 As Integer, wSocio3 As Integer, _
       wSocio4 As Integer, wSocio5 As Integer
   
   wCod = Val(txtCodofin.Text)
   wIns = Val(txtIns.Text)
   wNom = txtNombre.Text
   wCarPNP = Val(txtCarnetPNP.Text)
   wCarPIP = txtCarnetPIP.Text
   wNumDoc = txtNumdoc.Text
   wNomGra = cmbGrado.Text
   wNomSitu = cmbSitu.Text
   wNomCob = cmbTipCob.Text
   wObservac = txtObservac.Text
   wObserva2 = txtObserva2.Text
   wSocio1 = Val(txtSocio1.Text)
   wSocio2 = Val(txtSocio2.Text)
   wSocio3 = Val(txtSocio3.Text)
   wSocio4 = Val(txtSocio4.Text)
   wSocio5 = Val(txtSocio5.Text)
   
   wGrado = BuscaCodGrado(cmbGrado.List(cmbGrado.ListIndex))
   wSitua = BuscaCodSitu(cmbSitu.List(cmbSitu.ListIndex))
   wSituEsp = BuscaCodSituEsp(cmbSituEsp.List(cmbSituEsp.ListIndex))
   wTipCob = BuscaCodTipCob(cmbTipCob.List(cmbTipCob.ListIndex))
   
   If Len(Trim(wNom)) = 0 Then
      MsgBox "Nombre En Blanco", vbExclamation
      Exit Sub
   End If
   
   aa = Leerado8("SELECT * FROM MAEPNP " _
                & " WHERE CODIGO = " + Str(Val(wCod)) + " AND " _
                & "          INS = " + Str(Val(wIns)) + " ")
   If aa = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO MAEPNP " _
      & " (CODIGO, INS, NUMDOC, CARNETPNP, CARNETPIP, NOMBRE, GRADO, " _
      & "  SITU, SITUESP, TIPCOB, OBSERVAC, OBSERVA2, CODSOCIO1, " _
      & "  CODSOCIO2, CODSOCIO3, CODSOCIO4, CODSOCIO5 ) " _
      & " VALUES " _
      & " (" + Str(wCod) + ", " + Str(wIns) + ", '" + wNumDoc + "', " + Str(wCarPNP) + ", " _
      & "  '" + wCarPIP + "', '" + wNom + "', " + Str(wGrado) + ", " + Str(wSitua) + ", " _
      & "  " + Str(wSituEsp) + ", '" + wTipCob + "', " _
      & "  '" + wObservac + "', '" + wObserva2 + "', " + Str(wSocio1) + ", " _
      & "  " + Str(wSocio2) + ", " + Str(wSocio3) + ", " + Str(wSocio4) + ", " _
      & "  " + Str(wSocio5) + " ) ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAEPNP " _
      & " SET    NOMBRE = '" + wNom + "', NUMDOC = '" + wNumDoc + "', " _
      & "     CARNETPNP = " + Str(wCarPNP) + ", CARNETPIP = '" + wCarPIP + "', " _
      & "         GRADO = " + Str(wGrado) + ", SITU = " + Str(wSitua) + ", " _
      & "       SITUESP = " + Str(wSituEsp) + ", TIPCOB = '" + wTipCob + "', " _
      & "      OBSERVAC = '" + wObservac + "', OBSERVA2 = '" + wObserva2 + "', " _
      & "      CODSOCIO1 = " + Str(wSocio1) + ", CODSOCIO2 = " + Str(wSocio2) + ", " _
      & "      CODSOCIO3 = " + Str(wSocio3) + ", CODSOCIO4 = " + Str(wSocio4) + ", " _
      & "      CODSOCIO5 = " + Str(wSocio5) + " " _
      & " WHERE CODIGO = " + Str(Val(wCod)) + " AND " _
      & "          INS = " + Str(Val(wIns)) + " ")
      Db.CommitTrans
   End If
   Set ADO8 = Nothing
   
   aa = Leerado8("SELECT * FROM TMP_MAEPNP " _
                & " WHERE CODIGO = " + Str(Val(wCod)) + " AND " _
                & "          INS = " + Str(Val(wIns)) + " AND " _
                & "          USU = '" + wcodusu + "' ")
   If aa = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO TMP_MAEPNP " _
      & " (CODIGO, INS, NUMDOC, CARNETPNP, CARNETPIP, NOMBRE, GRADO, " _
      & "  SITU, SITUESP, TIPCOB, NOMGRA, NOMSITU, NOMCOB, " _
      & "  OBSERVAC, OBSERVA2, CODSOCIO1, CODSOCIO2, " _
      & "  CODSOCIO3, CODSOCIO4, CODSOCIO5, USU ) " _
      & " VALUES " _
      & " (" + Str(wCod) + ", " + Str(wIns) + ", '" + wNumDoc + "', " + Str(wCarPNP) + ", " _
      & "  '" + wCarPIP + "', '" + wNom + "', " + Str(wGrado) + ", " + Str(wSitua) + ", " _
      & "  " + Str(wSituEsp) + ", '" + wTipCob + "', " _
      & "  '" + wNomGra + "', '" + wNomSitu + "', '" + wNomCob + "', " _
      & "  '" + wObservac + "', '" + wObserva2 + "', " + Str(wSocio1) + ", " _
      & "  " + Str(wSocio2) + ", " + Str(wSocio3) + ", " + Str(wSocio4) + ", " _
      & "  " + Str(wSocio5) + ", '" + wcodusu + "' ) ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAEPNP " _
      & " SET NOMBRE = '" + wNom + "', INS = " + Str(wIns) + ", NUMDOC = '" + wNumDoc + "', " _
      & "     CARNETPNP = " + Str(wCarPNP) + ", CARNETPIP = '" + wCarPIP + "', " _
      & "     GRADO = " + Str(wGrado) + ", SITU = " + Str(wSitua) + ", SITUESP = " + Str(wSituEsp) + ", " _
      & "     TIPCOB = '" + wTipCob + "', " _
      & "     NOMGRA = '" + wNomGra + "', NOMSITU = '" + wNomSitu + "', " _
      & "     NOMCOB  = '" + wNomCob + "', " _
      & "     OBSERVAC = '" + wObservac + "', OBSERVA2 = '" + wObserva2 + "', " _
      & "      CODSOCIO1 = " + Str(wSocio1) + ", CODSOCIO2 = " + Str(wSocio2) + ", " _
      & "      CODSOCIO3 = " + Str(wSocio3) + ", CODSOCIO4 = " + Str(wSocio4) + ", " _
      & "      CODSOCIO5 = " + Str(wSocio5) + " " _
      & " WHERE CODIGO = " + Str(Val(wCod)) + " AND " _
      & "          INS = " + Str(Val(wIns)) + " AND " _
      & "          USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   Set ADO8 = Nothing
   
   ADO1.Requery
   LlenaCab1
   ADO1.Find "CODIGO = " + Str(Val(wCod)) + " "
   ADO1.Find "   INS = " + Str(Val(wIns)) + " "
   MsgBox "Codofin " + Str(wCod) + " " + wNom + vbNewLine + _
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

Private Function BuscaGrado(zCod As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer
   zRes = 0
   zz = Leerado5a("SELECT * FROM MAEGRADO WHERE GRADO = " + Str(Val(zCod)) + " ")
   If zz > 0 Then
      zRes = ADO5a!num - 1
   End If
   
   BuscaGrado = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Private Function BuscaCodGrado(zCod As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer
   zRes = 0
   zz = Leerado5a("SELECT * FROM MAEGRADO WHERE NOMBRE LIKE '" + zCod + "%' ")
   If zz > 0 Then
      zRes = ADO5a!grado
   End If
   
   BuscaCodGrado = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Private Function BuscaSitu(zSitu As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer
   zRes = 0
   zz = Leerado5a("SELECT * FROM MAESITU WHERE SITU = " + Str(Val(zSitu)) + " ")
   If zz > 0 Then
      zRes = ADO5a!situ
   End If
   
   BuscaSitu = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Private Function BuscaSituEsp(zSituEsp As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer
   zRes = 0
   zz = Leerado5a("SELECT * FROM MAESITUESP WHERE SITUESP = " + Str(Val(zSituEsp)) + " ")
   If zz > 0 Then
      zRes = ADO5a!situesp
   End If
   
   BuscaSituEsp = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Private Function BuscaCodSitu(zSitu As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer
   zRes = 0
   zz = Leerado5a("SELECT * FROM MAESITU WHERE NOMBRE LIKE '" + Trim(zSitu) + "%' ")
   If zz > 0 Then
      zRes = ADO5a!situ
   End If
   
   BuscaCodSitu = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Private Function BuscaCodSituEsp(zSituEsp As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer
   zRes = 0
   zz = Leerado5a("SELECT * FROM MAESITUESP WHERE NOMBRE LIKE '" + Trim(zSituEsp) + "%' ")
   If zz > 0 Then
      zRes = ADO5a!situesp
   End If
   
   BuscaCodSituEsp = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Private Function BuscaTipCob(zCod As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer
   zRes = 0
   zz = Leerado5a("SELECT * FROM MAETIPCOB WHERE TIPCOB = '" + zCod + "' ")
   If zz > 0 Then
      zRes = ADO5a!num - 1
   End If
   
   BuscaTipCob = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Private Function BuscaCodTipCob(zCod As String) As String
   On Error GoTo err
   
   Dim zz As Integer, zRes As String
   zRes = ""
   zz = Leerado5a("SELECT * FROM MAETIPCOB WHERE NOMBRE LIKE '" + Trim(zCod) + "%' ")
   If zz > 0 Then
      zRes = ADO5a!tipcob
   End If
   
   BuscaCodTipCob = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Private Sub cmbGrado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtCarnetPNP.SetFocus
   End If
End Sub

Private Sub cmbSitu_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmbSituEsp.SetFocus
   End If
End Sub

Private Sub cmbSituEsp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmbTipCob.SetFocus
   End If
End Sub

Private Sub cmbTipCob_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtObservac.SetFocus
   End If
End Sub

Private Sub cmdCerrar_Click()
   Unload Me
End Sub

Private Sub cmdEliminar_Click()
   On Error GoTo err
   
   Dim wCod As Long, wIns As Integer, wNom As String, _
       wCodNew As Long, wInsNew As Integer, aa As Integer
   wCod = ADO1!codigo
   wIns = ADO1!ins
   wNom = Trim(ADO1!nombre)
   wCodNew = 0: wInsNew = 0
   ADO1.MoveNext
   If Not ADO1.EOF Then
      wCodNew = ADO1!codigo
      wInsNew = ADO1!ins
   Else
      ADO1.MovePrevious
      ADO1.MovePrevious
      If ADO1.BOF Then
         wCodNew = 0: wInsNew = 0
      Else
         wCodNew = ADO1!codigo
         wInsNew = ADO1!ins
      End If
   End If
   
   If MsgBox("¿Esta seguro de borrar Codigo " + Format(wcon, "#######0") + "?", vbYesNo + vbDefaultButton2 + vbQuestion, "Advertencia") = vbYes Then
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEPNP " _
      & " WHERE CODIGO = " + Str(wCod) + " AND " _
      & "          INS = " + Str(wIns) + " ")
      Db.CommitTrans
      
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_MAEPNP " _
      & " WHERE CODIGO = " + Str(wCod) + " AND " _
      & "          INS = " + Str(wIns) + " AND " _
      & "          USU = '" + wcodusu + "' ")
      Db.CommitTrans
      
      ADO1.Requery
      LlenaCab
      LlenaCab1
      Limpiar
      refrescar
      
      MsgBox "PNP " + Str(wCod) + "-" + Str(wIns) + " " + wNom + vbNewLine + _
             "Eliminado OK", vbExclamation
      
      If wCodNew <> 0 Then
         ADO1.Find "CODIGO=" + Str(Val(wCodNew)) + ""
         ADO1.Find "   INS=" + Str(Val(wInsNew)) + ""
      End If
   End If
   ACCION = 0
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdGrabar_Click()
   On Error GoTo err
   Dim wCod As Integer, wIns As Integer
   If ACCION = 1 Then
      wSoc = Val(txtCodigo.Text)
      If Leerado2("SELECT * FROM MAEPNP " _
                & " WHERE CODIGO = " + Str(wCod) + " AND " _
                & "          INS = " + Str(wIns) + " ") > 0 Then
         MsgBox "Codofin Ya Existe", vbExclamation
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
   txtCodofin.Enabled = False
   txtIns.Enabled = False
   txtNumdoc.SetFocus
End Sub

Private Sub cmdNuevo_Click()
   ACCION = 1
   editar True
   Limpiar
   
   txtCodofin.SetFocus
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If ACCION = 0 Then
      Limpiar
      refrescar
   End If
End Sub

Private Sub Form_Activate()
   frmMaePNP.Left = (Screen.Width - Width) \ 2
   frmMaePNP.Top = 0
   optTodos.Value = True
   
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
   Set ADO1 = Nothing
End Sub

Private Sub LlenaCab()
   Dim a As Integer
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_MAEPNP WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_MAEPNP " _
   & " (CODIGO, INS, NUMDOC, CARNETPNP, CARNETPIP, NOMBRE, GRADO, SITU, SITUESP, " _
   & "  TIPCOB, OBSERVAC, OBSERVA2, " _
   & "  CODSOCIO1, CODSOCIO2, CODSOCIO3, CODSOCIO4, CODSOCIO5, USU ) " _
   & " SELECT " _
   & "  CODIGO, INS, NUMDOC, CARNETPNP, CARNETPIP, NOMBRE, GRADO, SITU, SITUESP, " _
   & "  TIPCOB, OBSERVAC, OBSERVA2, " _
   & "  CODSOCIO1, CODSOCIO2, CODSOCIO3, CODSOCIO4, CODSOCIO5, " _
   & "  '" + wcodusu + "' " _
   & " FROM MAEPNP ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_MAEPNP " _
   & " SET NOMGRA = M.NOMBRE " _
   & " FROM TMP_MAEPNP AS T INNER JOIN MAEGRADO AS M " _
   & "   ON T.GRADO = M.GRADO " _
   & " WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_MAEPNP " _
   & " SET NOMSITU = M.NOMBRE " _
   & " FROM TMP_MAEPNP AS T INNER JOIN MAESITU AS M " _
   & "   ON T.SITU = M.SITU " _
   & " WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_MAEPNP " _
   & " SET NOMSITUESP = M.NOMBRE " _
   & " FROM TMP_MAEPNP AS T INNER JOIN MAESITUESP AS M " _
   & "   ON T.SITUESP = M.SITUESP " _
   & " WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_MAEPNP " _
   & " SET NOMCOB = M.NOMBRE " _
   & " FROM TMP_MAEPNP AS T INNER JOIN MAETIPCOB AS M " _
   & "   ON T.TIPCOB = M.TIPCOB " _
   & " WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   txtFiltrar.Text = ""
   txtFiltrarCodofin.Text = ""
   
   a = Leerado("SELECT CODIGO, INS, NUMDOC, NOMBRE, NOMGRA, NOMSITU, NOMCOB, " _
                & "    CARNETPNP, CARNETPIP, TIPCOB, SITU, GRADO, " _
                & "    OBSERVAC, OBSERVA2, SITUESP, NOMSITUESP, " _
                & "    CODSOCIO1, CODSOCIO2, CODSOCIO3, CODSOCIO4, CODSOCIO5, USU  " _
                & " FROM TMP_MAEPNP " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY CODIGO, INS ")
   Set DataGrid1.DataSource = ADO1
End Sub

Private Sub LlenaCab1()
   DataGrid1.Columns(0).Width = 800
   DataGrid1.Columns(0).Alignment = dbgRight
   DataGrid1.Columns(0).Caption = "CODOFIN"
    
   DataGrid1.Columns(1).Width = 400
   DataGrid1.Columns(1).Alignment = dbgCenter
   DataGrid1.Columns(1).Caption = "INS"
    
   DataGrid1.Columns(2).Width = 900
   DataGrid1.Columns(2).Alignment = dbgCenter
   DataGrid1.Columns(2).Caption = "D.N.I."
    
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Width = 4400
   DataGrid1.Columns(3).Caption = "NOMBRE"

   DataGrid1.Columns(4).Width = 1500
   DataGrid1.Columns(4).Alignment = dbgCenter
   DataGrid1.Columns(4).Caption = "GRADO"

   DataGrid1.Columns(5).Width = 1500
   DataGrid1.Columns(5).Alignment = dbgCenter
   DataGrid1.Columns(5).Caption = "SITUAC"

   DataGrid1.Columns(6).Width = 1800
   DataGrid1.Columns(6).Alignment = dbgCenter
   DataGrid1.Columns(6).Caption = "TIP.COB"

   DataGrid1.Columns(7).Visible = False
   DataGrid1.Columns(8).Visible = False
   DataGrid1.Columns(9).Visible = False
   DataGrid1.Columns(10).Visible = False
   DataGrid1.Columns(11).Visible = False
   DataGrid1.Columns(12).Visible = False
   DataGrid1.Columns(13).Visible = False
   DataGrid1.Columns(14).Visible = False
   DataGrid1.Columns(15).Visible = False
   DataGrid1.Columns(16).Visible = False
End Sub

Private Sub txtCarnetPIP_GotFocus()
   txtCarnetPIP.SelStart = 0
   txtCarnetPIP.SelLength = Len(Trim(txtCarnetPIP.Text))
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
   txtCarnetPNP.SelLength = Len(Trim(txtCarnetPNP.Text))
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
   If KeyAscii = 13 Then
      If Len(Trim(txtIns.Text)) = 0 Then
         MsgBox "Institución En Blanco", vbExclamation
         txtIns.Text = ""
         Exit Sub
      End If
      
      txtNumdoc.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub
Private Sub txtNombre_GotFocus()
    txtNombre.SelStart = 0
    txtNombre.SelLength = Len(Trim(txtNombre))
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txtNombre)) = 0 Then
         MsgBox "Nombre En Blanco", vbExclamation
         Exit Sub
      End If
      cmbGrado.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtNumDoc_GotFocus()
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

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtNombre.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtObserva2_GotFocus()
   txtObserva2.SelStart = 0
   txtObserva2.SelLength = Len(Trim(txtObserva2.Text))
End Sub

Private Sub txtObserva2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtObservac.SetFocus
   Case 40
      
   End Select
End Sub

Private Sub txtObserva2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtSocio1.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtObservac_GotFocus()
   txtObservac.SelStart = 0
   txtObservac.SelLength = Len(Trim(txtObservac.Text))
End Sub

Private Sub txtObservac_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        cmbTipCob.SetFocus
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

Private Sub txtSocio1_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio1.Text)) + " ")
   If aa > 0 Then
      lblCodigo1.Caption = ADO6a!codigo
      lblIns1.Caption = ADO6a!ins
      lblSocio1.Caption = ADO6a!nombre
   Else
      lblCodigo1.Caption = ""
      lblIns1.Caption = ""
      lblSocio1.Caption = ""
   End If
   Set ADO6a = Nothing
End Sub

Private Sub txtSocio1_GotFocus()
   txtSocio1.SelStart = 0
   txtSocio1.SelLength = 8
End Sub

Private Sub txtSocio1_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
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
      If Len(Trim(txtSocio1.Text)) = 0 Then
         MsgBox "Codigo Socio En Blanco", vbExclamation
         txtSocio1.Text = ""
         Exit Sub
      End If
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio1.Text)) + " ")
      If aa = 0 Then
         MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
         txtSocio1.Text = ""
         Exit Sub
      End If
      lblCodigo1.Caption = ADO8!codigo
      lblIns1.Caption = ADO8!ins
      lblSocio1.Caption = ADO8!nombre
      
      txtSocio2.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtSocio2_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio2.Text)) + " ")
   If aa > 0 Then
      lblCodigo2.Caption = ADO6a!codigo
      lblIns2.Caption = ADO6a!ins
      lblSocio2.Caption = ADO6a!nombre
   Else
      lblCodigo2.Caption = ""
      lblIns2.Caption = ""
      lblSocio2.Caption = ""
   End If
   Set ADO6a = Nothing
End Sub

Private Sub txtSocio2_GotFocus()
   txtSocio2.SelStart = 0
   txtSocio2.SelLength = 8
End Sub

Private Sub txtSocio2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
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
      End If
      txtSocio3.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtSocio3_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio3.Text)) + " ")
   If aa > 0 Then
      lblCodigo3.Caption = ADO6a!codigo
      lblIns3.Caption = ADO6a!ins
      lblSocio3.Caption = ADO6a!nombre
   Else
      lblCodigo3.Caption = ""
      lblIns3.Caption = ""
      lblSocio3.Caption = ""
   End If
   Set ADO6a = Nothing
End Sub

Private Sub txtSocio3_GotFocus()
   txtSocio3.SelStart = 0
   txtSocio3.SelLength = 8
End Sub

Private Sub txtSocio3_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
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
      End If
      
      txtSocio4.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtSocio4_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio4.Text)) + " ")
   If aa > 0 Then
      lblCodigo4.Caption = ADO6a!codigo
      lblIns4.Caption = ADO6a!ins
      lblSocio4.Caption = ADO6a!nombre
   Else
      lblCodigo4.Caption = ""
      lblIns4.Caption = ""
      lblSocio4.Caption = ""
   End If
   Set ADO6a = Nothing
End Sub

Private Sub txtSocio4_GotFocus()
   txtSocio4.SelStart = 0
   txtSocio4.SelLength = 8
End Sub

Private Sub txtSocio4_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
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
      End If
      
      txtSocio5.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtSocio5_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio5.Text)) + " ")
   If aa > 0 Then
      lblCodigo5.Caption = ADO6a!codigo
      lblIns5.Caption = ADO6a!ins
      lblSocio5.Caption = ADO6a!nombre
   Else
      lblCodigo5.Caption = ""
      lblIns5.Caption = ""
      lblSocio5.Caption = ""
   End If
   Set ADO6a = Nothing
End Sub

Private Sub txtSocio5_GotFocus()
   txtSocio5.SelStart = 0
   txtSocio5.SelLength = 8
End Sub

Private Sub txtSocio5_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
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
      End If
      
      cmdGrabar.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub
