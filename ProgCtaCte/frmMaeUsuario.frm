VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMaeUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Usuarios"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   14220
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
      Left            =   11640
      TabIndex        =   32
      Top             =   1920
      Width           =   2535
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
         TabIndex        =   36
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
         Left            =   1320
         TabIndex        =   35
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
         Left            =   720
         TabIndex        =   34
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
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
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
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   11640
      TabIndex        =   26
      Top             =   120
      Width           =   2535
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
         Left            =   1320
         TabIndex        =   31
         Top             =   240
         Width           =   1095
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
         Left            =   1320
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   720
         Width           =   1095
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
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
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
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   720
         Width           =   1095
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
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtro"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   6840
      Width           =   8775
      Begin VB.TextBox txtFiltrar 
         Height          =   285
         Left            =   4080
         MaxLength       =   50
         TabIndex        =   11
         Top             =   240
         Width           =   4335
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Filtrar x Nombre"
         Height          =   255
         Left            =   2280
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Mostrar Todos"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
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
      TabIndex        =   2
      Top             =   120
      Width           =   11415
      Begin VB.CheckBox chkCtaCte 
         Alignment       =   1  'Right Justify
         Caption         =   "Ctas.Ctes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         TabIndex        =   56
         Top             =   3000
         Width           =   1530
      End
      Begin VB.TextBox txtCia05 
         Height          =   285
         Left            =   1440
         TabIndex        =   52
         Tag             =   "0"
         Top             =   3000
         Width           =   375
      End
      Begin VB.TextBox txtCia04 
         Height          =   285
         Left            =   1440
         TabIndex        =   49
         Tag             =   "0"
         Top             =   2700
         Width           =   375
      End
      Begin VB.TextBox txtCia03 
         Height          =   285
         Left            =   1440
         TabIndex        =   46
         Tag             =   "0"
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox txtCia02 
         Height          =   285
         Left            =   1440
         TabIndex        =   43
         Tag             =   "0"
         Top             =   2100
         Width           =   375
      End
      Begin VB.TextBox txtTipo 
         Height          =   285
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   40
         Top             =   1500
         Width           =   255
      End
      Begin VB.TextBox txtCia01 
         Height          =   285
         Left            =   1440
         TabIndex        =   37
         Tag             =   "0"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Frame Frame2 
         Caption         =   "Opciones Menú Principal"
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
         Height          =   2655
         Left            =   7800
         TabIndex        =   17
         Top             =   240
         Width           =   2415
         Begin VB.CheckBox chkConsulta 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   240
            TabIndex        =   55
            Top             =   1920
            Width           =   1770
         End
         Begin VB.CheckBox chkAporte 
            Alignment       =   1  'Right Justify
            Caption         =   "Aportes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   25
            Top             =   960
            Width           =   1770
         End
         Begin VB.CheckBox chkDIECO 
            Alignment       =   1  'Right Justify
            Caption         =   "DIECO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   24
            Top             =   1200
            Width           =   1770
         End
         Begin VB.CheckBox chkEleccion 
            Alignment       =   1  'Right Justify
            Caption         =   "Elecciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   23
            Top             =   480
            Width           =   1770
         End
         Begin VB.CheckBox chkMaestro 
            Alignment       =   1  'Right Justify
            Caption         =   "Maestros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   1770
         End
         Begin VB.CheckBox chkServicios 
            Alignment       =   1  'Right Justify
            Caption         =   "Servicio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   21
            Top             =   2160
            Width           =   1770
         End
         Begin VB.CheckBox chkGestion 
            Alignment       =   1  'Right Justify
            Caption         =   "Gestión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   20
            Top             =   720
            Width           =   1770
         End
         Begin VB.CheckBox chkCajaMP 
            Alignment       =   1  'Right Justify
            Caption         =   "CajaMP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   19
            Top             =   1440
            Width           =   1770
         End
         Begin VB.CheckBox chkTesor 
            Alignment       =   1  'Right Justify
            Caption         =   "Tesoreria"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   18
            Top             =   1680
            Width           =   1770
         End
      End
      Begin VB.TextBox txtAbrev 
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
         MaxLength       =   10
         TabIndex        =   14
         Top             =   900
         Width           =   1170
      End
      Begin VB.TextBox txtPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   1200
         Width           =   1050
      End
      Begin VB.CheckBox chkSupervisor 
         Alignment       =   1  'Right Justify
         Caption         =   "Supervisor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   12
         Top             =   1200
         Width           =   1530
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   4
         Top             =   300
         Width           =   570
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   3
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Empresa 05"
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
         TabIndex        =   54
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblCia05 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   53
         Top             =   3000
         Width           =   5175
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Empresa 04"
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
         TabIndex        =   51
         Top             =   2700
         Width           =   1215
      End
      Begin VB.Label lblCia04 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   50
         Top             =   2700
         Width           =   5175
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Empresa 03"
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
         TabIndex        =   48
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblCia03 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   47
         Top             =   2400
         Width           =   5175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Empresa 02"
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
         TabIndex        =   45
         Top             =   2100
         Width           =   1215
      End
      Begin VB.Label lblCia02 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   44
         Top             =   2100
         Width           =   5175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Usuario"
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
         TabIndex        =   42
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label lblTipo 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   41
         Top             =   1500
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Empresa 01"
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
         TabIndex        =   39
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblCia01 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   38
         Top             =   1800
         Width           =   5175
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Abreviado"
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
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Password"
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
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   1215
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
      Left            =   11040
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   4895
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
      Caption         =   "TABLA DE USUARIOS"
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
      Left            =   240
      TabIndex        =   1
      Top             =   7560
      Width           =   8655
   End
End
Attribute VB_Name = "frmMaeUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACCION As Byte
Dim marca As Variant
    
Sub Limpiar()
   txtCodigo.Text = ""
   txtNombre.Text = ""
   txtAbrev.Text = ""
   txtPassword.Text = ""
   txtTipo.Text = ""
   txtCia01.Text = ""
   txtCia02.Text = ""
   txtCia03.Text = ""
   txtCia04.Text = ""
   txtCia05.Text = ""
   chkSupervisor.Value = vbUnchecked
   chkMaestro.Value = vbUnchecked
   chkEleccion.Value = vbUnchecked
   chkGestion.Value = vbUnchecked
   chkAporte.Value = vbUnchecked
   chkDIECO.Value = vbUnchecked
   chkCajaMP.Value = vbUnchecked
   chkTesor.Value = vbUnchecked
   chkConsulta.Value = vbUnchecked
   chkServicios.Value = vbUnchecked
   chkCtaCte.Value = vbUnchecked
End Sub

Sub refrescar()
   If ADOMaster.BOF Then Exit Sub
   If ADOMaster.EOF Then Exit Sub
   txtCodigo.Text = ADOMaster!codigo
   txtNombre.Text = IIf(IsNull(ADOMaster!nombre), "", ADOMaster!nombre)
   txtAbrev.Text = IIf(IsNull(ADOMaster!abrev), "", ADOMaster!abrev)
   txtPassword.Text = IIf(IsNull(ADOMaster!Password), "", ADOMaster!Password)
   txtTipo.Text = IIf(IsNull(ADOMaster!tipo), "", ADOMaster!tipo)
   txtCia01.Text = IIf(IsNull(ADOMaster!cia01), "", ADOMaster!cia01)
   txtCia02.Text = IIf(IsNull(ADOMaster!cia02), "", ADOMaster!cia02)
   txtCia03.Text = IIf(IsNull(ADOMaster!cia03), "", ADOMaster!cia03)
   txtCia04.Text = IIf(IsNull(ADOMaster!cia04), "", ADOMaster!cia04)
   txtCia05.Text = IIf(IsNull(ADOMaster!cia05), "", ADOMaster!cia05)
   If ADOMaster!SUPERVISOR = True Then
      chkSupervisor.Value = vbChecked
   Else
      chkSupervisor.Value = vbUnchecked
   End If
   
   If ADOMaster!cxc_maestro = True Then
      chkMaestro.Value = vbChecked
   Else
      chkMaestro.Value = vbUnchecked
   End If
   
   If ADOMaster!cxc_eleccion = True Then
      chkEleccion.Value = vbChecked
   Else
      chkEleccion.Value = vbUnchecked
   End If
   
   If ADOMaster!cxc_gestion = True Then
      chkGestion.Value = vbChecked
   Else
      chkGestion.Value = vbUnchecked
   End If
   If ADOMaster!cxc_aporte = True Then
      chkAporte.Value = vbChecked
   Else
      chkAporte.Value = vbUnchecked
   End If
   If ADOMaster!cxc_dieco = True Then
      chkDIECO.Value = vbChecked
   Else
      chkDIECO.Value = vbUnchecked
   End If
   If ADOMaster!cxc_cajamp = True Then
      chkCajaMP.Value = vbChecked
   Else
      chkCajaMP.Value = vbUnchecked
   End If
   If ADOMaster!cxc_tesor = True Then
      chkTesor.Value = vbChecked
   Else
      chkTesor.Value = vbUnchecked
   End If
   If ADOMaster!cxc_consulta = True Then
      chkConsulta.Value = vbChecked
   Else
      chkConsulta.Value = vbUnchecked
   End If
   
   If ADOMaster!cxc_servicios = True Then
      chkServicios.Value = vbChecked
   Else
      chkServicios.Value = vbUnchecked
   End If

   If ADOMaster!swctacte = True Then
      chkCtaCte.Value = vbChecked
   Else
      chkCtaCte.Value = vbUnchecked
   End If
End Sub

Sub grabar()
   On Error GoTo err
   
   ADOMaster!codigo = txtCodigo.Text
   ADOMaster!nombre = txtNombre.Text
   ADOMaster!abrev = txtAbrev.Text
   ADOMaster!Password = txtPassword.Text
   ADOMaster!tipo = txtTipo.Text
   ADOMaster!cia01 = txtCia01.Text
   ADOMaster!cia02 = txtCia02.Text
   ADOMaster!cia03 = txtCia03.Text
   ADOMaster!cia04 = txtCia04.Text
   ADOMaster!cia05 = txtCia05.Text
   If chkSupervisor.Value = vbChecked Then
      ADOMaster!SUPERVISOR = True
      ADOMaster!cia = "99"
      ADOMaster!ruc = wruccia
      ADOMaster!mes = wmescia
      ADOMaster!ano = wanocia
   Else
      ADOMaster!SUPERVISOR = False
      ADOMaster!cia = wcodcia
      ADOMaster!ruc = wruccia
      ADOMaster!mes = wmescia
      ADOMaster!ano = wanocia
   End If
   If chkMaestro.Value = vbChecked Then
      ADOMaster!cxc_maestro = True
   Else
      ADOMaster!cxc_maestro = False
   End If
   
   If chkEleccion.Value = vbChecked Then
      ADOMaster!cxc_eleccion = True
   Else
      ADOMaster!cxc_eleccion = False
   End If
   
   If chkGestion.Value = vbChecked Then
      ADOMaster!cxc_gestion = True
   Else
      ADOMaster!cxc_gestion = False
   End If
   
   If chkAporte.Value = vbChecked Then
      ADOMaster!cxc_aporte = True
   Else
      ADOMaster!cxc_aporte = False
   End If
   
   If chkDIECO.Value = vbChecked Then
      ADOMaster!cxc_dieco = True
   Else
      ADOMaster!cxc_dieco = False
   End If
   
   If chkCajaMP.Value = vbChecked Then
      ADOMaster!cxc_cajamp = True
   Else
      ADOMaster!cxc_cajamp = False
   End If
   
   If chkTesor.Value = vbChecked Then
      ADOMaster!cxc_tesor = True
   Else
      ADOMaster!cxc_tesor = False
   End If
   
   If chkServicios.Value = vbChecked Then
      ADOMaster!cxc_servicios = True
   Else
      ADOMaster!cxc_servicios = False
   End If
   
   If chkCtaCte.Value = vbChecked Then
      ADOMaster!swctacte = True
   Else
      ADOMaster!swctacte = False
   End If
   
   ADOMaster!estado = 1
   ADOMaster.Update
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub editar(estado As Boolean)
   txtCodigo.Enabled = estado
   txtNombre.Enabled = estado
   txtAbrev.Enabled = estado
   txtPassword.Enabled = estado
   txtTipo.Enabled = estado
   txtCia01.Enabled = estado
   txtCia02.Enabled = estado
   txtCia03.Enabled = estado
   txtCia04.Enabled = estado
   txtCia05.Enabled = estado
   chkSupervisor.Enabled = estado
   chkMaestro.Enabled = estado
   chkEleccion.Enabled = estado
   chkGestion.Enabled = estado
   chkAporte.Enabled = estado
   chkDIECO.Enabled = estado
   chkCajaMP.Enabled = estado
   chkTesor.Enabled = estado
   chkConsulta.Enabled = estado
   chkServicios.Enabled = estado
   chkCtaCte.Enabled = estado
   
   cmdNuevo.Visible = Not estado
   cmdModificar.Visible = Not estado
   cmdEliminar.Visible = Not estado
   
   DataGrid1.Enabled = Not estado
   fraDesplaza.Enabled = Not estado
   fraFiltro.Enabled = Not estado
   
   cmdGrabar.Visible = estado
   cmdDeshacer.Visible = estado
   cmdCerrar.Visible = Not estado
End Sub

Private Sub chkAporte_Click()
   chkAporte_KeyPress (13)
End Sub

Private Sub chkAporte_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If chkDIECO.Enabled = True Then
         chkDIECO.SetFocus
      End If
   End If
End Sub

Private Sub chkCajaMP_Click()
   chkCajaMP_KeyPress (13)
End Sub

Private Sub chkCajaMP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If chkTesor.Enabled = True Then
         chkTesor.SetFocus
      End If
   End If
End Sub

Private Sub chkConsulta_Click()
   chkConsulta_KeyPress (13)
End Sub

Private Sub chkConsulta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If chkServicios.Enabled = True Then
         chkServicios.SetFocus
      End If
   End If
End Sub

Private Sub chkCtaCte_Click()
   chkCtaCte_KeyPress (13)
End Sub

Private Sub chkCtaCte_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmdGrabar.Visible = True Then
         cmdGrabar.SetFocus
      End If
   End If
End Sub

Private Sub chkDIECO_Click()
   chkDIECO_KeyPress (13)
End Sub

Private Sub chkDIECO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If chkCajaMP.Enabled = True Then
         chkCajaMP.SetFocus
      End If
   End If
End Sub

Private Sub chkEleccion_Click()
   chkEleccion_KeyPress (13)
End Sub

Private Sub chkEleccion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If chkGestion.Enabled = True Then
         chkGestion.SetFocus
      End If
   End If
End Sub

Private Sub chkGestion_Click()
   chkGestion_KeyPress (13)
End Sub

Private Sub chkGestion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If chkAporte.Enabled = True Then
         chkAporte.SetFocus
      End If
   End If
End Sub

Private Sub chkMaestro_Click()
   chkMaestro_KeyPress (13)
End Sub

Private Sub chkMaestro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If chkEleccion.Enabled = True Then
         chkEleccion.SetFocus
      End If
   End If
End Sub

Private Sub chkServicios_Click()
   chkServicios_KeyPress (13)
End Sub

Private Sub chkServicios_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If chkCtaCte.Enabled = True Then
         chkCtaCte.SetFocus
      End If
   End If
End Sub

Private Sub chkSupervisor_Click()
   chkSupervisor_KeyPress (13)
End Sub

Private Sub chkSupervisor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtTipo.Enabled = True Then
         txtTipo.SetFocus
      End If
   End If
End Sub

Private Sub chkTesor_Click()
   chkTesor_KeyPress (13)
End Sub

Private Sub chkTesor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If chkConsulta.Enabled = True Then
         chkConsulta.SetFocus
      End If
   End If
End Sub

Private Sub cmdCerrar_Click()
   Unload Me
End Sub

Private Sub cmdDeshacer_Click()
   MsgBox "Los Cambios Efectuados Se Perderán", vbExclamation
   ACCION = 0
   
   editar (False)
   
   refrescar
End Sub

Private Sub cmdEliminar_Click()
   On Error GoTo err
   
   Dim wcon As String, wNom As String, wNew As String, aa As Integer
   wcon = ADOMaster!codigo
   wNom = Trim(ADOMaster!nombre)
   wNew = ""
   ADOMaster.MoveNext
   If Not ADOMaster.EOF Then
      wNew = ADOMaster!codigo
   Else
      ADOMaster.MovePrevious
      ADOMaster.MovePrevious
      If ADOMaster.BOF Then
         wNew = ""
      Else
         wNew = ADOMaster!codigo
      End If
   End If
   If MsgBox("¿Esta seguro de borrar Registro?", vbYesNo + vbDefaultButton2 + vbQuestion, "Advertencia") = vbYes Then
      DbMaster.BeginTrans
      DbMaster.Execute ("DELETE FROM USUARIOS WHERE CODIGO = '" + wcon + "' ")
      DbMaster.CommitTrans
      
      ADOMaster.Requery
      Limpiar
      LlenaCab
      LlenaCab1
      
      If wNew <> "" Then
         ADOMaster.Find "CODIGO='" + wNew + "'"
      End If
      MsgBox "Usuario " + wcon + " " + wNom + vbNewLine + _
             "Eliminado OK", vbExclamation
   End If
   ACCION = 0
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdGrabar_Click()
   On Error GoTo err
   Dim wCod As String
   If ACCION = 1 Then
      wCod = txtCodigo.Text
      If LeeradoMaster2("SELECT * FROM USUARIOS WHERE CODIGO = '" + wCod + "'") > 0 Then
         MsgBox "Codigo Ya Existe", vbExclamation
         Limpiar
         txtCodigo.SetFocus
         Exit Sub
      End If
      ADOMaster.AddNew
   End If
   grabar
   ADOMaster.Update
   editar False
   MsgBox "Usuario Grabado OK", vbExclamation
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
   txtCodigo.Enabled = False
   txtNombre.SetFocus
End Sub

Private Sub cmdMover_Click(Index As Integer)
    With ADOMaster
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
   Dim aa As Integer, wNew As String
   
   ACCION = 1
   editar True
   Limpiar
   
   wNew = "000"
   aa = LeeradoMaster3("SELECT MAX(CODIGO) AS CODIGO FROM USUARIOS WHERE CODIGO < '999'")
   If aa > 0 Then
      wNew = IIf(IsNull(ADOMaster3!codigo), "000", ADOMaster3!codigo)
   End If
   Set ADOMaster3 = Nothing
   wNew = Format(Val(wNew) + 1, "000")
   
   
   txtCodigo.Text = wNew
   txtCodigo.Enabled = False
     
   txtNombre.SetFocus
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   Select Case ColIndex
   Case 0
        ADOMaster.Sort = "CODIGO"
   Case 1
        ADOMaster.Sort = "NOMBRE"
   End Select
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If ACCION = 0 Then
      refrescar
   End If
End Sub

Private Sub Form_Activate()
   frmMaeUsuario.Left = (Screen.Width - Width) \ 2
   frmMaeUsuario.Top = 0
   optTodos.Value = True
   
   editar (False)
   LlenaCab
   LlenaCab1
   refrescar
   DataGrid1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set ADOMaster = Nothing
End Sub

Private Sub estado(sw As Boolean)
End Sub

Private Sub LlenaCab()
   Dim a As Integer
   a = LeeradoMaster("SELECT CODIGO, NOMBRE, ABREV, PASSWORD, SUPERVISOR, " _
                & "           CXC_MAESTRO, CXC_ELECCION, CXC_GESTION, CXC_APORTE, CXC_DIECO, " _
                & "           CXC_CAJAMP, CXC_TESOR, CXC_CONSULTA, CXC_SERVICIOS, SWCTACTE, ESTADO, " _
                & "           CIA, RUC, MES, ANO, TIPO, CIA01, CIA02, CIA03, " _
                & "           CIA04, CIA05 " _
                & " FROM USUARIOS " _
                & " ORDER BY CODIGO ")
   Set DataGrid1.DataSource = ADOMaster
End Sub

Private Sub LlenaCab1()
   DataGrid1.Columns(0).Alignment = dbgLeft
   DataGrid1.Columns(0).Width = 680
   DataGrid1.Columns(0).Caption = "CODIGO"
    
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Width = 3350
   DataGrid1.Columns(1).Caption = "NOMBRE"

   DataGrid1.Columns(2).Alignment = dbgLeft
   DataGrid1.Columns(2).Width = 900
   DataGrid1.Columns(2).Caption = "ABREV"
   DataGrid1.Columns(2).Visible = False
   
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Width = 850
   DataGrid1.Columns(3).Caption = "PASSWORD"
   DataGrid1.Columns(3).Visible = False

   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Width = 850
   DataGrid1.Columns(4).Caption = "SUPERVISOR"

   DataGrid1.Columns(5).Alignment = dbgLeft
   DataGrid1.Columns(5).Width = 750
   DataGrid1.Columns(5).Caption = "MAESTROS"

   DataGrid1.Columns(6).Alignment = dbgLeft
   DataGrid1.Columns(6).Width = 750
   DataGrid1.Columns(6).Caption = "ELECCION"

   DataGrid1.Columns(7).Alignment = dbgLeft
   DataGrid1.Columns(7).Width = 750
   DataGrid1.Columns(7).Caption = "GESTION"

   DataGrid1.Columns(8).Alignment = dbgLeft
   DataGrid1.Columns(8).Width = 750
   DataGrid1.Columns(8).Caption = "APORTE"

   DataGrid1.Columns(9).Alignment = dbgLeft
   DataGrid1.Columns(9).Width = 750
   DataGrid1.Columns(9).Caption = "DIECO"

   DataGrid1.Columns(10).Alignment = dbgLeft
   DataGrid1.Columns(10).Width = 750
   DataGrid1.Columns(10).Caption = "CAJAMP"

   DataGrid1.Columns(11).Alignment = dbgLeft
   DataGrid1.Columns(11).Width = 750
   DataGrid1.Columns(11).Caption = "TESOR"

   DataGrid1.Columns(12).Alignment = dbgLeft
   DataGrid1.Columns(12).Width = 750
   DataGrid1.Columns(12).Caption = "CONSULTA"

   DataGrid1.Columns(13).Alignment = dbgLeft
   DataGrid1.Columns(13).Width = 750
   DataGrid1.Columns(13).Caption = "SERVIC."

   DataGrid1.Columns(14).Alignment = dbgLeft
   DataGrid1.Columns(14).Width = 750
   DataGrid1.Columns(14).Caption = "SW"

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
End Sub

Private Sub optAmbas_Click()
   optAmbas_KeyPress (13)
End Sub

Private Sub optAmbas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmdGrabar.Enabled = True And (ACCION = 1 Or ACCION = 2) Then
         cmdGrabar.SetFocus
      End If
   End If
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

Private Sub OptLaser_Click()
   optLaser_KeyPress (13)
End Sub

Private Sub optLaser_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmdGrabar.Enabled = True Then
         cmdGrabar.SetFocus
      End If
   End If
End Sub

Private Sub optMatricial_Click()
   optMatricial_KeyPress (13)
End Sub

Private Sub optMatricial_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmdGrabar.Visible = True Then
         cmdGrabar.SetFocus
      End If
   End If
End Sub

Private Sub optTodos_Click()
   If optTodos.Value = True Then
      txtFiltrar.Text = ""
      txtFiltrar.Enabled = False
      LlenaCab
      LlenaCab1
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
         ADOMaster.Filter = ""
         Set DataGrid1.DataSource = ADOMaster
         DataGrid1.SetFocus
      Else
         txtFiltrar.Enabled = True
         optFiltro.Value = True
         txtFiltrar.SetFocus
      End If
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If PreviousTab = 0 Then
      If ACCION = 0 Then
      End If
   Else
      ACCION = 0
   End If
End Sub

Private Sub txtAbrev_GotFocus()
   txtAbrev.SelStart = 0
   txtAbrev.SelLength = Len(Trim(txtAbrev.Text))
End Sub

Private Sub txtAbrev_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtAbrev.Text = "" Then
         MsgBox "Nombre Abreviado Esta En Blanco", vbExclamation
         txtAbrev.Text = ""
         Exit Sub
      End If
      txtPassword.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtCia01_Change()
   Dim zz As Integer
   zz = LeeradoMaster3("SELECT * FROM COMPANIAS WHERE CODIGOCIA = '" + txtCia01.Text + "' ")
   If zz > 0 Then
      lblCia01.Caption = ADOMaster3!NombreCia
   Else
      lblCia01.Caption = ""
   End If
   Set ADOMaster3 = Nothing
End Sub

Private Sub txtCia01_GotFocus()
   txtCia01.SelStart = 0
   txtCia01.SelLength = Len(Trim(txtCia01.Text))
End Sub

Private Sub txtCia01_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtTipo.SetFocus
   Case 40
        txtCia02.SetFocus
   Case 116
        xlista = "01"
        frmSeleccion.Show 1
        If xseleccion <> "" Then
           txtCia01.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtCia01_KeyPress(KeyAscii As Integer)
   Dim zz As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtCia01.Text)) = 0 Then
         txtCia01.Text = ""
         Exit Sub
      End If
      zz = LeeradoMaster3("SELECT * FROM COMPANIAS WHERE CODIGOCIA = '" + txtCia01.Text + "' ")
      If zz = 0 Then
         MsgBox "Compañia Digitada No Existe", vbExclamation
         txtCia01.Text = ""
         Exit Sub
      End If
      txtCia02.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtCia02_Change()
   Dim zz As Integer
   zz = LeeradoMaster3("SELECT * FROM COMPANIAS WHERE CODIGOCIA = '" + txtCia02.Text + "' ")
   If zz > 0 Then
      lblCia02.Caption = ADOMaster3!NombreCia
   Else
      lblCia02.Caption = ""
   End If
   Set ADOMaster3 = Nothing
End Sub

Private Sub txtCia02_GotFocus()
   txtCia02.SelStart = 0
   txtCia02.SelLength = Len(Trim(txtCia02.Text))
End Sub

Private Sub txtCia02_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtCia01.SetFocus
   Case 40
        txtCia03.SetFocus
   Case 116
        xlista = "01"
        frmSeleccion.Show 1
        If xseleccion <> "" Then
           txtCia02.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtCia02_KeyPress(KeyAscii As Integer)
   Dim zz As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtCia02.Text)) <> 0 Then
         zz = LeeradoMaster3("SELECT * FROM COMPANIAS WHERE CODIGOCIA = '" + txtCia02.Text + "' ")
         If zz = 0 Then
            MsgBox "Compañia Digitada No Existe", vbExclamation
            txtCia02.Text = ""
            Exit Sub
         End If
      End If
      txtCia03.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtCia03_Change()
   Dim zz As Integer
   zz = LeeradoMaster3("SELECT * FROM COMPANIAS WHERE CODIGOCIA = '" + txtCia03.Text + "' ")
   If zz > 0 Then
      lblCia03.Caption = ADOMaster3!NombreCia
   Else
      lblCia03.Caption = ""
   End If
   Set ADOMaster3 = Nothing
End Sub

Private Sub txtCia03_GotFocus()
   txtCia03.SelStart = 0
   txtCia03.SelLength = Len(Trim(txtCia03.Text))
End Sub

Private Sub txtCia03_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtCia02.SetFocus
   Case 40
        txtCia04.SetFocus
   Case 116
        xlista = "01"
        frmSeleccion.Show 1
        If xseleccion <> "" Then
           txtCia03.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtCia03_KeyPress(KeyAscii As Integer)
   Dim zz As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtCia03.Text)) <> 0 Then
         zz = LeeradoMaster3("SELECT * FROM COMPANIAS WHERE CODIGOCIA = '" + txtCia03.Text + "' ")
         If zz = 0 Then
            MsgBox "Compañia Digitada No Existe", vbExclamation
            txtCia03.Text = ""
            Exit Sub
         End If
      End If
      txtCia04.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtCia04_Change()
   Dim zz As Integer
   zz = LeeradoMaster3("SELECT * FROM COMPANIAS WHERE CODIGOCIA = '" + txtCia04.Text + "' ")
   If zz > 0 Then
      lblCia04.Caption = ADOMaster3!NombreCia
   Else
      lblCia04.Caption = ""
   End If
   Set ADOMaster3 = Nothing
End Sub

Private Sub txtCia04_GotFocus()
   txtCia04.SelStart = 0
   txtCia04.SelLength = Len(Trim(txtCia04.Text))
End Sub

Private Sub txtCia04_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtCia03.SetFocus
   Case 40
        txtCia05.SetFocus
   Case 116
        xlista = "01"
        frmSeleccion.Show 1
        If xseleccion <> "" Then
           txtCia04.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtCia04_KeyPress(KeyAscii As Integer)
   Dim zz As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtCia04.Text)) <> 0 Then
         zz = LeeradoMaster3("SELECT * FROM COMPANIAS WHERE CODIGOCIA = '" + txtCia04.Text + "' ")
         If zz = 0 Then
            MsgBox "Compañia Digitada No Existe", vbExclamation
            txtCia04.Text = ""
            Exit Sub
         End If
      End If
      txtCia05.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtCia05_Change()
   Dim zz As Integer
   zz = LeeradoMaster3("SELECT * FROM COMPANIAS WHERE CODIGOCIA = '" + txtCia05.Text + "' ")
   If zz > 0 Then
      lblCia05.Caption = ADOMaster3!NombreCia
   Else
      lblCia05.Caption = ""
   End If
   Set ADOMaster3 = Nothing
End Sub

Private Sub txtCia05_GotFocus()
   txtCia05.SelStart = 0
   txtCia05.SelLength = Len(Trim(txtCia05.Text))
End Sub

Private Sub txtCia05_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtCia04.SetFocus
   Case 40
        chkMaestro.SetFocus
   Case 116
        xlista = "01"
        frmSeleccion.Show 1
        If xseleccion <> "" Then
           txtCia05.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtCia05_KeyPress(KeyAscii As Integer)
   Dim zz As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtCia05.Text)) <> 0 Then
         zz = LeeradoMaster3("SELECT * FROM COMPANIAS WHERE CODIGOCIA = '" + txtCia05.Text + "' ")
         If zz = 0 Then
            MsgBox "Compañia Digitada No Existe", vbExclamation
            txtCia05.Text = ""
            Exit Sub
         End If
      End If
      chkMaestro.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtCodigo_GotFocus()
    txtCodigo.SelStart = 0
    txtCodigo.SelLength = Len(Trim(txtCodigo))
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If txtCodigo = "" Then
         MsgBox "Codigo En Blanco", vbExclamation
         Exit Sub
      End If
      txtCodigo.Text = Format(txtCodigo.Text, "000")
      aa = LeeradoMaster3("SELECT * FROM USUARIOS WHERE CODIGO = '" + txtCodigo.Text + "' ")
      If aa > 0 Then
         MsgBox "Codigo de Usuario Ya Existe", vbExclamation
         txtCodigo.Text = ""
         Exit Sub
      End If
      Set ADOMaster3 = Nothing
      
      txtNombre.SetFocus
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
      ADOMaster.Filter = "NOMBRE LIKE '" & Trim(txtFiltrar) & "%' "
      refrescar
      DataGrid1.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtNombre_GotFocus()
    txtNombre.SelStart = 0
    txtNombre.SelLength = Len(Trim(txtNombre))
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtNombre = "" Then
         MsgBox "Nombre En Blanco", vbExclamation
         Exit Sub
      End If
      txtAbrev.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtPassword_GotFocus()
   txtPassword.SelStart = 0
   txtPassword.SelLength = Len(Trim(txtPassword.Text))
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtPassword.Text = "" Then
         MsgBox "Password Esta En Blanco", vbExclamation
         Exit Sub
      End If
      chkSupervisor.SetFocus
   End If
End Sub

Private Sub txtTipo_Change()
   Dim zz As Integer
   zz = LeeradoMaster3("SELECT * FROM MAETIPOUSUARIO WHERE CODIGO = '" + txtTipo.Text + "' ")
   If zz > 0 Then
      lblTipo.Caption = ADOMaster3!nombre
   Else
      lblTipo.Caption = ""
   End If
   Set ADOMaster3 = Nothing
End Sub

Private Sub txtTipo_GotFocus()
   txtTipo.SelStart = 0
   txtTipo.SelLength = Len(Trim(txtTipo.Text))
End Sub

Private Sub txtTipo_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        chkSupervisor.SetFocus
   Case 40
        If txtTipo.Text = "U" Then
           txtCia01.SetFocus
        Else
           txtCia01.Text = ""
           txtCia02.Text = ""
           txtCia03.Text = ""
           txtCia04.Text = ""
           txtCia05.Text = ""
           chkMaestro.SetFocus
        End If
   Case 116
        xlista = "02"
        frmSeleccion.Show 1
        If xseleccion <> "" Then
           txtTipo.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtTipo_KeyPress(KeyAscii As Integer)
   Dim zz As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtTipo.Text)) = 0 Then
         txtTipo.Text = "U"
         Exit Sub
      End If
      zz = LeeradoMaster3("SELECT * FROM MAETIPOUSUARIO WHERE CODIGO = '" + txtTipo.Text + "' ")
      If zz = 0 Then
         MsgBox "Tipo de Usuario No Existe", vbExclamation
         txtTipo.Text = "U"
         Exit Sub
      End If
      If txtTipo.Text = "U" Then
         txtCia01.SetFocus
      Else
         txtCia01.Text = ""
         txtCia02.Text = ""
         txtCia03.Text = ""
         txtCia04.Text = ""
         txtCia05.Text = ""
         
         chkMaestro.SetFocus
      End If
   Else
      If InStr(1, "US" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub
