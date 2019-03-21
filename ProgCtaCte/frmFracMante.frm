VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFracMante 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Fraccionamientos"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   15225
   Begin VB.Frame Frame1 
      Caption         =   "EXCEL"
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
      Height          =   855
      Left            =   12840
      TabIndex        =   68
      Top             =   3480
      Width           =   2295
      Begin VB.CommandButton Exportar 
         Caption         =   "&Exportar Relación"
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
         Left            =   120
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdExpMoroso 
         Caption         =   "&Exportar Morosos"
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
         Left            =   1200
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
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
      Left            =   13680
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8040
      Width           =   975
   End
   Begin VB.Frame FraDetalles 
      Caption         =   "Detalles Del Fraccionamiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   12615
      Begin VB.TextBox txtNumCob 
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   59
         Top             =   3495
         Width           =   1095
      End
      Begin VB.TextBox txtSerCob 
         Height          =   285
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   58
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox txtCanMes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   8
         TabIndex        =   54
         Top             =   2115
         Width           =   615
      End
      Begin VB.TextBox txtCuoMes 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         MaxLength       =   8
         TabIndex        =   52
         Top             =   2115
         Width           =   975
      End
      Begin VB.TextBox txtSdoCob 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         MaxLength       =   8
         TabIndex        =   50
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtCuoIni 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         MaxLength       =   8
         TabIndex        =   48
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtGlosa2 
         Height          =   285
         Left            =   120
         MaxLength       =   100
         TabIndex        =   47
         Top             =   2835
         Width           =   5415
      End
      Begin VB.TextBox txtGrado 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         MaxLength       =   3
         TabIndex        =   39
         Top             =   830
         Width           =   495
      End
      Begin VB.TextBox txtNumdoc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         MaxLength       =   8
         TabIndex        =   38
         Top             =   380
         Width           =   975
      End
      Begin VB.TextBox txtIns 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   37
         Top             =   380
         Width           =   375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         MaxLength       =   8
         TabIndex        =   36
         Top             =   380
         Width           =   975
      End
      Begin VB.TextBox txtE_socio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   35
         Top             =   830
         Width           =   495
      End
      Begin VB.TextBox txtGlosa1 
         Height          =   285
         Left            =   120
         MaxLength       =   100
         TabIndex        =   29
         Top             =   2550
         Width           =   5415
      End
      Begin VB.TextBox txtSdoPen 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         MaxLength       =   8
         TabIndex        =   25
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtMoneda 
         Height          =   285
         Left            =   120
         MaxLength       =   1
         TabIndex        =   24
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox txtCodSocio 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         MaxLength       =   9
         TabIndex        =   20
         Top             =   380
         Width           =   975
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   16
         Top             =   380
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   285
         Left            =   1200
         TabIndex        =   17
         Top             =   380
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   3375
         Left            =   5880
         TabIndex        =   31
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
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
         Caption         =   "Detalles"
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
      Begin MSMask.MaskEdBox txtVcmto 
         Height          =   285
         Left            =   1680
         TabIndex        =   56
         Top             =   2115
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecCob 
         Height          =   285
         Left            =   4440
         TabIndex        =   60
         Top             =   3495
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "Serie"
         Height          =   255
         Left            =   2880
         TabIndex        =   64
         Top             =   3315
         Width           =   615
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "Nro.Recibo"
         Height          =   255
         Left            =   3360
         TabIndex        =   63
         Top             =   3315
         Width           =   1215
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "Fecha"
         Height          =   255
         Left            =   4440
         TabIndex        =   62
         Top             =   3315
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "--------Datos del Cobro Inicial--------"
         Height          =   255
         Left            =   2880
         TabIndex        =   61
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label Label12 
         Caption         =   "1er Vcmto"
         Height          =   255
         Left            =   1680
         TabIndex        =   57
         Top             =   1935
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Meses"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   1935
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Cuota"
         Height          =   195
         Left            =   840
         TabIndex        =   53
         Top             =   1935
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Sdo.xCobrar"
         Height          =   195
         Left            =   4560
         TabIndex        =   51
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Cuota Inicial"
         Height          =   195
         Left            =   3480
         TabIndex        =   49
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label lblGrado 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3720
         TabIndex        =   46
         Top             =   830
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "Grado"
         Height          =   195
         Left            =   3480
         TabIndex        =   45
         Top             =   650
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "D.N.I."
         Height          =   195
         Left            =   4560
         TabIndex        =   44
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Ins"
         Height          =   195
         Left            =   3240
         TabIndex        =   43
         Top             =   200
         Width           =   375
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Codofin"
         Height          =   195
         Left            =   2280
         TabIndex        =   42
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Estado del Socio"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   650
         Width           =   1335
      End
      Begin VB.Label lblE_socio 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   600
         TabIndex        =   40
         Top             =   830
         Width           =   2655
      End
      Begin VB.Label lblFormato 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Left            =   360
         TabIndex        =   34
         Top             =   3360
         Width           =   1815
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
         Left            =   11040
         TabIndex        =   33
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Total"
         Height          =   255
         Left            =   10560
         TabIndex        =   32
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   2370
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Deuda a Fracc."
         Height          =   195
         Left            =   2160
         TabIndex        =   28
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Tipo de Moneda"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1500
         Width           =   1455
      End
      Begin VB.Label lblMoneda 
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Nombre de Socio"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblCodSocio 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1260
         Width           =   5415
      End
      Begin VB.Label Label5 
         Caption         =   "Cod.Socio"
         Height          =   195
         Left            =   3600
         TabIndex        =   21
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Numero"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   200
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Fecha"
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   200
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
      Left            =   12840
      TabIndex        =   9
      Top             =   4320
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
         Left            =   1680
         TabIndex        =   13
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
         Left            =   1200
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
      Height          =   1455
      Left            =   12840
      TabIndex        =   7
      Top             =   2040
      Width           =   2295
      Begin VB.CommandButton cmdMorosos 
         Caption         =   "&Imprimir Morosos"
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
         Left            =   1200
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdUnFrac 
         Caption         =   "&Imprimir Un Fracc."
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
         Left            =   120
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdRelacion 
         Caption         =   "&Imprimir Relación"
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
         Left            =   120
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir Cronog"
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
         Left            =   1200
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
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
      Height          =   2055
      Left            =   12840
      TabIndex        =   1
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton cmdDeshacer 
         Caption         =   "&Deshacer <ESC>"
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
         Left            =   1200
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar <F4>"
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
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar <DEL>"
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
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar <ENTER>"
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
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo <Ins>"
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
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   7858
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
      Caption         =   "RELACION DE FRACCIONAMIENTOS"
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
   Begin Crystal.CrystalReport Crys1 
      Left            =   14520
      Top             =   6240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Estado de Cuenta"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin Crystal.CrystalReport Crys2 
      Left            =   13680
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Estado de Cuenta"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
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
      Left            =   3480
      TabIndex        =   71
      Top             =   8760
      Width           =   7575
   End
End
Attribute VB_Name = "frmFracMante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACCION As Byte, wcia As String

Sub Limpiar()
   txtNumero.Text = ""
   txtFecha.Text = "__/__/____"
   txtCodSocio.Text = ""
   txtCodigo.Text = ""
   txtIns.Text = ""
   txtNumDoc.Text = ""
   txtE_socio.Text = ""
   txtGrado.Text = ""
   txtMoneda.Text = ""
   txtSdoPen.Text = ""
   txtCuoIni.Text = ""
   txtSdoCob.Text = ""
   txtCanMes.Text = ""
   txtCuoMes.Text = ""
   txtVcmto.Text = "__/__/____"
   txtSerCob.Text = ""
   txtNumCob.Text = ""
   txtFecCob.Text = "__/__/____"
   txtGlosa1.Text = ""
   txtGlosa2.Text = ""
    
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_FRACDET WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Set DataGrid2.DataSource = Nothing
'   llenadet1
End Sub

Sub refrescar()
   If ADO1.BOF Then Exit Sub
   If ADO1.EOF Then Exit Sub
   
   Dim zFec As Date, zDia As Integer, zMes As Integer, zAno As Integer, zVcm As Date
   
   txtNumero.Text = ADO1!numero
   If IsDate(ADO1!fecha) Then
      txtFecha.Text = Format(ADO1!fecha, "dd/mm/yyyy")
   
      zFec = Format(txtFecha.Text, "dd/mm/yyyy")
      zDia = Day(zFec)
      zMes = Month(zFec)
      zAno = Year(zFec)
    
      If zMes = 12 Then
         zMes = 1
         zAno = zAno + 1
      Else
         zMes = zMes + 1
      End If
   
      zVcm = fundiames(Format(zMes, "00")) + "/" + Format(zMes, "00") + "/" + Format(zAno, "0000")
      
      txtVcmto.Text = Format(zVcm, "dd/mm/yyyy")
   Else
      txtFecha.Text = "__/__/____"
      txtVcmto.Text = "__/__/____"
   End If
   
   txtCodSocio.Text = IIf(IsNull(ADO1!codsocio), "", ADO1!codsocio)
   txtCodigo.Text = IIf(IsNull(ADO1!codigo), "", ADO1!codigo)
   txtIns.Text = IIf(IsNull(ADO1!ins), "", ADO1!ins)
   txtMoneda.Text = IIf(IsNull(ADO1!moneda), "", ADO1!moneda)
   txtSdoPen.Text = IIf(IsNull(ADO1!sdopen), 0, Format(ADO1!sdopen, "#####0.00;;\ "))
   txtCuoIni.Text = IIf(IsNull(ADO1!cuoini), 0, Format(ADO1!cuoini, "#####0.00;;\ "))
   txtSdoCob.Text = IIf(IsNull(ADO1!sdocob), 0, Format(ADO1!sdocob, "#####0.00;;\ "))
   txtCanMes.Text = IIf(IsNull(ADO1!canmes), 0, Format(ADO1!canmes, "#0;;\ "))
   txtCuoMes.Text = IIf(IsNull(ADO1!cuomes), 0, Format(ADO1!cuomes, "#####0.00;;\ "))
'   If IsDate(ADO1!vcmto) Then
'      txtVcmto.Text = Format(ADO1!vcmto, "dd/mm/yyyy")
'   Else
'      txtVcmto.Text = "__/__/____"
'   End If
   txtGlosa1.Text = IIf(IsNull(ADO1!glosa1), "", ADO1!glosa1)
   txtGlosa2.Text = IIf(IsNull(ADO1!glosa2), "", ADO1!glosa2)
   txtSerCob.Text = IIf(IsNull(ADO1!sercob), "", ADO1!sercob)
   txtNumCob.Text = IIf(IsNull(ADO1!numcob), "", ADO1!numcob)
   
   If Format(Year(ADO1!feccob), "0000") >= Format(Val(wanocia) - 2, "0000") Then
      txtFecCob.Text = Format(ADO1!feccob, "dd/mm/yyyy")
   Else
      txtFecCob.Text = "__/__/____"
   End If
   
   llenadet
   llenadet1
End Sub

Sub grabar()
   On Error GoTo err
   
   Dim wSoc As Integer, wCod As Long, wIns As Integer, wNum As String, wFec As Date, _
       wMon As String, wSdoPen As Currency, wCuoIni As Currency, wSdoCob As Currency, _
       wCanMes As Integer, wCuoMes As Currency, wVcm As Date, wGlo1 As String, _
       wGlo2 As String, wSerCob As String, wNumCob As String, wFecCob As Date, _
       wNom As String, aa As Integer, wLin As String, wMes As String, wSdoNew As Currency, _
       wCargos As Currency, wAbonos As Currency, wTot As Currency, wDolar As Currency, wSoles As Currency, _
       wE_S As String
   
   wSoc = Val(txtCodSocio.Text)
   wCod = Val(txtCodigo.Text)
   wIns = Val(txtIns.Text)
   wNum = txtNumero.Text
   wE_S = txtE_socio.Text
   wNom = Trim(lblCodSocio.Caption)
   wFec = Format(txtFecha.Text, "dd/mm/yyyy")
   wMon = txtMoneda.Text
   wSdoPen = Val(txtSdoPen.Text)
   wCuoIni = Val(txtCuoIni.Text)
   wSdoCob = Val(txtSdoCob.Text)
   wCanMes = Val(txtCanMes.Text)
   wCuoMes = Val(txtCuoMes.Text)
   wVcm = Format(txtVcmto.Text, "dd/mm/yyyy")
   wGlo1 = txtGlosa1.Text
   wGlo2 = txtGlosa2.Text
   wSerCob = txtSerCob.Text
   wNumCob = txtNumCob.Text
   wTot = Val(txtSdoPen)
   If IsDate(txtFecCob.Text) Then
      wFecCob = Format(txtFecCob.Text, "dd/mm/yyyy")
   End If

   aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
   If aa > 0 Then
      wCod = ADO8!codigo
      wIns = ADO8!ins
   End If
   
   aa = Leerado8("SELECT * FROM TMP_FRACCAB " _
                & " WHERE    USU = '" + wcodusu + "' AND " _
                & "       NUMERO = '" + wNum + "' ")
   If aa = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO TMP_FRACCAB " _
      & " (NUMERO, FECHA, CODSOCIO, CODIGO, INS, NOMBRE, MONEDA, SDOPEN, CUOINI, SDOCOB, " _
      & "  CANMES, CUOMES, VCMTO, GLOSA1, GLOSA2, SERCOB, NUMCOB, FECCOB, USU ) " _
      & " VALUES " _
      & "  ('" + wNum + "', '" + Format(wFec, "dd/mm/yyyy") + "', " + Str(wSoc) + ", " _
      & "   " + Str(wCod) + ", " + Str(wIns) + ", '" + wNom + "', '" + wMon + "', " _
      & "   " + Str(wSdoPen) + ", " + Str(wCuoIni) + ", " + Str(wSdoCob) + ", " _
      & "   " + Str(wCanMes) + ", " + Str(wCuoMes) + ", '" + Format(wVcm, "dd/mm/yyyy") + "', " _
      & "   '" + wGlo1 + "', '" + wGlo2 + "', '" + wSerCob + "', '" + wNumCob + "', " _
      & "   '" + Format(wFecCob, "dd/mm/yyyy") + "', '" + wcodusu + "' ) ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_FRACCAB " _
      & " SET    FECHA = '" + Format(wFec, "dd/mm/yyyy") + "', " _
      & "     CODSOCIO = " + Str(wSoc) + ",    CODIGO = " + Str(wCod) + ", " _
      & "          INS = " + Str(wIns) + ",    NOMBRE = '" + wNom + "', " _
      & "       MONEDA = '" + wMon + "',       SDOPEN = " + Str(wSdoPen) + ", " _
      & "       CUOINI = " + Str(wCuoIni) + ", SDOCOB = " + Str(wSdoCob) + ", " _
      & "       CANMES = " + Str(wCanMes) + ", CUOMES = " + Str(wCuoMes) + ", " _
      & "        VCMTO = '" + Format(wVcm, "dd/mm/yyyy") + "', " _
      & "       GLOSA1 = '" + wGlo1 + "',       GLOSA2 = '" + wGlo2 + "', " _
      & "       SERCOB = '" + wSerCob + "',     NUMCOB = '" + wNumCob + "', " _
      & "       FECCOB = '" + Format(wFecCob, "dd/mm/yyyy") + "' " _
      & " WHERE    USU = '" + wcodusu + "' AND " _
      & "       NUMERO = '" + wNum + "' ")
      Db.CommitTrans
   End If
   
   aa = Leerado8("SELECT * FROM FRACCAB " _
                & " WHERE NUMERO = '" + wNum + "' ")
   If aa = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO FRACCAB " _
      & " (NUMERO, FECHA, CODSOCIO, CODIGO, INS, MONEDA, SDOPEN, CUOINI, SDOCOB, " _
      & "  CANMES, CUOMES, VCMTO, GLOSA1, GLOSA2, SERCOB, NUMCOB, FECCOB ) " _
      & " VALUES " _
      & "  ('" + wNum + "', '" + Format(wFec, "dd/mm/yyyy") + "', " + Str(wSoc) + ", " _
      & "   " + Str(wCod) + ", " + Str(wIns) + ", '" + wMon + "', " _
      & "   " + Str(wSdoPen) + ", " + Str(wCuoIni) + ", " + Str(wSdoCob) + ", " _
      & "   " + Str(wCanMes) + ", " + Str(wCuoMes) + ", '" + Format(wVcm, "dd/mm/yyyy") + "', " _
      & "   '" + wGlo1 + "', '" + wGlo2 + "', '" + wSerCob + "', '" + wNumCob + "', " _
      & "   '" + Format(wFecCob, "dd/mm/yyyy") + "' ) ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE FRACCAB " _
      & " SET    FECHA = '" + Format(wFec, "dd/mm/yyyy") + "', " _
      & "     CODSOCIO = " + Str(wSoc) + ", " _
      & "       CODIGO = " + Str(wCod) + ",    INS = " + Str(wIns) + ", " _
      & "    MONEDA = '" + wMon + "', " _
      & "    SDOPEN = " + Str(wSdoPen) + ",                   CUOINI = " + Str(wCuoIni) + ", " _
      & "    SDOCOB = " + Str(wSdoCob) + ",                   CANMES = " + Str(wCanMes) + ", " _
      & "    CUOMES = " + Str(wCuoMes) + ",                    VCMTO = '" + Format(wVcm, "dd/mm/yyyy") + "', " _
      & "    GLOSA1 = '" + wGlo1 + "',                        GLOSA2 = '" + wGlo2 + "', " _
      & "    SERCOB = '" + wSerCob + "',                      NUMCOB = '" + wNumCob + "', " _
      & "    FECCOB = '" + Format(wFecCob, "dd/mm/yyyy") + "' " _
      & " WHERE NUMERO = '" + wNum + "' ")
      Db.CommitTrans
   End If
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM FRACDET " _
   & " WHERE NUMERO = '" + wNum + "' ")
   Db.CommitTrans
   
   aa = Leerado8("SELECT * FROM TMP_FRACDET " _
                & " WHERE NUMERO = '" + wNum + "' AND USU = '" + wcodusu + "' ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wLin = ADO8!linea
         wVcm = Format(ADO8!vcmto, "dd/mm/yyyy")
         wCargos = ADO8!cargos
         wAbonos = ADO8!abonos
         wSdoNew = ADO8!sdonew
         wSerCob = ADO8!sercob
         wNumCob = ADO8!numcob
         If IsDate(ADO8!feccob) Then
            wFecCob = Format(ADO8!feccob, "dd/mm/yyyy")
         Else
            wFecCob = Format("01/01/1900", "dd/mm/yyyy")
         End If
         
         Db.BeginTrans
         Db.Execute ("INSERT INTO FRACDET " _
         & " (NUMERO, LINEA, VCMTO, CARGOS, ABONOS, SDONEW, " _
         & "  SERCOB, NUMCOB ) " _
         & " VALUES " _
         & " ('" + wNum + "', '" + wLin + "', '" + Format(wVcm, "dd/mm/yyyy") + "', " _
         & "   " + Str(wCargos) + ", " + Str(wAbonos) + ", " + Str(wSdoNew) + ", " _
         & "   '" + wSerCob + "', '" + wNumCob + "' ) ")
         Db.CommitTrans
                    
         If Format(wFecCob, "dd/mm/yyyy") <> Format("01/01/1900", "dd/mm/yyyy") Then
            Db.BeginTrans
            Db.Execute ("UPDATE FRACDET " _
            & " SET FECCOB = '" + Format(wFecCob, "dd/mm/yyyy") + "' " _
            & " WHERE NUMERO = '" + wNum + "' AND " _
            & "        LINEA = '" + wLin + "' ")
            Db.CommitTrans
         End If
                    
         ADO8.MoveNext
      Loop
   End If
   
   wDolar = 0: wSoles = 0
   If wMon = "S" Then
      wSoles = wTot
   Else
      wDolar = wTot
   End If
   wMes = Format(Year(wFec), "0000") + "/" + _
          Format(Month(wFec), "00")
    
'   Dim wApoOld As Currency
'   wApoOld = 0
   
'   aa = leeado6("SELECT SUM(CARGOS) AS CARGOS " _
'                & " FROM CTASXDET " _
'                & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
'                & "       CONCEPTO = '01' AND " _
'                & "       MES <= '" + wMes + "' AND " _
'                & "       CARGOS >0 ")
'   If aa > 0 Then
'      wApoOld = IIf(IsNull(ADO6!cargos), 0, ADO6!cargos)
'   End If
'   Set ADO6 = Nothing
      
'   aa = leeado6("SELECT SUM(CARGOS) AS CARGOS " _
'                & " FROM CTASXDET " _
'                & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
'                & "       CONCEPTO = '01' AND " _
'                & "       MES < '" + wMes + "' AND " _
'                & "       ABONOS > ")
'   If aa > 0 Then
'      wApoOld = wApoOld - IIf(IsNull(ADO6!cargos), 0, ADO6!cargos)
'   End If
'   Set ADO6 = Nothing
   
   
'' SE CANCELA APORTACIONES
   aa = Leerado6("SELECT * FROM CTASXCAB " _
             & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
             & "            MES = '" + wMes + "' AND " _
             & "       CONCEPTO = '01' ")
   If aa = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO CTASXCAB " _
      & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, " _
      & "  CARGOS, ABONOS, SDONEW ) " _
      & " VALUES " _
      & " (" + Str(wSoc) + ", '" + wMes + "', '01', '" + wE_S + "', '" + wMon + "', " _
      & "  0, " + Str(wTot) + ", " + Str(-wTot) + ")  ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE CTASXCAB " _
      & " SET CARGOS = 0, " _
      & "     ABONOS = " + Str(wTot) + ", " _
      & "     SDONEW = " + Str(-wTot) + " " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "            MES = '" + wMes + "' AND " _
      & "       CONCEPTO = '01' ")
      Db.CommitTrans
   End If
   
   aa = Leerado6("SELECT * FROM CTASXDET " _
             & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
             & "            MES = '" + wMes + "' AND " _
             & "       CONCEPTO = '01' AND " _
             & "         TIPCOB = '04' AND " _
             & "         SERCOB = '001' AND " _
             & "         NUMCOB = '0000000001' AND " _
             & "         LINCOB = '01'  ")
   If aa = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO CTASXDET " _
      & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
      & "  TIPCAM, DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW ) " _
      & " VALUES " _
      & " (" + Str(wSoc) + ", '" + wMes + "', '01', " _
      & "  '04', '001', '0000000001', '01', " _
      & "  '2', '" + Format(wFec, "dd/mm/yyyy") + "', 0, " _
      & "  " + Str(wDolar) + ", " + Str(wSoles) + ", " _
      & "  0, 0, " + Str(wTot) + ", " + Str(-wTot) + " ) ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE CTASXDET " _
      & " SET TIPMOV = '2', FECHA = '" + Format(wFec, "dd/mm/yyyy") + "', " _
      & "     TIPCAM = 0, DOLARE = " + Str(wDolar) + ", SOLESS = " + Str(wSoles) + ", " _
      & "     SDOOLD = 0, CARGOS = 0, ABONOS = " + Str(wTot) + ", SDONEW = " + Str(-wTot) + " " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "            MES = '" + wMes + "' AND " _
      & "       CONCEPTO = '01' AND " _
      & "         TIPCOB = '04' AND " _
      & "         SERCOB = '001' AND " _
      & "         NUMCOB = '0000000001' AND " _
      & "         LINCOB = '01'  ")
      Db.CommitTrans
   End If
         
'' SE CREA EL FRACCIONAMIENTO POR COBRAR
   
   aa = Leerado7("SELECT * FROM TMP_FRACDET " _
                & " WHERE NUMERO = '" + wNum + "' AND " _
                & "          USU = '" + wcodusu + "' ")
   If aa > 0 Then
      ADO7.MoveFirst
      Do While Not ADO7.EOF
         wMes = Format(Year(ADO7!vcmto), "0000") + "/" + _
                Format(Month(ADO7!vcmto), "00")
         wTot = ADO7!cargos
            
         aa = Leerado6("SELECT * FROM CTASXCAB " _
                   & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                   & "            MES = '" + wMes + "' AND " _
                   & "       CONCEPTO = '03' ")
         If aa = 0 Then
            Db.BeginTrans
            Db.Execute ("INSERT INTO CTASXCAB " _
            & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, " _
            & "  CARGOS, ABONOS, SDONEW ) " _
            & " VALUES " _
            & " (" + Str(wSoc) + ", '" + wMes + "', '03', '" + wE_S + "', '" + wMon + "', " _
            & "  " + Str(wTot) + ", 0, 0)  ")
            Db.CommitTrans
         Else
            Db.BeginTrans
            Db.Execute ("UPDATE CTASXCAB " _
            & " SET CARGOS = " + Str(wTot) + ", " _
            & "     ABONOS = 0 " _
            & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
            & "            MES = '" + wMes + "' AND " _
            & "       CONCEPTO = '03' ")
            Db.CommitTrans
         End If
         Set ADO6 = Nothing
   
         aa = Leerado6("SELECT * FROM CTASXDET " _
                   & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                   & "            MES = '" + wMes + "' AND " _
                   & "       CONCEPTO = '03' AND " _
                   & "         TIPCOB = '00' ")
         If aa = 0 Then
            Db.BeginTrans
            Db.Execute ("INSERT INTO CTASXDET " _
            & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
            & "  TIPCAM, DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW, OBS ) " _
            & " VALUES " _
            & " (" + Str(wSoc) + ", '" + wMes + "', '03', '00', '', '', '', '1', " _
            & "  '" + Format(wFec, "dd/mm/yyyy") + "', " _
            & "  0, 0, 0, 0, " + Str(wTot) + ", 0, 0, '')  ")
            Db.CommitTrans
         Else
            Db.BeginTrans
            Db.Execute ("UPDATE CTASXDET " _
            & " SET CARGOS = " + Str(wTot) + ", " _
            & "     ABONOS = 0 " _
            & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
            & "            MES = '" + wMes + "' AND " _
            & "       CONCEPTO = '03' AND " _
            & "         TIPCOB = '00' ")
            Db.CommitTrans
         End If
         Set ADO6 = Nothing
   
         Call ActualizaSaldos(wSoc, wMes, "03")

         ADO7.MoveNext
      Loop
   End If
   
   
   ADO1.Requery
   LlenaCab
   LlenaCab1
   
   ACCION = 0
   
   ADO1.Find "NUMERO='" + wNum + "'"
   If ADO1.EOF Then
      ADO1.MoveFirst
   End If
   Limpiar
   refrescar
   llenadet
   llenadet1
   TotalDet
   
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub editar(estado As Boolean)
   FraDetalles.Enabled = estado
      
   txtSerCob.Enabled = False
   
   cmdNuevo.Visible = Not estado
   cmdModificar.Visible = Not estado
   cmdEliminar.Visible = Not estado
   
   DataGrid1.Enabled = Not estado
   fraDesplaza.Enabled = Not estado
   
   cmdGrabar.Visible = estado
   cmdDeshacer.Visible = estado
   cmdSalir.Visible = Not estado
   cmdImprimir.Visible = Not estado
End Sub

Private Sub cmdDeshacer_Click()
   MsgBox "Los Cambios Efectuados Se Perderán", vbExclamation
   ACCION = 0
   
   editar (False)
   
   Limpiar
   llenadet
   llenadet1
   refrescar
   DataGrid1.SetFocus
End Sub

Private Sub cmdEliminar_Click()
   On Error GoTo err
   
   Dim wNum As String, wNew As String, _
       aa As Long, wMesOri As String, wMes As String, wSoc As Integer
   
   If ADO1.BOF Or ADO1.EOF Then
      Exit Sub
   End If
   wNum = ADO1!numero
   wSoc = ADO1!codsocio
   wMesOri = Format(Year(ADO1!fecha), "0000") + "/" + _
             Format(Month(ADO1!fecha), "00")
   
   If Len(Trim(txtSerCob.Text)) > 0 And _
      Len(Trim(txtNumCob.Text)) > 0 Then
      MsgBox "Fraccionamiento NO Se Puede Eliminar" + vbNewLine + _
             "Tiene Cobros Efectuados" + vbNewLine + _
             "Coordinar Con INFORMATICA", vbExclamation
      Exit Sub
   End If
   
   If MsgBox("¿Esta seguro de borrar Registro?", vbYesNo + vbDefaultButton2 + vbQuestion, "Advertencia") = vbYes Then
      
      ADO1.MoveNext
      If Not ADO1.EOF Then
         wNew = ADO1!numero
      Else
         ADO1.MovePrevious
         ADO1.MovePrevious
         If ADO1.BOF Then
            wNew = ""
         Else
            wNew = ADO1!numero
         End If
      End If
      
      aa = Leerado8("SELECT * FROM FRACDET " _
                & " WHERE NUMERO = '" + wNum + "'  " _
                & "           " _
                & " ORDER BY LINEA ")
      If aa > 0 Then
         ADO8.MoveFirst
         Do While Not ADO8.EOF
            wMes = Format(Year(ADO8!vcmto), "0000") + "/" + _
                   Format(Month(ADO8!vcmto), "00")
      
            Db.BeginTrans
            Db.Execute ("DELETE FROM CTASXCAB " _
            & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
            & "            MES = '" + wMes + "' AND " _
            & "       CONCEPTO = '03' ")
            Db.CommitTrans
         
            Db.BeginTrans
            Db.Execute ("DELETE FROM CTASXDET " _
            & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
            & "            MES = '" + wMes + "' AND " _
            & "       CONCEPTO = '03' ")
            Db.CommitTrans
         
            Call ActualizaSaldos(wSoc, wMes, "03")
      
            ADO8.MoveNext
         Loop
      End If
      
      Db.BeginTrans
      Db.Execute ("DELETE FROM FRACDET " _
      & " WHERE NUMERO = '" + wNum + "' ")
      Db.CommitTrans
         
      Db.BeginTrans
      Db.Execute ("DELETE FROM FRACCAB " _
      & " WHERE NUMERO = '" + wNum + "' ")
      Db.CommitTrans
         
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_FRACDET " _
      & " WHERE NUMERO = '" + wNum + "' AND " _
      & "          USU = '" + wcodusu + "' ")
      Db.CommitTrans
         
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_FRACCAB " _
      & " WHERE NUMERO = '" + wNum + "' AND " _
      & "          USU = '" + wcodusu + "' ")
      Db.CommitTrans
         
      Db.BeginTrans
      Db.Execute ("DELETE FROM CTASXCAB " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "            MES = '" + wMesOri + "' AND " _
      & "       CONCEPTO = '01' ")
      Db.CommitTrans
         
      Db.BeginTrans
      Db.Execute ("DELETE FROM CTASXDET " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "            MES = '" + wMesOri + "' AND " _
      & "       CONCEPTO = '01' ")
      Db.CommitTrans
         
      Call ActualizaSaldos(wSoc, wMesOri, "01")
               
      ADO1.Requery
      LlenaCab1
      If wNew <> "" Then
         ADO1.Find "NUMERO = '" + wNew + "'"
      End If
      Limpiar
      refrescar
   End If
   ACCION = 0
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdExpMoroso_Click()
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_FRACCAX WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_FRACDEX WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_FRACCAX " _
   & " (NUMERO, FECHA, CODSOCIO, CODIGO, INS, NOMBRE, MONEDA, SDOPEN, CUOINI, SDOCOB, " _
   & "  CANMES, CUOMES, VCMTO, SERCOB, NUMCOB, FECCOB, GLOSA1, GLOSA2, USU ) " _
   & " SELECT " _
   & "  C.NUMERO, C.FECHA, M.CODSOCIO, M.CODIGO, M.INS, M.NOMBRE, C.MONEDA, C.SDOPEN, " _
   & "  C.CUOINI, C.SDOCOB, C.CANMES, C.CUOMES, C.VCMTO, C.SERCOB, C.NUMCOB, C.FECCOB, " _
   & "  C.GLOSA1, C.GLOSA2, '" + wcodusu + "' " _
   & " FROM FRACCAB AS C INNER JOIN MAESOCIO AS M " _
   & "   ON C.CODSOCIO = M.CODSOCIO ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_FRACDEX " _
   & " (NUMERO, LINEA, VCMTO, CARGOS, ABONOS, SDONEW, SERCOB, NUMCOB, LINCOB, FECCOB, USU ) " _
   & " SELECT " _
   & "  NUMERO, LINEA, VCMTO, CARGOS, ABONOS, SDONEW, SERCOB, NUMCOB, LINCOB, FECCOB, '" + wcodusu + "' " _
   & " FROM FRACDET " _
   & " WHERE SDONEW > 0 AND VCMTO < GETDATE() ")
   Db.CommitTrans

   Dim aa As Integer, I As Integer, Heading(14) As String, wreg As Integer, wTot As Integer
   Dim wFec As Date, wVcm As Date, wFecCob As Date, wNum As String
   Heading(0) = "NUMERO"
   Heading(1) = "FECHA"
   Heading(2) = "SOCIO"
   Heading(3) = "CODOFIN"
   Heading(4) = "NOMBRE SOCIO"
   Heading(5) = "MONEDA"
   Heading(6) = "LIN"
   Heading(7) = "VCMTO"
   Heading(8) = "FEC.COB"
   Heading(9) = "SERIE"
   Heading(10) = "NUM.COB"
   Heading(11) = "CUOTA"
   Heading(12) = "COBROS"
   Heading(13) = " SALDO "
   Heading(14) = "SDO.ACUM"
   aa = Leerado3("SELECT * FROM TMP_FRACDEX WHERE USU = '" + wcodusu + "' ORDER BY NUMERO, LINEA ")
   If aa > 0 Then
      wTot = aa
      Set objExcel = New Excel.Application
      objExcel.SheetsInNewWorkbook = 1
      objExcel.Workbooks.Add
      With objExcel
           .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
           .Range(.Cells(3, 1), .Cells(3, 15)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 15)).Font.Bold = True
           .Cells(1, 1) = wnomcia
           .Cells(2, 1) = "RELACION DE FRACCIONAMIENTOS MOROSOS"
           For I = 1 To 15 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           objExcel.Columns("A").ColumnWidth = 12
           objExcel.Columns("B").ColumnWidth = 11
           objExcel.Columns("C").ColumnWidth = 8
           objExcel.Columns("D").ColumnWidth = 10
           objExcel.Columns("E").ColumnWidth = 50
           objExcel.Columns("F").ColumnWidth = 4
           objExcel.Columns("G").ColumnWidth = 5
           objExcel.Columns("H").ColumnWidth = 11
           objExcel.Columns("I").ColumnWidth = 11
           objExcel.Columns("J").ColumnWidth = 5
           objExcel.Columns("K").ColumnWidth = 10
           objExcel.Columns("L").ColumnWidth = 11
           objExcel.Columns("M").ColumnWidth = 11
           objExcel.Columns("N").ColumnWidth = 11
           objExcel.Columns("O").ColumnWidth = 11
      End With
      V = 4
      H = 1
      wreg = 1
      Do While Not ADO3.EOF
         wNum = ADO3!numero
         aa = Leerado4("SELECT * FROM TMP_FRACCAX WHERE NUMERO = '" + wNum + "' ")
         If aa > 0 Then
            wFec = ADO4!fecha
            objExcel.Cells(V, H + 0) = ADO4!numero
            objExcel.Cells(V, H + 1) = wFec
            objExcel.Cells(V, H + 2) = ADO4!codsocio
            objExcel.Cells(V, H + 3) = Format(ADO4!codigo, "#######0") + "-" + Format(ADO4!ins, "0")
            objExcel.Cells(V, H + 4) = IIf(IsNull(ADO4!nombre), "", ADO4!nombre)
            objExcel.Cells(V, H + 5) = ADO4!moneda
         End If
         Set ADO4 = Nothing
         
         Do While ADO3!numero = wNum
            DoEvents
            lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wreg, "####0") + " / " + Format(wTot, "####0")
            lblMensaje.Refresh
         
            objExcel.Range(objExcel.Cells(V, H + 11), objExcel.Cells(V, H + 14)).NumberFormat = "####,##0.00;;\ "
            
            objExcel.Cells(V, H + 6) = ADO3!linea
            If IsDate(ADO3!vcmto) Then
               wVcm = ADO3!vcmto
               objExcel.Cells(V, H + 7) = wVcm
            End If
            If IsDate(ADO3!feccob) Then
               wFecCob = ADO3!feccob
               objExcel.Cells(V, H + 8) = wFecCob
            End If
            objExcel.Cells(V, H + 9) = ADO3!sercob
            objExcel.Cells(V, H + 10) = ADO3!numcob
            objExcel.Cells(V, H + 11) = ADO3!cargos
            objExcel.Cells(V, H + 12) = ADO3!abonos
            objExcel.Cells(V, H + 13) = ADO3!sdonew
            objExcel.Cells(V, H + 14) = ADO3!sdogra
         
            V = V + 1
            wreg = wreg + 1
            ADO3.MoveNext
            If ADO3.EOF Then
               Exit Do
            End If
         Loop
         
         V = V + 1
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

Private Sub cmdGrabar_Click()
   On Error GoTo err
   Dim wNum As String, wNew As String, wOld As String, aa As Integer
   If ACCION = 1 Then
      wNum = txtNumero.Text
      If Leerado8("SELECT * FROM FRACCAB " _
                & " WHERE NUMERO = '" + wNum + "' ") > 0 Then
         wOld = wNum
         aa = Leerado8("SELECT MAX(NUMERO) AS NUMERO " _
                    & " FROM FRACCAB ")
         If aa > 0 Then
            wNew = Format(Val(IIf(IsNull(ADO8!numero), "000000000", ADO8!numero)) + 1, "000000000")
         End If
         txtNumero.Text = wNew
      End If
   End If
   If validaCob Then
      MsgBox "Fraccionamiento Con Errores, No Se Graba", vbExclamation
      Exit Sub
   End If
   grabar
   editar False
   MsgBox "Fraccionamiento Grabada OK", vbExclamation
   ACCION = 0
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdImprimir_Click()
   Dim wNombre As String, wDni As String, wCodofin As String, _
       wDirec As String, wDist As String, wCorreo As String, _
       wTelefono As String, wCelular As String, wSdoPen As String, _
       wCuoIni As String, wCanMes As String, wCuoMes As String, _
       wFecDia As String, wFecMes As String, wFecAno As String, _
       zz As Integer, wSoc As Integer, wCod As Long, wIns As Integer
   
   wSoc = Val(txtCodSocio.Text)
   wNombre = Trim(lblCodSocio.Caption)
   wDni = txtNumDoc.Text
   wCodofin = Trim(txtCodigo.Text) + "-" + Trim(txtIns.Text)
   wDirec = "": wDist = "": wCorreo = "": wTelefono = "": wCelular = ""
   
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
   End If
   Set ADO8 = Nothing
   
   If Len(Trim(wDist)) > 0 Then
      zz = Leerado8("SELECT * FROM MAEUBIGEO WHERE CODIGO = '" + wDist + "' ")
      If zz > 0 Then
         wDist = ADO8!nombre
      End If
      Set ADO8 = Nothing
   End If
   wSdoPen = Format(txtSdoPen.Text, "###,##0.00")
   wCuoIni = Format(txtCuoIni.Text, "###,##0.00")
   wCanMes = Format(txtCanMes.Text, "##0")
   wCuoMes = Format(txtCuoMes.Text, "###,##0.00")
   wFecDia = Format(Day(txtFecha.Text), "00")
   wFecMes = Trim(funnommes(Format(Month(txtFecha.Text), "00")))
   wFecAno = Right(Format(Year(txtFecha.Text), "0000"), 2)
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\Fraccionamiento.RPT"
   Crys1.Formulas(0) = "NOMBRE= '" + wNombre + "' "
   Crys1.Formulas(1) = "DNI= '" + wDni + "' "
   Crys1.Formulas(2) = "CODOFIN= '" + wCodofin + "' "
   Crys1.Formulas(3) = "DIREC= '" + wDirec + "' "
   Crys1.Formulas(4) = "DIST= '" + wDist + "' "
   Crys1.Formulas(5) = "CORREO= '" + wCorreo + "' "
   Crys1.Formulas(6) = "TELEFONO= '" + wTelefono + "' "
   Crys1.Formulas(7) = "CELULAR= '" + wCelular + "' "
   Crys1.Formulas(8) = "SDOPEN= '" + wSdoPen + "' "
   Crys1.Formulas(9) = "CUOINI= '" + wCuoIni + "' "
   Crys1.Formulas(10) = "CANMES= '" + wCanMes + "' "
   Crys1.Formulas(11) = "CUOMES= '" + wCuoMes + "' "
   Crys1.Formulas(12) = "FECDIA= '" + wFecDia + "' "
   Crys1.Formulas(13) = "FECMES= '" + wFecMes + "' "
   Crys1.Formulas(14) = "FECANO= '" + wFecAno + "' "
   Crys1.SelectionFormula = " {TMP_FRACDET.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdModificar_Click()
   ACCION = 2
   lblFormato.Caption = "MODIFICAR"
   DataGrid2.AllowDelete = True
   DataGrid2.AllowUpdate = True
   
   editar True
   refrescar
   llenadet
   llenadet1
'   LabelDet
   TotalDet
   
   txtNumCob.Enabled = False
   txtFecha.SetFocus
End Sub

Private Sub cmdMorosos_Click()
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_FRACCAX WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_FRACDEX WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_FRACCAX " _
   & " (NUMERO, FECHA, CODSOCIO, CODIGO, INS, NOMBRE, MONEDA, SDOPEN, CUOINI, SDOCOB, " _
   & "  CANMES, CUOMES, VCMTO, SERCOB, NUMCOB, FECCOB, GLOSA1, GLOSA2, USU ) " _
   & " SELECT " _
   & "  C.NUMERO, C.FECHA, M.CODSOCIO, M.CODIGO, M.INS, M.NOMBRE, C.MONEDA, C.SDOPEN, " _
   & "  C.CUOINI, C.SDOCOB, C.CANMES, C.CUOMES, C.VCMTO, C.SERCOB, C.NUMCOB, C.FECCOB, " _
   & "  C.GLOSA1, C.GLOSA2, '" + wcodusu + "' " _
   & " FROM FRACCAB AS C INNER JOIN MAESOCIO AS M " _
   & "   ON C.CODSOCIO = M.CODSOCIO ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_FRACDEX " _
   & " (NUMERO, LINEA, VCMTO, CARGOS, ABONOS, SDONEW, SERCOB, NUMCOB, LINCOB, FECCOB, USU ) " _
   & " SELECT " _
   & "  NUMERO, LINEA, VCMTO, CARGOS, ABONOS, SDONEW, SERCOB, NUMCOB, LINCOB, FECCOB, '" + wcodusu + "' " _
   & " FROM FRACDET " _
   & " WHERE SDONEW > 0 AND VCMTO < GETDATE() ")
   Db.CommitTrans

   Crys2.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys2.ReportFileName = xraiz + "ReportCtaCte\FracMorosos.RPT"
   Crys2.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys2.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys2.Formulas(2) = "NOMMES= 'AL " + Format(Date, "dd/mm/yyyy") + "' "
   Crys2.SelectionFormula = " {TMP_FRACDEX.USU}='" + wcodusu + "' "
   Crys2.WindowState = crptMaximized
   Crys2.Action = 1
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
   Limpiar
   refrescar
End Sub

Private Sub cmdNuevo_Click()
   ACCION = 1
   lblFormato.Caption = "NUEVO"
   DataGrid2.AllowDelete = True
   DataGrid2.AllowUpdate = True
   
   DataGrid2.Refresh
   
   Limpiar
   
   Dim wFec As Date, wVcm As Date, aa As Integer, wNew As String, _
      wDia As Integer, wMes As Integer, wAno As Integer, _
      zDia As String, zMes As String, zAno As String, zVcm As String
   wNew = "0000000000"
   aa = Leerado8("SELECT MAX(NUMERO) AS NUMERO FROM FRACCAB ")
   If aa > 0 Then
      wNew = IIf(IsNull(ADO8!numero), "0000000000", ADO8!numero)
   End If
   Set ADO8 = Nothing
   wNew = Format(Val(wNew) + 1, "0000000000")
   
   wFec = Format(Date, "dd/mm/yyyy")
   wDia = Day(wFec)
   wMes = Month(wFec)
   wAno = Year(wFec)
   
   If wMes = 12 Then
      wMes = 1
      wAno = wAno + 1
   Else
      wMes = wMes + 1
   End If
   zDia = Format(wDia, "00")
   zMes = Format(wMes, "00")
   zAno = Format(wAno, "0000")
   zVcm = Format(zDia + "/" + zMes + "/" + zAno, "dd/mm/yyyy")
   
   zVcm = Format(fundiames(zMes) + "/" + zMes + "/" + zAno, "dd/mm/yyyy")
   
   wVcm = Format(zVcm, "dd/mm/yyyy")
   
   txtNumero.Text = wNew
   txtFecha.Text = Format(wFec, "dd/mm/yyyy")
   txtCodigo.Text = ""
   txtCodSocio.Text = ""
   txtIns.Text = ""
   txtNumDoc.Text = ""
   txtE_socio.Text = ""
   txtGrado.Text = ""
   txtMoneda.Text = "S"
   txtSdoPen.Text = ""
   txtCuoIni.Text = ""
   txtSdoCob.Text = ""
   txtCanMes.Text = ""
   txtCuoMes.Text = ""
   txtVcmto.Text = Format(wVcm, "dd/mm/yyyy")
   txtGlosa1.Text = ""
   txtGlosa2.Text = ""
   txtSerCob.Text = ""
   txtNumCob.Text = ""
   txtFecCob.Text = "__/__/____"
   
   llenadet
   llenadet1
       
   editar True
   
   txtFecha.SetFocus
End Sub

Private Sub LlenaPreRelacion()
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_FRACCAX WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_FRACDEX WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_FRACCAX " _
   & " (NUMERO, FECHA, CODSOCIO, CODIGO, INS, NOMBRE, MONEDA, SDOPEN, CUOINI, SDOCOB, " _
   & "  CANMES, CUOMES, VCMTO, SERCOB, NUMCOB, FECCOB, GLOSA1, GLOSA2, USU ) " _
   & " SELECT " _
   & "  C.NUMERO, C.FECHA, M.CODSOCIO, M.CODIGO, M.INS, M.NOMBRE, C.MONEDA, C.SDOPEN, " _
   & "  C.CUOINI, C.SDOCOB, C.CANMES, C.CUOMES, C.VCMTO, C.SERCOB, C.NUMCOB, C.FECCOB, " _
   & "  C.GLOSA1, C.GLOSA2, '" + wcodusu + "' " _
   & " FROM FRACCAB AS C INNER JOIN MAESOCIO AS M " _
   & "   ON C.CODSOCIO = M.CODSOCIO ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_FRACDEX " _
   & " (NUMERO, LINEA, VCMTO, CARGOS, ABONOS, SDONEW, SERCOB, NUMCOB, LINCOB, FECCOB, USU ) " _
   & " SELECT " _
   & "  NUMERO, LINEA, VCMTO, CARGOS, ABONOS, SDONEW, SERCOB, NUMCOB, LINCOB, FECCOB, '" + wcodusu + "' " _
   & " FROM FRACDET ")
   Db.CommitTrans

   Dim aa As Integer, wNum As String, wLin As String, wSdoOld As Currency, wSdoGra As Currency

   aa = Leerado8("SELECT * FROM TMP_FRACCAX " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY NUMERO ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wNum = ADO8!numero
         wSdoOld = ADO8!sdopen
         
         aa = Leerado7("SELECT * FROM TMP_FRACDEX " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "       NUMERO = '" + wNum + "' " _
                & " ORDER BY LINEA ")
         If aa > 0 Then
            ADO7.MoveFirst
            Do While Not ADO7.EOF
               wLin = ADO7!linea
               wSdoGra = wSdoOld - ADO7!abonos
   
               Db.BeginTrans
               Db.Execute ("UPDATE TMP_FRACDEX  " _
               & " SET SDOGRA = " + Str(wSdoGra) + " " _
               & " WHERE NUMERO = '" + wNum + "' AND " _
               & "        LINEA = '" + wLin + "' ")
               Db.CommitTrans
   
               wSdoOld = wSdoGra
   
               ADO7.MoveNext
            Loop
         End If
         Set ADO7 = Nothing
         ADO8.MoveNext
      Loop
   End If
End Sub

Private Sub cmdRelacion_Click()
   LlenaPreRelacion

   Crys2.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys2.ReportFileName = xraiz + "ReportCtaCte\FracRelacion.RPT"
   Crys2.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys2.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys2.Formulas(2) = "NOMMES= '" + Format(Date, "dd/mm/yyyy") + "' "
   Crys2.SelectionFormula = " {TMP_FRACDEX.USU}='" + wcodusu + "' "
   Crys2.WindowState = crptMaximized
   Crys2.Action = 1
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdUnFrac_Click()
   Dim wNum As String
   
   wNum = ADO1!numero
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_FRACCAX WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_FRACDEX WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_FRACCAX " _
   & " (NUMERO, FECHA, CODSOCIO, CODIGO, INS, NOMBRE, MONEDA, SDOPEN, CUOINI, SDOCOB, " _
   & "  CANMES, CUOMES, VCMTO, SERCOB, NUMCOB, FECCOB, GLOSA1, GLOSA2, USU ) " _
   & " SELECT " _
   & "  C.NUMERO, C.FECHA, M.CODSOCIO, M.CODIGO, M.INS, M.NOMBRE, C.MONEDA, C.SDOPEN, " _
   & "  C.CUOINI, C.SDOCOB, C.CANMES, C.CUOMES, C.VCMTO, C.SERCOB, C.NUMCOB, C.FECCOB, " _
   & "  C.GLOSA1, C.GLOSA2, '" + wcodusu + "' " _
   & " FROM FRACCAB AS C INNER JOIN MAESOCIO AS M " _
   & "   ON C.CODSOCIO = M.CODSOCIO " _
   & " WHERE NUMERO = '" + wNum + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_FRACDEX " _
   & " (NUMERO, LINEA, VCMTO, CARGOS, ABONOS, SDONEW, SERCOB, NUMCOB, FECCOB, USU ) " _
   & " SELECT " _
   & "  NUMERO, LINEA, VCMTO, CARGOS, ABONOS, SDONEW, SERCOB, NUMCOB, FECCOB, '" + wcodusu + "' " _
   & " FROM FRACDET " _
   & " WHERE NUMERO = '" + wNum + "' ")
   Db.CommitTrans

   Dim aa As Integer, wLin As String, wSdoOld As Currency, wSdoGra As Currency

   aa = Leerado8("SELECT * FROM TMP_FRACCAX " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY NUMERO ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wNum = ADO8!numero
         wSdoOld = ADO8!sdopen
         
         aa = Leerado7("SELECT * FROM TMP_FRACDEX " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "       NUMERO = '" + wNum + "' " _
                & " ORDER BY LINEA ")
         If aa > 0 Then
            ADO7.MoveFirst
            Do While Not ADO7.EOF
               wLin = ADO7!linea
               wSdoGra = wSdoOld - ADO7!abonos
   
               Db.BeginTrans
               Db.Execute ("UPDATE TMP_FRACDEX  " _
               & " SET SDOGRA = " + Str(wSdoGra) + " " _
               & " WHERE NUMERO = '" + wNum + "' AND " _
               & "        LINEA = '" + wLin + "' ")
               Db.CommitTrans
   
               wSdoOld = wSdoGra
   
               ADO7.MoveNext
            Loop
         End If
         Set ADO7 = Nothing
         ADO8.MoveNext
      Loop
   End If

   Crys2.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys2.ReportFileName = xraiz + "ReportCtaCte\FracRelacion.RPT"
   Crys2.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys2.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys2.Formulas(2) = "NOMMES= '" + Format(Date, "dd/mm/yyyy") + "' "
   Crys2.SelectionFormula = " {TMP_FRACDEX.USU}='" + wcodusu + "' "
   Crys2.WindowState = crptMaximized
   Crys2.Action = 1
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   Select Case ColIndex
   Case 0
        ADO1.Sort = "NUMERO"
   Case 1
        ADO1.Sort = "FECHA"
   Case 2
        ADO1.Sort = "CODSOCIO"
   Case 3
        ADO1.Sort = "CODIGO"
   Case 5
        ADO1.Sort = "NOMBRE"
   End Select
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   llenadet
   llenadet1
   TotalDet
   Limpiar
   refrescar
End Sub

Private Sub DataGrid2_AfterDelete()
   TotalDet
End Sub

Private Sub DataGrid2_BeforeDelete(Cancel As Integer)
   If ADO2.RecordCount > 1 Then
      If MsgBox("Esta Seguro de Eliminar Registro?", vbExclamation + vbYesNo, "Eliminar Registro") = vbYes Then
         Cancel = 0
         TotalDet
      Else
         Cancel = 1
      End If
   Else
      Cancel = 1
   End If
End Sub

Private Sub DataGrid2_GotFocus()
'   If ADO2.RecordCount = 0 Then
'      CreaDetalle
'      DataGrid2.Row = 0
'      DataGrid2.col = 1
'      DataGrid2.Text = IIf(IsNull(ADO2!concepto), 0, ADO2!concepto)
'   End If
   DataGrid2.col = 1
   DataGrid2.SelStart = 0
   If Len(Trim(DataGrid2.Text)) > 0 Then
      DataGrid2.SelLength = Len(Trim(DataGrid2.Text))
   End If
   DataGrid2.Refresh
End Sub

Private Sub DataGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim wvariable As String
    
    On Error GoTo err
    Select Case KeyCode
    Case 40  ' DOWN
         If ACCION = 1 Or ACCION = 2 Then
            
         wvariable = DataGrid2.Text
         
         Select Case DataGrid2.col
         Case 1
              ADO2!vcmto = IIf(IsNull(wvariable), "", wvariable)
         Case 2
              ADO2!cargos = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 3
              ADO2!abonos = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 4
              ADO2!sdonew = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 5
              ADO2!sercob = IIf(IsNull(wvariable), "", wvariable)
         Case 6
              ADO2!numcob = IIf(IsNull(wvariable), "", wvariable)
         Case 7
              ADO2!feccob = IIf(IsNull(wvariable), "", wvariable)
         End Select
    
         Select Case DataGrid2.col
         Case 0
              DataGrid2.Text = ADO2!linea
         Case 1
              DataGrid2.Text = IIf(IsNull(ADO2!vcmto), "", ADO2!vcmto)
         Case 2
              DataGrid2.Text = IIf(IsNull(ADO2!cargos), 0, ADO2!cargos)
         Case 3
              DataGrid2.Text = IIf(IsNull(ADO2!abonos), 0, ADO2!abonos)
         Case 4
              DataGrid2.Text = IIf(IsNull(ADO2!sdonew), 0, ADO2!sdonew)
         Case 5
              DataGrid2.Text = IIf(IsNull(ADO2!sercob), "", ADO2!sercob)
         Case 6
              DataGrid2.Text = IIf(IsNull(ADO2!numcob), "", ADO2!numcob)
         Case 7
              DataGrid2.Text = IIf(IsNull(ADO2!feccob), "", ADO2!feccob)
         End Select
         End If
    
    Case 37 ' Retroceder
         wvariable = DataGrid2.Text
         
         Select Case DataGrid2.col
         Case 1
              ADO2!vcmto = IIf(IsNull(wvariable), "", wvariable)
         Case 2
              ADO2!cargos = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 3
              ADO2!abonos = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 4
              ADO2!sdonew = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 5
              ADO2!sercob = IIf(IsNull(wvariable), "", wvariable)
         Case 6
              ADO2!numcob = IIf(IsNull(wvariable), "", wvariable)
         Case 7
              ADO2!feccob = IIf(IsNull(wvariable), "", wvariable)
         End Select
         
         If DataGrid2.col = 1 Then
            If DataGrid2.Row > 0 Then
               DataGrid2.Row = DataGrid2.Row - 1
            End If
            DataGrid2.col = 0
         End If
         
         Select Case DataGrid2.col
         Case 0
              DataGrid2.Text = ADO2!linea
         Case 1
              DataGrid2.Text = IIf(IsNull(ADO2!vcmto), "", ADO2!vcmto)
         Case 2
              DataGrid2.Text = IIf(IsNull(ADO2!cargos), 0, ADO2!cargos)
         Case 3
              DataGrid2.Text = IIf(IsNull(ADO2!abonos), 0, ADO2!abonos)
         Case 4
              DataGrid2.Text = IIf(IsNull(ADO2!sdonew), 0, ADO2!sdonew)
         Case 5
              DataGrid2.Text = IIf(IsNull(ADO2!sercob), "", ADO2!sercob)
         Case 6
              DataGrid2.Text = IIf(IsNull(ADO2!numcob), "", ADO2!numcob)
         Case 7
              DataGrid2.Text = IIf(IsNull(ADO2!feccob), "", ADO2!feccob)
         End Select
         
    Case 38 ' Subir
         wvariable = DataGrid2.Text
         
         Select Case DataGrid2.col
         Case 1
              ADO2!vcmto = IIf(IsNull(wvariable), "", wvariable)
         Case 2
              ADO2!cargos = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 3
              ADO2!abonos = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 4
              ADO2!sdonew = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 5
              ADO2!sercob = IIf(IsNull(wvariable), "", wvariable)
         Case 6
              ADO2!numcob = IIf(IsNull(wvariable), "", wvariable)
         Case 7
              ADO2!feccob = IIf(IsNull(wvariable), "", wvariable)
         End Select
    
         Select Case DataGrid2.col
         Case 0
              DataGrid2.Text = ADO2!linea
         Case 1
              DataGrid2.Text = IIf(IsNull(ADO2!vcmto), "", ADO2!vcmto)
         Case 2
              DataGrid2.Text = IIf(IsNull(ADO2!cargos), 0, ADO2!cargos)
         Case 3
              DataGrid2.Text = IIf(IsNull(ADO2!abonos), 0, ADO2!abonos)
         Case 4
              DataGrid2.Text = IIf(IsNull(ADO2!sdonew), 0, ADO2!sdonew)
         Case 5
              DataGrid2.Text = IIf(IsNull(ADO2!sercob), "", ADO2!sercob)
         Case 6
              DataGrid2.Text = IIf(IsNull(ADO2!numcob), "", ADO2!numcob)
         Case 7
              DataGrid2.Text = IIf(IsNull(ADO2!feccob), "", ADO2!feccob)
         End Select
    
    Case 39 ' Avanzar
         wvariable = DataGrid2.Text
         
         Select Case DataGrid2.col
         Case 1
              ADO2!vcmto = IIf(IsNull(wvariable), "", wvariable)
         Case 2
              ADO2!cargos = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 3
              ADO2!abonos = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 4
              ADO2!sdonew = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 5
              ADO2!sercob = IIf(IsNull(wvariable), "", wvariable)
         Case 6
              ADO2!numcob = IIf(IsNull(wvariable), "", wvariable)
         Case 7
              ADO2!feccob = IIf(IsNull(wvariable), "", wvariable)
         End Select
         
         If DataGrid2.col = 7 Then
            If Val(ADO2!lincob) < ADO2.RecordCount Then
               DataGrid2.Row = DataGrid2.Row + 1
            End If
            DataGrid2.col = 0
         End If
          
         Select Case DataGrid2.col
         Case 0
              DataGrid2.Text = ADO2!linea
         Case 1
              DataGrid2.Text = IIf(IsNull(ADO2!vcmto), "", ADO2!vcmto)
         Case 2
              DataGrid2.Text = IIf(IsNull(ADO2!cargos), 0, ADO2!cargos)
         Case 3
              DataGrid2.Text = IIf(IsNull(ADO2!abonos), 0, ADO2!abonos)
         Case 4
              DataGrid2.Text = IIf(IsNull(ADO2!sdonew), 0, ADO2!sdonew)
         Case 5
              DataGrid2.Text = IIf(IsNull(ADO2!sercob), "", ADO2!sercob)
         Case 6
              DataGrid2.Text = IIf(IsNull(ADO2!numcob), "", ADO2!numcob)
         Case 7
              DataGrid2.Text = IIf(IsNull(ADO2!feccob), "", ADO2!feccob)
         End Select
    
    End Select
    Exit Sub
err:
    MsgBox Format(err.Number, "00000000000") + " " + err.Description
    Resume Next
End Sub

Private Sub DataGrid2_KeyPress(KeyAscii As Integer)
    Dim c As Integer
    Dim wvariable As String, wvariable2 As String, wlll As Integer, wlll2 As Integer
    Dim wvariaold As String
    Dim wSoles As Currency, zTdo As String, zSer As String, zDoc As String, _
        wOld As Currency, wlinold As Integer, _
        waaa As String, wmmm As String
    
    On Error GoTo err
    Select Case KeyAscii
    Case 13
       Select Case DataGrid2.col
       Case 0  ' Linea
            DataGrid2.col = 1
       Case 1  ' Vcmto
            wvariable = DataGrid2.Text
            If Not IsDate(wvariable) Then
               MsgBox "Vencimiento Es Invalido", vbInformation
               ADO2!vcmto = Format(txtVcmto.Text, "dd/mm/yyyy")
               Exit Sub
            End If
            DataGrid2.Text = wvariable
            ADO2!vcmto = Format(wvariable, "dd/mm/yyyy")
            DataGrid2.col = 2
       Case 2  ' Cargos
            wvariable = DataGrid2.Text
            If Not IsNumeric(wvariable) Then
               MsgBox "Importe Cargo Es Invalido", vbInformation
               ADO2!cargos = 0
               Exit Sub
            End If
            DataGrid2.Text = wvariable
            ADO2!cargos = Format(wvariable, "####0.00")
            DataGrid2.col = 2
       Case 3  ' Abonos
       Case 4  ' SdoNew
       Case 5  ' SerCob
       Case 6  ' NumCob
       Case 7  ' FecCob
       End Select
       wvariable2 = IIf(IsNull(ADO2.Fields(DataGrid2.col)), "", Trim(ADO2.Fields(DataGrid2.col)))
       DataGrid2.Text = wvariable2
       ADO2.Update
       DataGrid2.Refresh
    Case Else
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
    Exit Sub
err:
    MsgBox Format(err.Number, "00000000000") + " " + err.Description
    Resume Next
End Sub

Private Sub DataGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim wvariable As String
    
   On Error GoTo err
   Select Case KeyCode
   Case 37  ' RETROCEDER
         If DataGrid2.col = 0 Then
            DataGrid2.col = 7
         End If
         
         Select Case DataGrid2.col
         Case 0
              DataGrid2.Text = ADO2!linea
         Case 1
              DataGrid2.Text = IIf(IsNull(ADO2!vcmto), "", ADO2!vcmto)
         Case 2
              DataGrid2.Text = IIf(IsNull(ADO2!cargos), 0, ADO2!cargos)
         Case 3
              DataGrid2.Text = IIf(IsNull(ADO2!abonos), 0, ADO2!abonos)
         Case 4
              DataGrid2.Text = IIf(IsNull(ADO2!sdonew), 0, ADO2!sdonew)
         Case 5
              DataGrid2.Text = IIf(IsNull(ADO2!sercob), "", ADO2!sercob)
         Case 6
              DataGrid2.Text = IIf(IsNull(ADO2!numcob), "", ADO2!numcob)
         Case 7
              DataGrid2.Text = IIf(IsNull(ADO2!feccob), "", ADO2!feccob)
         End Select
   Case 38  ' UP
   
   Case 39  ' AVANZAR

        If DataGrid2.col = 0 Then
           DataGrid2.col = 1
        End If
          
         Select Case DataGrid2.col
         Case 0
              DataGrid2.Text = ADO2!linea
         Case 1
              DataGrid2.Text = IIf(IsNull(ADO2!vcmto), "", ADO2!vcmto)
         Case 2
              DataGrid2.Text = IIf(IsNull(ADO2!cargos), 0, ADO2!cargos)
         Case 3
              DataGrid2.Text = IIf(IsNull(ADO2!abonos), 0, ADO2!abonos)
         Case 4
              DataGrid2.Text = IIf(IsNull(ADO2!sdonew), 0, ADO2!sdonew)
         Case 5
              DataGrid2.Text = IIf(IsNull(ADO2!sercob), "", ADO2!sercob)
         Case 6
              DataGrid2.Text = IIf(IsNull(ADO2!numcob), "", ADO2!numcob)
         Case 7
              DataGrid2.Text = IIf(IsNull(ADO2!feccob), "", ADO2!feccob)
         End Select
        
   Case 40  ' DOWN
   
   End Select
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub Exportar_Click()
   On Error GoTo err
   
   LlenaPreRelacion
   
   Dim aa As Integer, I As Integer, Heading(14) As String, wreg As Integer, wTot As Integer
   Dim wFec As Date, wVcm As Date, wFecCob As Date, wNum As String
   Heading(0) = "NUMERO"
   Heading(1) = "FECHA"
   Heading(2) = "SOCIO"
   Heading(3) = "CODOFIN"
   Heading(4) = "NOMBRE SOCIO"
   Heading(5) = "MONEDA"
   Heading(6) = "LIN"
   Heading(7) = "VCMTO"
   Heading(8) = "FEC.COB"
   Heading(9) = "SERIE"
   Heading(10) = "NUM.COB"
   Heading(11) = "CUOTA"
   Heading(12) = "COBROS"
   Heading(13) = " SALDO "
   Heading(14) = "SDO.ACUM"
   aa = Leerado3("SELECT * FROM TMP_FRACDEX WHERE USU = '" + wcodusu + "' ORDER BY NUMERO, LINEA ")
   If aa > 0 Then
      wTot = aa
      Set objExcel = New Excel.Application
      objExcel.SheetsInNewWorkbook = 1
      objExcel.Workbooks.Add
      With objExcel
           .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
           .Range(.Cells(3, 1), .Cells(3, 15)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 15)).Font.Bold = True
           .Cells(1, 1) = wnomcia
           .Cells(2, 1) = "RELACION DE FRACCIONAMIENTOS"
           For I = 1 To 15 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           objExcel.Columns("A").ColumnWidth = 12
           objExcel.Columns("B").ColumnWidth = 11
           objExcel.Columns("C").ColumnWidth = 8
           objExcel.Columns("D").ColumnWidth = 10
           objExcel.Columns("E").ColumnWidth = 50
           objExcel.Columns("F").ColumnWidth = 4
           objExcel.Columns("G").ColumnWidth = 5
           objExcel.Columns("H").ColumnWidth = 11
           objExcel.Columns("I").ColumnWidth = 11
           objExcel.Columns("J").ColumnWidth = 5
           objExcel.Columns("K").ColumnWidth = 10
           objExcel.Columns("L").ColumnWidth = 11
           objExcel.Columns("M").ColumnWidth = 11
           objExcel.Columns("N").ColumnWidth = 11
           objExcel.Columns("O").ColumnWidth = 11
      End With
      V = 4
      H = 1
      wreg = 1
      Do While Not ADO3.EOF
         wNum = ADO3!numero
         aa = Leerado4("SELECT * FROM TMP_FRACCAX WHERE NUMERO = '" + wNum + "' ")
         If aa > 0 Then
            wFec = ADO4!fecha
            objExcel.Cells(V, H + 0) = ADO4!numero
            objExcel.Cells(V, H + 1) = wFec
            objExcel.Cells(V, H + 2) = ADO4!codsocio
            objExcel.Cells(V, H + 3) = Format(ADO4!codigo, "#######0") + "-" + Format(ADO4!ins, "0")
            objExcel.Cells(V, H + 4) = IIf(IsNull(ADO4!nombre), "", ADO4!nombre)
            objExcel.Cells(V, H + 5) = ADO4!moneda
         End If
         Set ADO4 = Nothing
         
         Do While ADO3!numero = wNum
            DoEvents
            lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wreg, "####0") + " / " + Format(wTot, "####0")
            lblMensaje.Refresh
         
            objExcel.Range(objExcel.Cells(V, H + 11), objExcel.Cells(V, H + 14)).NumberFormat = "####,##0.00;;\ "
            
            objExcel.Cells(V, H + 6) = ADO3!linea
            If IsDate(ADO3!vcmto) Then
               wVcm = ADO3!vcmto
               objExcel.Cells(V, H + 7) = wVcm
            End If
            If IsDate(ADO3!feccob) Then
               wFecCob = ADO3!feccob
               objExcel.Cells(V, H + 8) = wFecCob
            End If
            objExcel.Cells(V, H + 9) = ADO3!sercob
            objExcel.Cells(V, H + 10) = ADO3!numcob
            objExcel.Cells(V, H + 11) = ADO3!cargos
            objExcel.Cells(V, H + 12) = ADO3!abonos
            objExcel.Cells(V, H + 13) = ADO3!sdonew
            objExcel.Cells(V, H + 14) = ADO3!sdogra
         
            V = V + 1
            wreg = wreg + 1
            ADO3.MoveNext
            If ADO3.EOF Then
               Exit Do
            End If
         Loop
         
         V = V + 1
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

Private Sub Form_Activate()
   ACCION = 0
'   fraMantenimiento.Enabled = False
     
   Dim a As Integer
   editar (False)
   
   DataGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
   
   Limpiar
   LlenaCab
   LlenaCab1
   llenadet
   
   llenadet1
   Limpiar
   refrescar

End Sub

Private Sub Form_Load()
   frmFracMante.Left = (Screen.Width - Width) \ 2
   frmFracMante.Top = 0
   
   Set DataGrid1.DataSource = Nothing
   Set DataGrid2.DataSource = Nothing
'   Limpiar
End Sub

Private Sub LlenaCab()
   Dim c As Integer, waaa As String, wmmm As String, wFec As Date
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_FRACCAB WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
    
   Db.BeginTrans
      Db.Execute ("INSERT INTO TMP_FRACCAB " _
      & " (NUMERO, FECHA, CODSOCIO, CODIGO, INS, NOMBRE, MONEDA, SDOPEN, CUOINI, SDOCOB, " _
      & "  CANMES, CUOMES, VCMTO, SERCOB, NUMCOB, FECCOB, GLOSA1, GLOSA2, USU ) " _
      & " SELECT " _
      & "  C.NUMERO, C.FECHA, C.CODSOCIO, C.CODIGO, C.INS, M.NOMBRE, C.MONEDA, C.SDOPEN, " _
      & "  C.CUOINI, C.SDOCOB, C.CANMES, C.CUOMES, C.VCMTO, C.SERCOB, C.NUMCOB, C.FECCOB, " _
      & "  C.GLOSA1, C.GLOSA2, '" + wcodusu + "' " _
      & " FROM FRACCAB AS C INNER JOIN MAESOCIO AS M ON C.CODSOCIO = M.CODSOCIO ")
      Db.CommitTrans
   
   c = Leerado("SELECT NUMERO, FECHA, CODSOCIO, CODIGO, INS, NOMBRE, MONEDA, SDOPEN, " _
              & "      CUOINI, SDOCOB, CANMES, CUOMES, VCMTO, SERCOB, NUMCOB, FECCOB, " _
              & "      GLOSA1, GLOSA2, USU  " _
              & " FROM TMP_FRACCAB " _
              & " WHERE USU = '" + wcodusu + "' " _
              & " ORDER BY NUMERO ")
   Set DataGrid1.DataSource = ADO1
End Sub
    
Private Sub LlenaCab1()
    DataGrid1.Columns(0).Width = 1070   ' NUMERO
    DataGrid1.Columns(0).Alignment = dbgCenter
    DataGrid1.Columns(0).Caption = "NUMERO"
    
    DataGrid1.Columns(1).Width = 1050   ' Fecha
    DataGrid1.Columns(1).Alignment = dbgCenter
    DataGrid1.Columns(1).NumberFormat = "dd/mm/yyyy"
    DataGrid1.Columns(1).Caption = "FECHA"
    
    DataGrid1.Columns(2).Width = 650   ' CODSOCIO
    DataGrid1.Columns(2).Alignment = dbgRight
    DataGrid1.Columns(2).Caption = "SOCIO"
       
    DataGrid1.Columns(3).Width = 850   ' CODIGO
    DataGrid1.Columns(3).Alignment = dbgRight
    DataGrid1.Columns(3).Caption = "CODIGO"
       
    DataGrid1.Columns(4).Width = 300   ' INS
    DataGrid1.Columns(4).Alignment = dbgRight
    DataGrid1.Columns(4).Caption = "INS"
       
    DataGrid1.Columns(5).Width = 3700  ' NOMBRE
    DataGrid1.Columns(5).Alignment = dbgLeft
    DataGrid1.Columns(5).Caption = "NOMBRE"
    
    DataGrid1.Columns(6).Width = 360   ' MONEDA
    DataGrid1.Columns(6).Alignment = dbgLeft
    DataGrid1.Columns(6).Caption = "MON"
    
    DataGrid1.Columns(7).Width = 900   ' SDOPEN
    DataGrid1.Columns(7).Alignment = dbgRight
    DataGrid1.Columns(7).Caption = " DEUDA "
    DataGrid1.Columns(7).NumberFormat = "###,##0.00;;\ "
    
    DataGrid1.Columns(8).Width = 900   ' CUOINI
    DataGrid1.Columns(8).Alignment = dbgRight
    DataGrid1.Columns(8).Caption = "INICIAL"
    DataGrid1.Columns(8).NumberFormat = "###,##0.00;;\ "
    
    DataGrid1.Columns(9).Width = 900   ' SDOCOB
    DataGrid1.Columns(9).Alignment = dbgRight
    DataGrid1.Columns(9).Caption = "SDOxCOB"
    DataGrid1.Columns(9).NumberFormat = "###,##0.00;;\ "
    
    DataGrid1.Columns(10).Width = 400   ' CANMES
    DataGrid1.Columns(10).Alignment = dbgRight
    DataGrid1.Columns(10).Caption = "MESES"
    DataGrid1.Columns(10).NumberFormat = "#0;;\ "
    
    DataGrid1.Columns(11).Width = 900   ' CUOMES
    DataGrid1.Columns(11).Alignment = dbgRight
    DataGrid1.Columns(11).Caption = "CUOTA"
    DataGrid1.Columns(11).NumberFormat = "###,##0.00;;\ "
    
    DataGrid1.Columns(12).Visible = False
    DataGrid1.Columns(13).Visible = False
    DataGrid1.Columns(14).Visible = False
    DataGrid1.Columns(15).Visible = False
    DataGrid1.Columns(16).Visible = False
    DataGrid1.Columns(17).Visible = False
    DataGrid1.Columns(18).Visible = False
      
    DataGrid1.Refresh

   DataGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
End Sub

Private Sub llenadet()
   Dim wNum As String, _
       wDia As Integer, wMes As Integer, wAno As Integer, wVcm As Date, wFec As Date, _
       zDia As String, zMes As String, zAno As String, zVcm As String
    
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_FRACDET WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
    
    If ACCION = 1 Then
    
    Else
       
       If Not ADO1.BOF And Not ADO1.EOF Then
          wNum = ADO1!numero
       Else
          wNum = ""
       End If
       
       If wNum <> "" Then
          Db.BeginTrans
          Db.Execute ("INSERT INTO TMP_FRACDET " _
                & " ( NUMERO, LINEA  , VCMTO  , CARGOS , ABONOS, " _
                & "   SDONEW, SERCOB , NUMCOB , FECCOB , USU ) " _
                & " SELECT " _
                & "   NUMERO, LINEA  , VCMTO  , CARGOS , ABONOS, " _
                & "   SDONEW, SERCOB , NUMCOB , FECCOB , '" + wcodusu + "' " _
                & " From FRACDET " _
                & " WHERE NUMERO = '" + wNum + "' ")
          Db.CommitTrans
       
          txtNumero.Text = ADO1!numero
          If IsDate(ADO1!fecha) Then
             txtFecha.Text = Format(ADO1!fecha, "dd/mm/yyyy")
          
             wFec = Format(txtFecha.Text, "dd/mm/yyyy")
             wDia = Day(wFec)
             wMes = Month(wFec)
             wAno = Year(wFec)
          
             zDia = Format(wDia, "00")
             zMes = Format(wMes, "00")
             zAno = Format(wAno, "0000")
             zVcm = Format(zDia + "/" + zMes + "/" + zAno, "dd/mm/yyyy")
          
             wVcm = Format(fundiames(zMes) + "/" + zMes + "/" + zAno, "dd/mm/yyyy")
          Else
             txtFecha.Text = "__/__/____"
             txtVcmto.Text = "__/__/____"
          End If
          txtCodSocio.Text = IIf(IsNull(ADO1!codsocio), "", ADO1!codsocio)
          txtCodigo.Text = IIf(IsNull(ADO1!codigo), "", ADO1!codigo)
          txtIns.Text = IIf(IsNull(ADO1!ins), "", ADO1!ins)
          txtMoneda.Text = IIf(IsNull(ADO1!moneda), "", ADO1!moneda)
          txtSdoPen.Text = IIf(IsNull(ADO1!sdopen), 0, Format(ADO1!sdopen, "#####0.00;;\ "))
          txtCuoIni.Text = IIf(IsNull(ADO1!cuoini), 0, Format(ADO1!cuoini, "#####0.00;;\ "))
          txtSdoCob.Text = IIf(IsNull(ADO1!sdocob), 0, Format(ADO1!sdocob, "#####0.00;;\ "))
          txtCanMes.Text = IIf(IsNull(ADO1!canmes), 0, Format(ADO1!canmes, "#0;;\ "))
          txtCuoMes.Text = IIf(IsNull(ADO1!cuomes), 0, Format(ADO1!cuomes, "#####0.00;;\ "))
'          If IsDate(ADO1!vcmto) Then
'             txtVcmto.Text = Format(ADO1!vcmto, "dd/mm/yyyy")
'          Else
'          End If
          txtGlosa1.Text = IIf(IsNull(ADO1!glosa1), "", ADO1!glosa1)
          txtGlosa2.Text = IIf(IsNull(ADO1!glosa2), "", ADO1!glosa2)
          txtSerCob.Text = IIf(IsNull(ADO1!sercob), "", ADO1!sercob)
          txtNumCob.Text = IIf(IsNull(ADO1!numcob), "", ADO1!numcob)
          
          If ACCION = 2 Then
             txtNumero.Enabled = False
          End If
       
       End If
       
    End If
    Dim c As Integer
    c = Leerado2("SELECT LINEA, VCMTO, CARGOS, ABONOS, SDONEW, " _
                 & "     SERCOB, NUMCOB, FECCOB, NUMERO, USU " _
                 & " FROM TMP_FRACDET " _
                 & " WHERE USU = '" + wcodusu + "' " _
                 & " ORDER BY LINEA ")
    Set DataGrid2.DataSource = ADO2
End Sub

Private Sub llenadet1()
   DataGrid2.Columns(0).Width = 290   ' Linea'
   DataGrid2.Columns(0).Alignment = dbgCenter
   DataGrid2.Columns(0).Caption = "LIN"
       
   DataGrid2.Columns(1).Width = 1020  ' VCMTO
   DataGrid2.Columns(1).Alignment = dbgLeft
   DataGrid2.Columns(1).Caption = "VCMTO"
   DataGrid2.Columns(1).NumberFormat = "dd/mm/yyyy"
   
   DataGrid2.Columns(2).Width = 780    ' CARGOS
   DataGrid2.Columns(2).Alignment = dbgRight
   DataGrid2.Columns(2).Caption = "CARGOS"
   DataGrid2.Columns(2).NumberFormat = "##,##0.00;;\ "
      
   DataGrid2.Columns(3).Width = 780    ' ABONOS
   DataGrid2.Columns(3).Alignment = dbgRight
   DataGrid2.Columns(3).Caption = "ABONOS"
   DataGrid2.Columns(3).NumberFormat = "##,##0.00;;\ "
      
   DataGrid2.Columns(4).Width = 780    ' SDONEW
   DataGrid2.Columns(4).Alignment = dbgRight
   DataGrid2.Columns(4).Caption = "SALDO"
   DataGrid2.Columns(4).NumberFormat = "##,##0.00;;\ "
      
   DataGrid2.Columns(5).Width = 420    ' SERCOB
   DataGrid2.Columns(5).Alignment = dbgRight
   DataGrid2.Columns(5).Caption = "SERIE"
   
   DataGrid2.Columns(6).Width = 970    ' NUMCOB
   DataGrid2.Columns(6).Alignment = dbgRight
   DataGrid2.Columns(6).Caption = "NUM.COB"
   
   DataGrid2.Columns(7).Width = 1020  ' FECCOB
   DataGrid2.Columns(7).Alignment = dbgLeft
   DataGrid2.Columns(7).Caption = "FEC.COB"
   DataGrid2.Columns(7).NumberFormat = "dd/mm/yyyy"
   
   DataGrid2.Columns(0).Locked = True
       
   DataGrid2.Columns(3).Locked = True
   DataGrid2.Columns(4).Locked = True
   DataGrid2.Columns(5).Locked = True
   DataGrid2.Columns(6).Locked = True
   DataGrid2.Columns(7).Locked = True
       
   DataGrid2.Columns(8).Visible = False
   DataGrid2.Columns(9).Visible = False
   DataGrid2.col = 1
          
   DataGrid2.Refresh
End Sub

Private Sub TotalDet()
   Dim aa As Integer, wTot As Currency, wNum As String
   wNum = txtNumero.Text
   wTot = 0
   aa = Leerado8("SELECT SUM(CARGOS) AS CARGOS FROM TMP_FRACDET " _
                & " WHERE NUMERO = '" + wNum + "' AND " _
                & "          USU = '" + wcodusu + "' ")
   If aa > 0 Then
      wTot = IIf(IsNull(ADO8!cargos), 0, ADO8!cargos)
   End If
   Set ADO8 = Nothing

   lblTotal.Caption = Format(wTot, "#####0.00;;\ ")
End Sub

Private Sub txtCanMes_GotFocus()
   txtCanMes.SelStart = 0
   txtCanMes.SelLength = Len(Trim(txtCanMes.Text))
End Sub

Private Sub txtCanMes_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtCuoIni.SetFocus
   Case 40
        txtCuoMes.SetFocus
   End Select
End Sub

Private Sub txtCanMes_KeyPress(KeyAscii As Integer)
   Dim wCanMes As Currency, wCuoMes As Currency, wSdoCob As Currency
   
   If KeyAscii = 13 Then
      If Len(Trim(txtCanMes.Text)) = 0 Then
         MsgBox "Cantidad En Cero", vbExclamation
         txtCanMes = 0
         Exit Sub
      End If
      wCanMes = Val(txtCanMes.Text)
      wSdoCob = Val(txtSdoCob.Text)
      wCuoMes = Round((wSdoCob / wCanMes), 2)
      txtCuoMes.Text = Format(wCuoMes, "####0.00;;\ ")
    
      txtVcmto.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtCodigo_Change()
   Dim aa As Integer
   If Len(Trim(txtCodigo.Text)) > 0 Then
      aa = Leerado5a("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(Val(txtCodigo.Text)) + " ")
      If aa > 0 Then
         lblCodSocio.Caption = ADO5a!nombre
         txtCodSocio.Text = ADO5a!codsocio
         txtCodigo.Text = ADO5a!codigo
         txtIns.Text = ADO5a!ins
         txtNumDoc.Text = ADO5a!numdoc
         txtE_socio.Text = ADO5a!e_socio
         txtGrado.Text = ADO5a!grado
      Else
         lblCodSocio.Caption = ""
         txtCodSocio.Text = ""
         txtCodigo.Text = ""
         txtIns.Text = ""
         txtNumDoc.Text = ""
         txtE_socio.Text = ""
         txtGrado.Text = ""
      End If
      Set ADO8a = Nothing
   Else
      lblCodSocio.Caption = ""
      txtCodSocio.Text = ""
      txtCodigo.Text = ""
      txtIns.Text = ""
      txtNumDoc.Text = ""
      txtE_socio.Text = ""
      txtGrado.Text = ""
   End If
End Sub

Private Sub txtCodigo_GotFocus()
   txtCodigo.SelStart = 0
   txtCodigo.SelLength = Len(Trim(txtCodigo.Text))
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim KeyNew As Integer
   KeyNew = Asc(UCase(Chr(KeyCode)))
   
   Select Case KeyCode
   Case 38
        txtFecha.SetFocus
   Case 40
        txtMoneda.SetFocus
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocioCodigo.Show 1
        If xseleccion <> "" Then
           txtCodigo.Text = xseleccion
        End If
   Case 65 To 90
        xlista = "1"
        xseleccion = Chr(KeyNew)
        frmSelecSocioCodigo.Show 1
        If xseleccion <> "" Then
           txtCodigo.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   Dim aa As Integer, wSoc As Integer, _
       wSdoPen As Currency, wCuoIni As Currency, wSdoCob As Currency, _
       wCuoMes As Currency, wCanMes As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtCodigo.Text)) = 0 Then
         MsgBox "CodOFIN En Blanco", vbExclamation
         txtCodigo.Text = ""
         Exit Sub
      End If
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(Val(txtCodigo.Text)) + " ")
      If aa = 0 Then
         MsgBox "Codofin Digitado NO Existe", vbExclamation
         txtCodigo.Text = ""
         Exit Sub
      End If
      
      aa = Leerado8a("SELECT C.NUMERO, C.FECHA, C.CODSOCIO, D.SDONEW " _
                    & " FROM FRACCAB AS C INNER JOIN FRACDET AS D " _
                    & "   ON C.NUMERO = D.NUMERO " _
                    & " WHERE C.CODIGO = " + Str(Val(txtCodigo.Text)) + " AND " _
                    & "       D.SDONEW > 0 ")
      If aa > 0 Then
         MsgBox "Socio Tiene Frac." + ADO8a!numero + " Por Pagar" + vbNewLine + vbNewLine + _
                "No Se Puede Crear Otro Nuevo", vbExclamation
         txtCodigo.Text = ""
         Exit Sub
      End If
      Set ADO8a = Nothing
      
      If Len(Trim(txtSdoPen.Text)) = 0 Then
         wSoc = Val(txtCodSocio.Text)
      
         wSdoPen = SaldoFoto(wSoc, zMesTope)
         If wSdoPen < 0 Then
            MsgBox "No Existe Saldo Por Cobrar", vbExclamation
            Exit Sub
         End If
      Else
         wSdoPen = Val(txtSdoPen.Text)
      End If
      If Len(Trim(txtCuoIni.Text)) = 0 Then
         wCuoIni = Round((wSdoPen * 0.2), 2)
      Else
         wCuoIni = Val(txtCuoIni.Text)
      End If
      wSdoCob = wSdoPen - wCuoIni
      If Len(Trim(txtCanMes.Text)) = 0 Then
         wCanMes = 5
      Else
         wCanMes = Val(txtCanMes.Text)
      End If
      wCuoMes = Round((wSdoCob / wCanMes), 2)
      
      txtCodSocio.Text = ADO8!codsocio
      txtIns.Text = ADO8!ins
      txtNumDoc.Text = ADO8!numdoc
      txtE_socio.Text = ADO8!e_socio
      txtGrado.Text = ADO8!grado
      txtSdoPen.Text = Format(wSdoPen, "#####0.00;;\ ")
      txtCuoIni.Text = Format(wCuoIni, "#####0.00;;\ ")
      txtSdoCob.Text = Format(wSdoCob, "#####0.00;;\ ")
      txtCanMes.Text = Format(wCanMes, "#0;;\ ")
      txtCuoMes.Text = Format(wCuoMes, "#####0.00;;\ ")
      
      txtMoneda.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtCodSocio_Change()
   Dim aa As Integer, wSoc As Integer
   If Len(Trim(txtCodSocio.Text)) > 0 Then
      aa = Leerado5a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
      If aa > 0 Then
         wSoc = Val(txtCodSocio.Text)
         txtCodigo.Text = ADO5a!codigo
         txtIns.Text = ADO5a!ins
         txtNumDoc.Text = ADO5a!numdoc
         txtE_socio.Text = ADO5a!e_socio
         txtGrado.Text = ADO5a!grado
      End If
   Else
      lblCodSocio.Caption = ""
   End If
End Sub

Private Sub txtCuoIni_GotFocus()
   txtCuoIni.SelStart = 0
   txtCuoIni.SelLength = Len(Trim(txtCuoIni.Text))
End Sub

Private Sub txtCuoIni_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtSdoPen.SetFocus
   Case 40
        txtCanMes.SetFocus
   End Select
End Sub

Private Sub txtCuoIni_KeyPress(KeyAscii As Integer)
   Dim wSdoPen As Currency, wCuoIni As Currency, wSdoCob As Currency, _
       wCuoMes As Currency, wCanMes As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtCuoIni.Text)) = 0 Then
         MsgBox "Cuota Inicial En Cero", vbExclamation
         txtCuoIni.Text = ""
         Exit Sub
      End If
      wSdoPen = Val(txtSdoPen.Text)
      wCuoIni = Val(txtCuoIni.Text)
      wSdoCob = wSdoPen - wCuoIni
      If Len(Trim(txtCanMes.Text)) = 0 Then
         wCanMes = 5
      Else
         wCanMes = Val(txtCanMes.Text)
      End If
      If wCanMes <> 0 Then
         wCuoMes = Round((wSdoCob / wCanMes), 2)
      
         txtCuoMes.Text = Format(wCuoMes, "#####0.00;;\ ")
      End If
      txtCuoIni.Text = Format(wCuoIni, "#####0.00;;\ ")
      txtSdoCob.Text = Format(wSdoCob, "#####0.00;;\ ")
      txtCanMes.SetFocus
   Else
      If InStr(1, "0123456789." + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtE_socio_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + txtE_socio.Text + "' ")
   If aa > 0 Then
      lblE_socio.Caption = ADO6a!nombre
   Else
      lblE_socio.Caption = ""
   End If
   Set ADO6a = Nothing
End Sub

Private Sub txtGlosa1_GotFocus()
   txtGlosa1.SelStart = 0
   txtGlosa1.SelLength = Len(Trim(txtGlosa1.Text))
End Sub

Private Sub txtGlosa1_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtVcmto.SetFocus
   Case 40
        txtGlosa2.SetFocus
   End Select
End Sub

Private Sub txtGlosa1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtGlosa2.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtGlosa2_GotFocus()
   txtGlosa2.SelStart = 0
   txtGlosa2.SelLength = Len(Trim(txtGlosa2.Text))
End Sub

Private Sub txtGlosa2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtGlosa1.SetFocus
   End Select
End Sub

Private Sub txtGlosa2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      CreaDetalle
      
      DataGrid2.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtGrado_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAEGRADO WHERE GRADO = " + Str(Val(txtGrado.Text)) + " ")
   If aa > 0 Then
      lblGrado.Caption = ADO6a!nombre
   Else
      lblGrado.Caption = ""
   End If
   Set ADO6a = Nothing
End Sub

Private Sub txtFecha_GotFocus()
   txtFecha.SelStart = 0
   txtFecha.SelLength = 10
End Sub

Private Sub txtFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 38
         If txtNumero.Enabled = True Then
            txtNumero.SetFocus
         End If
    Case 40
         txtCodigo.SetFocus
    End Select
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFecha.Text = "__/__/____" Then
         MsgBox "Fecha En Blanco", vbExclamation
         txtFecha.SetFocus
         Exit Sub
      End If
      If Not IsDate(Trim(txtFecha)) Then
         MsgBox "Campo Digitado No Es Fecha Valida", vbExclamation
         txtFecha.Text = "__/__/____"
         txtFecha.SetFocus
         Exit Sub
      End If
      txtFecha.Text = Format(txtFecha.Text, "dd/mm/yyyy")
      
      txtCodigo.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtMoneda_Change()
   Select Case txtMoneda.Text
   Case "S"
        lblMoneda.Caption = "S/. Nuevos Soles"
   Case "D"
        lblMoneda.Caption = "US$ Dolares USA "
   Case Else
        lblMoneda.Caption = ""
   End Select
End Sub

Private Sub txtMoneda_GotFocus()
   txtMoneda.SelStart = 0
   txtMoneda.SelLength = Len(Trim(txtMoneda.Text))
End Sub

Private Sub txtMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtCodSocio.SetFocus
   Case 40
        txtSdoPen.SetFocus
   End Select
End Sub

Private Sub txtMoneda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txtMoneda.Text)) = 0 Then
         MsgBox "Tipo de Moneda En Blanco", vbExclamation
         txtMoneda.Text = "S"
         Exit Sub
      End If
      If txtMoneda.Text <> "S" And _
         txtMoneda.Text <> "D" Then
         MsgBox "Tipo de Moneda Digitada Es Invalida", vbExclamation
         txtMoneda.Text = "S"
         Exit Sub
      End If
      txtSdoPen.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Function validaCob()
   On Error GoTo err
   Dim aa As Integer
   Dim wImp As Currency, wTip As String
   Dim wlleane As Boolean, wllecen As Boolean
   Dim wdolcar As Currency, wdolabo As Currency
   Dim wsolcar As Currency, wsolabo As Currency
   Dim autom1 As String, autom2 As String, wtotdet As Currency
   
   If Len(Trim(txtNumero.Text)) = 0 Then
      MsgBox "Numero de Fraccionamiento En Blanco", vbExclamation
      txtNumero.SetFocus
      validaCob = True
      Exit Function
   End If
   If txtFecha.Text = "__/__/____" Then
      MsgBox "Fecha En Blanco", vbExclamation
      txtFecha.SetFocus
      validaCob = True
      Exit Function
   Else
      If Not IsDate(txtFecha.Text) Then
         MsgBox "Fecha Digitada Es Invalida", vbExclamation
         txtFecha.SetFocus
         validaCob = True
         Exit Function
      End If
   End If
   If txtVcmto.Text = "__/__/____" Then
      MsgBox "Vcmto En Blanco", vbExclamation
      txtVcmto.SetFocus
      validaCob = True
      Exit Function
   Else
      If Not IsDate(txtVcmto.Text) Then
         MsgBox "Vcmto Digitado Es Invalido", vbExclamation
         txtVcmto.SetFocus
         validaCob = True
         Exit Function
      End If
   End If
   If Len(Trim(txtCodSocio.Text)) = 0 Then
      MsgBox "Codigo de Socio En Blanco", vbExclamation
      txtCodSocio.SetFocus
      validaCob = True
      Exit Function
   Else
      aa = Leerado7("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
      If aa = 0 Then
         MsgBox "Codigo de Socio No Existe", vbExclamation
         txtCodSocio.Text = ""
         txtCodSocio.SetFocus
         validaCob = True
         Exit Function
      End If
   End If
   If txtMoneda <> "S" And _
      txtMoneda <> "D" Then
      MsgBox "Tipo de Moneda", vbExclamation
      txtMoneda.SetFocus
      validaCob = True
      Exit Function
   End If
   If Len(Trim(txtSdoPen.Text)) > 0 Then
      If Not IsNumeric(txtSdoPen.Text) Then
         MsgBox "Importe No Es Numerico", vbExclamation
         txtSdoPen.Text = ""
         txtSdoPen.SetFocus
         validaCob = True
         Exit Function
      End If
   End If
   If Len(Trim(txtCuoIni.Text)) > 0 Then
      If Not IsNumeric(txtCuoIni.Text) Then
         MsgBox "Cuota Inicial No Es Numerico", vbExclamation
         txtCuoIni.Text = ""
         txtCuoIni.SetFocus
         validaCob = True
         Exit Function
      End If
   End If
   If Len(Trim(txtSdoCob.Text)) > 0 Then
      If Not IsNumeric(txtSdoCob.Text) Then
         MsgBox "Saldo x Cobrar No Es Numerico", vbExclamation
         txtSdoCob.Text = ""
         txtSdoCob.SetFocus
         validaCob = True
         Exit Function
      End If
   End If
   If Len(Trim(txtCuoMes.Text)) > 0 Then
      If Not IsNumeric(txtCuoMes.Text) Then
         MsgBox "Cuota Mes No Es Numerico", vbExclamation
         txtCuoMes.Text = ""
         txtCuoMes.SetFocus
         validaCob = True
         Exit Function
      End If
   End If
   If Len(Trim(txtCanMes.Text)) > 0 Then
      If Not IsNumeric(txtCanMes.Text) Then
         MsgBox "Cantidad deCuotas No Es Numerico", vbExclamation
         txtCanMes.Text = ""
         txtCanMes.SetFocus
         validaCob = True
         Exit Function
      End If
   End If
   
   
   validaCob = False
   Exit Function
err:
   MsgBox Format(err.Number, "000000000000") + " " + err.Description
   Resume Next
End Function

Private Sub txtSdoPen_GotFocus()
   txtSdoPen.SelStart = 0
   txtSdoPen.SelLength = Len(Trim(txtSdoPen.Text))
End Sub

Private Sub txtSdoPen_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtMoneda.SetFocus
   Case 40
        txtCuoIni.SetFocus
   End Select
End Sub

Private Sub txtSdoPen_KeyPress(KeyAscii As Integer)
   Dim wSdoPen As Currency, wCuoIni As Currency, wSdoCob As Currency, _
       wCanMes As Integer, wCuoMes As Currency
   If KeyAscii = 13 Then
      If Len(Trim(txtSdoPen.Text)) = 0 Then
         MsgBox "Saldo a Fraccionar En Cero", vbExclamation
         txtSdoPen.Text = ""
         Exit Sub
      End If
      wSdoPen = Val(txtSdoPen.Text)
      
      If Len(Trim(txtCuoIni.Text)) = 0 Then
         wCuoIni = Round((wSdoPen * 0.2), 2)
      Else
         wCuoIni = Val(txtCuoIni.Text)
      End If
      wSdoCob = wSdoPen - wCuoIni
      If Len(Trim(txtCanMes.Text)) = 0 Then
         wCanMes = 5
      Else
         wCanMes = Val(txtCanMes.Text)
      End If
      wCuoMes = Round((wSdoCob / wCanMes), 2)
      
      txtSdoPen = Format(wSdoPen, "####0.00;;\ ")
      txtCuoIni = Format(wCuoIni, "####0.00;;\ ")
      txtSdoCob = Format(wSdoCob, "####0.00;;\ ")
      txtCanMes = Format(wCanMes, "#0;;\ ")
      txtCuoMes = Format(wCuoMes, "####0.00;;\ ")
      
      txtCuoIni.SetFocus
   Else
      If InStr(1, "0123456789." + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtVcmto_GotFocus()
   txtVcmto.SelStart = 0
   txtVcmto.SelLength = 10
End Sub

Private Sub txtVcmto_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtCanMes.SetFocus
   Case 40
        txtGlosa1.SetFocus
   End Select
End Sub

Private Sub txtVcmto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtVcmto.Text = "__/__/____" Then
         MsgBox "Primer Vcmto En Blanco", vbExclamation
         txtVcmto.Text = "__/__/____"
         Exit Sub
      End If
      If Not IsDate(txtVcmto.Text) Then
         MsgBox "Primer Vcmto Digitado Es Invalido", vbExclamation
         txtVcmto.Text = "__/__/____"
         Exit Sub
      End If
      
      CreaDetalle
      
      txtGlosa1.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub CreaDetalle()
   Dim aa As Integer, II As Integer, zCan As Integer, _
       zNum As String, zLin As String, zVcm As Date, _
       zCar As Currency, zAbo As Currency, zSdo As Currency, _
       zTot As Currency, zDif As Currency, zDia As Integer, zMes As Integer, zAno As Integer
   
   zNum = txtNumero.Text
   zVcm = Format(txtFecha.Text, "dd/mm/yyyy")
   zCar = Val(txtCuoIni.Text)
   
   If zCar > 0 Then
      aa = Leerado8("SELECT * FROM TMP_FRACDET " _
                       & " WHERE NUMERO = '" + zNum + "' AND " _
                       & "        LINEA = ' 0' AND " _
                       & "          USU = '" + wcodusu + "' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO TMP_FRACDET " _
         & " (NUMERO, LINEA, VCMTO, CARGOS, ABONOS, SDONEW, " _
         & "  SERCOB, NUMCOB, USU ) " _
         & " VALUES " _
         & " ('" + zNum + "', ' 0', '" + Format(zVcm, "dd/mm/yyyy") + "', " _
         & "  " + Str(zCar) + ", 0, " + Str(zCar) + ", '', '', '" + wcodusu + "'  ) ")
         Db.CommitTrans
      Else
         zAbo = ADO8!abonos
         zSdo = zCar + zAbo
          
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_FRACDET " _
         & " SET  VCMTO = '" + Format(zVcm, "dd/mm/yyyy") + "', " _
         & "     CARGOS = " + Str(zCar) + ", " _
         & "     ABONOS = " + Str(zAbo) + ", " _
         & "     SDONEW = " + Str(zSdo) + " " _
         & " WHERE NUMERO = '" + zNum + "' AND " _
         & "        LINEA = ' 0' AND " _
         & "          USU = '" + wcodusu + "' ")
         Db.CommitTrans
      End If
      Set ADO8 = Nothing
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_FRACDET " _
      & " WHERE    USU = '" + wcodusu + "' AND " _
      & "       NUMERO = '" + zNum + "' AND " _
      & "        LINEA = ' 0' ")
      Db.CommitTrans
   End If

   zVcm = Format(txtVcmto.Text, "dd/mm/yyyy")
   zCan = Val(txtCanMes.Text)
   zDia = Day(zVcm)
   zMes = Month(zVcm)
   zAno = Year(zVcm)
   zCar = Val(txtCuoMes.Text)
   zTot = 0

   For II = 1 To zCan
       zLin = Format(II, "@@")
   
       aa = Leerado8("SELECT * FROM TMP_FRACDET " _
                    & " WHERE NUMERO = '" + zNum + "' AND " _
                    & "        LINEA = '" + zLin + "' AND " _
                    & "          USU = '" + wcodusu + "' ")
       If aa = 0 Then
          Db.BeginTrans
          Db.Execute ("INSERT INTO TMP_FRACDET " _
          & " (NUMERO, LINEA, VCMTO, CARGOS, ABONOS, SDONEW, " _
          & "  SERCOB, NUMCOB, USU ) " _
          & " VALUES " _
          & " ('" + zNum + "', '" + zLin + "', '" + Format(zVcm, "dd/mm/yyyy") + "', " _
          & "  " + Str(zCar) + ", 0, " + Str(zCar) + ", '', '', '" + wcodusu + "'  ) ")
          Db.CommitTrans
       Else
          zAbo = ADO8!abonos
          zSdo = zCar + zAbo
          
          Db.BeginTrans
          Db.Execute ("UPDATE TMP_FRACDET " _
          & " SET  VCMTO = '" + Format(zVcm, "dd/mm/yyyy") + "', " _
          & "     CARGOS = " + Str(zCar) + ", " _
          & "     ABONOS = " + Str(zAbo) + ", " _
          & "     SDONEW = " + Str(zSdo) + " " _
          & " WHERE NUMERO = '" + zNum + "' AND " _
          & "        LINEA = '" + zLin + "' AND " _
          & "          USU = '" + wcodusu + "' ")
          Db.CommitTrans
       End If
       zTot = zTot + zCar
       
       If zMes >= 12 Then
          zMes = 1
          zAno = zAno + 1
       Else
          zMes = zMes + 1
       End If
       
       zVcm = fundiames(Format(zMes, "00")) + "/" + Format(zMes, "00") + "/" + Format(zAno, "0000")
       
   Next II

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_FRACDET " _
   & " WHERE    USU = '" + wcodusu + "' AND " _
   & "       NUMERO = '" + zNum + "' AND " _
   & "        LINEA > '" + Format(zCan, "@@") + "' ")
   Db.CommitTrans

   If zTot <> Val(lblTotal.Caption) Then
      If zTot > Val(txtSdoCob.Text) Then
         zDif = zTot - Val(txtSdoCob.Text)
      
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_FRACDET " _
         & " SET CARGOS = CARGOS - " + Str(zDif) + ", " _
         & "     SDONEW = SDONEW - " + Str(zDif) + " " _
         & " WHERE NUMERO = '" + zNum + "' AND " _
         & "        LINEA = '" + zLin + "' AND " _
         & "          USU = '" + wcodusu + "' ")
         Db.CommitTrans
      
      Else
         zDif = Val(txtSdoCob.Text) - zTot
      
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_FRACDET " _
         & " SET CARGOS = CARGOS + " + Str(zDif) + ", " _
         & "     SDONEW = SDONEW + " + Str(zDif) + " " _
         & " WHERE NUMERO = '" + zNum + "' AND " _
         & "        LINEA = '" + zLin + "' AND " _
         & "          USU = '" + wcodusu + "' ")
         Db.CommitTrans
      
      End If
   End If

   ADO2.Requery
   llenadet1
   TotalDet

End Sub

