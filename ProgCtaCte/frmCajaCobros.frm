VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajaCobros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobros x Caja"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   16305
   Begin VB.Frame fraAnula 
      Caption         =   "Anulación de Cobros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   13920
      TabIndex        =   78
      Top             =   3960
      Width           =   2295
      Begin VB.CommandButton cmdCancelaAnula 
         Caption         =   "&Cancelar Anulación"
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
         Left            =   120
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar Anulación"
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
         Left            =   1200
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtSerCobAnula 
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   83
         Top             =   900
         Width           =   495
      End
      Begin VB.TextBox txtNumCobAnula 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   81
         Top             =   900
         Width           =   1335
      End
      Begin VB.ComboBox cmbTipoAnula 
         Height          =   315
         ItemData        =   "frmCajaCobros.frx":0000
         Left            =   120
         List            =   "frmCajaCobros.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   420
         Width           =   2055
      End
      Begin VB.Label Label24 
         Caption         =   "Serie"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label22 
         Caption         =   "Numero"
         Height          =   255
         Left            =   1200
         TabIndex        =   82
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cobro"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   80
         Top             =   240
         Width           =   780
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
      Left            =   14640
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8160
      Width           =   975
   End
   Begin VB.Frame FraDetalles 
      Caption         =   "Detalles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   20
      Top             =   480
      Width           =   13695
      Begin VB.TextBox txtAnterio2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   74
         Top             =   1860
         Width           =   6255
      End
      Begin VB.TextBox txtAnterior 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   69
         Top             =   1560
         Width           =   6255
      End
      Begin VB.TextBox txtGrado 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         MaxLength       =   3
         TabIndex        =   57
         Top             =   860
         Width           =   495
      End
      Begin VB.TextBox txtNumdoc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   12480
         MaxLength       =   8
         TabIndex        =   56
         Top             =   380
         Width           =   975
      End
      Begin VB.TextBox txtIns 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6480
         MaxLength       =   1
         TabIndex        =   55
         Top             =   380
         Width           =   375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5520
         MaxLength       =   8
         TabIndex        =   54
         Top             =   380
         Width           =   975
      End
      Begin VB.TextBox txtTipCob 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         MaxLength       =   3
         TabIndex        =   53
         Top             =   860
         Width           =   495
      End
      Begin VB.TextBox txtE_socio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   52
         Top             =   860
         Width           =   495
      End
      Begin VB.TextBox txtForPag 
         Height          =   285
         Left            =   7800
         MaxLength       =   2
         TabIndex        =   49
         Top             =   1300
         Width           =   375
      End
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         ItemData        =   "frmCajaCobros.frx":0004
         Left            =   120
         List            =   "frmCajaCobros.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   380
         Width           =   2655
      End
      Begin VB.TextBox txtTipCam 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9000
         MaxLength       =   8
         TabIndex        =   40
         Top             =   860
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtGlosa 
         Height          =   285
         Left            =   120
         MaxLength       =   100
         TabIndex        =   37
         Top             =   1300
         Width           =   7695
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   12360
         MaxLength       =   8
         TabIndex        =   33
         Top             =   860
         Width           =   1095
      End
      Begin VB.TextBox txtMoneda 
         Height          =   285
         Left            =   9840
         MaxLength       =   1
         TabIndex        =   32
         Top             =   860
         Width           =   255
      End
      Begin VB.TextBox txtCodSocio 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   6960
         MaxLength       =   9
         TabIndex        =   28
         Top             =   380
         Width           =   975
      End
      Begin VB.TextBox txtSerCob 
         Height          =   285
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   26
         Top             =   380
         Width           =   495
      End
      Begin VB.TextBox txtNumCob 
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   22
         Top             =   380
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   285
         Left            =   4440
         TabIndex        =   23
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
         Height          =   2055
         Left            =   240
         TabIndex        =   39
         Top             =   2280
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   3625
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
      Begin VB.Label lblSdoAporte 
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
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   7800
         TabIndex        =   76
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Glosa Ant2"
         Height          =   195
         Left            =   120
         TabIndex        =   75
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label lblRenov 
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
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   10800
         TabIndex        =   73
         Top             =   1905
         Width           =   2775
      End
      Begin VB.Label lblFrac 
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
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   7800
         TabIndex        =   72
         Top             =   1600
         Width           =   5775
      End
      Begin VB.Label lblAporte 
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
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   10800
         TabIndex        =   71
         Top             =   1300
         Width           =   2775
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Glosa Ant1"
         Height          =   195
         Left            =   120
         TabIndex        =   70
         Top             =   1605
         Width           =   855
      End
      Begin VB.Label lblGrado 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3720
         TabIndex        =   66
         Top             =   860
         Width           =   2175
      End
      Begin VB.Label Label18 
         Caption         =   "Grado"
         Height          =   195
         Left            =   3480
         TabIndex        =   65
         Top             =   680
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "D.N.I."
         Height          =   195
         Left            =   12480
         TabIndex        =   64
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Ins"
         Height          =   195
         Left            =   6480
         TabIndex        =   63
         Top             =   200
         Width           =   375
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Codofin"
         Height          =   195
         Left            =   5520
         TabIndex        =   62
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Tipo de Cobro"
         Height          =   195
         Left            =   6120
         TabIndex        =   61
         Top             =   680
         Width           =   1335
      End
      Begin VB.Label lblTipCob 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6360
         TabIndex        =   60
         Top             =   860
         Width           =   2655
      End
      Begin VB.Label Label11 
         Caption         =   "Estado del Socio"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   680
         Width           =   1335
      End
      Begin VB.Label lblE_socio 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   600
         TabIndex        =   58
         Top             =   860
         Width           =   2655
      End
      Begin VB.Label lblForPag 
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
         Left            =   8160
         TabIndex        =   51
         Top             =   1300
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "Forma de Pago"
         Height          =   195
         Left            =   8040
         TabIndex        =   50
         Top             =   1125
         Width           =   2055
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cobro"
         Height          =   195
         Index           =   9
         Left            =   780
         TabIndex        =   48
         Top             =   200
         Width           =   780
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
         Left            =   480
         TabIndex        =   46
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label lblTotDol 
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
         Left            =   5640
         TabIndex        =   45
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label lblTotSol 
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
         Left            =   7920
         TabIndex        =   44
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "Total US$"
         Height          =   255
         Left            =   4680
         TabIndex        =   43
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Total S/."
         Height          =   255
         Left            =   7080
         TabIndex        =   42
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label34 
         Caption         =   "Tip.Cam"
         Height          =   195
         Left            =   9000
         TabIndex        =   41
         Top             =   680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   1120
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Importe Total"
         Height          =   195
         Left            =   12360
         TabIndex        =   36
         Top             =   680
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Tipo de Moneda"
         Height          =   195
         Left            =   9960
         TabIndex        =   35
         Top             =   680
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
         Left            =   10200
         TabIndex        =   34
         Top             =   860
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Nombre de Socio"
         Height          =   195
         Left            =   8040
         TabIndex        =   31
         Top             =   200
         Width           =   1695
      End
      Begin VB.Label lblCodSocio 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7920
         TabIndex        =   30
         Top             =   380
         Width           =   4575
      End
      Begin VB.Label Label5 
         Caption         =   "Cod.Socio"
         Height          =   195
         Left            =   6960
         TabIndex        =   29
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Serie"
         Height          =   255
         Left            =   2880
         TabIndex        =   27
         Top             =   200
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Numero"
         Height          =   255
         Left            =   3480
         TabIndex        =   25
         Top             =   200
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   4440
         TabIndex        =   24
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
      Left            =   13920
      TabIndex        =   15
      Top             =   2760
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
      Left            =   13920
      TabIndex        =   13
      Top             =   2040
      Width           =   2295
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
         Height          =   375
         Left            =   1200
         TabIndex        =   14
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
      Left            =   13920
      TabIndex        =   7
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton cmdAnular 
         Caption         =   "&Anular"
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
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   1440
         Width           =   975
      End
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtAno 
      Enabled         =   0   'False
      Height          =   305
      Left            =   7800
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox cmbMeses 
      Height          =   315
      ItemData        =   "frmCajaCobros.frx":0008
      Left            =   9000
      List            =   "frmCajaCobros.frx":0033
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmCajaCobros.frx":00CD
      Left            =   1080
      List            =   "frmCajaCobros.frx":00CF
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3615
      Left            =   720
      TabIndex        =   6
      Top             =   5280
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   6376
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
      Caption         =   "RELACION DE COBROS "
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
      Left            =   14880
      Top             =   7440
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
   Begin MSMask.MaskEdBox txtFecCab 
      Height          =   285
      Left            =   12000
      TabIndex        =   67
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha"
      Height          =   255
      Left            =   11160
      TabIndex        =   68
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Año"
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
      Left            =   7320
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Mes"
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
      Left            =   8520
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label38 
      Caption         =   "Compañia"
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
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmCajaCobros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACCION As Byte, wcia As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub Limpiar()
   txtSerCob.Text = ""
   txtNumCob.Text = ""
   txtFecha.Text = "__/__/____"
   txtCodSocio.Text = ""
   txtCodigo.Text = ""
   txtIns.Text = ""
   txtMoneda.Text = ""
   txtTipCam.Text = ""
   txtImporte.Text = ""
   txtGlosa.Text = ""
   txtForPag.Text = ""
   cmbTipo.ListIndex = 0
   lblCodSocio.Caption = ""
   lblAporte.Caption = ""
   lblFrac.Caption = ""
     
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_COBRODET WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Set DataGrid2.DataSource = Nothing
'   llenadet1
End Sub

Sub refrescar()
   If ADO1.BOF Then Exit Sub
   If ADO1.EOF Then Exit Sub
   cmbTipo.ListIndex = Val(ADO1!tipcob) - 1
   
   txtSerCob.Text = IIf(IsNull(ADO1!sercob), "", ADO1!sercob)
   txtNumCob.Text = IIf(IsNull(ADO1!numcob), "", ADO1!numcob)
   If IsDate(ADO1!fecha) Then
      txtFecha.Text = Format(ADO1!fecha, "dd/mm/yyYy")
   Else
      txtFecha.Text = "__/__/____"
   End If
   txtCodSocio.Text = IIf(IsNull(ADO1!codsocio), "", ADO1!codsocio)
   txtMoneda.Text = IIf(IsNull(ADO1!moneda), "", ADO1!moneda)
   txtImporte.Text = IIf(IsNull(ADO1!importe), 0, ADO1!importe)
   txtTipCam.Text = IIf(IsNull(ADO1!tipcam), 0, ADO1!tipcam)
   txtGlosa.Text = IIf(IsNull(ADO1!glosa), "", ADO1!glosa)
   txtForPag.Text = IIf(IsNull(ADO1!forpag), "", ADO1!forpag)
   
   llenadet
   llenadet1
End Sub

Sub grabar()
   On Error GoTo err
   
   Dim waaa As String, wmmm As String, wTip As String, wSer As String, wNum As String, aa As Integer
   Dim wLin As String, _
       wFec As String, wCam As Currency, _
       wSoc As Integer, wCod As Long, wIns As Integer, _
       wcon As String, wccc As String, wMes As String, _
       wAde As Currency, wDeu As Currency, wSdo As Currency, _
       wForPag As String, wGlo As String, _
       wTipoPariente As String, wNomPariente As String, wLinPariente As String, wqqq As Variant, _
       wUlt As String, wNumOpe As String
   waaa = txtAno.Text
   If Left(cmbMeses.Text, 2) <> "00" Then
      wmmm = Left(cmbMeses.Text, 2)
   Else
      wmmm = Format(Month(txtFecCab.Text), "00")
   End If
   wTip = Format(cmbTipo.ListIndex + 1, "0")
   wSer = txtSerCob.Text
   wNum = txtNumCob.Text
   wFec = Format(txtFecha.Text, "dd/mm/yyyy")
   wCam = Val(txtTipCam.Text)
   wSoc = Val(txtCodSocio.Text)
   wForPag = txtForPag.Text
   wGlo = txtGlosa.Text
   
''
'' Verifica que Cobro No Exista
''
   Dim wVeces As Integer
   
   wVeces = 1
   Do While True
      aa = Leerado8("SELECT * FROM CTRL_GRABAR ")
      If aa = 0 Then
         
         Db.BeginTrans
         Db.Execute ("INSERT INTO CTRL_GRABAR " _
         & " (USU, SERCOB, NUMCOB) " _
         & " VALUES " _
         & " ('" + wcodusu + "', '" + wSer + "', '" + wNum + "') ")
         Db.CommitTrans
         
         Exit Do
      End If
      wVeces = wVeces + 1
      If wVeces > 5 Then
         Db.BeginTrans
         Db.Execute ("DELETE FROM CTRL_GRABAR ")
         Db.CommitTrans
      End If
      Call Sleep(2000) 'espera por 1 segundo
   Loop
   
   If ACCION = 1 Then
      aa = Leerado8("SELECT * FROM COBROCAB " _
                & " WHERE    ANO = '" + waaa + "' AND " _
                & "          MES = '" + wmmm + "' AND " _
                & "       TIPCOB = '" + wTip + "' AND " _
                & "       SERCOB = '" + wSer + "' AND " _
                & "       NUMCOB = '" + wNum + "' ")
      If aa > 0 Then
      
         wUlt = "0000000000"
         aa = Leerado7("SELECT MAX(CAST(NUMCOB AS INTEGER)) AS NUMCOB " _
                   & " FROM COBROCAB " _
                   & " WHERE ANO = '" + waaa + "' AND " _
                   & "       TIPCOB = '" + wTip + "' AND " _
                   & "       SERCOB = '" + wSer + "' ")
         If aa > 0 Then
            wUlt = IIf(IsNull(ADO7!numcob), "0000000000", ADO7!numcob)
         End If
         Set ADO7 = Nothing
         wUlt = Format(Val(wUlt) + 1, "0000000000")
               
         MsgBox "Numero de Recibo Ya Existe" + vbNewLine + _
                "Se renumera Por " + wUlt
      
         txtNumCob.Text = wUlt
      
         wNum = wUlt
      End If
   End If
   
   If Len(Trim(wGlo)) = 0 And Left(wSer, 3) = "004" Then
      wGlo = "APORTE "
      aa = Leerado8("SELECT * FROM TMP_COBRODET " _
                & " WHERE    ANO = '" + waaa + "' AND " _
                & "          MES = '" + wmmm + "' AND " _
                & "       TIPCOB = '" + wTip + "' AND " _
                & "       SERCOB = '" + wSer + "' AND " _
                & "       NUMCOB = '" + wNum + "' AND " _
                & "          USU = '" + wcodusu + "' " _
                & " ORDER BY LINCOB ")
      If aa > 0 Then
         ADO8.MoveFirst
         Do While Not ADO8.EOF
            wGlo = wGlo + ADO8!mescob + " "
            ADO8.MoveNext
         Loop
      End If
      Set ADO8 = Nothing
   End If
   
   aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
   If aa > 0 Then
      wCod = ADO8!codigo
      wIns = ADO8!ins
   End If
   
   aa = Leerado8("SELECT * FROM TMP_COBROCAB " _
                & " WHERE    USU = '" + wcodusu + "' AND " _
                & "          ANO = '" + waaa + "' AND " _
                & "          MES = '" + wmmm + "' AND " _
                & "       TIPCOB = '" + wTip + "' AND " _
                & "       SERCOB = '" + wSer + "' AND " _
                & "       NUMCOB = '" + wNum + "' ")
   If aa = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO TMP_COBROCAB " _
      & " (ANO, MES, TIPCOB, SERCOB, NUMCOB, FECHA, MONEDA, IMPORTE, GLOSA, CODSOCIO, NOMBRE, " _
      & "  TIPCAM, DOLARE, SOLESS, FORPAG, USU ) " _
      & " VALUES " _
      & "  ('" + waaa + "', '" + wmmm + "', '" + wTip + "', '" + wSer + "', '" + wNum + "', " _
      & "   '" + Format(txtFecha.Text, "dd/mm/yyyy") + "', '" + txtMoneda.Text + "', " _
      & "   " + Str(Val(txtImporte.Text)) + ", '" + wGlo + "', " _
      & "   " + Str(Val(txtCodSocio.Text)) + ", '" + Trim(lblCodSocio.Caption) + "', " + Str(Val(txtTipCam.Text)) + ", " _
      & "   " + Str(Val(lblTotDol.Caption)) + ", " _
      & "   " + Str(Val(lblTotSol.Caption)) + ", '" + txtForPag.Text + "', '" + wcodusu + "' ) ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_COBROCAB " _
      & " SET FECHA = '" + Format(txtFecha.Text, "dd/mm/yyyy") + "', NOMBRE = '" + lblCodSocio.Caption + "', " _
      & "     CODSOCIO = " + Str(Val(txtCodSocio.Text)) + ", TIPCAM = " + Str(txtTipCam.Text) + ", " _
      & "     DOLARE = " + Str(Val(lblTotDol.Caption)) + ", " _
      & "     SOLESS = " + Str(Val(lblTotSol.Caption)) + ", " _
      & "      GLOSA = '" + wGlo + "', MONEDA = '" + txtMoneda.Text + "', " _
      & "     IMPORTE = " + Str(Val(txtImporte.Text)) + ", FORPAG = '" + txtForPag.Text + "' " _
      & " WHERE    USU = '" + wcodusu + "' AND " _
      & "          ANO = '" + waaa + "' AND " _
      & "          MES = '" + wmmm + "' AND " _
      & "       TIPCOB = '" + wTip + "' AND " _
      & "       SERCOB = '" + wSer + "' AND " _
      & "       NUMCOB = '" + wNum + "' ")
      Db.CommitTrans
   End If
   
   aa = Leerado8("SELECT * FROM COBROCAB " _
                & " WHERE    ANO = '" + waaa + "' AND " _
                & "          MES = '" + wmmm + "' AND " _
                & "       TIPCOB = '" + wTip + "' AND " _
                & "       SERCOB = '" + wSer + "' AND " _
                & "       NUMCOB = '" + wNum + "' ")
   If aa = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO COBROCAB " _
      & " (ANO, MES, TIPCOB, SERCOB, NUMCOB, FECHA, MONEDA, IMPORTE, GLOSA, CODSOCIO, " _
      & "  TIPCAM, DOLARE, SOLESS, FORPAG, USU ) " _
      & " VALUES " _
      & "  ('" + waaa + "', '" + wmmm + "', '" + wTip + "', '" + wSer + "', '" + wNum + "', " _
      & "   '" + Format(txtFecha.Text, "dd/mm/yyyy") + "', '" + txtMoneda.Text + "', " _
      & "   " + Str(Val(txtImporte.Text)) + ", '" + wGlo + "', " _
      & "   " + Str(Val(txtCodSocio.Text)) + ", " + Str(Val(txtTipCam.Text)) + ", " _
      & "   " + Str(Val(lblTotDol.Caption)) + ", " _
      & "   " + Str(Val(lblTotSol.Caption)) + ", '" + txtForPag.Text + "', '" + wcodusu + "' ) ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE COBROCAB " _
      & " SET FECHA = '" + Format(txtFecha.Text, "dd/mm/yyyy") + "', " _
      & "     CODSOCIO = " + Str(Val(txtCodSocio.Text)) + ", TIPCAM = " + Str(txtTipCam.Text) + ", " _
      & "     DOLARE = " + Str(Val(lblTotDol.Caption)) + ", " _
      & "     SOLESS = " + Str(Val(lblTotSol.Caption)) + ", " _
      & "      GLOSA = '" + wGlo + "', MONEDA = '" + txtMoneda.Text + "', " _
      & "     IMPORTE = " + Str(Val(txtImporte.Text)) + ", FORPAG = '" + txtForPag.Text + "', USU = '" + wcodusu + "' " _
      & " WHERE    ANO = '" + waaa + "' AND " _
      & "          MES = '" + wmmm + "' AND " _
      & "       TIPCOB = '" + wTip + "' AND " _
      & "       SERCOB = '" + wSer + "' AND " _
      & "       NUMCOB = '" + wNum + "' ")
      Db.CommitTrans
   End If
   
   aa = Leerado8("SELECT * FROM COBRODET " _
                & " WHERE    ANO = '" + waaa + "' AND " _
                & "          MES = '" + wmmm + "' AND " _
                & "       TIPCOB = '" + wTip + "' AND " _
                & "       SERCOB = '" + wSer + "' AND " _
                & "       NUMCOB = '" + wNum + "' ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wLin = ADO8!lincob
         wMes = ADO8!mescob
         wcon = ADO8!conpago
         wccc = ADO8!concepto

         Db.BeginTrans
         Db.Execute ("DELETE FROM COBRODET " _
         & " WHERE    ANO = '" + waaa + "' AND " _
         & "          MES = '" + wmmm + "' AND " _
         & "       TIPCOB = '" + wTip + "' AND " _
         & "       SERCOB = '" + wSer + "' AND " _
         & "       NUMCOB = '" + wNum + "' AND " _
         & "       LINCOB = '" + wLin + "' ")
         Db.CommitTrans
         
         If Len(Trim(wccc)) > 0 Then
            Db.BeginTrans
            Db.Execute ("DELETE FROM CTASXDET " _
            & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
            & "            MES = '" + wMes + "' AND " _
            & "       CONCEPTO = '" + wccc + "' AND " _
            & "         TIPMOV = '2' AND " _
            & "         TIPCOB = '03' AND " _
            & "         SERCOB = '" + wSer + "' AND " _
            & "         NUMCOB = '" + wNum + "' AND " _
            & "         LINCOB = '" + wLin + "' ")
            Db.CommitTrans
         
            Call ActualizaSaldos(wSoc, wMes, wccc)
         End If
                    
         ADO8.MoveNext
      Loop
   End If
   
   Dim wNumFra As String, wLinFra As String
   
   aa = Leerado8("SELECT * FROM TMP_COBRODET " _
                & " WHERE    ANO = '" + waaa + "' AND " _
                & "          MES = '" + wmmm + "' AND " _
                & "       TIPCOB = '" + wTip + "' AND " _
                & "       SERCOB = '" + wSer + "' AND " _
                & "       NUMCOB = '" + wNum + "' AND " _
                & "          USU = '" + wcodusu + "' " _
                & " ORDER BY LINCOB ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wLin = ADO8!lincob
         wMes = ADO8!mescob
         wcon = ADO8!conpago
         wccc = ADO8!concepto
         wNumFra = IIf(IsNull(ADO8!numfra), "", ADO8!numfra)
         wLinFra = IIf(IsNull(ADO8!linfra), "", ADO8!linfra)
         wNumOpe = IIf(IsNull(ADO8!numope), "", ADO8!numope)

         If Len(Trim(wcon)) > 0 Then
            Db.BeginTrans
            Db.Execute ("INSERT INTO COBRODET " _
            & " (ANO, MES, TIPCOB, SERCOB, NUMCOB, LINCOB, MESCOB, CONPAGO, DOLARE, SOLESS, " _
            & "  MONDOC, SDOOLD, ABONOS, SDONEW, IMPORTE, CONCEPTO, PARIENTE, NOMBRE, NUMFRA, LINFRA, NUMOPE) " _
            & " VALUES " _
            & " ('" + waaa + "', '" + wmmm + "', '" + wTip + "', '" + wSer + "', '" + wNum + "', '" + wLin + "', " _
            & "  '" + ADO8!mescob + "', '" + ADO8!conpago + "', " _
            & "  " + Str(ADO8!dolare) + ", " + Str(ADO8!soless) + ", " _
            & "  '" + ADO8!mondoc + "', " + Str(ADO8!sdoold) + ", " + Str(ADO8!abonos) + ", " _
            & "  " + Str(ADO8!sdonew) + ", " + Str(ADO8!importe) + ", '" + wccc + "', " _
            & "  '" + IIf(IsNull(ADO8!pariente), "", ADO8!pariente) + "', " _
            & "  '" + IIf(IsNull(ADO8!nombre), "", ADO8!nombre) + "', " _
            & "  '" + IIf(IsNull(ADO8!numfra), "", ADO8!numfra) + "', " _
            & "  '" + IIf(IsNull(ADO8!linfra), "", ADO8!linfra) + "', '" + wNumOpe + "' ) ")
            Db.CommitTrans
        
            If Len(Trim(wccc)) > 0 Then
               aa = Leerado7a("SELECT * FROM CTASXCAB " _
                            & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                            & "            MES = '" + wMes + "' AND " _
                            & "       CONCEPTO = '" + wccc + "' ")
               If aa = 0 Then
                  wqqq = CreaAporteMes(wSoc, wMes, wccc, 1)
               End If
               
               Db.BeginTrans
               Db.Execute ("INSERT INTO CTASXDET " _
               & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
               & "  TIPCAM, DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW ) " _
               & " VALUES " _
               & " (" + Str(wSoc) + ", '" + wMes + "', '" + wccc + "', " _
               & "  '03', '" + wSer + "', '" + wNum + "', '" + wLin + "', " _
               & "  '2', '" + Format(wFec, "dd/mm/yyyy") + "', " + Str(wCam) + ", " _
               & "  " + Str(ADO8!dolare) + ", " + Str(ADO8!soless) + ", " _
               & "  0, 0, " + Str(ADO8!abonos) + ", 0 ) ")
               Db.CommitTrans
         
               Call ActualizaSaldos(wSoc, wMes, wccc)
            
            Else
               If Len(Trim(ADO8!pariente)) > 0 And _
                  Len(Trim(ADO8!linparie)) > 0 And _
                  Len(Trim(ADO8!nombre)) > 0 And _
                  (ADO8!conpago = "161" Or ADO8!conpago = "162" Or _
                   ADO8!conpago = "163" Or ADO8!conpago = "164" Or _
                   ADO8!conpago = "165") Then
                  wTipoPariente = ADO8!pariente
                  wNomPariente = ADO8!nombre
                  wLinPariente = ADO8!linparie
                      
                  aa = Leerado8a("SELECT * FROM MAEFAMILIA " _
                           & " WHERE     CODSOCIO = " + Str(wSoc) + " AND " _
                           & "       TIPOPARIENTE = '" + wTipoPariente + "' AND " _
                           & "                LIN = '" + wLinPariente + "' ")
                  If aa = 0 Then
                     Db.BeginTrans
                     Db.Execute ("INSERT INTO MAEFAMILIA " _
                     & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, NUMREC, SERREC ) " _
                     & " VALUES " _
                     & " (" + Str(wSoc) + ", '" + wTipoPariente + "', '" + wLinPariente + "', " _
                     & "  '', '" + wNomPariente + "', '" + Right(wNum, 9) + "', '" + wSer + "'  ) ")
                     Db.CommitTrans
                  Else
                     Db.BeginTrans
                     Db.Execute ("UPDATE MAEFAMILIA " _
                     & " SET NOMBRE = '" + wNomPariente + "', " _
                     & "     NUMREC = '" + Right(wNum, 9) + "', " _
                     & "     SERREC = '" + wSer + "' " _
                     & " WHERE     CODSOCIO = " + Str(wSoc) + " AND " _
                     & "       TIPOPARIENTE = '" + wTipoPariente + "' AND " _
                     & "                LIN = '" + wLinPariente + "' ")
                     Db.CommitTrans
                  End If
                  Set ADO8a = Nothing
               End If
            End If
         End If
         
         If wcon = "128" Then
            Db.BeginTrans
            Db.Execute ("UPDATE FRACDET " _
            & " SET ABONOS = " + Str(ADO8!importe) + ", " _
            & "     SDONEW = CARGOS - " + Str(ADO8!importe) + ", " _
            & "     SERCOB = '" + wSer + "', NUMCOB = '" + wNum + "', " _
            & "     FECCOB = '" + Format(wFec, "dd/mm/yyyy") + "' " _
            & " WHERE NUMERO = '" + wNumFra + "' AND " _
            & "        LINEA = '" + Format(Val(wLinFra), "@@") + "' ")
            Db.CommitTrans
         End If
         
         ADO8.MoveNext
      Loop
   End If
   
'   wAde = 0: wDeu = 0: wSdo = 0
'   aa = Leerado8("SELECT SUM(SDONEW) AS SDONEW " _
'                & " FROM CTASXCAB " _
'                & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
'                & "       MES <= '" + zMesTope + "' AND " _
'                & "       CONCEPTO = '01' ")
'   If aa > 0 Then
'      wSdo = IIf(IsNull(ADO8!sdonew), 0, ADO8!sdonew)
'   End If
'   Set ADO8 = Nothing
   
'   If wSdo > 0 Then
'      wDeu = wSdo
'   Else
'      wAde = -wSdo
'   End If
   
'   aa = Leerado8("SELECT SUM(ABONOS) AS ABONOS " _
'                & " FROM CTASXCAB " _
'                & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
'                & "       MES >= '" + zMesTope + "' AND " _
'                & "       CONCEPTO = '01' ")
'   If aa > 0 Then
'      wAde = wAde + IIf(IsNull(ADO8!abonos), 0, ADO8!abonos)
'   End If
'   Set ADO8 = Nothing
   
'   Db.BeginTrans
'   Db.Execute ("UPDATE MAESOCIO " _
'   & " SET DEUDA_PT2 = " + Str(wDeu) + ", " _
'   & "      ADELANTO = " + Str(wAde) + " " _
'   & " WHERE CODSOCIO = " + Str(wSoc) + " ")
'   Db.CommitTrans
   
'   Db.BeginTrans
'   Db.Execute ("UPDATE ZZZ_MAESTRO " _
'   & " SET DEUDA_PT2 = " + Str(wDeu) + ", " _
'   & "      ADELANTO = " + Str(wAde) + " " _
'   & " WHERE CODSOCIO = " + Str(wSoc) + " ")
'   Db.CommitTrans
       
   wDeu = SaldoFoto(wSoc, zMesTope)
   
   If Trim(wSer) = "004" Then
      Db.BeginTrans
      Db.Execute ("DELETE FROM ZZZ_MRECIBOS " _
      & " WHERE YEAR(FECHA_PAGO) = " + Str(Year(wFec)) + " AND " _
      & "       SERIE = '" + wSer + "' AND " _
      & "       NRO_COMP = " + Str(Val(wNum)) + " ")
      Db.CommitTrans
      
      Dim wImp As Currency
      
      aa = Leerado8a("SELECT * FROM TMP_COBRODET " _
                & " WHERE    ANO = '" + waaa + "' AND " _
                & "          MES = '" + wmmm + "' AND " _
                & "       TIPCOB = '" + wTip + "' AND " _
                & "       SERCOB = '" + wSer + "' AND " _
                & "       NUMCOB = '" + wNum + "' AND " _
                & "          USU = '" + wcodusu + "' " _
                & " ORDER BY LINCOB ")
      If aa > 0 Then
         ADO8a.MoveFirst
         Do While Not ADO8a.EOF
            wLin = ADO8a!lincob
            wMes = ADO8a!mescob
            wcon = ADO8a!conpago
            wccc = ADO8a!concepto
            wImp = ADO8a!importe
      
            aa = Leerado7a("SELECT * FROM ZZZ_MRECIBOS " _
                   & " WHERE    SERIE = '" + wSer + "' AND " _
                   & "       NRO_COMP = " + Str(Val(wNum)) + " AND " _
                   & "       YEAR(FECHA_PAGO) = " + Str(Val(wFec)) + " AND " _
                   & "      LINCOB = '" + wLin + "' ")
            If aa = 0 Then
               Db.BeginTrans
               Db.Execute ("INSERT INTO ZZZ_MRECIBOS " _
               & " (CODIGO, INS, CONCEPTO, SERIE, AUXILIAR, NRO_COMP, MONTO, MONEDA, T_CAMBIO, " _
               & "  FECHA_PAGO, FECHA_CADU, OBS, D_IMPOR, DEUDA_PT2, DINS_CER, ADELANTO, " _
               & "  MARCA1, MARCA2, MARCA3, MARCA4, OBS1, LINCOB ) " _
               & " VALUES " _
               & " (" + Str(wCod) + ", " + Str(wIns) + ", " + Str(Val(wcon)) + ", '" + wSer + "', 0, " _
               & "  " + Str(Val(wNum)) + ", " + Str(wImp) + ", " _
               & "  '" + IIf(txtMoneda.Text = "S", "S/.", "$") + "', " + Str(wCam) + ", " _
               & "  '" + Format(wFec, "dd/mm/yyyy") + "', null, '" + Left(wGlo, 50) + "', " _
               & "  '', " + Str(wDeu) + ", 0, " + Str(wAde) + ", " _
               & "  '" + Format(Date, "dd/mm/yyyy") + "', 'N', '" + wcodusu + "', " _
               & "  '" + Format(Time, "hh:mm:ss") + "', '', '" + wLin + "' ) ")
               Db.CommitTrans
            Else
               Db.BeginTrans
               Db.Execute ("UPDATE ZZZ_MRECIBOS " _
               & " SET CODIGO = " + Str(wCod) + ", INS = " + Str(wIns) + ", CONCEPTO = " + Str(Val(wcon)) + ", " _
               & "     AUXILIAR = 0, MONTO = " + Str(wImp) + ", " _
               & "     MONEDA = '" + IIf(txtMoneda.Text = "S", "S/.", "$") + "', " _
               & "     T_CAMBIO = " + Str(wCam) + ", FECHA_PAGO = '" + Format(wFec, "dd/mm/yyyy") + "', " _
               & "     FECHA_CADU = null, OBS = '" + Left(wGlo, 50) + "', D_IMPOR = '', " _
               & "     DEUDA_PT2 = " + Str(wDeu) + ", DINS_CER = 0, ADELANTO = " + Str(wAde) + ", " _
               & "     MARCA1 = '" + Format(Date, "dd/mm/yyyy") + "', MARCA2 = 'N', " _
               & "     MARCA3 = '" + wcodusu + "', MARCA4 = '" + Format(Time, "hh:mm:ss") + "', " _
               & "     OBS1 = '' " _
               & " WHERE    SERIE = '" + wSer + "' AND " _
               & "       NRO_COMP = " + Str(Val(wNum)) + " AND " _
               & "       YEAR(FECHA_PAGO) = " + Str(Val(wanocia)) + " AND " _
               & "       LINCOB = '" + wLin + "' ")
               Db.CommitTrans
            End If
            Set ADO7a = Nothing
      
            ADO8a.MoveNext
         Loop
      End If
   
      
      
'      aa = Leerado8("SELECT * FROM ZZZ_MRECIBOS " _
'                   & " WHERE    SERIE = '" + wSer + "' AND " _
'                   & "       NRO_COMP = " + Str(Val(wNum)) + " AND " _
'                   & "       YEAR(FECHA_PAGO) = " + Str(Val(wFec)) + " ")
'      If aa = 0 Then
'         Db.BeginTrans
'         Db.Execute ("INSERT INTO ZZZ_MRECIBOS " _
'         & " (CODIGO, INS, CONCEPTO, SERIE, AUXILIAR, NRO_COMP, MONTO, MONEDA, T_CAMBIO, " _
'         & "  FECHA_PAGO, FECHA_CADU, OBS, D_IMPOR, DEUDA_PT2, DINS_CER, ADELANTO, " _
'         & "  MARCA1, MARCA2, MARCA3, MARCA4, OBS1 ) " _
'         & " VALUES " _
'         & " (" + Str(wCod) + ", " + Str(wIns) + ", " + Str(Val(wcon)) + ", '" + wSer + "', 0, " _
'         & "  " + Str(Val(wNum)) + ", " + Str(Val(txtImporte.Text)) + ", " _
'         & "  '" + IIf(txtMoneda.Text = "S", "S/.", "$") + "', " + Str(wcam) + ", " _
'         & "  '" + Format(wFec, "dd/mm/yyyy") + "', null, '" + Left(wGlo, 50) + "', " _
'         & "  '', " + Str(wDeu) + ", 0, " + Str(wAde) + ", " _
'         & "  '" + Format(Date, "dd/mm/yyyy") + "', 'N', '" + wcodusu + "', " _
'         & "  '" + Format(Time, "hh:mm:ss") + "', '' ) ")
'         Db.CommitTrans
'      Else
'         Db.BeginTrans
'         Db.Execute ("UPDATE ZZZ_MRECIBOS " _
'         & " SET CODIGO = " + Str(wCod) + ", INS = " + Str(wIns) + ", CONCEPTO = " + Str(Val(wcon)) + ", " _
'         & "     AUXILIAR = 0, MONTO = " + Str(Val(txtImporte.Text)) + ", " _
'         & "     MONEDA = '" + IIf(txtMoneda.Text = "S", "S/.", "$") + "', " _
'         & "     T_CAMBIO = " + Str(wcam) + ", FECHA_PAGO = '" + Format(wFec, "dd/mm/yyyy") + "', " _
'         & "     FECHA_CADU = null, OBS = '" + Left(wGlo, 50) + "', D_IMPOR = '', " _
'         & "     DEUDA_PT2 = " + Str(wDeu) + ", DINS_CER = 0, ADELANTO = " + Str(wAde) + ", " _
'         & "     MARCA1 = '" + Format(Date, "dd/mm/yyyy") + "', MARCA2 = 'N', " _
'         & "     MARCA3 = '" + wcodusu + "', MARCA4 = '" + Format(Time, "hh:mm:ss") + "', " _
'         & "     OBS1 = '' " _
'         & " WHERE    SERIE = '" + wSer + "' AND " _
'         & "       NRO_COMP = " + Str(Val(wNum)) + " AND " _
'         & "       YEAR(FECHA_PAGO) = " + Str(Val(wanocia)) + " ")
'         Db.CommitTrans
'      End If
   End If
   
'   DoEvents
'   lblFormato.Caption = ""
'   lblFormato.Refresh
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM CTRL_GRABAR WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   ADO1.Requery
   LlenaCab
   LlenaCab1
   
   ACCION = 0
   
   ADO1.Find "TIPCOB='" + wTip + "'"
   ADO1.Find "SERCOB='" + wSer + "'"
   ADO1.Find "NUMCOB='" + wNum + "'"
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
'   cmdEliminar.Visible = Not estado
   cmdEliminar.Visible = False
   cmdAnular.Enabled = Not estado
   
   DataGrid1.Enabled = Not estado
   fraDesplaza.Enabled = Not estado
   
   cmdGrabar.Visible = estado
   cmdDeshacer.Visible = estado
   cmdSalir.Visible = Not estado
   cmdImprimir.Visible = Not estado
End Sub

Private Sub cmbMeses_Click()
   cmbMeses_KeyPress (13)
End Sub

Private Sub cmbMeses_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmbMeses.Text <> "" Then
         If Left(cmbMeses.Text, 2) <> "00" Then
            txtFecCab.Text = "__/__/____"
            fraMantenimiento.Enabled = True
            editar (False)
            Limpiar
            LlenaCab
            LlenaCab1
            refrescar
            cmdNuevo.SetFocus
         Else
            txtFecCab.SetFocus
         End If
      End If
   End If
End Sub

Private Sub cmbTipo_Click()
   cmbTipo_KeyPress (13)
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
   Dim wTip As String, wSer As String, wNew As String, a As Integer
   If KeyAscii = 13 Then
      
      wTip = Format(cmbTipo.ListIndex + 1, "0")
      If cmbTipo.ListIndex = 0 Then
         wSer = "005"
         txtGlosa.Text = "PAGO POR CARNET"
         txtMoneda.Text = "S"
      Else
         wSer = "004"
         txtGlosa.Text = ""
         txtMoneda.Text = "S"
      End If
      
      wNew = "0000000000"
      a = Leerado8("SELECT MAX(NUMCOB) AS NUMCOB " _
                & " FROM COBROCAB " _
                & " WHERE ANO = '" + wanocia + "' AND " _
                & "       TIPCOB = '" + wTip + "' AND " _
                & "       SERCOB = '" + wSer + "' ")
      If a > 0 Then
         wNew = IIf(IsNull(ADO8!numcob), "0000000000", ADO8!numcob)
      End If
      Set ADO8 = Nothing
   
      a = Leerado8("SELECT MAX(NRO_COMP) AS NRO_COMP FROM ZZZ_MRECIBOS " _
                & " WHERE SERIE = '" + wSer + "' AND " _
                & "       YEAR(FECHA_PAGO) = " + Str(Val(wanocia)) + " ")
      If a > 0 Then
         If ADO8!nro_comp > Val(wNew) Then
            wNew = IIf(IsNull(ADO8!nro_comp), "0000000000", ADO8!nro_comp)
         End If
      End If
      Set ADO8 = Nothing
      wNew = Format(Val(wNew) + 1, "0000000000")
      
      txtSerCob.Text = wSer
      txtSerCob.Enabled = False
      txtNumCob.Text = wNew
      txtNumCob.Enabled = False
      If IsDate(txtFecCab.Text) Then
         txtFecha.Text = Format(txtFecCab.Text, "dd/mm/yyyy")
         txtFecha.Enabled = False
      End If
            
      If txtFecha.Enabled = True Then
         txtFecha.SetFocus
      Else
         If txtCodigo.Enabled = True Then
            txtCodigo.SetFocus
         End If
      End If
   End If
End Sub

Private Sub cmbTipoAnula_Click()
   cmbTipoAnula_KeyPress (13)
End Sub

Private Sub cmbTipoAnula_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmbTipoAnula.ListIndex = 0 Then
         txtSerCobAnula.Text = "005"
      Else
         txtSerCobAnula.Text = "004"
      End If
      
      
      txtSerCobAnula.SetFocus
   End If
End Sub

Private Sub cmdAceptar_Click()
   
   If Len(Trim(txtSerCobAnula.Text)) = 0 Then
      MsgBox "Serie Cobranza En Blanco", vbExclamation
      Exit Sub
   End If
   
   If Len(Trim(txtNumCobAnula.Text)) = 0 Then
      MsgBox "Numero Cobranza En Blanco", vbExclamation
      Exit Sub
   End If
   
   fraAnula.Visible = False
   editar True
   
   Dim waaa As String, wmmm As String, wTip As String, wSer As String, wNum As String, _
       aa As Long, wreg As Variant, wLin As String, wSoc As Integer, wFec As Date, _
       wDeu As Currency, wAde As Currency, wSdo As Currency, wCod As Long, wIns As Integer, _
       wNumFra As String, wLinFra As String
   
   waaa = txtAno.Text
   If Left(cmbMeses.Text, 2) <> "00" Then
      wmmm = Left(cmbMeses.Text, 2)
   Else
      wmmm = Format(Month(txtFecCab.Text), "00")
   End If
   wTip = Format(cmbTipoAnula.ListIndex + 1, "0")
   wSer = txtSerCobAnula.Text
   wNum = txtNumCobAnula.Text
   
   If MsgBox("¿Esta seguro de Anular Registro?", vbYesNo + vbDefaultButton2 + vbQuestion, "Advertencia") = vbYes Then
      
      aa = Leerado8("SELECT * FROM TMP_COBROCAB " _
                      & " WHERE    ANO = '" + waaa + "' AND " _
                      & "          MES = '" + wmmm + "' AND " _
                      & "       TIPCOB = '" + wTip + "' AND " _
                      & "       SERCOB = '" + wSer + "' AND " _
                      & "       NUMCOB = '" + wNum + "' ")
      If aa > 0 Then
         wSoc = ADO1!codsocio
         wFec = Format(ADO1!fecha, "dd/mm/yyyy")
      
         aa = Leerado4a("SELECT * FROM COBRODET " _
                   & " WHERE    ANO = '" + waaa + "' AND " _
                   & "          MES = '" + wmmm + "' AND " _
                   & "       TIPCOB = '" + wTip + "' AND " _
                   & "       SERCOB = '" + wSer + "' AND " _
                   & "       NUMCOB = '" + wNum + "' " _
                   & " ORDER BY LINCOB ")
         If aa > 0 Then
            ADO4a.MoveFirst
            Do While Not ADO4a.EOF
               wLin = ADO4a!lincob
               If ADO4a!conpago = "128" Then
                  wNumFra = ADO4a!numfra
                  wLinFra = ADO4a!linfra
               
                  Db.BeginTrans
                  Db.Execute ("UPDATE FRACDET " _
                  & " SET ABONOS = 0, SDONEW = CARGOS, SERCOB = '', NUMCOB = '', LINCOB = '', FECCOB = NULL " _
                  & " WHERE NUMCOB = '" + wNum + "' AND " _
                  & "       NUMERO = '" + wNumFra + "' ")
                  Db.CommitTrans
               End If
         
               Db.BeginTrans
               Db.Execute ("DELETE FROM COBRODET " _
               & " WHERE    ANO = '" + waaa + "' AND " _
               & "          MES = '" + wmmm + "' AND " _
               & "       TIPCOB = '" + wTip + "' AND " _
               & "       SERCOB = '" + wSer + "' AND " _
               & "       NUMCOB = '" + wNum + "' AND " _
               & "       LINCOB = '" + wLin + "' ")
               Db.CommitTrans
         
               Db.BeginTrans
               Db.Execute ("DELETE FROM CTASXDET " _
               & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
               & "            MES = '" + ADO4a!mescob + "' AND " _
               & "       CONCEPTO = '" + ADO4a!concepto + "' AND " _
               & "         TIPCOB = '03' AND " _
               & "         SERCOB = '" + ADO4a!sercob + "' AND " _
               & "         NUMCOB = '" + ADO4a!numcob + "' AND " _
               & "         LINCOB = '" + ADO4a!lincob + "' ")
               Db.CommitTrans
         
               Call ActualizaSaldos(wSoc, ADO4a!mescob, ADO4a!concepto)
          
               ADO4a.MoveNext
            Loop
         End If
      
         Db.BeginTrans
         Db.Execute ("DELETE FROM ZZZ_MRECIBOS " _
         & " WHERE    SERIE = '" + wSer + "' AND " _
         & "       NRO_COMP = " + Str(Val(wNum)) + " AND " _
         & "       YEAR(FECHA_PAGO) = " + Str(Val(wanocia)) + " ")
         Db.CommitTrans
      
         Db.BeginTrans
         Db.Execute ("DELETE FROM COBROCAB " _
         & " WHERE    ANO = '" + waaa + "' AND " _
         & "          MES = '" + wmmm + "' AND " _
         & "       TIPCOB = '" + wTip + "' AND " _
         & "       SERCOB = '" + wSer + "' AND " _
         & "       NUMCOB = '" + wNum + "' ")
         Db.CommitTrans
   
         Db.BeginTrans
         Db.Execute ("DELETE FROM TMP_COBROCAB " _
         & " WHERE    ANO = '" + waaa + "' AND " _
         & "          MES = '" + wmmm + "' AND " _
         & "          USU = '" + wcodusu + "' AND " _
         & "       TIPCOB = '" + wTip + "' AND " _
         & "       SERCOB = '" + wSer + "' AND " _
         & "       NUMCOB = '" + wNum + "' ")
         Db.CommitTrans
   
         wAde = 0: wDeu = 0: wSdo = 0
         aa = Leerado8("SELECT SUM(SDONEW) AS SDONEW " _
                & " FROM CTASXCAB " _
                & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                & "       MES < '" + zMesTope + "' AND " _
                & "       CONCEPTO = '01' ")
         If aa > 0 Then
            wSdo = IIf(IsNull(ADO8!sdonew), 0, ADO8!sdonew)
         End If
         Set ADO8 = Nothing
   
         If wSdo > 0 Then
            wDeu = wSdo
         Else
            wAde = -wSdo
         End If
      
         Db.BeginTrans
         Db.Execute ("UPDATE MAESOCIO " _
         & " SET DEUDA_PT2 = " + Str(wDeu) + ", " _
         & "      ADELANTO = " + Str(wAde) + " " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " ")
         Db.CommitTrans
   
         Db.BeginTrans
         Db.Execute ("UPDATE ZZZ_MAESTRO " _
         & " SET DEUDA_PT2 = " + Str(wDeu) + ", " _
         & "      ADELANTO = " + Str(wAde) + " " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " ")
         Db.CommitTrans
      End If
      
      Db.BeginTrans
      Db.Execute ("INSERT INTO COBRODET " _
      & " (ANO, MES, TIPCOB, SERCOB, NUMCOB, LINCOB, MESCOB, CONPAGO, DOLARE, SOLESS, " _
      & "  MONDOC, SDOOLD, ABONOS, SDONEW, IMPORTE, CONCEPTO, PARIENTE, NOMBRE, NUMFRA, LINFRA, NUMOPE) " _
      & " VALUES " _
      & " ('" + waaa + "', '" + wmmm + "', '" + wTip + "', '" + wSer + "', '" + wNum + "', '01', " _
      & "  '', '', 0, 0, '', 0, 0, 0, 0, '', '', '', '', '', '' ) ")
      Db.CommitTrans

      Db.BeginTrans
      Db.Execute ("INSERT INTO COBROCAB " _
      & " (ANO, MES, TIPCOB, SERCOB, NUMCOB, FECHA, MONEDA, IMPORTE, GLOSA, CODSOCIO, " _
      & "  TIPCAM, DOLARE, SOLESS, FORPAG, USU ) " _
      & " VALUES " _
      & "  ('" + waaa + "', '" + wmmm + "', '" + wTip + "', '" + wSer + "', '" + wNum + "', " _
      & "   '" + Format(txtFecha.Text, "dd/mm/yyyy") + "', 'S', " _
      & "   0, 'DOCUMENTO ANULADO', 0, 0, 0, 0, '', '" + wcodusu + "' ) ")
      Db.CommitTrans
      
      Db.BeginTrans
      Db.Execute ("INSERT INTO TMP_COBROCAB " _
      & " (ANO, MES, TIPCOB, SERCOB, NUMCOB, FECHA, MONEDA, IMPORTE, GLOSA, CODSOCIO, NOMBRE, " _
      & "  TIPCAM, DOLARE, SOLESS, FORPAG, USU ) " _
      & " VALUES " _
      & "  ('" + waaa + "', '" + wmmm + "', '" + wTip + "', '" + wSer + "', '" + wNum + "', " _
      & "   '" + Format(txtFecha.Text, "dd/mm/yyyy") + "', 'S', " _
      & "   0, 'DOCUMENTO ANULADO', 0, '', 0, 0, 0, '', '" + wcodusu + "' ) ")
      Db.CommitTrans
      
      ADO1.Requery
      LlenaCab1
      Limpiar
      ADO1.Find "TIPCOB='" + wTip + "'"
      ADO1.Find "SERCOB='" + wSer + "'"
      ADO1.Find "NUMCOB='" + wNum + "'"
      refrescar
   End If
   
   editar False
   fraMantenimiento.Visible = True
   
   DataGrid1.SetFocus
   
   ACCION = 0
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdAnular_Click()
   On Error GoTo err
   
   If Not IsDate(txtFecCab.Text) Then
      MsgBox "Fecha de Proceso Invalida", vbExclamation
      txtFecCab.SetFocus
      Exit Sub
   End If
   
   fraAnula.Visible = True
   editar True
   fraMantenimiento.Visible = False
   
   cmbTipoAnula.SetFocus
   cmbTipoAnula.ListIndex = Val(ADO1!tipcob) - 1
   txtSerCob.Text = IIf(IsNull(ADO1!sercob), "", ADO1!sercob)
   txtNumCob.Text = IIf(IsNull(ADO1!numcob), "", ADO1!numcob)
   
   cmbTipoAnula.SetFocus
   Exit Sub
err:
   MsgBox err.Description
   Resume Next
End Sub

Private Sub cmdCancelaAnula_Click()
   MsgBox "Los Cambios Efectuados Se Perderán", vbExclamation
   ACCION = 0
   
   editar (False)
   
   fraAnula.Visible = False
   fraMantenimiento.Visible = True
   
   DataGrid1.SetFocus
End Sub

Private Sub cmdDeshacer_Click()
   MsgBox "Los Cambios Efectuados Se Perderán", vbExclamation
   ACCION = 0
   
   editar (False)
   
   Limpiar
   llenadet
   llenadet1
   TotalDet
   refrescar
   DataGrid1.SetFocus
End Sub

Private Sub cmdEliminar_Click()
   On Error GoTo err
   
   If Not IsDate(txtFecCab.Text) Then
      MsgBox "Fecha de Proceso Invalida", vbExclamation
      txtFecCab.SetFocus
      Exit Sub
   End If
   
   Dim waaa As String, wmmm As String, wTip As String, wSer As String, wNum As String, _
       aa As Long, wreg As Variant, wLin As String, wSoc As Integer, wFec As Date, _
       wDeu As Currency, wAde As Currency, wSdo As Currency, wCod As Long, wIns As Integer, _
       wNumFra As String, wLinFra As String
   
   If ADO1.BOF Or ADO1.EOF Then
      Exit Sub
   End If
   waaa = txtAno.Text
   If Left(cmbMeses.Text, 2) <> "00" Then
      wmmm = Left(cmbMeses.Text, 2)
   Else
      wmmm = Format(Month(txtFecCab.Text), "00")
   End If
   wTip = Format(cmbTipo.ListIndex + 1, "0")
   wSer = ADO1!sercob
   wNum = ADO1!numcob
   wSoc = ADO1!codsocio
   wFec = Format(ADO1!fecha, "dd/mm/yyyy")
   
   If MsgBox("¿Esta seguro de borrar Registro?", vbYesNo + vbDefaultButton2 + vbQuestion, "Advertencia") = vbYes Then
      
      ADO1.MoveNext
      If Not ADO1.EOF Then
         wreg = ADO1.Bookmark
      Else
         ADO1.MovePrevious
         ADO1.MovePrevious
         If ADO1.BOF Then
            wreg = 0
         Else
            wreg = ADO1.Bookmark
         End If
      End If
      
      aa = Leerado4a("SELECT * FROM COBRODET " _
                   & " WHERE    ANO = '" + waaa + "' AND " _
                   & "          MES = '" + wmmm + "' AND " _
                   & "       TIPCOB = '" + wTip + "' AND " _
                   & "       SERCOB = '" + wSer + "' AND " _
                   & "       NUMCOB = '" + wNum + "' " _
                   & " ORDER BY LINCOB ")
      If aa > 0 Then
         ADO4a.MoveFirst
         Do While Not ADO4a.EOF
            wLin = ADO4a!lincob
            If ADO4a!conpago = "128" Then
               wNumFra = ADO4a!numfra
               wLinFra = ADO4a!linfra
               
               Db.BeginTrans
               Db.Execute ("UPDATE FRACDET " _
               & " SET ABONOS = 0, SDONEW = CARGOS, SERCOB = '', NUMCOB = '', LINCOB = '', FECCOB = NULL " _
               & " WHERE NUMCOB = '" + wNum + "' AND " _
               & "       NUMERO = '" + wNumFra + "' ")
               Db.CommitTrans
               
            End If
         
            Db.BeginTrans
            Db.Execute ("DELETE FROM COBRODET " _
            & " WHERE    ANO = '" + waaa + "' AND " _
            & "          MES = '" + wmmm + "' AND " _
            & "       TIPCOB = '" + wTip + "' AND " _
            & "       SERCOB = '" + wSer + "' AND " _
            & "       NUMCOB = '" + wNum + "' AND " _
            & "       LINCOB = '" + wLin + "' ")
            Db.CommitTrans
         
            Db.BeginTrans
            Db.Execute ("DELETE FROM CTASXDET " _
            & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
            & "            MES = '" + ADO4a!mescob + "' AND " _
            & "       CONCEPTO = '" + ADO4a!concepto + "' AND " _
            & "         TIPCOB = '03' AND " _
            & "         SERCOB = '" + ADO4a!sercob + "' AND " _
            & "         NUMCOB = '" + ADO4a!numcob + "' AND " _
            & "         LINCOB = '" + ADO4a!lincob + "' ")
            Db.CommitTrans
         
'            Db.BeginTrans
'            Db.Execute ("DELETE FROM ZZZ_MRECIBOS " _
'            & " WHERE CODIGO =  ")
'            Db.CommitTrans
         
            Call ActualizaSaldos(wSoc, ADO4a!mescob, ADO4a!concepto)
         
            ADO4a.MoveNext
         Loop
      End If
      
      Db.BeginTrans
      Db.Execute ("DELETE FROM ZZZ_MRECIBOS " _
      & " WHERE    SERIE = '" + wSer + "' AND " _
      & "       NRO_COMP = " + Str(Val(wNum)) + " AND " _
      & "       YEAR(FECHA_PAGO) = " + Str(Val(wanocia)) + " ")
      Db.CommitTrans
      
      Db.BeginTrans
      Db.Execute ("DELETE FROM COBROCAB " _
      & " WHERE    ANO = '" + waaa + "' AND " _
      & "          MES = '" + wmmm + "' AND " _
      & "       TIPCOB = '" + wTip + "' AND " _
      & "       SERCOB = '" + wSer + "' AND " _
      & "       NUMCOB = '" + wNum + "' ")
      Db.CommitTrans
   
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_COBROCAB " _
      & " WHERE    ANO = '" + waaa + "' AND " _
      & "          MES = '" + wmmm + "' AND " _
      & "          USU = '" + wcodusu + "' AND " _
      & "       TIPCOB = '" + wTip + "' AND " _
      & "       SERCOB = '" + wSer + "' AND " _
      & "       NUMCOB = '" + wNum + "' ")
      Db.CommitTrans
   
      wAde = 0: wDeu = 0: wSdo = 0
      aa = Leerado8("SELECT SUM(SDONEW) AS SDONEW " _
                & " FROM CTASXCAB " _
                & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                & "       MES < '" + zMesTope + "' AND " _
                & "       CONCEPTO = '01' ")
      If aa > 0 Then
         wSdo = IIf(IsNull(ADO8!sdonew), 0, ADO8!sdonew)
      End If
      Set ADO8 = Nothing
   
      If wSdo > 0 Then
         wDeu = wSdo
      Else
         wAde = -wSdo
      End If
      
      Db.BeginTrans
      Db.Execute ("UPDATE MAESOCIO " _
      & " SET DEUDA_PT2 = " + Str(wDeu) + ", " _
      & "      ADELANTO = " + Str(wAde) + " " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   
      Db.BeginTrans
      Db.Execute ("UPDATE ZZZ_MAESTRO " _
      & " SET DEUDA_PT2 = " + Str(wDeu) + ", " _
      & "      ADELANTO = " + Str(wAde) + " " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
      
      ADO1.Requery
      LlenaCab1
      If wreg <> 0 Then
         ADO1.Bookmark = wreg
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

Private Sub cmdGrabar_Click()
   On Error GoTo err
   Dim aa As Integer, wmmm As String, _
       wTip As String, wSer As String, wCob As String, wcobnew As String, wcobold As String
'   If ACCION = 1 Then
'      If Left(cmbMeses.Text, 2) <> "00" Then
'         wmmm = Left(cmbMeses.Text, 2)
'      Else
'         wmmm = Format(Month(txtFecCab.Text), "00")
'      End If
'      wTip = Format(cmbTipo.ListIndex + 1, "0")
'      wSer = txtSerCob.Text
'      wCob = txtNumCob.Text
'      If Leerado8("SELECT * FROM COBROCAB " _
'                & " WHERE NUMCOB = '" + wCob + "' AND " _
'                & "       TIPCOB = '" + wTip + "' AND " _
'                & "       SERCOB = '" + wSer + "' ") > 0 Then
'         wcobold = wCob
'         aa = Leerado8("SELECT MAX(NUMCOB) AS NUMCOB " _
'                    & " FROM COBROCAB " _
'                    & " WHERE    ANO = '" + wanocia + "' AND " _
'                    & "          MES = '" + wmmm + "' AND " _
'                    & "       TIPCOB = '" + wTip + "' AND SERCOB = '" + wSer + "'")
'         If aa > 0 Then
'            wcobnew = Format(Val(IIf(IsNull(ADO8!numcob), "000000000", ADO8!numcob)) + 1, "000000000")
'         End If
'         txtNumCob.Text = wcobnew
'      End If
'   End If
   If validaCob Then
      MsgBox "Cobranza Con Errores, No Se Graba", vbExclamation
      Exit Sub
   End If
   
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   
   grabar
   
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   
   editar False
   MsgBox "Cobranza Grabada OK", vbExclamation
   ACCION = 0
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdImprimir_Click()
   Dim wSer As String, wNum As String, wNumLet As String, wMonto As String
   wSer = txtSerCob.Text
   wNum = txtNumCob.Text
   wNumLet = NumLetras(ADO1!importe) + " " + IIf(ADO1!moneda = "S", "SOLES", "DOLARES USA")
   wMonto = IIf(ADO1!moneda = "S", "S/.", "US$") + Format(ADO1!importe, "###,##0.00")
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\Recibo.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "USU= '" + wnomusu + "' "
   Crys1.Formulas(3) = "NUMLET= 'SON " + wNumLet + "' "
   Crys1.Formulas(4) = "MONTO= '" + wMonto + "' "
   Crys1.SelectionFormula = " {TMP_COBRODET.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdModificar_Click()
   
   If Not IsDate(txtFecCab.Text) Then
      MsgBox "Fecha de Proceso Invalida", vbExclamation
      txtFecCab.SetFocus
      Exit Sub
   End If
   
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
   If txtFecha.Enabled = True Then
      txtFecha.SetFocus
   End If
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
   
   If Not IsDate(txtFecCab.Text) Then
      MsgBox "Fecha de Proceso Invalida", vbExclamation
      txtFecCab.SetFocus
      Exit Sub
   End If
   
   ACCION = 1
   lblFormato.Caption = "NUEVO"
   DataGrid2.AllowDelete = True
   DataGrid2.AllowUpdate = True
   
   DataGrid2.Refresh
   
   Limpiar
   
   If Val(wanocia) = Year(Date) And Val(Mid(cmbMeses.Text, 1, 2)) = Month(Date) Then
      txtFecha.Text = Format(Date, "dd/mm/yyyy")
   Else
      txtFecha.Text = "__/__/____"
   End If
   txtCodigo.Text = ""
   txtCodSocio.Text = ""
   txtIns.Text = ""
   txtSerCob.Text = ""
   txtNumCob.Text = ""
   txtTipCam.Text = ""
   txtMoneda.Text = "S"
   txtImporte.Text = ""
   lblTotDol.Caption = ""
   lblTotSol.Caption = ""
   txtForPag.Text = "01"
   

   llenadet
   llenadet1
       
   editar True
   
   cmbTipo.SetFocus
End Sub

Private Sub cmdSalir_Click()
   Unload Me
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
   If ADO2.RecordCount = 0 Then
      CreaDet
      TotalDet
      DataGrid2.Row = 0
      DataGrid2.col = 1
      DataGrid2.Text = IIf(IsNull(ADO2!conpago), 0, ADO2!conpago)
   End If
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
    Case 116  '' F5
         Select Case DataGrid2.col
         Case 1
              xlista = "CO"
              xseleccion = Trim(DataGrid2.Text)
              zSerCaj = txtSerCob.Text
              zMonCaj = txtMoneda.Text
              frmSeleccion.Show 1
              If xseleccion <> "" Then
                 DataGrid2.col = 1
                 DataGrid2.Text = xseleccion
                 ADO2!conpago = xseleccion
                 ADO2.Update
              End If
         Case 7
              xlista = "TP"
              xseleccion = Trim(DataGrid2.Text)
              frmSeleccion.Show 1
              If xseleccion <> "" Then
                 DataGrid2.col = 5
                 DataGrid2.Text = xseleccion
                 ADO2!pariente = xseleccion
                 ADO2.Update
              End If
         Case 9
              xlista = "FA"
              zFamSocio = Val(txtCodSocio.Text)
              zFamParie = ADO2!pariente
              zLinParie = ADO2!linparie
              xseleccion = Trim(DataGrid2.Text)
              frmSelecFam.Show 1
              If xseleccion <> "" Then
                 DataGrid2.col = 7
                 DataGrid2.Text = Trim(xseleccion)
                 ADO2!nombre = Trim(xseleccion)
                 ADO2.Update
              End If
         End Select
    Case 40  ' DOWN
         If ACCION = 1 Or ACCION = 2 Then
            
         wvariable = DataGrid2.Text
         
         Select Case DataGrid2.col
         Case 1
              ADO2!conpago = IIf(IsNull(wvariable), "", wvariable)
         Case 2
              ADO2!nomcon = IIf(IsNull(wvariable), "", wvariable)
         Case 3
              ADO2!numfra = IIf(IsNull(wvariable), "", wvariable)
         Case 4
              ADO2!linfra = IIf(IsNull(wvariable), "", wvariable)
         Case 5
              ADO2!mescob = IIf(IsNull(wvariable), "", wvariable)
         Case 6
              ADO2!importe = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 7
              ADO2!pariente = IIf(IsNull(wvariable), "", wvariable)
         Case 8
              ADO2!linparie = IIf(IsNull(wvariable), "", wvariable)
         Case 9
              ADO2!nombre = IIf(IsNull(wvariable), "", wvariable)
         End Select
    
         TotalDet
         
         Select Case DataGrid2.col
         Case 0
              DataGrid2.Text = ADO2!lincob
         Case 1
              DataGrid2.Text = IIf(IsNull(ADO2!conpago), "", ADO2!conpago)
         Case 2
              DataGrid2.Text = IIf(IsNull(ADO2!nomcon), "", ADO2!nomcon)
         Case 3
              DataGrid2.Text = IIf(IsNull(ADO2!numfra), "", ADO2!numfra)
         Case 4
              DataGrid2.Text = IIf(IsNull(ADO2!linfra), "", ADO2!linfra)
         Case 5
              DataGrid2.Text = IIf(IsNull(ADO2!mescob), "", ADO2!mescob)
         Case 6
              DataGrid2.Text = IIf(IsNull(ADO2!importe), 0, ADO2!importe)
         Case 7
              DataGrid2.Text = IIf(IsNull(ADO2!pariente), "", ADO2!pariente)
         Case 8
              DataGrid2.Text = IIf(IsNull(ADO2!linparie), "", ADO2!linparie)
         Case 9
              DataGrid2.Text = IIf(IsNull(ADO2!nombre), "", ADO2!nombre)
         End Select
            
         If ADO2.AbsolutePosition = ADO2.RecordCount And Len(Trim(ADO2!conpago)) > 0 Then
            CreaDet
            TotalDet
         End If
         End If
    Case 37 ' Retroceder
         wvariable = DataGrid2.Text
         
         Select Case DataGrid2.col
         Case 1
              ADO2!conpago = IIf(IsNull(wvariable), "", wvariable)
         Case 2
              ADO2!nomcon = IIf(IsNull(wvariable), "", wvariable)
         Case 3
              ADO2!numfra = IIf(IsNull(wvariable), "", wvariable)
         Case 4
              ADO2!linfra = IIf(IsNull(wvariable), "", wvariable)
         Case 5
              ADO2!mescob = IIf(IsNull(wvariable), "", wvariable)
         Case 6
              ADO2!importe = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 7
              ADO2!pariente = IIf(IsNull(wvariable), "", wvariable)
         Case 8
              ADO2!linparie = IIf(IsNull(wvariable), "", wvariable)
         Case 9
              ADO2!nombre = IIf(IsNull(wvariable), "", wvariable)
         End Select
         
         TotalDet
         
         If DataGrid2.col = 1 Then
            If DataGrid2.Row > 0 Then
               DataGrid2.Row = DataGrid2.Row - 1
            End If
            DataGrid2.col = 0
         End If
         
         Select Case DataGrid2.col
         Case 0
              DataGrid2.Text = ADO2!lincob
         Case 1
              DataGrid2.Text = IIf(IsNull(ADO2!conpago), "", ADO2!conpago)
         Case 2
              DataGrid2.Text = IIf(IsNull(ADO2!nomcon), "", ADO2!nomcon)
         Case 3
              DataGrid2.Text = IIf(IsNull(ADO2!numfra), "", ADO2!numfra)
         Case 4
              DataGrid2.Text = IIf(IsNull(ADO2!linfra), "", ADO2!linfra)
         Case 5
              DataGrid2.Text = IIf(IsNull(ADO2!mescob), "", ADO2!mescob)
         Case 6
              DataGrid2.Text = IIf(IsNull(ADO2!importe), 0, ADO2!importe)
         Case 7
              DataGrid2.Text = IIf(IsNull(ADO2!pariente), "", ADO2!pariente)
         Case 8
              DataGrid2.Text = IIf(IsNull(ADO2!linparie), "", ADO2!linparie)
         Case 9
              DataGrid2.Text = IIf(IsNull(ADO2!nombre), "", ADO2!nombre)
         End Select
         
    Case 38 ' Subir
         wvariable = DataGrid2.Text
         
         Select Case DataGrid2.col
         Case 1
              ADO2!conpago = IIf(IsNull(wvariable), "", wvariable)
         Case 2
              ADO2!nomcon = IIf(IsNull(wvariable), "", wvariable)
         Case 3
              ADO2!numfra = IIf(IsNull(wvariable), "", wvariable)
         Case 4
              ADO2!linfra = IIf(IsNull(wvariable), "", wvariable)
         Case 5
              ADO2!mescob = IIf(IsNull(wvariable), "", wvariable)
         Case 6
              ADO2!importe = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 7
              ADO2!pariente = IIf(IsNull(wvariable), "", wvariable)
         Case 8
              ADO2!linparie = IIf(IsNull(wvariable), "", wvariable)
         Case 9
              ADO2!nombre = IIf(IsNull(wvariable), "", wvariable)
         End Select
    
         TotalDet
         
         Select Case DataGrid2.col
         Case 0
              DataGrid2.Text = ADO2!lincob
         Case 1
              DataGrid2.Text = IIf(IsNull(ADO2!conpago), "", ADO2!conpago)
         Case 2
              DataGrid2.Text = IIf(IsNull(ADO2!nomcon), "", ADO2!nomcon)
         Case 3
              DataGrid2.Text = IIf(IsNull(ADO2!numfra), "", ADO2!numfra)
         Case 4
              DataGrid2.Text = IIf(IsNull(ADO2!linfra), "", ADO2!linfra)
         Case 5
              DataGrid2.Text = IIf(IsNull(ADO2!mescob), "", ADO2!mescob)
         Case 6
              DataGrid2.Text = IIf(IsNull(ADO2!importe), 0, ADO2!importe)
         Case 7
              DataGrid2.Text = IIf(IsNull(ADO2!pariente), "", ADO2!pariente)
         Case 8
              DataGrid2.Text = IIf(IsNull(ADO2!linparie), "", ADO2!linparie)
         Case 9
              DataGrid2.Text = IIf(IsNull(ADO2!nombre), "", ADO2!nombre)
         End Select
    
    Case 39 ' Avanzar
         wvariable = DataGrid2.Text
         
         Select Case DataGrid2.col
         Case 1
              ADO2!conpago = IIf(IsNull(wvariable), "", wvariable)
         Case 2
              ADO2!nomcon = IIf(IsNull(wvariable), "", wvariable)
         Case 3
              ADO2!numfra = IIf(IsNull(wvariable), "", wvariable)
         Case 4
              ADO2!linfra = IIf(IsNull(wvariable), "", wvariable)
         Case 5
              ADO2!mescob = IIf(IsNull(wvariable), "", wvariable)
         Case 6
              ADO2!importe = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 7
              ADO2!pariente = IIf(IsNull(wvariable), "", wvariable)
         Case 8
              ADO2!linparie = IIf(IsNull(wvariable), "", wvariable)
         Case 9
              ADO2!nombre = IIf(IsNull(wvariable), "", wvariable)
         End Select
         
         TotalDet
         
         If DataGrid2.col = 9 Then
            If Val(ADO2!lincob) < ADO2.RecordCount Then
               DataGrid2.Row = DataGrid2.Row + 1
            End If
            DataGrid2.col = 0
         End If
          
         Select Case DataGrid2.col
         Case 0
              DataGrid2.Text = ADO2!lincob
         Case 1
              DataGrid2.Text = IIf(IsNull(ADO2!conpago), "", ADO2!conpago)
         Case 2
              DataGrid2.Text = IIf(IsNull(ADO2!nomcon), "", ADO2!nomcon)
         Case 3
              DataGrid2.Text = IIf(IsNull(ADO2!numfra), "", ADO2!numfra)
         Case 4
              DataGrid2.Text = IIf(IsNull(ADO2!linfra), "", ADO2!linfra)
         Case 5
              DataGrid2.Text = IIf(IsNull(ADO2!mescob), "", ADO2!mescob)
         Case 6
              DataGrid2.Text = IIf(IsNull(ADO2!importe), 0, ADO2!importe)
         Case 7
              DataGrid2.Text = IIf(IsNull(ADO2!pariente), "", ADO2!pariente)
         Case 8
              DataGrid2.Text = IIf(IsNull(ADO2!linparie), "", ADO2!linparie)
         Case 9
              DataGrid2.Text = IIf(IsNull(ADO2!nombre), "", ADO2!nombre)
         End Select
    
    Case 45 ' Insertar
         If Len(Trim(ADO2!conpago)) > 0 Then
            insertlinea ADO2!lincob
            TotalDet
            DataGrid2.col = 1
            DataGrid2.SelStart = 0
            DataGrid2.SelLength = Len(Trim(DataGrid2.Text))
         End If
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
    
    Dim zSoc As Integer
    
    zSoc = Val(txtCodSocio.Text)
    waaa = Format(Year(txtFecCab.Text), "0000")
    wmmm = Format(Month(txtFecCab.Text), "00")
    
    On Error GoTo err
    Select Case KeyAscii
    Case 13
       Select Case DataGrid2.col
       Case 0  ' Linea
            DataGrid2.col = 1
       Case 1  ' ConPago
            wvariable = DataGrid2.Text
            wlll = Len(Trim(wvariable))
            If wlll = 0 Then
               MsgBox "Concepto En Blanco", vbInformation
               xlista = "CO"
               frmSeleccion.Show 1
               If xseleccion <> "" Then
                  DataGrid2.Text = xseleccion
               End If
               Exit Sub
            End If
            Select Case txtSerCob
            Case "004"
                 If wvariable = "160" Or wvariable = "161" Or wvariable = "162" Or _
                    wvariable = "163" Or wvariable = "164" Or wvariable = "164" Then
            
                    MsgBox "Los Conceptos de CARNET son en Serie 005", vbExclamation
                    wvariable = ""
                    DataGrid2.Text = ""
                    ADO2!conpago = ""
                    ADO2!nomcon = ""
                    Exit Sub
                 
                 End If
            Case "005"
                 If wvariable <> "160" And wvariable <> "161" And wvariable <> "162" And _
                    wvariable <> "163" And wvariable <> "164" And wvariable <> "164" Then
            
                    MsgBox "Serie 005 Es Solo Para CARNET, Usar Serie 004", vbExclamation
                    wvariable = ""
                    DataGrid2.Text = ""
                    ADO2!conpago = ""
                    ADO2!nomcon = ""
                    Exit Sub
            
                 End If
            End Select
            
            If ACCION = 1 And (Len(Trim(ADO2!numfra)) = 0 Or Len(Trim(ADO2!linfra)) = 0) Then
            If wvariable = "128" Then
               ADO2!nomcon = "PAGO X FRACCIONAM."
               If Len(Trim(ADO2!numfra)) = 0 Then
                  ADO2!numfra = BuscaFrac(zSoc, ADO2!lincob, 1)
               End If
               If Len(Trim(ADO2!linfra)) = 0 Then
                  ADO2!linfra = Format(BuscaFrac(zSoc, ADO2!lincob, 2), "@@")
               End If
               ADO2!importe = Val(BuscaFrac(zSoc, ADO2!lincob, 3))
               ADO2!mescob = BuscaFrac(zSoc, ADO2!lincob, 4)
               If ADO2!numfra = "" Then
                  ADO2!conpago = ""
                  ADO2!concepto = ""
                  ADO2!nomcon = ""
               End If
               ADO2!sdoold = ADO2!importe
               ADO2!abonos = ADO2!importe
               ADO2!sdonew = 0
            Else
               ADO2!numfra = ""
               ADO2!linfra = ""
            End If
            End If
            
            c = Leerado8("SELECT * FROM ZZZ_CONCEPTO WHERE CONCEPTO = '" + wvariable + " ' ")
            If c = 0 Then
               MsgBox ("Concepto " + wvariable + " No Existe")
               xlista = "CO"
               frmSeleccion.Show 1
               If xseleccion <> "" Then
                  DataGrid2.Text = xseleccion
               End If
               Exit Sub
            End If
            DataGrid2.col = 1
            DataGrid2.Text = wvariable
            ADO2!conpago = wvariable
            ADO2!concepto = IIf(IsNull(ADO8!ctasxcab), "", ADO8!ctasxcab)
            
            Select Case ADO2!concepto
            Case "01"
                 ADO2!nomcon = ADO8!desconce
'                 ADO2!mescob = BuscaUltimoMes(zSoc, ADO2!concepto, ADO2!lincob)
                                  
                 
                 ADO2!mescob = waaa + "/" + wmmm
                 ADO2!importe = BuscaUltimoApo(zSoc, ADO2!mescob, ADO2!concepto)
                 
                 ADO2!mondoc = txtMoneda.Text
                 ADO2!sdoold = ADO2!importe
                 ADO2!abonos = ADO2!importe
                 ADO2!sdonew = 0
            Case "02"
                 ADO2!nomcon = ADO8!desconce
                 ADO2!mescob = BuscaUltimoMes(zSoc, ADO2!concepto, ADO2!lincob)
                 ADO2!importe = BuscaUltimoApo(zSoc, ADO2!mescob, ADO2!concepto)
                 ADO2!mondoc = txtMoneda.Text
                 ADO2!sdoold = ADO2!importe
                 ADO2!abonos = ADO2!importe
                 ADO2!sdonew = 0
            Case "03"
                 ADO2!mondoc = txtMoneda.Text
            Case Else
                 ADO2!concepto = ""
                 ADO2!nomcon = ADO8!desconce
                 ADO2!mescob = ""
                 ADO2!numfra = ""
                 ADO2!linfra = ""
                 ADO2!importe = ADO8!importe
                 If ADO2!conpago = "160" Or ADO2!conpago = "161" Or ADO2!conpago = "162" Or _
                    ADO2!conpago = "163" Or ADO2!conpago = "164" Or ADO2!conpago = "165" Then
                    ADO2!mondoc = ""
                    ADO2!sdoold = 0
                    ADO2!abonos = 0
                    ADO2!sdonew = 0
                 Else
                    ADO2!mondoc = "S"
                    ADO2!sdoold = ADO8!importe
                    ADO2!abonos = ADO8!importe
                    ADO2!sdonew = 0
                 End If
            End Select
            If txtMoneda.Text = "S" Then
               ADO2!dolare = 0
               ADO2!soless = ADO2!importe
            Else
               ADO2!soless = 0
               ADO2!dolare = ADO2!importe
            End If
            Select Case ADO2!conpago
            Case "160"
                 ADO2!pariente = ""
                 ADO2!linparie = ""
                 ADO2!nombre = lblCodSocio.Caption
            Case "161"
                 ADO2!pariente = "E"
                 ADO2!linparie = "01"
            Case "162"
                 ADO2!pariente = "H"
                 ADO2!linparie = "01"
            Case "163"
                 ADO2!pariente = "P"
                 ADO2!linparie = "01"
            Case "164"
                 ADO2!pariente = "M"
                 ADO2!linparie = "01"
            Case "165"
                 ADO2!pariente = "N"
                 ADO2!linparie = "01"
            Case Else
                 ADO2!pariente = ""
                 ADO2!linparie = ""
                 ADO2!nombre = ""
            End Select
            ADO2!nombre = BuscaPariente(zSoc, ADO2!pariente, ADO2!linparie)
            
            If wvariable <> "128" Then
               If Len(Trim(ADO2!mescob)) > 0 Then
                  DataGrid2.col = 5
               Else
                  DataGrid2.col = 6
               End If
            End If
       Case 2  ' NonCon
            DataGrid2.col = 3
            
       Case 3  ' NumFra
            If ADO2!conpago = "128" Then
               wvariable = Trim(DataGrid2.Text)
               wlll = Len(Trim(wvariable))
               If wlll = 0 Then
                  MsgBox "Numero Frac.En Blanco", vbInformation
                  Exit Sub
               End If
               ADO2!numfra = IIf(IsNull(wvariable) Or Len(Trim(wvariable)) = 0, "", wvariable)
            Else
               ADO2!numfra = ""
               ADO2!linfra = ""
            End If
            DataGrid2.col = 4
       Case 4  ' LinFra
            If ADO2!conpago = "128" Then
               wvariable = Trim(DataGrid2.Text)
               wlll = Len(Trim(wvariable))
               If wlll = 0 Then
                  MsgBox "Linea Frac.En Blanco", vbInformation
                  Exit Sub
               End If
               ADO2!linfra = IIf(IsNull(wvariable) Or Len(Trim(wvariable)) = 0, "", wvariable)
            Else
               ADO2!numfra = ""
               ADO2!linfra = ""
            End If
            DataGrid2.col = 5
       Case 5  ' MesCob
            wvariable = Trim(DataGrid2.Text)
            wlll = Len(Trim(wvariable))
            If wlll = 0 Then
               MsgBox "Mes de Pago En Blanco", vbInformation
               Exit Sub
            End If
            If Len(Trim(wvariable)) <> 7 Then
               MsgBox "Longitud del Mes Invalido", vbExclamation
               Exit Sub
            End If
            waaa = Left(wvariable, 4)
            wmmm = Right(wvariable, 2)
            If wvariable <> waaa + "/" + wmmm Then
               MsgBox "Formato del Mes Invalido", vbExclamation
               Exit Sub
            End If
            DataGrid2.Text = wvariable
            ADO2!mescob = wvariable
            DataGrid2.col = 6
       Case 6  ' Importe
            If ADO2!conpago = "128" Then
               MsgBox "Importe No Se Puede Modificar", vbExclamation
               
               ADO2!importe = BuscaCargoFrac(zSoc, ADO2!numfra, ADO2!linfra)
               DataGrid2.Text = BuscaCargoFrac(zSoc, ADO2!numfra, ADO2!linfra)
                
               Exit Sub
            End If
            wvariable = Trim(DataGrid2.Text)
            DataGrid2.Text = wvariable
            ADO2!importe = IIf(IsNull(wvariable) Or Len(Trim(wvariable)) = 0, 0, wvariable)
            
            If ADO2!conpago = "160" Or _
               ADO2!conpago = "161" Or _
               ADO2!conpago = "162" Or _
               ADO2!conpago = "163" Or _
               ADO2!conpago = "164" Or _
               ADO2!conpago = "165" Then
               ADO2!mondoc = ""
               ADO2!abonos = 0
               ADO2!sdonew = 0
            Else
               ADO2!abonos = ADO2!importe
               ADO2!sdonew = ADO2!sdoold - ADO2!abonos
            End If
            
            If txtMoneda.Text = "S" Then
               ADO2!dolare = 0
               ADO2!soless = ADO2!importe
            Else
               ADO2!soless = 0
               ADO2!dolare = ADO2!importe
            End If
            
            If ADO2!conpago = "161" Or ADO2!conpago = "162" Or _
               ADO2!conpago = "163" Or ADO2!conpago = "164" Or _
               ADO2!conpago = "165" Then
               DataGrid2.col = 7
            Else
               ADO2!pariente = ""
               ADO2!linparie = ""
               ADO2!nombre = ""
               DataGrid2.col = 1
            End If
       Case 7  ' Pariente
            wvariable = Trim(DataGrid2.Text)
            If wvariable <> "E" And wvariable <> "H" And _
               wvariable <> "P" And wvariable <> "M" Then
               DataGrid2.Text = ""
               MsgBox "Pariente Digitado Esta Errado", vbInformation
               Exit Sub
            End If
            DataGrid2.Text = wvariable
            ADO2!pariente = wvariable
            ADO2!linparie = "01"
            
            ADO2!nombre = BuscaPariente(Val(txtCodSocio.Text), ADO2!pariente, ADO2!linparie)
                        
            DataGrid2.col = 6
       
       Case 8  ' Lin Pariente
            wvariable = Trim(DataGrid2.Text)
            If wvariable <> "01" And wvariable <> "02" And _
               wvariable <> "03" And wvariable <> "04" And _
               wvariable <> "05" And wvariable <> "06" And _
               wvariable <> "07" And wvariable <> "08" And _
               wvariable <> "09" And wvariable <> "10" Then
               DataGrid2.Text = ""
               MsgBox "Pariente Digitado Esta Errado", vbInformation
               Exit Sub
            End If
            DataGrid2.Text = wvariable
            ADO2!linparie = wvariable
       
            ADO2!nombre = BuscaPariente(Val(txtCodSocio.Text), ADO2!pariente, ADO2!linparie)
       
            DataGrid2.col = 7
       Case 9  ' Nombre
            wvariable = Trim(DataGrid2.Text)
            DataGrid2.Text = wvariable
            ADO2!nombre = wvariable
            DataGrid2.col = 1
       End Select
       wvariable2 = IIf(IsNull(ADO2.Fields(DataGrid2.col)), "", Trim(ADO2.Fields(DataGrid2.col)))
       DataGrid2.Text = wvariable2
       ADO2.Update
       TotalDet
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
            DataGrid2.col = 9
         End If
         
         Select Case DataGrid2.col
         Case 0
              DataGrid2.Text = ADO2!lincob
         Case 1
              DataGrid2.Text = IIf(IsNull(ADO2!conpago), "", ADO2!conpago)
         Case 2
              DataGrid2.Text = IIf(IsNull(ADO2!nomcon), "", ADO2!nomcon)
         Case 3
              DataGrid2.Text = IIf(IsNull(ADO2!numfra), "", ADO2!numfra)
         Case 4
              DataGrid2.Text = IIf(IsNull(ADO2!linfra), "", ADO2!linfra)
         Case 5
              DataGrid2.Text = IIf(IsNull(ADO2!mescob), "", ADO2!mescob)
         Case 6
              DataGrid2.Text = IIf(IsNull(ADO2!importe), 0, ADO2!importe)
         Case 7
              DataGrid2.Text = IIf(IsNull(ADO2!pariente), "", ADO2!pariente)
         Case 8
              DataGrid2.Text = IIf(IsNull(ADO2!linparie), "", ADO2!linparie)
         Case 9
              DataGrid2.Text = IIf(IsNull(ADO2!nombre), "", ADO2!nombre)
         End Select
   Case 38  ' UP
   
   Case 39  ' AVANZAR

        If DataGrid2.col = 0 Then
           DataGrid2.col = 1
        End If
          
         Select Case DataGrid2.col
         Case 0
              DataGrid2.Text = ADO2!lincob
         Case 1
              DataGrid2.Text = IIf(IsNull(ADO2!conpago), "", ADO2!conpago)
         Case 2
              DataGrid2.Text = IIf(IsNull(ADO2!nomcon), "", ADO2!nomcon)
         Case 3
              DataGrid2.Text = IIf(IsNull(ADO2!numfra), "", ADO2!numfra)
         Case 4
              DataGrid2.Text = IIf(IsNull(ADO2!linfra), "", ADO2!linfra)
         Case 5
              DataGrid2.Text = IIf(IsNull(ADO2!mescob), "", ADO2!mescob)
         Case 6
              DataGrid2.Text = IIf(IsNull(ADO2!importe), 0, ADO2!importe)
         Case 7
              DataGrid2.Text = IIf(IsNull(ADO2!pariente), "", ADO2!pariente)
         Case 8
              DataGrid2.Text = IIf(IsNull(ADO2!linparie), "", ADO2!linparie)
         Case 9
              DataGrid2.Text = IIf(IsNull(ADO2!nombre), "", ADO2!nombre)
         End Select
        
   Case 40  ' DOWN
   
   End Select
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub Form_Activate()
   txtAno.Text = wanocia
   txtAno.Enabled = False
   
   ACCION = 0
   fraMantenimiento.Enabled = False
     
   Dim a As Integer
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
   cmbCia.Text = wcodcia + " " + Trim(wnomcia)
   cmbCia.Enabled = False
   wcia = Left(cmbCia.Text, 2)
   
   cmbTipo.Clear
   cmbTipoAnula.Clear
   a = Leerado8("SELECT * FROM MAETIPORECIBO ORDER BY TIPORECIBO")
   If a > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
   
         cmbTipo.AddItem ADO8!nombre
         cmbTipoAnula.AddItem ADO8!nombre
   
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   
   editar (False)
   
   cmbMeses.Text = "00 POR FECHA"
   txtFecCab.Text = Format(Date, "dd/mm/yyyy")
   
   fraMantenimiento.Enabled = True
      
   editar (False)
   Limpiar
   LlenaCab
   LlenaCab1
   refrescar
      
   fraAnula.Visible = False
         
   DataGrid1.SetFocus
End Sub

Private Sub Form_Load()
   frmCajaCobros.Left = (Screen.Width - Width) \ 2
   frmCajaCobros.Top = 0
   
   Set DataGrid1.DataSource = Nothing
   Set DataGrid2.DataSource = Nothing
'   Limpiar
End Sub

Private Sub LlenaCab()
   Dim c As Integer, waaa As String, wmmm As String, wFec As Date
   waaa = txtAno.Text
   wmmm = Left(cmbMeses.Text, 2)
   If IsDate(txtFecCab.Text) Then
      wFec = Format(txtFecCab.Text, "dd/mm/yyyy")
   Else
      wFec = Format("01/01/1900", "dd/mm/yyyy")
   End If
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_COBROCAB WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
    
   If Format(wFec, "dd/mm/yyyy") <> Format("01/01/1900", "dd/mm/yyyy") Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO TMP_COBROCAB " _
      & " (ANO, MES, TIPCOB, SERCOB, NUMCOB, FECHA , CODSOCIO, TIPCAM, " _
      & "  MONEDA, IMPORTE , DOLARE, SOLESS, GLOSA   , NOMBRE, FORPAG, USU ) " _
      & " SELECT " _
      & "  ANO, MES, TIPCOB, SERCOB, NUMCOB, FECHA , CODSOCIO, TIPCAM, " _
      & "  MONEDA, IMPORTE , DOLARE, SOLESS, GLOSA   , ''    , FORPAG, '" + wcodusu + "' " _
      & " FROM COBROCAB " _
      & " WHERE FECHA = '" + Format(wFec, "dd/mm/yyyy") + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("INSERT INTO TMP_COBROCAB " _
      & " (ANO, MES, TIPCOB, SERCOB, NUMCOB, FECHA , CODSOCIO, TIPCAM, " _
      & "  MONEDA, IMPORTE , DOLARE, SOLESS, GLOSA   , NOMBRE, FORPAG, USU ) " _
      & " SELECT " _
      & "  ANO, MES, TIPCOB, SERCOB, NUMCOB, FECHA , CODSOCIO, TIPCAM, " _
      & "  MONEDA, IMPORTE , DOLARE, SOLESS, GLOSA   , ''    , FORPAG, '" + wcodusu + "' " _
      & " FROM COBROCAB " _
      & " WHERE MES = '" + wmmm + "' AND " _
      & "       ANO = '" + waaa + "' ")
      Db.CommitTrans
   End If
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_COBROCAB " _
   & " SET NOMBRE = M.NOMBRE " _
   & " FROM TMP_COBROCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODSOCIO = M.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_COBROCAB " _
   & " SET NOMBRE = 'DOCUMENTO ANULADO' " _
   & " WHERE CODSOCIO = 0 ")
   Db.CommitTrans
   
   c = Leerado("SELECT TIPCOB, SERCOB, NUMCOB, FECHA, CODSOCIO, NOMBRE, TIPCAM, MONEDA, " _
              & "       IMPORTE, ANO, MES, DOLARE, SOLESS, GLOSA, FORPAG, USU " _
              & " FROM TMP_COBROCAB " _
              & " WHERE USU = '" + wcodusu + "' " _
              & " ORDER BY SERCOB, NUMCOB DESC ")
   Set DataGrid1.DataSource = ADO1
   
   DataGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
End Sub
    
Private Sub LlenaCab1()
    DataGrid1.Columns(0).Width = 450   ' TIPCOB
    DataGrid1.Columns(0).Alignment = dbgCenter
    DataGrid1.Columns(0).Caption = "TIP"
    
    DataGrid1.Columns(1).Width = 450   ' SERCOB
    DataGrid1.Columns(1).Alignment = dbgCenter
    DataGrid1.Columns(1).Caption = "SERIE"
    
    DataGrid1.Columns(2).Width = 950   ' NUMCOB
    DataGrid1.Columns(2).Alignment = dbgCenter
    DataGrid1.Columns(2).Caption = "RECIBO"
    
    DataGrid1.Columns(3).Width = 1050   ' Fecha
    DataGrid1.Columns(3).Alignment = dbgCenter
    DataGrid1.Columns(3).NumberFormat = "dd/mm/yyyy"
    DataGrid1.Columns(3).Caption = "FECHA"
    
    DataGrid1.Columns(4).Width = 650   ' CODSOCIO
    DataGrid1.Columns(4).Alignment = dbgRight
    DataGrid1.Columns(4).Caption = "SOCIO"
       
    DataGrid1.Columns(5).Width = 6050  ' NOMBRE
    DataGrid1.Columns(5).Alignment = dbgLeft
    DataGrid1.Columns(5).Caption = "NOMBRE"
    
    DataGrid1.Columns(6).Width = 900    ' TIPCAM
    DataGrid1.Columns(6).Alignment = dbgCenter
    DataGrid1.Columns(6).Caption = "TIP.CAM"
    DataGrid1.Columns(6).NumberFormat = "###0.000"
    
    DataGrid1.Columns(7).Width = 400   ' MONEDA
    DataGrid1.Columns(7).Alignment = dbgLeft
    DataGrid1.Columns(7).Caption = "MON"
    
    DataGrid1.Columns(8).Width = 1000  ' IMPORTE
    DataGrid1.Columns(8).Alignment = dbgRight
    DataGrid1.Columns(8).Caption = "IMPORTE"
    DataGrid1.Columns(8).NumberFormat = "###,##0.00;;\ "
    
    DataGrid1.Columns(9).Visible = False
    DataGrid1.Columns(10).Visible = False
    DataGrid1.Columns(11).Visible = False
    DataGrid1.Columns(12).Visible = False
    DataGrid1.Columns(13).Visible = False
    DataGrid1.Columns(14).Visible = False
    DataGrid1.Columns(15).Visible = False
    DataGrid1.Refresh
End Sub

Private Sub llenadet()
   Dim wMes As String, wCob As String, waaa As String, wTip As String, wSer As String
   If Left(cmbMeses.Text, 2) <> "00" Then
      wMes = Left(cmbMeses.Text, 2)
   Else
      wMes = Format(Month(txtFecCab.Text), "00")
   End If
   waaa = txtAno.Text
    
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_COBRODET WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
    
    If ACCION = 1 Then
    
    Else
       
       If Not ADO1.BOF And Not ADO1.EOF Then
          wTip = ADO1!tipcob
          wSer = ADO1!sercob
          wCob = ADO1!numcob
       Else
          wTip = ""
          wSer = ""
          wCob = ""
       End If
       
       If wCob <> "" Then
          Db.BeginTrans
          Db.Execute ("INSERT INTO TMP_COBRODET " _
                & " ( ANO   , MES    , TIPCOB  , SERCOB  , NUMCOB, LINCOB, MESCOB, CONCEPTO, " _
                & "   NOMCON, DOLARE , SOLESS  , MONDOC  , SDOOLD, CARGOS, ABONOS  , " _
                & "   SDONEW, CONPAGO, PARIENTE, LINPARIE, NOMBRE, NUMFRA, LINFRA, NUMOPE, USU ) " _
                & " SELECT " _
                & "   ANO   , MES    , TIPCOB  , SERCOB  , NUMCOB, LINCOB, MESCOB, CONCEPTO, " _
                & "   ''    , DOLARE , SOLESS  , MONDOC  , SDOOLD, CARGOS, ABONOS  , " _
                & "   SDONEW, CONPAGO, PARIENTE, LINPARIE, NOMBRE, NUMFRA, LINFRA, NUMOPE, '" + wcodusu + "' " _
                & " From COBRODET " _
                & " WHERE (   ANO='" + waaa + "') AND " _
                & "       (   MES='" + wMes + "') AND " _
                & "       (NUMCOB='" + wCob + "') AND " _
                & "       (SERCOB='" + wSer + "') AND " _
                & "       (TIPCOB='" + wTip + "') ")
          Db.CommitTrans
       
          If ADO1!moneda = "S" Then
             Db.BeginTrans
             Db.Execute ("UPDATE TMP_COBRODET " _
             & " SET IMPORTE = SOLESS " _
             & " WHERE USU = '" + wcodusu + "' ")
             Db.CommitTrans
          Else
             Db.BeginTrans
             Db.Execute ("UPDATE TMP_COBRODET " _
             & " SET IMPORTE = DOLARE " _
             & " WHERE USU = '" + wcodusu + "' ")
             Db.CommitTrans
          End If
       
          Db.BeginTrans
          Db.Execute ("UPDATE TMP_COBRODET " _
          & " SET NOMCON = Z.DESCONCE " _
          & " FROM TMP_COBRODET AS T INNER JOIN ZZZ_CONCEPTO AS Z " _
          & "   ON T.CONPAGO = Z.CCONCE " _
          & " WHERE T.USU = '" + wcodusu + "' ")
          Db.CommitTrans
       
          txtSerCob.Text = ADO1!sercob
          txtNumCob.Text = ADO1!numcob
          If IsDate(ADO1!fecha) Then
             txtFecha.Text = Format(ADO1!fecha, "dd/mm/yyyy")
          End If
          txtTipCam.Text = IIf(IsNull(ADO1!tipcam), 0, ADO1!tipcam)
          txtCodSocio.Text = IIf(IsNull(ADO1!codsocio), "", ADO1!codsocio)
          txtMoneda.Text = IIf(IsNull(ADO1!moneda), "", ADO1!moneda)
          txtImporte.Text = Format(ADO1!importe, "#######0.00;;\ ")
          txtGlosa.Text = IIf(IsNull(ADO1!glosa), "", ADO1!glosa)
          txtForPag.Text = IIf(IsNull(ADO1!forpag), "", ADO1!forpag)
          lblTotDol.Caption = Format(ADO1!dolare, "#######0.00;;\ ")
          lblTotSol.Caption = Format(ADO1!soless, "#######0.00;;\ ")
          If ACCION = 2 Then
             txtNumCob.Enabled = False
          End If
       
       End If
       
    End If
    Dim c As Integer
    c = Leerado2("SELECT LINCOB, CONPAGO, NOMCON , NUMFRA, LINFRA, MESCOB, IMPORTE, PARIENTE, LINPARIE, NOMBRE, " _
                 & "     MONDOC, SDOOLD, ABONOS  , SDONEW, " _
                 & "     ANO   , MES   , TIPCOB  , SERCOB  , NUMCOB, " _
                 & "     DOLARE, SOLESS, CONCEPTO, NUMOPE, USU " _
                 & " FROM TMP_COBRODET " _
                 & " WHERE USU = '" + wcodusu + "' " _
                 & " ORDER BY LINCOB ")
    Set DataGrid2.DataSource = ADO2
End Sub

Private Sub llenadet1()
   DataGrid2.Columns(0).Width = 350   ' Linea'
   DataGrid2.Columns(0).Alignment = dbgCenter
   DataGrid2.Columns(0).Caption = "LIN"
       
   DataGrid2.Columns(1).Width = 530  ' CONCEPTO
   DataGrid2.Columns(1).Alignment = dbgLeft
   DataGrid2.Columns(1).Caption = "CONCEPTO"
       
   DataGrid2.Columns(2).Width = 2700     ' NOMCON
   DataGrid2.Columns(2).Alignment = dbgLeft
   DataGrid2.Columns(2).Caption = "DESCRIPCION CONCEPTO"
       
   DataGrid2.Columns(3).Width = 1100  ' NUMFRA '
   DataGrid2.Columns(3).Alignment = dbgLeft
   DataGrid2.Columns(3).Caption = "NUM.FRAC"
       
   DataGrid2.Columns(4).Width = 400   ' LINFRA '
   DataGrid2.Columns(4).Alignment = dbgLeft
   DataGrid2.Columns(4).Caption = "LIN"
       
   DataGrid2.Columns(5).Width = 770   ' MESCOB '
   DataGrid2.Columns(5).Alignment = dbgLeft
   DataGrid2.Columns(5).Caption = "MESCOB"
       
   DataGrid2.Columns(6).Width = 900    ' IMPORTE
   DataGrid2.Columns(6).Alignment = dbgRight
   DataGrid2.Columns(6).Caption = "IMPORTE"
   DataGrid2.Columns(6).NumberFormat = "####0.00;;\ "
      
   DataGrid2.Columns(7).Width = 530  ' PARIENTE
   DataGrid2.Columns(7).Alignment = dbgCenter
   DataGrid2.Columns(7).Caption = "PARIENTE"
       
   DataGrid2.Columns(8).Width = 530  ' LINPARIE
   DataGrid2.Columns(8).Alignment = dbgCenter
   DataGrid2.Columns(8).Caption = "LIN.PA"
       
   DataGrid2.Columns(9).Width = 4800 ' NOMBRE
   DataGrid2.Columns(9).Alignment = dbgLeft
   DataGrid2.Columns(9).Caption = "NOMBRE"
   
   DataGrid2.Columns(2).Locked = True
   DataGrid2.Columns(3).Locked = True
   DataGrid2.Columns(4).Locked = True
'   DataGrid2.Columns(5).Locked = True
'   DataGrid2.Columns(6).Locked = True
'   DataGrid2.Columns(7).Locked = True
'   DataGrid2.Columns(8).Locked = True
       
   DataGrid2.Columns(10).Visible = False
   DataGrid2.Columns(11).Visible = False
   DataGrid2.Columns(12).Visible = False
   DataGrid2.Columns(13).Visible = False
   DataGrid2.Columns(14).Visible = False
   DataGrid2.Columns(15).Visible = False
   DataGrid2.Columns(16).Visible = False
   DataGrid2.Columns(17).Visible = False
   DataGrid2.Columns(18).Visible = False
   DataGrid2.Columns(19).Visible = False
   DataGrid2.Columns(20).Visible = False
   DataGrid2.Columns(21).Visible = False
   DataGrid2.Columns(22).Visible = False
   DataGrid2.Columns(23).Visible = False
   DataGrid2.col = 1
          
   DataGrid2.Refresh
End Sub

Private Sub TotalDet()
   Dim numreg As Long
   Dim wdol As Currency, wsol As Currency, wImp As Currency
   wdol = 0: wsol = 0: wImp = 0
   numreg = Leerado8(" SELECT Sum(DOLARE) AS DOLARE, Sum(SOLESS) AS SOLESS " _
            & " FROM TMP_COBRODET " _
            & " WHERE USU = '" + wcodusu + "' ")
   If numreg > 0 Then
      wdol = IIf(IsNull(ADO8!dolare), 0, ADO8!dolare)
      wsol = IIf(IsNull(ADO8!soless), 0, ADO8!soless)
   End If
   lblTotDol.Caption = Format(wdol, "#####0.00;;\ ")
   lblTotSol.Caption = Format(wsol, "#####0.00;;\ ")
   
   If txtMoneda.Text = "S" Then
      wImp = wsol
   Else
      wImp = wdol
   End If
   txtImporte.Text = Format(wImp, "#####0.00;;\ ")
   
   Set ADO8 = Nothing
End Sub

Private Sub CreaDet()
   On Error GoTo err

   Dim wLin As Integer, wMes As String
   If Left(cmbMeses.Text, 2) <> "00" Then
      wMes = Left(cmbMeses.Text, 2)
   Else
      wMes = Format(Month(txtFecCab.Text), "00")
   End If
   
   If ADO2.BOF And ADO2.EOF Then
      wLin = 1
   Else
      wLin = Val(ADO2!lincob) + 1
   End If
   ADO2.AddNew
   ADO2!usu = wcodusu
   ADO2!ano = wanocia
   ADO2!mes = wMes
   ADO2!tipcob = Format(cmbTipo.ListIndex + 1, "0")
   ADO2!sercob = Trim(txtSerCob.Text)
   ADO2!numcob = Trim(txtNumCob.Text)
   ADO2!lincob = Format(wLin, "00")
   If lblFrac.Caption = "" Or lblFrac.Caption = "Fraccionamiento Cuotas Al Dia" Then
      ADO2!conpago = ""
      ADO2!nomcon = ""
   Else
      ADO2!conpago = "128"
      ADO2!nomcon = "PAGO X FRACCIONAM."
   End If
   ADO2!importe = 0
   ADO2!dolare = 0
   ADO2!soless = 0
   ADO2!mescob = ""
   ADO2!mondoc = ""
   ADO2!sdoold = 0
   ADO2!abonos = 0
   ADO2!sdonew = 0
   ADO2!concepto = ""
   ADO2!pariente = ""
   ADO2!nombre = ""
   ADO2!numfra = ""
   ADO2!linfra = ""
   ADO2!numope = ""
   ADO2.Update
   
   DataGrid2.col = 1
'  DataGrid2.Text = "   "
   DataGrid2.SelStart = 0
   DataGrid2.SelLength = 3
   
   
   Exit Sub
err:
   MsgBox err.Description, vbExclamation
   Resume Next
End Sub

Private Sub insertlinea(zLin As Integer)
   On Error GoTo err
   
   Dim zrrr As Variant, wMes As String
   If Left(cmbMeses.Text, 2) <> "00" Then
      wMes = Left(cmbMeses.Text, 2)
   Else
      wMes = Format(Month(txtFecCab.Text), "00")
   End If
   If Val(zLin) > 1 Then
      ADO2.MovePrevious
   End If
   zrrr = ADO2.Bookmark
    
   ADO2.MoveLast
   Do While ADO2!lincob >= zLin
      ADO2!lincob = Val(ADO2!lincob) + 1
      ADO2.MovePrevious
      If ADO2.BOF Then
         Exit Do
      End If
   Loop
    
   ADO2.AddNew
   ADO2!usu = wcodusu
   ADO2!ano = wanocia
   ADO2!mes = wMes
   ADO2!tipcob = Format(cmbTipo.ListIndex + 1, "0")
   ADO2!sercob = txtSerCob.Text
   ADO2!numcob = txtNumCob.Text
   ADO2!lincob = Format(zLin, "00")
   ADO2!conpago = ""
   ADO2!nomcon = ""
   ADO2!importe = 0
   ADO2!dolare = 0
   ADO2!soless = 0
   ADO2!mescob = ""
   ADO2!mondoc = ""
   ADO2!sdoold = 0
   ADO2!abonos = 0
   ADO2!sdonew = 0
   ADO2!concepto = ""
   ADO2!pariente = ""
   ADO2!nombre = ""
   ADO2!numfra = ""
   ADO2!linfra = ""
   ADO2!numope = ""
   ADO2.Update
    
   ADO2.Requery
   TotalDet
   llenadet1

   ADO2.Bookmark = zrrr + 1
   DataGrid2.col = 1
   DataGrid2.SelStart = 0
   DataGrid2.SelLength = 3
   Exit Sub
err:
   MsgBox err.Description, vbExclamation
   Resume Next
End Sub

Private Sub renumlinea()
   Dim zLin As Integer, zrrr As Variant
   ADO2.MoveFirst
   zLin = 1
   Do While Not ADO2.EOF
      ADO2!lincob = Format(zLin, "@@")
      zLin = zLin + 1
      ADO2.MoveNext
   Loop
   DataGrid2.col = 1
   ADO2.MoveFirst
End Sub

Private Sub txtCodigo_Change()
   Dim aa As Integer
   If Len(Trim(txtCodigo.Text)) > 0 Then
   aa = Leerado5a("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(Val(txtCodigo.Text)) + " ")
   If aa > 0 Then
      lblCodSocio.Caption = ADO5a!nombre
      txtCodSocio.Text = ADO5a!codsocio
      txtIns.Text = ADO5a!ins
      txtNumdoc.Text = ADO5a!numdoc
      txtE_socio.Text = ADO5a!e_socio
      txtGrado.Text = ADO5a!grado
      txtTipCob.Text = ADO5a!tipcob
   
   Else
      lblCodSocio.Caption = ""
      txtCodSocio.Text = ""
      txtIns.Text = ""
      txtNumdoc.Text = ""
      txtE_socio.Text = ""
      txtGrado.Text = ""
      txtTipCob.Text = ""
   End If
   Set ADO8a = Nothing
   Else
      lblCodSocio.Caption = ""
      txtCodSocio.Text = ""
      txtIns.Text = ""
      txtNumdoc.Text = ""
      txtE_socio.Text = ""
      txtGrado.Text = ""
      txtTipCob.Text = ""
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
        txtGlosa.SetFocus
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
   Dim aa As Integer, wSoc As Integer
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
      wSoc = Val(txtCodSocio.Text)
      txtCodSocio.Text = ADO8!codsocio
      txtIns.Text = ADO8!ins
      txtNumdoc.Text = ADO8!numdoc
      txtE_socio.Text = ADO8!e_socio
      txtGrado.Text = ADO8!grado
      txtTipCob.Text = ADO8!tipcob
      
      If txtE_socio.Text = "FAL" Or txtE_socio.Text = "REN" Or _
         txtE_socio.Text = "EXP" Or txtE_socio.Text = "EXC" Or _
         txtE_socio.Text = "SEP" Then
         MsgBox "Asociado No Esta Activo" + vbNewLine + vbNewLine + _
                "Su Estado Actual Es " + Trim(lblE_socio.Caption) + vbNewLine + vbNewLine + _
                "No Se Puede Cobrar", vbExclamation
         txtCodigo.Text = ""
         Exit Sub
      End If
      
      aa = Leerado8("SELECT * FROM ZZZ_MRECIBOS " _
                    & " WHERE CODIGO = " + Str(Val(txtCodigo.Text)) + " AND " _
                    & "       MONTO > 0 " _
                    & " ORDER BY FECHA_PAGO, NRO_COMP ")
      If aa > 0 Then
         ADO8.MoveLast
         txtAnterior.Text = "FECHA " + Format(ADO8!fecha_pago, "dd/mm/yyyy") + " " + IIf(IsNull(ADO8!obs), "", ADO8!obs)
         ADO8.MovePrevious
         If Not ADO8.BOF And Not ADO8.EOF Then
            txtAnterio2.Text = "FECHA " + Format(ADO8!fecha_pago, "dd/mm/yyyy") + " " + IIf(IsNull(ADO8!obs), "", ADO8!obs)
         End If
      End If
      
      txtGlosa.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtCodSocio_Change()
   Dim aa As Integer, wSoc As Integer, wDifRenov As Currency
   If Len(Trim(txtCodSocio.Text)) > 0 Then
      aa = Leerado5a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
      If aa > 0 Then
         wSoc = Val(txtCodSocio.Text)
         txtCodigo.Text = ADO5a!codigo
         txtIns.Text = ADO5a!ins
         txtNumdoc.Text = ADO5a!numdoc
         txtE_socio.Text = ADO5a!e_socio
         txtGrado.Text = ADO5a!grado
         txtTipCob.Text = ADO5a!tipcob
      
         aa = Leerado4a("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + txtE_socio + "' ")
         If aa > 0 Then
            If ADO4a!moneda = "S" Then
               lblAporte.Caption = "APORTE SOCIO S/." + Format(ADO4a!aporte, "###0.00")
            Else
               lblAporte.Caption = "APORTE SOCIO US$" + Format(ADO4a!aporte, "###0.00")
            End If
         End If
         Set ADO4a = Nothing
      End If
      
      wDifRenov = 0
      If txtE_socio = "TRA" Then
         aa = Leerado4a("SELECT SUM(CARGOS - ABONOS) AS SALDOS " _
                    & " FROM CTASXDET " _
                    & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                    & "       CONCEPTO = '02' ")
         If aa > 0 Then
            wDifRenov = IIf(IsNull(ADO4a!saldos), 0, ADO4a!saldos)
         End If
         Set ADO4a = Nothing
      
         lblRenov.Caption = "DEUDA RENOVACION US$" + Format(wDifRenov, "####0.00")
      End If
      
      aa = Leerado4a("select c.NUMERO ,c.CODSOCIO, c.CODIGO, c.moneda, c.ins, d.linea, D.VCMTO, D.CARGOS, D.ABONOS, D.SDONEW " _
                    & " from FRACCAB as c INNER JOIN FRACDET AS D " _
                    & "   ON C.NUMERO = D.NUMERO " _
                    & " where CODSOCIO = " + Str(wSoc) + " AND " _
                    & "       D.SDONEW > 0 AND " _
                    & "       D.VCMTO < '" + Format(txtFecha.Text, "dd/mm/yyyy") + "'")
      If aa > 0 Then
         ADO4a.MoveFirst
         lblFrac.Caption = "Cuota Frac.Vencida " + ADO4a!linea + _
                           " Vcmto " + Format(ADO4a!vcmto, "dd/mm/yyyy") + " " + _
                           IIf(ADO4a!moneda = "S", "S/", "US$") + Format(ADO4a!cargos, "####0.00")
      Else
         aa = Leerado4a("select c.NUMERO ,c.CODSOCIO, c.CODIGO, c.moneda, c.ins, d.linea, D.VCMTO, D.CARGOS, D.ABONOS, D.SDONEW " _
                    & " from FRACCAB as c INNER JOIN FRACDET AS D " _
                    & "   ON C.NUMERO = D.NUMERO " _
                    & " where CODSOCIO = " + Str(wSoc) + " AND " _
                    & "       D.SDONEW > 0 ")
         If aa > 0 Then
            lblFrac.Caption = "Fraccionamiento Cuotas Al Dia"
         End If
      End If
   
      lblSdoAporte.Caption = "SDO APORT " + Format(SaldoFoto(wSoc, zMesTope), "####0.00")
   
   Else
      lblCodSocio.Caption = ""
   End If
End Sub

Private Sub txtE_socio_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + txtE_socio.Text + "' ")
   If aa > 0 Then
      lblE_socio.Caption = ADO6a!nombre
      lblAporte.Caption = "APORTE SOCIO " + IIf(ADO6a!moneda = "S", "S/.", "US$") + Format(ADO6a!aporte, "###0.00")
   Else
      lblE_socio.Caption = ""
      lblAporte.Caption = ""
   End If
   Set ADO6a = Nothing
End Sub

Private Sub txtFecCab_GotFocus()
   txtFecCab.SelStart = 0
   txtFecCab.SelLength = 10
End Sub

Private Sub txtFecCab_KeyPress(KeyAscii As Integer)
   Dim wmmm As String
   
   If KeyAscii = 13 Then
      If txtFecCab.Text = "__/__/____" Then
         MsgBox "Fecha En Blanco", vbExclamation
         txtFecCab.Text = "__/__/____"
         Exit Sub
      End If
      If Not IsDate(txtFecCab.Text) Then
         MsgBox "Fecha Digitada Es Invalida", vbExclamation
         txtFecCab.Text = "__/__/____"
         Exit Sub
      End If
      cmbMeses.Text = "00 POR FECHA"
      
      txtAno.Text = Format(Year(txtFecCab.Text), "0000")
      
      fraMantenimiento.Enabled = True
      
      editar (False)
      Limpiar
      LlenaCab
      LlenaCab1
      refrescar
      
      DataGrid1.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
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

Private Sub txtNumCobAnula_GotFocus()
   txtSerCobAnula.SelStart = 0
   txtSerCobAnula.SelLength = 10
End Sub

Private Sub txtNumCobAnula_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtSerCobAnula.SetFocus
   End Select
End Sub

Private Sub txtNumCobAnula_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txtNumCobAnula.Text)) = 0 Then
         MsgBox "Documento a Anular En Blanco", vbExclamation
         txtNumCobAnula.Text = ""
         Exit Sub
      End If
      txtNumCobAnula.Text = Format(txtNumCobAnula.Text, "0000000000")
   
      cmdAceptar.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtNumCobAnula_LostFocus()
   txtNumCobAnula.Text = Format(txtNumCobAnula.Text, "0000000000")
End Sub

Private Sub txtSerCobAnula_GotFocus()
   txtSerCobAnula.SelStart = 0
   txtSerCobAnula.SelLength = 3
End Sub

Private Sub txtSerCobAnula_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        cmbTipoAnula.SetFocus
   Case 40
        txtNumCobAnula.SetFocus
   End Select
End Sub

Private Sub txtSerCobAnula_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txtSerCobAnula.Text)) = 0 Then
         MsgBox "Serie de Anulación En Blanca", vbExclamation
         txtSerCobAnula.Text = ""
         Exit Sub
      End If
      txtSerCobAnula.Text = Format(txtSerCobAnula.Text, "000")
   
      txtNumCobAnula.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtTipCob_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAETIPCOB WHERE TIPCOB = '" + txtTipCob.Text + "' ")
   If aa > 0 Then
      lblTipCob.Caption = ADO6a!nombre
   Else
      lblTipCob.Caption = ""
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
         If txtNumCob.Enabled = True Then
            txtNumCob.SetFocus
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
      
      If Not IsDate(txtFecCab.Text) Then
         If Format(Month(txtFecha.Text), "00") <> Mid(cmbMeses.Text, 1, 2) Then
            MsgBox "Mes Digitado No Corresponde", vbInformation
            txtFecha.Text = "__/__/____"
            txtFecha.SetFocus
            Exit Sub
         End If
         If Format(Year(txtFecha.Text), "0000") <> wanocia Then
            MsgBox "Año Digitado No Corresponde", vbInformation
            txtFecha.Text = "__/__/____"
            txtFecha.SetFocus
            Exit Sub
         End If
      End If
      
      If Val(txtTipCam.Text) = 0 Then
         Dim a As Integer
         a = LeeradoMaster3("SELECT * FROM MAECAMBIO WHERE FECHA='" + Format(txtFecha.Text, "dd/mm/yyyy") + "'")
         If a > 0 Then
            txtTipCam.Text = Format(ADOMaster3!venta, "###0.000")
         End If
      End If
      txtCodigo.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtForPag_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAEFORPAG WHERE FORPAG = '" + txtForPag.Text + "' ")
   If aa > 0 Then
      lblForPag.Caption = ADO6a!nombre
   Else
      lblForPag.Caption = ""
   End If
   Set ADO8 = Nothing
End Sub

Private Sub txtForPag_GotFocus()
   txtForPag.SelStart = 0
   txtForPag.SelLength = Len(Trim(txtForPag.Text))
End Sub

Private Sub txtForPag_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtGlosa.SetFocus
   Case 116
        xlista = "FP"
        xseleccion = ""
        frmSeleccion.Show 1
        If xseleccion <> "" Then
           txtForPag.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtForPag_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtForPag.Text)) = 0 Then
         MsgBox "Forma de Pago En Blanco", vbExclamation
         txtForPag.Text = ""
         Exit Sub
      End If
      aa = Leerado6a("SELECT * FROM MAEFORPAG WHERE FORPAG = '" + txtForPag.Text + "' ")
      If aa = 0 Then
         MsgBox "Forma de Pago No Existe", vbExclamation
         txtForPag.Text = ""
         Exit Sub
      End If
      
      DataGrid2.SetFocus
   Else
      If InStr(1, "0123456789." + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtGlosa_GotFocus()
   txtGlosa.SelStart = 0
   txtGlosa.SelLength = Len(Trim(txtGlosa.Text))
End Sub

Private Sub txtGlosa_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtMoneda.SetFocus
   Case 40
        txtForPag.SetFocus
   End Select
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtForPag.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
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
        txtGlosa.SetFocus
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
      txtGlosa.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtNumCob_GotFocus()
   txtNumCob.SelStart = 0
   txtNumCob.SelLength = Len(Trim(txtNumCob.Text))
End Sub

Private Sub txtNumCob_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 40
        txtFecha.SetFocus
   End Select
End Sub

Private Sub txtNumCob_KeyPress(KeyAscii As Integer)
   Dim aa As Integer, wSer As String, wNum As String, wTip As String
   If KeyAscii = 13 Then
      If Len(Trim(txtNumCob.Text)) = 0 Then
         MsgBox "Numero Recibo En Blanco", vbExclamation
         txtNumCob.Text = ""
         Exit Sub
      End If
      wTip = Format(cmbTipo.ListIndex + 1, "0")
      wSer = txtSerCob.Text
      txtNumCob.Text = Format(txtNumCob.Text, "0000000000")
      wNum = txtNumCob.Text
   
      aa = Leerado8("SELECT * FROM COBROCAB " _
                    & " WHERE TIPCOB = '" + wTip + "' AND " _
                    & "       SERCOB = '" + wSer + "' AND " _
                    & "       NUMCOB = '" + wNum + "' ")
      If aa > 0 Then
         MsgBox "Numero de Recibo Ya Existe", vbExclamation
         txtNumCob.Text = ""
         Exit Sub
      End If
      txtFecha.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtNumCob_LostFocus()
   txtNumCob.Text = Format(txtNumCob.Text, "0000000000")
End Sub

Private Sub txtTipCam_GotFocus()
   txtTipCam.SelStart = 0
   txtTipCam.SelLength = Len(Trim(txtTipCam.Text))
End Sub

Private Sub txtTipCam_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtCodigo.SetFocus
   Case 40
        txtMoneda.SetFocus
   End Select
End Sub

Private Sub txtTipCam_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtTipCam.Text = "" Then
         MsgBox "Tipo de Cambio Es Obligatorio", vbInformation
         Exit Sub
      End If
      If Not IsNumeric(Trim(txtTipCam.Text)) Then
         MsgBox "Campo Digitado No Es Numerico", vbExclamation
         txtTipCam.Text = ""
         Exit Sub
      End If
      txtTipCam.Text = Format(txtTipCam.Text, "###0.000")
      txtMoneda.SetFocus
   End If
End Sub

'Private Function BuscaUltimoMes(zSoc As Integer, zCon As String, zLin As String) As String
'   On Error GoTo err
   
'   Dim zz As Integer, zCod As Long, zIns As Integer, zE_s As String, _
'       zMon As String, zApo As Currency, zMes As String

'   zCod = 0: zIns = 0: zMes = ""
'   zz = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(zSoc) + " ")
'   If zz = 0 Then
'      MsgBox "Codigo de Socio " + Str(zSoc) + " No Existe", vbExclamation
'      Exit Function
'   End If
'   zCod = ADO8!codigo
'   zIns = ADO8!ins
'   zE_s = ADO8!e_socio
'   Set ADO8 = Nothing

'   zz = Leerado8("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + zE_s + "' ")
'   If zz > 0 Then
'      zMon = ADO8!moneda
'      zApo = ADO8!aporte
'   End If
'   Set ADO8 = Nothing
      
'   zMes = wanocia + "/01"
'   zz = Leerado8("SELECT * FROM CTASXCAB " _
'                & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
'                & "       CONCEPTO = '" + zCon + "' AND " _
'                & "       SDONEW > 0 " _
'                & " ORDER BY MES ")
'   If zz > 0 Then
'      ADO8.MoveFirst
'      zMes = ADO8!MES
'   End If
   
'   Do While True
'      zz = Leerado8("SELECT * FROM TMP_COBRODET " _
'                & " WHERE    USU = '" + wcodusu + "' AND " _
'                & "       MESCOB = '" + zMes + "' AND " _
'                & "       LINCOB <> '" + zLin + "' ")
'      If zz = 0 Then
'         Exit Do
'      End If
'      If Mid(zMes, 6, 2) = "12" Then
'         zMes = Format(Val(Mid(zMes, 1, 4)) + 1, "0000") + "/" + "01"
'      Else
'         zMes = Mid(zMes, 1, 4) + "/" + Format(Val(Mid(zMes, 6, 2)) + 1, "00")
'      End If
'   Loop
   
'   BuscaUltimoMes = zMes
'   Exit Function
'err:
'   MsgBox err.Description
'   Resume Next
'End Function

'Private Function BuscaUltimoApo(zSoc As Integer, zCon As String, zMes As String) As Currency
'   On Error GoTo err

'   Dim zz As Integer, zCod As Long, zIns As Integer, zE_s As String, _
'       zMon As String, zApo As Currency

'   If Len(Trim(zMes)) = 0 Then
'      BuscaUltimoApo = 0
'      Exit Function
'   End If

'   zCod = 0: zIns = 0
'   zz = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(zSoc) + " ")
'   If zz = 0 Then
'      MsgBox "Codigo de Socio " + Str(zSoc) + " No Existe", vbExclamation
'      Exit Function
'   End If
'   zCod = ADO8!codigo
'   zIns = ADO8!ins
'   zE_s = ADO8!e_socio
'   Set ADO8 = Nothing

'   zz = Leerado8("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + zE_s + "' ")
'   If zz > 0 Then
'      zMon = ADO8!moneda
'      zApo = ADO8!aporte
'   End If
'   Set ADO8 = Nothing
      
'   zz = Leerado8("SELECT * FROM CTASXCAB " _
'                & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
'                & "            MES = '" + Trim(zMes) + "' AND " _
'                & "       CONCEPTO = '" + zCon + "' ")
'   If zz = 0 Then
'      zApo = CreaAporteMes(zSoc, zMes, zCon, 2)
'   Else
'      If ADO8!sdonew > 0 Then
'         zApo = ADO8!sdonew
'      Else
'         MsgBox "Mes Digitado NO Tiene Saldos Por Cobrar", vbExclamation
'         zApo = 0
'      End If
'   End If
      
'   BuscaUltimoApo = zApo
'   Exit Function
'err:
'   MsgBox err.Description
'   Resume Next
'End Function

Private Function validaCob()
   On Error GoTo err
   Dim aa As Integer
   Dim wImp As Currency, wTip As String
   Dim wlleane As Boolean, wllecen As Boolean
   Dim wdolcar As Currency, wdolabo As Currency
   Dim wsolcar As Currency, wsolabo As Currency
   Dim autom1 As String, autom2 As String, wtotdet As Currency
   
   wTip = Format(cmbTipo.ListIndex + 1, "0")
   If Len(Trim(txtNumCob.Text)) = 0 Then
      MsgBox "Numero de Cobranza En Blanco", vbExclamation
      txtNumCob.SetFocus
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
   If Len(Trim(txtImporte.Text)) > 0 Then
      If Not IsNumeric(txtImporte.Text) Then
         MsgBox "Importe No Es Numerico", vbExclamation
         txtImporte.Text = ""
         txtImporte.SetFocus
         validaCob = True
         Exit Function
      End If
   End If
   wImp = Round(Val(txtImporte.Text), 2)
   
   wtotdet = 0
   aa = Leerado8("SELECT SUM(SOLESS) AS SOLESS, SUM(DOLARE) AS DOLARE FROM TMP_COBRODET " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "       TIPCOB = '" + wTip + "' AND " _
                & "       NUMCOB = '" + txtNumCob.Text + "' ")
   If aa > 0 Then
      If txtMoneda.Text = "S" Then
         wtotdet = IIf(IsNull(ADO8!soless), 0, ADO8!soless)
      Else
         wtotdet = IIf(IsNull(ADO8!dolare), 0, ADO8!dolare)
      End If
   End If
   
   If wtotdet <> wImp Then
'      MsgBox "Importe Detalle No Coinciden Con Cabecera" + vbNewLine + "Se Modifica Cabecera", vbExclamation
      txtImporte.Text = Format(wtotdet, "######0.00")
'
'      DataGrid2.SetFocus
'      validaCob = True
'      Exit Function
   End If
   
   validaCob = False
   Exit Function
err:
   MsgBox Format(err.Number, "000000000000") + " " + err.Description
   Resume Next
End Function

Private Function BuscaFrac(zSoc As Integer, zLinCob As String, sw As Byte) As String
   On Error GoTo err
   
   Dim zz As Integer, zNumFra As String, zLinFra As String, zAntFra As String, _
                      zNumAnt As String, zLinAnt As String, _
       zTipCob As String, zSerCob As String, zNumCob As String, _
       zSdoFra As Currency, zMesFra As String, zError As Boolean, _
       zCarFra As Currency, zAboFra As Currency
   
   zTipCob = Format(cmbTipo.ListIndex + 1, "0")
   zSerCob = txtSerCob.Text
   zNumCob = txtNumCob.Text
   zNumFra = "": zLinFra = "": zAntFra = "": zSdoFra = 0: zMesFra = ""
   zError = False
   zz = Leerado8("select c.numero, c.FECHA, c.codsocio, " _
            & "          sum(d.CARGOS) as cargos, SUM(d.abonos) as abonos, " _
            & "          SUM(d.cargos - d.abonos) as sdonew " _
            & " from FRACDET as d inner join FRACCAB as c on d.NUMERO = c.NUMERO " _
            & " Where CODSOCIO = " + Str(zSoc) + " " _
            & " group by c.numero, c.CODSOCIO, c.fecha " _
            & " Having sum(d.cargos - d.abonos) > 0 " _
            & " order by c.NUMERO, c.FECHA, c.codsocio")
   If zz > 0 Then
      zNumFra = IIf(IsNull(ADO8!numero), "", ADO8!numero)
   End If
   Set ADO8 = Nothing
  
   If Len(Trim(zNumFra)) = 0 Then
      MsgBox "NO Existen Fraccionamientos Pendientes", vbExclamation
      zError = True
   End If

   If zNumFra <> "" And sw <> 1 Then
      zz = Leerado8("SELECT * FROM TMP_COBRODET " _
               & " WHERE    USU = '" + wcodusu + "' AND " _
               & "       TIPCOB = '" + zTipCob + "' AND " _
               & "       SERCOB = '" + zSerCob + "' AND " _
               & "       NUMCOB = '" + zNumCob + "' AND " _
               & "       LINCOB < '" + zLinCob + "' AND " _
               & "       CONPAGO = '128' AND " _
               & "       NUMFRA = '" + zNumFra + "' " _
               & " ORDER BY LINCOB DESC ")
      If zz > 0 Then
         zLinAnt = IIf(IsNull(ADO8!linfra), "", ADO8!linfra)
      End If
      Set ADO8 = Nothing
   
      zz = Leerado8("select c.numero, c.fecha, d.vcmto, c.codsocio, d.linea, d.cargos, d.sdonew " _
                & " from FRACDET as d inner join FRACCAB as c on d.NUMERO = c.numero " _
                & " Where c.CODSOCIO = " + Str(zSoc) + " and " _
                & "       c.numero = '" + zNumFra + "' And " _
                & "       d.SDONEW > 0 AND " _
                & "       d.linea > '" + zLinAnt + "' " _
                & " order by linea")
      If zz > 0 Then
         zMesFra = Format(ADO8!vcmto, "yyyy/mm")
         zLinFra = Format(ADO8!linea, "@@")
         zCarFra = ADO8!cargos
         zSdoFra = ADO8!sdonew
      End If
   
      
      If Len(Trim(zLinFra)) = 0 Then
         MsgBox "NO Existen Cuotas Con Saldo Frac." + zNumFra, vbExclamation
         zError = True
      End If
   
   End If

   If zError = False Then
      Select Case sw
      Case 1
           BuscaFrac = zNumFra
      Case 2
           BuscaFrac = zLinFra
      Case 3
           BuscaFrac = Format(zSdoFra, "########0.00")
      Case 4
           BuscaFrac = zMesFra
      Case 5
           BuscaFrac = Format(zCarFra, "########0.00")
      End Select
   Else
      Select Case sw
      Case 1
           BuscaFrac = ""
      Case 2
           BuscaFrac = ""
      Case 3
           BuscaFrac = ""
      Case 4
           BuscaFrac = ""
      Case 5
           BuscaFrac = ""
      End Select
   End If

   Exit Function
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Function

Private Function BuscaCargoFrac(zSoc As Integer, zNumFra As String, zLinFra As String) As Currency
   On Error GoTo err
   
   Dim zCargos As Currency, zz As Integer
   
   zCargos = 0
   zz = Leerado8("select c.numero, c.fecha, d.vcmto, c.codsocio, d.linea, d.cargos, d.sdonew " _
             & " from FRACDET as d inner join FRACCAB as c on d.NUMERO = c.numero " _
             & " Where c.CODSOCIO = " + Str(zSoc) + " and " _
             & "       d.numero = '" + zNumFra + "' And " _
             & "       d.linea = '" + zLinFra + "' ")
   If zz > 0 Then
      zCargos = ADO8!cargos
   End If

   BuscaCargoFrac = zCargos

   Exit Function
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Function

