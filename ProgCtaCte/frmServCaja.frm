VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmServCaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar Pagos de Socio En Caja"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   12360
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
      Left            =   10800
      TabIndex        =   12
      Top             =   6960
      Width           =   1215
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
      Left            =   9480
      TabIndex        =   11
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox txtCodSocio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      MaxLength       =   9
      TabIndex        =   3
      Top             =   180
      Width           =   975
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   6840
      MaxLength       =   8
      TabIndex        =   2
      Top             =   180
      Width           =   975
   End
   Begin VB.TextBox txtIns 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7800
      MaxLength       =   1
      TabIndex        =   1
      Top             =   180
      Width           =   375
   End
   Begin VB.TextBox txtNumdoc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8160
      MaxLength       =   8
      TabIndex        =   0
      Top             =   180
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6015
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   10610
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
      Caption         =   "TESORERIA"
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
   Begin VB.Label Label3 
      Caption         =   "Cod.Socio"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblCodSocio 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   180
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de Socio"
      Height          =   195
      Left            =   1200
      TabIndex        =   7
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Codofin"
      Height          =   195
      Left            =   6840
      TabIndex        =   6
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Ins"
      Height          =   195
      Left            =   7800
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "D.N.I."
      Height          =   195
      Left            =   8160
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmServCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
