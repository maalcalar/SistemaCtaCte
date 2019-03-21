VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMaeFoto 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   12225
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
      Left            =   10560
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9763
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
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2775
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmMaeFoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   On Error GoTo err

   Dim wreg As Integer, zReg As String
   
   zReg = Trim(Format(ADO2!codsocio, "00000"))
   
   file = Left(App.Path, 2) + "\fotos\" + zReg + "jpg.jpg"
      
   If Len(Dir$(file)) Then
      Image1.Picture = LoadPicture(file)
   Else
      Image1.Picture = LoadPicture(Left(App.Path, 2) + "\fotos\SinFoto.jpg")
   End If
   Image1.Refresh
   

   Exit Sub
err:
   Resume Next
End Sub

Private Sub Form_Load()
   Dim aa As Integer

   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO " _
                & " FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'TIT' " _
                & " ORDER BY NOMBRE ")
   DataGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
   Set DataGrid1.DataSource = ADO2

   DataGrid1.Columns(0).Width = 800
   DataGrid1.Columns(0).Alignment = dbgRight
   DataGrid1.Columns(0).Caption = "SOCIO"
    
   DataGrid1.Columns(1).Width = 900
   DataGrid1.Columns(1).Alignment = dbgRight
   DataGrid1.Columns(1).Caption = "CODIGO"
    
   DataGrid1.Columns(2).Width = 400
   DataGrid1.Columns(2).Alignment = dbgCenter
   DataGrid1.Columns(2).Caption = "INS"
    
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Width = 5200
   DataGrid1.Columns(3).Caption = "NOMBRE"

   DataGrid1.Columns(4).Width = 600
   DataGrid1.Columns(4).Alignment = dbgCenter
   DataGrid1.Columns(4).Caption = "ESTADO"
   
   Call DataGrid1_RowColChange(0, 0)
End Sub
