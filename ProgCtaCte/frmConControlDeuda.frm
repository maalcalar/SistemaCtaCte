VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConControlDeuda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Deudas x Asociado"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   13800
   Begin VB.OptionButton optAdelanto 
      Caption         =   "Solo Asociados Con Adelanto"
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
      Left            =   7680
      TabIndex        =   27
      Top             =   720
      Width           =   3015
   End
   Begin VB.OptionButton optSaldo 
      Caption         =   "Solo Asociados Con Saldo"
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
      Left            =   7680
      TabIndex        =   14
      Top             =   480
      Width           =   3015
   End
   Begin VB.OptionButton optTodos 
      Caption         =   "Todos los Activos"
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
      Left            =   7680
      TabIndex        =   13
      Top             =   240
      Value           =   -1  'True
      Width           =   3015
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
      Left            =   12120
      TabIndex        =   11
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdExportar 
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
      Height          =   495
      Left            =   9480
      TabIndex        =   10
      Top             =   8280
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5895
      Left            =   480
      TabIndex        =   9
      Top             =   1320
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   10398
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
   Begin VB.ComboBox cmbE_Socio 
      Height          =   315
      ItemData        =   "frmConControlDeuda.frx":0000
      Left            =   1560
      List            =   "frmConControlDeuda.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
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
      Left            =   10920
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtCodSocio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      MaxLength       =   9
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin MSMask.MaskEdBox txtMes 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin Crystal.CrystalReport Crys1 
      Left            =   12600
      Top             =   7440
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
   Begin VB.Label lblImp1 
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   10200
      TabIndex        =   26
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label lblImp2 
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   10200
      TabIndex        =   25
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label lblImp3 
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   10200
      TabIndex        =   24
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label lblFec1 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9120
      TabIndex        =   23
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label lblFec2 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9120
      TabIndex        =   22
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label lblFec3 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9120
      TabIndex        =   21
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Glosa 1"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label lblGlosa1 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      TabIndex        =   19
      Top             =   7680
      Width           =   7935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Glosa 2"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label lblGlosa2 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Top             =   7440
      Width           =   7935
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      Caption         =   "Glosa3"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label lblGlosa3 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      TabIndex        =   15
      Top             =   7200
      Width           =   7935
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
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
      TabIndex        =   12
      Top             =   8280
      Width           =   7575
   End
   Begin VB.Label lblMes 
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Mes Consulta"
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
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Estado Socio"
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
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   1140
   End
   Begin VB.Label lblCodSocio 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "Cod.Socio"
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
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmConControlDeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbE_Socio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtCodSocio.SetFocus
   End If
End Sub

Private Sub cmdBuscar_Click()
   LlenaCab
   LabelCab

   DataGrid1.SetFocus
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(21) As String, _
       wreg As Integer, wTot As Integer, wMes As String, wFecIng As Date, wFec As Date
       
   
   Heading(0) = "SOCIO"
   Heading(1) = "CODOFIN"
   Heading(2) = "INS"
   Heading(3) = "NOMBRE SOCIO"
   Heading(4) = "ESTADO"
   Heading(5) = "FEC.ING"
   
   Heading(6) = "MON"
   Heading(7) = "APORTE"
   Heading(8) = "SDO.ACTUAL"
   Heading(9) = "ENVIO MES"
   
   Heading(10) = "GLOSA 1"
   Heading(11) = "FECHA"
   Heading(12) = "MONEDA"
   Heading(13) = "IMPORTE"
   
   Heading(14) = "GLOSA 2"
   Heading(15) = "FECHA"
   Heading(16) = "MONEDA"
   Heading(17) = "IMPORTE"
   
   Heading(18) = "GLOSA 3"
   Heading(19) = "FECHA"
   Heading(20) = "MONEDA"
   Heading(21) = "IMPORTE"
   
   
   aa = Leerado3("SELECT * FROM TMP_CTRLDEU " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY NOMBRE ")
   If aa > 0 Then
      wTot = aa
      Set objExcel = New Excel.Application
      objExcel.SheetsInNewWorkbook = 1
      objExcel.Workbooks.Add
      With objExcel
           .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
           .Range(.Cells(3, 1), .Cells(3, 22)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 22)).Font.Bold = True
           .Cells(1, 1) = wnomcia
           .Cells(2, 1) = "CONTROL DE DEUDAS - MES " + wMes
           For I = 1 To 22 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           objExcel.Columns("A").ColumnWidth = 7
           objExcel.Columns("B").ColumnWidth = 11
           objExcel.Columns("C").ColumnWidth = 5
           objExcel.Columns("D").ColumnWidth = 60
           objExcel.Columns("E").ColumnWidth = 5
           objExcel.Columns("F").ColumnWidth = 11
           objExcel.Columns("G").ColumnWidth = 5
           objExcel.Columns("H").ColumnWidth = 9
           objExcel.Columns("I").ColumnWidth = 11
           objExcel.Columns("J").ColumnWidth = 11
           objExcel.Columns("K").ColumnWidth = 50
           objExcel.Columns("L").ColumnWidth = 11
           objExcel.Columns("M").ColumnWidth = 5
           objExcel.Columns("N").ColumnWidth = 10
           objExcel.Columns("O").ColumnWidth = 50
           objExcel.Columns("P").ColumnWidth = 11
           objExcel.Columns("Q").ColumnWidth = 5
           objExcel.Columns("R").ColumnWidth = 10
           objExcel.Columns("S").ColumnWidth = 50
           objExcel.Columns("T").ColumnWidth = 11
           objExcel.Columns("U").ColumnWidth = 5
           objExcel.Columns("V").ColumnWidth = 10
      End With
      V = 4
      H = 1
      wreg = 1
      Do While Not ADO3.EOF
         DoEvents
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + _
                              Format(wreg, "####0") + " / " + _
                              Format(wTot, "####0")
         lblMensaje.Refresh
         
         objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 10)).NumberFormat = "######0.00;;\ "
         objExcel.Range(objExcel.Cells(V, H + 12), objExcel.Cells(V, H + 14)).NumberFormat = "######0.00"
            
         objExcel.Cells(V, H + 0) = ADO3!codsocio
         objExcel.Cells(V, H + 1) = ADO3!codigo
         objExcel.Cells(V, H + 2) = ADO3!ins
         objExcel.Cells(V, H + 3) = ADO3!nombre
         objExcel.Cells(V, H + 4) = ADO3!e_socio
         If IsDate(ADO3!fecing) Then
            wFecIng = Format(ADO3!fecing, "dd/mm/yyyy")
            objExcel.Cells(V, H + 5) = wFecIng
         End If
         objExcel.Cells(V, H + 6) = ADO3!moneda
         objExcel.Cells(V, H + 7) = ADO3!aporte
         objExcel.Cells(V, H + 8) = ADO3!sdonew
         objExcel.Cells(V, H + 9) = ADO3!proces
            
         objExcel.Cells(V, H + 10) = IIf(IsNull(ADO3!glosa1), "", ADO3!glosa1)
         If IsDate(ADO3!fec1) Then
            wFec = Format(ADO3!fec1, "dd/mm/yyyy")
            objExcel.Cells(V, H + 11) = wFec
         End If
         If ADO3!mon1 <> "" Then
            objExcel.Cells(V, H + 12) = IIf(ADO3!mon1 = "S", "S/.", "US$")
         End If
         objExcel.Cells(V, H + 13) = ADO3!imp1
            
         objExcel.Cells(V, H + 14) = IIf(IsNull(ADO3!glosa2), "", ADO3!glosa2)
         If IsDate(ADO3!fec2) Then
            wFec = Format(ADO3!fec2, "dd/mm/yyyy")
            objExcel.Cells(V, H + 15) = wFec
         End If
         If ADO3!mon2 <> "" Then
            objExcel.Cells(V, H + 16) = IIf(ADO3!mon2 = "S", "S/.", "US$")
         End If
         objExcel.Cells(V, H + 17) = ADO3!imp2
            
         objExcel.Cells(V, H + 18) = IIf(IsNull(ADO3!glosa3), "", ADO3!glosa3)
         If IsDate(ADO3!fec3) Then
            wFec = Format(ADO3!fec3, "dd/mm/yyyy")
            objExcel.Cells(V, H + 19) = wFec
         End If
         If ADO3!mon3 <> "" Then
            objExcel.Cells(V, H + 20) = IIf(ADO3!mon3 = "S", "S/.", "US$")
         End If
         objExcel.Cells(V, H + 21) = ADO3!imp3
            
         wreg = wreg + 1
         V = V + 1
         ADO3.MoveNext
         If ADO3.EOF Then
            Exit Do
         End If
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

Private Sub cmdImprimir_Click()
   Dim wMes As String
   wMes = Right(txtMes.Text, 2) + " DEL " + Left(txtMes.Text, 4)
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\ControlDeuda.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= 'AL MES DE " + wMes + "' "
   Crys1.SelectionFormula = " {TMP_CTRLDEU.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   LabelCab
End Sub

Private Sub Form_Activate()
   Dim a As Integer, wAno As String, wMes As String
   
   cmbE_Socio.Clear
   cmbE_Socio.AddItem "Todos Los Estados de Socio"
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
   cmbE_Socio.ListIndex = 0
   
'   wAno = wanocia
'   wMes = Format(Month(Date), "00")
   
   txtMes.Text = Left(zMesTope, 4) + "/" + Right(zMesTope, 2)
   
   txtMes.SetFocus
End Sub

Private Sub Form_Initialize()
   frmConControlDeuda.Left = (Screen.Width - Width) \ 2
   frmConControlDeuda.Top = 0
End Sub

Private Sub LlenaCab()
   Dim aa As Long, _
       wcon As String, wE_S As String, zMes As String, _
       zSoc As Integer, sw As String, zSdo As Currency, zPro As Currency, _
       zGlo1 As String, zGlo2 As String, zGlo3 As String, _
       zMon1 As String, zMon2 As String, zMon3 As String, _
       zFec1 As Date, zFec2 As Date, zFec3 As Date, _
       zImp1 As Currency, zImp2 As Currency, zImp3 As Currency, _
       zCod As Long, zIns As Integer
   
   Set DataGrid1.DataSource = Nothing
   
   zMes = txtMes.Text
   wE_S = BuscaCodEsocio(cmbE_Socio.List(cmbE_Socio.ListIndex))
   wSoc = Val(txtCodSocio.Text)
   
   lblGlosa1.Caption = ""
   lblGlosa2.Caption = ""
   lblGlosa3.Caption = ""
   
   lblFec1.Caption = ""
   lblFec2.Caption = ""
   lblFec3.Caption = ""
   
   lblImp1.Caption = ""
   lblImp2.Caption = ""
   lblImp3.Caption = ""
   
   Set DataGrid1.DataSource = Nothing
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_CTRLDEU WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   lblMensaje.Caption = "Preparando Archivo...."
   lblMensaje.Refresh
   
   sw = "WHERE E.APORTE > 0 "
   
   If Len(Trim(wE_S)) > 0 Then
      If Len(Trim(sw)) = 0 Then
         sw = sw + "WHERE "
      Else
         sw = sw + " AND "
      End If
      sw = sw + "M.E_SOCIO = '" + wE_S + "'"
   End If
   
   If wSoc > 0 Then
      If Len(Trim(sw)) = 0 Then
         sw = "WHERE "
      Else
         sw = sw + " AND "
      End If
      sw = sw + " M.CODSOCIO = " + Str(wSoc) + " "
   End If
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_CTRLDEU " _
   & " (CODSOCIO, CODIGO, INS, E_SOCIO, NOMBRE, FECING, MONEDA, APORTE, " _
   & "  TIPCOB, NOMCOB, SDONEW, PROCES, " _
   & "  GLOSA1, FEC1, MON1, IMP1, " _
   & "  GLOSA2, FEC2, MON2, IMP2, " _
   & "  GLOSA3, FEC3, MON3, IMP3, USU) " _
   & " SELECT M.CODSOCIO, M.CODIGO, M.INS, M.E_SOCIO, M.NOMBRE, M.FECING, " _
   & "        E.MONEDA, E.APORTE, M.TIPCOB, T.NOMBRE, 0, 0, " _
   & "        '', NULL, '', 0, " _
   & "        '', NULL, '', 0, " _
   & "        '', NULL, '', 0, '" + wcodusu + "' " _
   & " FROM MAESOCIO AS M INNER JOIN MAETIPCOB  AS T ON M.TIPCOB = T.TIPCOB " _
   & "                    INNER JOIN MAEE_SOCIO AS E ON M.E_SOCIO = E.E_SOCIO " _
   & " " + sw + " ")
   Db.CommitTrans

   Dim wreg As Long, wTot As Long

   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, " _
            & "          FECING, MONEDA, APORTE, SDONEW, PROCES, USU, " _
            & "          TIPCOB, NOMCOB, " _
            & "          GLOSA1, FEC1, MON1, IMP1, " _
            & "          GLOSA2, FEC2, MON2, IMP2, " _
            & "          GLOSA3, FEC3, MON3, IMP3 " _
            & " FROM TMP_CTRLDEU " _
            & " WHERE USU = '" + wcodusu + "' " _
            & " ORDER BY NOMBRE ")
   If aa > 0 Then
      ADO2.MoveFirst
      wreg = wreg + 1
      wTot = aa
      Do While Not ADO2.EOF
         DoEvents
         lblMensaje.Caption = "Registro " + _
                              Trim(Format(wreg, "####0")) + " / " + _
                              Trim(Format(wTot, "####0"))
         lblMensaje.Refresh
         
         zSoc = ADO2!codsocio
         zCod = ADO2!codigo
         zIns = ADO2!ins
         zSdo = 0: zPro = 0:
         zGlo1 = "": zGlo2 = "": zGlo3 = ""
         zMon1 = "": zMon2 = "": zMon3 = ""
         zImp1 = 0: zImp2 = 0: zMon3 = 0
         zFec1 = Format("01/01/1900", "dd/mm/yyyy")
         zFec2 = Format("01/01/1900", "dd/mm/yyyy")
         zFec3 = Format("01/01/1900", "dd/mm/yyyy")
   
         zSdo = SaldoFoto(zSoc, zMes)
         zPro = BuscaEnvioDieco(zSoc, Left(zMes, 4) + Right(zMes, 2))
         If zPro = 0 Then
            zPro = BuscaEnvioCajMP(zSoc, Left(zMes, 4) + Right(zMes, 2))
         End If
         
         aa = Leerado8a("SELECT * FROM ZZZ_MRECIBOS " _
                    & " WHERE CODIGO = " + Str(zCod) + " AND " _
                    & "          INS = " + Str(zIns) + " AND " _
                    & "      (MARCA2 <> 'A' OR MARCA2 IS NULL) " _
                    & " ORDER BY FECHA_PAGO DESC  ")
         If aa > 0 Then
            zGlo1 = IIf(IsNull(ADO8a!obs), "", ADO8a!obs)
            zFec1 = Format(ADO8a!fecha_pago, "dd/mm/yyyy")
            zMon1 = IIf(ADO8a!moneda = "S/.", "S", "D")
            zImp1 = ADO8a!monto
            
            ADO8a.MoveNext
            If Not ADO8a.EOF And Not ADO8a.BOF Then
               zGlo2 = IIf(IsNull(ADO8a!obs), "", ADO8a!obs)
               zFec2 = Format(ADO8a!fecha_pago, "dd/mm/yyyy")
               zMon2 = IIf(ADO8a!moneda = "S/.", "S", "D")
               zImp2 = ADO8a!monto
            
               ADO8a.MoveNext
               If Not ADO8a.EOF And Not ADO8a.BOF Then
                  zGlo3 = IIf(IsNull(ADO8a!obs), "", ADO8a!obs)
                  zFec3 = Format(ADO8a!fecha_pago, "dd/mm/yyyy")
                  zMon3 = IIf(ADO8a!moneda = "S/.", "S", "D")
                  zImp3 = ADO8a!monto
               End If
            End If
         
            Db.BeginTrans
            Db.Execute ("UPDATE TMP_CTRLDEU " _
            & " SET SDONEW = " + Str(zSdo) + ", " _
            & "     PROCES = " + Str(zPro) + ", " _
            & "     GLOSA1='" + GlosaLibre(zGlo1) + "', FEC1='" + Format(zFec1, "dd/mm/yyyy") + "', MON1='" + zMon1 + "', IMP1=" + Str(zImp1) + ", " _
            & "     GLOSA2='" + GlosaLibre(zGlo2) + "', FEC2='" + Format(zFec2, "dd/mm/yyyy") + "', MON2='" + zMon2 + "', IMP2=" + Str(zImp2) + ", " _
            & "     GLOSA3='" + GlosaLibre(zGlo3) + "', FEC3='" + Format(zFec3, "dd/mm/yyyy") + "', MON3='" + zMon3 + "', IMP3=" + Str(zImp3) + " " _
            & " WHERE      USU = '" + wcodusu + "' AND " _
            & "       CODSOCIO = " + Str(zSoc) + " ")
            Db.CommitTrans
         
         End If
         Set ADO8a = Nothing
            
         wreg = wreg + 1
         ADO2.MoveNext
      Loop
   End If
   If ADO2.RecordCount > 0 Then
      ADO2.MoveFirst
   End If
   
   If optSaldo.Value = True Then
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_CTRLDEU WHERE USU = '" + wcodusu + "' AND SDONEW <= 0 ")
      Db.CommitTrans
   End If
   
   If optAdelanto.Value = True Then
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_CTRLDEU WHERE USU = '" + wcodusu + "' AND SDONEW >= 0 ")
      Db.CommitTrans
   End If
   
   lblMensaje.Caption = ""
   lblMensaje.Refresh
   
   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, " _
            & "          FECING, MONEDA, APORTE, SDONEW, PROCES, USU, " _
            & "          TIPCOB, NOMCOB, " _
            & "          GLOSA1, FEC1, MON1, IMP1, " _
            & "          GLOSA2, FEC2, MON2, IMP2, " _
            & "          GLOSA3, FEC3, MON3, IMP3 " _
            & " FROM TMP_CTRLDEU " _
            & " WHERE USU = '" + wcodusu + "' " _
            & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO2
   
   DataGrid1.Columns(0).Width = 700   ' CODSOCIO
   DataGrid1.Columns(0).Alignment = dbgRight
   DataGrid1.Columns(0).Caption = "COD.SOCIO"
    
   DataGrid1.Columns(1).Width = 1000  ' CODIGO
   DataGrid1.Columns(1).Alignment = dbgRight
   DataGrid1.Columns(1).Caption = "CODIGO"
    
   DataGrid1.Columns(2).Width = 400   ' INS
   DataGrid1.Columns(2).Alignment = dbgCenter
   DataGrid1.Columns(2).Caption = "INS"
    
   DataGrid1.Columns(3).Width = 4800  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE"
    
   DataGrid1.Columns(4).Width = 750   ' E_SOCIO
   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Caption = "ESTADO"
    
   DataGrid1.Columns(5).Width = 1050  ' FECING
   DataGrid1.Columns(5).Alignment = dbgLeft
   DataGrid1.Columns(5).Caption = "FEC.ING"
   DataGrid1.Columns(5).NumberFormat = "dd/mm/yyyy"
    
   DataGrid1.Columns(6).Width = 350   ' MONEDA
   DataGrid1.Columns(6).Alignment = dbgCenter
   DataGrid1.Columns(6).Caption = "MON"
    
   DataGrid1.Columns(7).Width = 650   ' APORTE
   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).Caption = "APORTE"
   DataGrid1.Columns(7).NumberFormat = "####0.00"
    
   DataGrid1.Columns(8).Width = 850   ' SDONEW
   DataGrid1.Columns(8).Alignment = dbgRight
   DataGrid1.Columns(8).Caption = "SALDO"
   DataGrid1.Columns(8).NumberFormat = "###,##0.00;-###,##0.00;\ "
    
   DataGrid1.Columns(9).Width = 850   ' PROCES
   DataGrid1.Columns(9).Alignment = dbgRight
   DataGrid1.Columns(9).Caption = "TRANSITO"
   DataGrid1.Columns(9).NumberFormat = "###,##0.00"
    
   DataGrid1.Columns(10).Visible = False
   DataGrid1.Columns(11).Visible = False
   DataGrid1.Columns(12).Visible = False
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
End Sub

Private Sub LabelCab()
   If Not ADO2.BOF And Not ADO2.EOF Then
      lblGlosa1.Caption = ADO2!glosa1
      lblGlosa2.Caption = ADO2!glosa2
      lblGlosa3.Caption = ADO2!glosa3
   
      lblFec1.Caption = Format(ADO2!fec1, "dd/mm/yyyy")
      lblFec2.Caption = Format(ADO2!fec2, "dd/mm/yyyy")
      lblFec3.Caption = Format(ADO2!fec3, "dd/mm/yyyy")
   
      lblImp1.Caption = IIf(ADO2!mon1 = "S", "S/.", "US$") + Format(ADO2!imp1, "####0.00")
      lblImp2.Caption = IIf(ADO2!mon2 = "S", "S/.", "US$") + Format(ADO2!imp2, "####0.00")
      lblImp3.Caption = IIf(ADO2!mon3 = "S", "S/.", "US$") + Format(ADO2!imp3, "####0.00")
   End If
End Sub

Private Sub txtCodSocio_Change()
   Dim aa As Integer
   aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
   If aa > 0 Then
      lblCodSocio.Caption = ADO8!nombre
   Else
      lblCodSocio.Caption = ""
   End If
   Set ADO8 = Nothing
End Sub

Private Sub txtCodSocio_GotFocus()
   txtCodSocio.SelStart = 0
   txtCodSocio.SelLength = Len(Trim(txtCodSocio.Text))
End Sub

Private Sub txtCodSocio_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim KeyNew As Integer
   KeyNew = Asc(UCase(Chr(KeyCode)))
   
   Select Case KeyCode
   Case 38
        cmbE_Socio.SetFocus
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtCodSocio.Text = xseleccion
        End If
   Case 65 To 90
        xlista = "1"
        xseleccion = Chr(KeyNew)
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtCodSocio.Text = xseleccion
        End If
          
   End Select
End Sub

Private Sub txtCodSocio_KeyPress(KeyAscii As Integer)
   Dim aa As Integer, wSoc As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtCodSocio.Text)) <> 0 Then
         aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
         If aa = 0 Then
            MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
            txtCodSocio.Text = ""
            Exit Sub
         End If
         wSoc = Val(txtCodSocio.Text)
      End If
      
      cmdBuscar.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
'         KeyAscii = 0
'      End If
   End If
End Sub

Private Sub txtMes_Change()
   Dim waaa As String, wmmm As String

   waaa = Left(txtMes.Text, 4)
   wmmm = Right(txtMes.Text, 2)

   lblMes.Caption = Trim(funnommes(wmmm)) + " " + waaa
End Sub

Private Sub txtMes_GotFocus()
   txtMes.SelStart = 0
   txtMes.SelLength = 7
End Sub

Private Sub txtMes_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        cmbConcepto.SetFocus
   Case 40
        cmbE_Socio.SetFocus
   End Select
End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)
   Dim waaa As String, wmmm As String
   If KeyAscii = 13 Then
      If txtMes.Text = "____-__" Then
         MsgBox "Mes de Corte En Blanco", vbExclamation
         txtMes.Text = "____/__"
         Exit Sub
      End If
      waaa = Left(txtMes.Text, 4)
      wmmm = Right(txtMes.Text, 2)
          
      If wmmm <> "01" And wmmm <> "02" And wmmm <> "03" And wmmm <> "04" And _
         wmmm <> "05" And wmmm <> "06" And wmmm <> "07" And wmmm <> "08" And _
         wmmm <> "09" And wmmm <> "10" And wmmm <> "11" And wmmm <> "12" Then
         MsgBox "Mes Digitado Es Invalido", vbQuestion
         txtMes.Text = "____/__"
         Exit Sub
      End If
      If waaa < "2010" And waaa > "2040" Then
         MsgBox "Año Digitado Es Invalido", vbQuestion
         txtMes.Text = "____/__"
         Exit Sub
      End If
      cmbE_Socio.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub


