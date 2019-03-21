VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConSdosxMes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Saldos Por Mes"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   13920
   Begin VB.OptionButton optSaldo 
      Caption         =   "Solo Meses Con Saldo"
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
      Left            =   8400
      TabIndex        =   20
      Top             =   1200
      Value           =   -1  'True
      Width           =   3015
   End
   Begin VB.OptionButton optTodos 
      Caption         =   "Mostrar Todos Los Meses"
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
      Left            =   8400
      TabIndex        =   19
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtCodSocio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   16
      Top             =   1200
      Width           =   975
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
      Left            =   6960
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox cmbE_Socio 
      Height          =   315
      ItemData        =   "frmConSdosxMes.frx":0000
      Left            =   1320
      List            =   "frmConSdosxMes.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   840
      Width           =   3375
   End
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
      Height          =   495
      Left            =   10920
      TabIndex        =   10
      Top             =   7680
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
      Left            =   9600
      TabIndex        =   9
      Top             =   7680
      Width           =   1095
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
      Left            =   12240
      TabIndex        =   8
      Top             =   7680
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5535
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   13695
      _ExtentX        =   24156
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
   Begin VB.ComboBox cmbConcepto 
      Height          =   315
      ItemData        =   "frmConSdosxMes.frx":0004
      Left            =   1320
      List            =   "frmConSdosxMes.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmConSdosxMes.frx":0008
      Left            =   1320
      List            =   "frmConSdosxMes.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   165
      Width           =   7335
   End
   Begin MSMask.MaskEdBox txtMes 
      Height          =   285
      Left            =   6000
      TabIndex        =   5
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin Crystal.CrystalReport Crys1 
      Left            =   9600
      Top             =   0
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
      Left            =   840
      TabIndex        =   22
      Top             =   7800
      Width           =   7575
   End
   Begin VB.Label lblAbonos 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   11400
      TabIndex        =   21
      Top             =   7200
      Width           =   1095
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
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblCodSocio 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   1200
      Width           =   4575
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
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   1140
   End
   Begin VB.Label lblCargos 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   10320
      TabIndex        =   12
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label lblSdoNew 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   12480
      TabIndex        =   11
      Top             =   7200
      Width           =   1095
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
      Left            =   4800
      TabIndex        =   7
      Top             =   480
      Width           =   1215
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
      Left            =   6840
      TabIndex        =   6
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Concepto"
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
      Index           =   23
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   825
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
      TabIndex        =   1
      Top             =   165
      Width           =   855
   End
End
Attribute VB_Name = "frmConSdosxMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbConcepto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtMes.SetFocus
   End If
End Sub

Private Sub cmbE_Socio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtCodSocio.SetFocus
   End If
End Sub

Private Sub cmdBuscar_Click()
   LlenaCab
   TotalCab

   DataGrid1.SetFocus
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Long, I As Integer, Heading(10) As String, wRegAct As Long, wRegTot As Long
   Dim wNom As String, wSoc As Long, wMes As String, wcon As String, wMon As String, _
       wCargos As Currency, wAbonos As Currency, wSdoNew As Currency, _
       zCargos As Currency, zAbonos As Currency, zSdoNew As Currency
   Heading(0) = "MES"
   Heading(1) = "SOCIO"
   Heading(2) = "CODIGO"
   Heading(3) = "INS"
   Heading(4) = "NOMBRE"
   Heading(5) = "ESTADO"
   Heading(6) = "CONCEPTO"
   Heading(7) = "MON"
   Heading(8) = "CUOTA"
   Heading(9) = "COBRO"
   Heading(10) = "SALDO"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 11)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 11)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "REPORTE DE SALDOS POR SOCIO - MES " + txtMes.Text
        For I = 1 To 11 Step 1
            .Cells(3, I) = Heading(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 9
        objExcel.Columns("B").ColumnWidth = 8
        objExcel.Columns("C").ColumnWidth = 10
        objExcel.Columns("D").ColumnWidth = 3
        objExcel.Columns("E").ColumnWidth = 50
        objExcel.Columns("F").ColumnWidth = 5
        objExcel.Columns("G").ColumnWidth = 35
        objExcel.Columns("H").ColumnWidth = 8
        objExcel.Columns("I").ColumnWidth = 10
        objExcel.Columns("J").ColumnWidth = 10
        objExcel.Columns("K").ColumnWidth = 10
   End With
   
   aa = Leerado3("SELECT * FROM TMP_CTASXCAB WHERE USU = '" + wcodusu + "' ORDER BY MES, MONEDA, NOMBRE, CONCEPTO ")
   If aa > 0 Then
      V = 4
      H = 1
      wRegAct = 1
      wRegTot = aa
      zCargos = 0: zAbonos = 0: zSdoNew = 0
      Do While Not ADO3.EOF
         wSoc = ADO3!codsocio
         wMes = ADO3!mes
         wcon = ADO3!concepto
         wNom = ADO3!nombre
         wMon = ADO3!moneda
         wCargos = 0: wAbonos = 0: wSdoNew = 0
                             
         
         Do While ADO3!mes = wMes And ADO3!moneda = wMon
            DoEvents
            lblMensaje.Caption = "Traslando a EXCEL - Registro " + _
                                 Format(wRegAct, "####0") + " / " + _
                                 Format(wRegTot, "####0")
            lblMensaje.Refresh
         
            objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 10)).NumberFormat = "####,##0.00;-####,##0.00;\ "
         
            objExcel.Cells(V, H + 0) = ADO3!mes
            objExcel.Cells(V, H + 1) = ADO3!codsocio
            objExcel.Cells(V, H + 2) = ADO3!codigo
            objExcel.Cells(V, H + 3) = ADO3!ins
            objExcel.Cells(V, H + 4) = ADO3!nombre
            objExcel.Cells(V, H + 5) = ADO3!e_socio
            objExcel.Cells(V, H + 6) = ADO3!nomcon
            objExcel.Cells(V, H + 7) = ADO3!moneda
            objExcel.Cells(V, H + 8) = ADO3!cargos
            objExcel.Cells(V, H + 9) = ADO3!abonos
            objExcel.Cells(V, H + 10) = ADO3!sdonew
         
            wCargos = wCargos + ADO3!cargos
            wAbonos = wAbonos + ADO3!abonos
            wSdoNew = wSdoNew + ADO3!sdonew
         
            zCargos = zCargos + ADO3!cargos
            zAbonos = zAbonos + ADO3!abonos
            zSdoNew = zSdoNew + ADO3!sdonew
         
            V = V + 1
            wRegAct = wRegAct + 1
            ADO3.MoveNext
            If ADO3.EOF Then
               Exit Do
            End If
         Loop
         objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 10)).NumberFormat = "####,##0.00;-####,##0.00;\ "
         objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 10)).Font.Bold = True
         objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 10)).Font.Color = RGB(255, 0, 0)
         objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 10)).Borders.LineStyle = xlContinuous
         objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 10)).Borders.Color = RGB(255, 0, 0)
                  
         objExcel.Cells(V, H + 6) = "TOTALES X MES " + IIf(wMon = "S", "S/.", "US$")
         objExcel.Cells(V, H + 8) = wCargos
         objExcel.Cells(V, H + 9) = wAbonos
         objExcel.Cells(V, H + 10) = wSdoNew
         V = V + 2
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
   Dim wFec As Date, wMes As String, wAno As String
   wAno = Left(txtMes.Text, 4)
   wMes = Trim(funnommes(Right(txtMes.Text, 2))) + " DEL " + wAno
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\SdosxMES.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "MESCORTE= 'DEL MES DE " + wMes + "' "
   Crys1.SelectionFormula = " {TMP_CTASXCAB.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Dim a As Integer, wAno As String, wMes As String
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
   
   cmbConcepto.Clear
   cmbConcepto.AddItem "Todos Los Conceptos"
   a = Leerado8("SELECT * FROM MAECONCEPTO ORDER BY CONCEPTO ")
   If a > 0 Then
      ADO8.MoveFirst
      I = 0
      Do While Not ADO8.EOF
         cmbConcepto.AddItem ADO8!nombre
         I = I + 1
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   cmbConcepto.ListIndex = 0
   
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
   
   wAno = wanocia
   wMes = Format(Month(Date), "00")
   
   txtMes.Text = wAno + "/" + wMes
   
   txtMes.SetFocus
End Sub

Private Sub Form_Load()
   frmConSdosxMes.Left = (Screen.Width - Width) \ 2
   frmConSdosxMes.Top = 0
End Sub

Private Sub LlenaCab()
   Dim aa As Long, _
       wcon As String, WE_S As String, wMes As String, _
       wSoc As Integer, sw As String
   
   Set DataGrid1.DataSource = Nothing
   
   wMes = txtMes.Text
   wcon = BuscaCodConcepto(cmbConcepto.List(cmbConcepto.ListIndex))
   WE_S = BuscaCodEsocio(cmbE_Socio.List(cmbE_Socio.ListIndex))
   wSoc = Val(txtCodSocio.Text)
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_CTASXCAB WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   If optSaldo.Value = True Then
      sw = " WHERE C.SDONEW <> 0 AND "
   Else
      sw = ""
   End If
   
   If Len(Trim(wMes)) <> 0 Then
      If Len(Trim(sw)) = 0 Then
         sw = "WHERE "
      End If
      sw = sw + "C.MES = '" + wMes + "'"
   End If
   If Len(Trim(wcon)) > 0 Then
      If Len(Trim(sw)) = 0 Then
         sw = "WHERE "
      Else
         sw = sw + " AND "
      End If
      sw = sw + "C.CONCEPTO = '" + wcon + "'"
   End If
   If Len(Trim(WE_S)) > 0 Then
      If Len(Trim(sw)) = 0 Then
         sw = "WHERE "
      Else
         sw = sw + " AND "
      End If
      sw = sw + "S.E_SOCIO = '" + WE_S + "'"
   End If
   If wSoc > 0 Then
      If Len(Trim(sw)) = 0 Then
         sw = "WHERE "
      Else
         sw = sw + " AND "
      End If
      sw = sw + " S.CODSOCIO = " + Str(wSoc) + " "
   End If
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_CTASXCAB " _
   & " (CODSOCIO, CODIGO, INS, MES, CONCEPTO, NOMBRE, NOMCON, E_SOCIO, MONEDA, " _
   & "  CARGOS, ABONOS, SDONEW, USU ) " _
   & " SELECT " _
   & "  C.CODSOCIO, S.CODIGO, S.INS, C.MES, C.CONCEPTO, S.NOMBRE, M.NOMBRE, S.E_SOCIO, " _
   & "  C.MONEDA, C.CARGOS, C.ABONOS, C.SDONEW, '" + wcodusu + "'  " _
   & " FROM CTASXCAB AS C INNER JOIN MAECONCEPTO AS M ON C.CONCEPTO = M.CONCEPTO " _
   & "                    INNER JOIN MAESOCIO    AS S ON C.CODSOCIO = S.CODSOCIO " _
   & " " + sw + "")
   Db.CommitTrans

   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, MES, NOMCON, E_SOCIO, " _
            & "          MONEDA, CARGOS, ABONOS, SDONEW, USU, " _
            & "          CONCEPTO, CODSOCIO " _
            & " FROM TMP_CTASXCAB " _
            & " WHERE USU = '" + wcodusu + "' " _
            & " ORDER BY NOMBRE, MES, CONCEPTO ")
   Set DataGrid1.DataSource = ADO2
   
   DataGrid1.Columns(0).Width = 700   ' CODSOCIO
   DataGrid1.Columns(0).Alignment = dbgLeft
   DataGrid1.Columns(0).Caption = "COD.SOCIO"
    
   DataGrid1.Columns(1).Width = 1000  ' NOMBRE
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Caption = "CODIGO"
    
   DataGrid1.Columns(2).Width = 400   ' INS
   DataGrid1.Columns(2).Alignment = dbgLeft
   DataGrid1.Columns(2).Caption = "INS"
    
   DataGrid1.Columns(3).Width = 4150  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE"
    
   DataGrid1.Columns(4).Width = 750   ' MESCOB
   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Caption = "MES"
    
   DataGrid1.Columns(5).Width = 2500  ' NOMCON
   DataGrid1.Columns(5).Alignment = dbgLeft
   DataGrid1.Columns(5).Caption = "CONCEPTO"
    
   DataGrid1.Columns(6).Width = 500   ' E_SOCIO
   DataGrid1.Columns(6).Alignment = dbgLeft
   DataGrid1.Columns(6).Caption = "E_SOCIO"
    
   DataGrid1.Columns(7).Width = 350   ' MONEDA
   DataGrid1.Columns(7).Alignment = dbgCenter
   DataGrid1.Columns(7).Caption = "MON"
    
   DataGrid1.Columns(8).Width = 850   ' CARGOS
   DataGrid1.Columns(8).Alignment = dbgRight
   DataGrid1.Columns(8).Caption = "CARGOS"
   DataGrid1.Columns(8).NumberFormat = "###,##0.00;-###,##0.00;\ "
    
   DataGrid1.Columns(9).Width = 850   ' CARGOS
   DataGrid1.Columns(9).Alignment = dbgRight
   DataGrid1.Columns(9).Caption = "ABONOS"
   DataGrid1.Columns(9).NumberFormat = "###,##0.00;-###,##0.00;\ "
    
   DataGrid1.Columns(10).Width = 850   ' SDONEW
   DataGrid1.Columns(10).Alignment = dbgRight
   DataGrid1.Columns(10).Caption = "SALDOS"
   DataGrid1.Columns(10).NumberFormat = "###,##0.00;-###,##0.00;###,##0.00"
    
   DataGrid1.Columns(11).Visible = False
   DataGrid1.Columns(12).Visible = False
   DataGrid1.Columns(13).Visible = False
End Sub

Private Sub TotalCab()
   Dim zz As Integer, _
       wCargos As Currency, wAbonos As Currency, wSdoNew As Currency
   
   wCargos = 0: wAbonos = 0: wSdoNew = 0
   zz = Leerado7a("SELECT SUM(CARGOS) AS CARGOS, SUM(ABONOS) AS ABONOS, " _
                & "       SUM(SDONEW) AS SDONEW " _
                & " FROM TMP_CTASXCAB " _
                & " WHERE USU = '" + wcodusu + "' ")
   If zz > 0 Then
      wCargos = IIf(IsNull(ADO7a!cargos), 0, ADO7a!cargos)
      wAbonos = IIf(IsNull(ADO7a!abonos), 0, ADO7a!abonos)
      wSdoNew = IIf(IsNull(ADO7a!sdonew), 0, ADO7a!sdonew)
   End If
   Set ADO7a = Nothing
   
   lblCargos.Caption = Format(wCargos, "####,##0.00;;\ ")
   lblAbonos.Caption = Format(wAbonos, "####,##0.00;;\ ")
   lblSdoNew.Caption = Format(wSdoNew, "####,##0.00;;\ ")
End Sub

Private Sub optSaldo_Click()
   LlenaCab
End Sub

Private Sub optTodos_Click()
   LlenaCab
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

Private Sub txtFecha_GotFocus()
   txtFecha.SelStart = 0
   txtFecha.SelLength = 10
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFecha.Text = "__/__/____" Then
         MsgBox "Fecha En Blanco", vbExclamation
         txtFecha.Text = Format(Date, "dd/mm/yyyy")
         Exit Sub
      End If
      txtCodSocio.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii), 1) = 0 Then
         KeyAscii = 0
      End If
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
