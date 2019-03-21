VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmElePadronGral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Padron General"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   13095
   Begin VB.OptionButton Option1 
      Caption         =   "Solo Socios Activos Habiles"
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
      Left            =   1080
      TabIndex        =   16
      Top             =   1440
      Value           =   -1  'True
      Width           =   3735
   End
   Begin VB.OptionButton optActivos 
      Caption         =   "Solo Socios Activos"
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
      Left            =   1080
      TabIndex        =   15
      Top             =   1200
      Width           =   3735
   End
   Begin VB.OptionButton optTodos 
      Caption         =   "Mostrar Todos Los Socios"
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
      Left            =   1080
      TabIndex        =   14
      Top             =   960
      Width           =   3615
   End
   Begin VB.CommandButton cmdLegalizar 
      Caption         =   "Padrón Legalizar"
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
      Left            =   10320
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
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
      Left            =   5400
      TabIndex        =   12
      Top             =   360
      Width           =   975
   End
   Begin VB.ComboBox cmbE_socio 
      Height          =   315
      ItemData        =   "frmElePadronGral.frx":0000
      Left            =   1320
      List            =   "frmElePadronGral.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   540
      Width           =   3615
   End
   Begin VB.ComboBox cmbGrado 
      Height          =   315
      ItemData        =   "frmElePadronGral.frx":0004
      Left            =   1320
      List            =   "frmElePadronGral.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   180
      Width           =   3615
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
      Left            =   11520
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6960
      Width           =   975
   End
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
      Height          =   495
      Left            =   9120
      TabIndex        =   2
      Top             =   6960
      Width           =   975
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
      Left            =   10320
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6960
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   8281
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
   Begin Crystal.CrystalReport Crys1 
      Left            =   12480
      Top             =   840
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
   Begin Crystal.CrystalReport Crys2 
      Left            =   11760
      Top             =   960
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
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Estado Socio"
      Height          =   195
      Index           =   1
      Left            =   210
      TabIndex        =   11
      Top             =   540
      Width           =   945
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Grado"
      Height          =   195
      Index           =   2
      Left            =   720
      TabIndex        =   9
      Top             =   240
      Width           =   435
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmElePadronGral.frx":0008
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
      Height          =   615
      Left            =   6960
      TabIndex        =   7
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   11160
      TabIndex        =   6
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total de Asociados"
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
      Left            =   9390
      TabIndex        =   5
      Top             =   6480
      Width           =   1650
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
      Left            =   120
      TabIndex        =   4
      Top             =   6840
      Width           =   5775
   End
End
Attribute VB_Name = "frmElePadronGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbE_Socio_KeyPress(KeyAscii As Integer)
   cmdBuscar.SetFocus
End Sub

Private Sub cmbGrado_KeyPress(KeyAscii As Integer)
   cmdBuscar.SetFocus
End Sub

Private Sub cmdBuscar_Click()
   LlenaCab
   LlenaCab1
   TotalCab
End Sub

Private Sub cmdExporta_Click()
   On Error GoTo err
   
   Dim aa As Integer, wRegAct As Integer, wRegTot As Integer, _
       I As Integer, Heading(8) As String, _
       wNum As Integer, wFec As Date, wEso As String, wNomEso As String, _
       wDni As String
       
   Heading(0) = "ESTADO DE SOCIO"
   Heading(1) = "NUM"
   Heading(2) = "GRADO"
   Heading(3) = "NOMBRE ASOCIADO"
   Heading(4) = "FEC.ING"
   Heading(5) = "D.N.I."
   Heading(6) = "DEUDA"
   Heading(7) = "FIRMA"
   Heading(8) = "IMPRESION DIGITAL"
   
   aa = Leerado3("SELECT * FROM TMP_PADRON WHERE USU = '" + wcodusu + "' ORDER BY E_SOCIO, NOMBRE ")
   If aa > 0 Then
      wRegAct = 1
      wRegTot = aa
      Set objExcel = New Excel.Application
      objExcel.SheetsInNewWorkbook = 1
      objExcel.Workbooks.Add
      With objExcel
           .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
           .Range(.Cells(3, 1), .Cells(3, 9)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 9)).Font.Bold = True
           .Cells(1, 1) = wnomcia
           .Cells(2, 1) = "PADRON GENERAL DE ASOCIADOS - ORDENADO POR ESTADO DE SOCIO"
           For I = 1 To 9 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           objExcel.Columns("A").ColumnWidth = 12
           objExcel.Columns("B").ColumnWidth = 6
           objExcel.Columns("C").ColumnWidth = 15
           objExcel.Columns("D").ColumnWidth = 50
           objExcel.Columns("E").ColumnWidth = 11
           objExcel.Columns("F").ColumnWidth = 10
           objExcel.Columns("G").ColumnWidth = 11
           objExcel.Columns("H").ColumnWidth = 18
           objExcel.Columns("I").ColumnWidth = 18
      End With
      V = 4
      H = 1
      wNum = 1
      Do While Not ADO3.EOF
         wEso = ADO3!e_socio
         wNomEso = ADO3!nome_socio
         wNum = 1
         objExcel.Cells(V, H + 0) = wEso + " " + wNomEso
         V = V + 1
         Do While ADO3!e_socio = wEso
            
            DoEvents
            lblMensaje.Caption = "Traslando a EXCEL - Registro " + _
                                 Trim(Format(wRegAct, "####0")) + " / " + _
                                 Trim(Format(wRegTot, "####0"))
            lblMensaje.Refresh
            objExcel.Range(objExcel.Cells(V, H + 6), objExcel.Cells(V, H + 6)).NumberFormat = "####,##0.00"
         
            objExcel.Cells(V, H + 1) = wNum
            objExcel.Cells(V, H + 2) = Trim(ADO3!nomgra)
            objExcel.Cells(V, H + 3) = Trim(ADO3!nombre)
            If IsDate(ADO3!fecing) Then
               wFec = Format(ADO3!fecing, "dd/mm/yyyy")
               objExcel.Cells(V, H + 4) = wFec
            End If
            objExcel.Cells(V, H + 5) = ADO3!numdoc
            objExcel.Cells(V, H + 6) = ADO3!deuda_pt2
            objExcel.Cells(V, H + 7) = "_______________"
            objExcel.Cells(V, H + 8) = "_______________"
         
            wRegAct = wRegAct + 1
            V = V + 1
            wNum = wNum + 1
            ADO3.MoveNext
            If ADO3.EOF Then
               Exit Do
            End If
         Loop
         
      Loop
      
      V = V + 1
      objExcel.Cells(V, H + 3) = "TOTAL GENERAL " + Format(wRegTot, "##,##0")
      
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
   Crys2.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys2.ReportFileName = xraiz + "ReportCtaCte\PadronCompleto.RPT"
   Crys2.SelectionFormula = " {TMP_PADRON.USU}='" + wcodusu + "' "
   Crys2.WindowState = crptMaximized
   Crys2.Action = 1
End Sub

Private Sub cmdLegalizar_Click()
   Dim wNomMes As String
   wNomMes = Trim(funnommes(Format(Month(Date), "00"))) + " DEL " + Format(Year(Date), "0000")
   
   wNomMes = "AL " + Format(Day(Date), "00") + " DE " + Trim(funnommes(Format(Month(Date), "00"))) + " DEL " + Format(Year(Date), "0000")
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\PadronLegalizar.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= '" + wNomMes + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Form_Initialize
End Sub

Private Sub Form_Initialize()
   frmElePadronGral.Left = (Screen.Width - Width) \ 2
   frmElePadronGral.Top = 0
   
   cmbGrado.Clear
   cmbGrado.AddItem "000 TODOS LOS GRADOS"
   aa = Leerado8("SELECT * FROM MAEGRADO ORDER BY GRADO ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
   
         cmbGrado.AddItem Format(ADO8!grado, "@@@") + " " + ADO8!nombre
   
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   cmbGrado.ListIndex = 0
   
   cmbE_Socio.Clear
   cmbE_Socio.AddItem "000 TODOS LOS ESTADOS"
   aa = Leerado8("SELECT * FROM MAEE_SOCIO ORDER BY E_SOCIO ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
   
         cmbE_Socio.AddItem ADO8!e_socio + " " + ADO8!nombre
   
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   cmbE_Socio.ListIndex = 0
   
   cmdBuscar.SetFocus
End Sub

Private Sub LlenaCab()
   Dim aa As Long, wGra As Integer, wEso As String

   lblMensaje.Caption = "Preparando Archivo......Espere"
   lblMensaje.Refresh

   Set DataGrid1.DataSource = Nothing
   
   If cmbGrado.ListIndex = 0 Then
      wGra = 0
   Else
      wGra = Val(Left(cmbGrado.Text, 3))
   End If

   If cmbE_Socio.ListIndex = 0 Then
      wEso = ""
   Else
      wEso = Left(cmbE_Socio.Text, 3)
   End If

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_PADRON WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_PADRON " _
   & " (CODSOCIO, CODIGO, INS, TIPDOC, NUMDOC, CARNETPNP, CARNETPIP, GRADO, SITU, " _
   & "  NOMBRE, E_SOCIO, FECRET, UNIPAG, REGION, FECING, DEUDA_PT2, NOMGRA, " _
   & "  NOME_SOCIO, NOMUNIPAG, NOMREGION, GRADOGRUPO, NOMGRADOGRUPO, " _
   & "  REGIONGRUPO, NOMREGIONGRUPO, USU ) " _
   & " SELECT " _
   & "  CODSOCIO, CODIGO, INS, '01', NUMDOC, CARNETPNP, CARNETPIP, GRADO, SITU, " _
   & "  NOMBRE, E_SOCIO, FECRET, UNIPAG, REGION, FECING, DEUDA_PT2, " _
   & "  '', '', '', '', 0, '', '', '', '" + wcodusu + "' " _
   & " FROM MAESOCIO ")
   Db.CommitTrans

   If optActivos.Value = True Then
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_PADRON " _
      & " WHERE USU = '" + wcodusu + "' AND " _
      & "       E_SOCIO <> 'TIT' AND " _
      & "       E_SOCIO <> 'HIJ' AND " _
      & "       E_SOCIO <> 'NIE' AND " _
      & "       E_SOCIO <> 'HER' AND " _
      & "       E_SOCIO <> 'VIU' AND " _
      & "       E_SOCIO <> 'CIV' AND " _
      & "       E_SOCIO <> 'CI1' AND " _
      & "       E_SOCIO <> 'TRA' AND " _
      & "       E_SOCIO <> 'ADH' AND " _
      & "       E_SOCIO <> 'PNP' ")
      Db.CommitTrans
   End If
   
   If wGra <> 0 Then
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_PADRON WHERE USU = '" + wcodusu + "' AND GRADO <> " + Str(wGra) + " ")
      Db.CommitTrans
   End If

   If Len(Trim(wEso)) <> 0 Then
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_PADRON WHERE USU = '" + wcodusu + "' AND E_SOCIO <> '" + wEso + "' ")
      Db.CommitTrans
   End If

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_PADRON " _
   & " SET NOMGRA = G.NOMBRE, GRADOGRUPO = G.GRADOGRUPO " _
   & " FROM TMP_PADRON AS T INNER JOIN MAEGRADO AS G " _
   & "   ON T.GRADO = G.GRADO " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_PADRON " _
   & " SET NOMGRADOGRUPO = G.NOMBRE " _
   & " FROM TMP_PADRON AS T INNER JOIN MAEGRADOGRUPO AS G " _
   & "   ON T.GRADOGRUPO = G.GRADOGRUPO " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_PADRON " _
   & " SET NOME_SOCIO = E.NOMBRE " _
   & " FROM TMP_PADRON AS T INNER JOIN MAEE_SOCIO AS E " _
   & "   ON T.E_SOCIO = E.E_SOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_PADRON " _
   & " SET NOMUNIPAG = U.NOMBRE, REGION = U.REGION " _
   & " FROM TMP_PADRON AS T INNER JOIN MAEUNIDAD AS U " _
   & "   ON T.UNIPAG = U.UNIDAD " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_PADRON " _
   & " SET NOMREGION = R.NOMBRE, REGIONGRUPO = R.REGIONGRUPO " _
   & " FROM TMP_PADRON AS T INNER JOIN MAEREGION AS R " _
   & "   ON T.REGION = R.REGION " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_PADRON " _
   & " SET NOMREGIONGRUPO = R.NOMBRE " _
   & " FROM TMP_PADRON AS T INNER JOIN MAEREGIONGRUPO AS R " _
   & "   ON T.REGIONGRUPO = R.REGIONGRUPO " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans

   lblMensaje.Caption = ""
   lblMensaje.Refresh

   aa = Leerado("SELECT CODSOCIO, NOMGRA, CODIGO, INS, NOMBRE, " _
                & "     FECING, NUMDOC, DEUDA_PT2, NOME_SOCIO, " _
                & "     NOMUNIPAG, NOMREGION " _
                & " FROM TMP_PADRON " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY NOME_SOCIO, NOMBRE ")
   Set DataGrid1.DataSource = ADO1
End Sub

Private Sub LlenaCab1()
   DataGrid1.Columns(0).Alignment = dbgLeft
   DataGrid1.Columns(0).Width = 800
   DataGrid1.Columns(0).Caption = "CODIGO"

   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Width = 1200
   DataGrid1.Columns(1).Caption = "GRADO"

   DataGrid1.Columns(2).Alignment = dbgLeft
   DataGrid1.Columns(2).Width = 800
   DataGrid1.Columns(2).Caption = "CODOFIN"

   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Width = 370
   DataGrid1.Columns(3).Caption = "INS"

   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Width = 3500
   DataGrid1.Columns(4).Caption = "NOMBRE"

   DataGrid1.Columns(5).Alignment = dbgCenter
   DataGrid1.Columns(5).Width = 1080
   DataGrid1.Columns(5).Caption = "FEC.ING"
   DataGrid1.Columns(5).NumberFormat = "dd/mm/yyyy"

   DataGrid1.Columns(6).Alignment = dbgLeft
   DataGrid1.Columns(6).Width = 870
   DataGrid1.Columns(6).Caption = "D.N.I."

   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).Width = 1080
   DataGrid1.Columns(7).Caption = "DEUDA"
   DataGrid1.Columns(7).NumberFormat = "###,##0.00;;\ "

   DataGrid1.Columns(8).Width = 1500
   DataGrid1.Columns(8).Alignment = dbgLeft
   DataGrid1.Columns(8).Caption = "ESTADO"
    
   DataGrid1.Columns(9).Width = 1800
   DataGrid1.Columns(9).Alignment = dbgLeft
   DataGrid1.Columns(9).Caption = "UNI.PAG"
    
   DataGrid1.Columns(10).Width = 1500
   DataGrid1.Columns(10).Alignment = dbgLeft
   DataGrid1.Columns(10).Caption = "REGION"
End Sub

Private Sub TotalCab()
   Dim zz As Integer, wTot As Long

   zz = Leerado8("SELECT COUNT(*) AS TOTAL FROM TMP_PADRON " _
                & " WHERE USU = '" + wcodusu + "' ")
   If zz > 0 Then
      wTot = IIf(IsNull(ADO8!Total), 0, ADO8!Total)
   End If
   Set ADO8 = Nothing

   lblTotal.Caption = Format(wTot, "###,##0")
End Sub

Private Sub optActivos_Click()
   LlenaCab
   LlenaCab1
   TotalCab
End Sub

Private Sub optTodos_Click()
   LlenaCab
   LlenaCab1
   TotalCab
End Sub
