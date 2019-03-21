VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConListadoCeo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Alfabetico para CEO"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   12930
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
      Left            =   7080
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox cmbE_Socio 
      Height          =   315
      ItemData        =   "frmConListadoCeo.frx":0000
      Left            =   1440
      List            =   "frmConListadoCeo.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   920
      Width           =   3375
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
      Left            =   11160
      TabIndex        =   5
      Top             =   7320
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
      Left            =   8520
      TabIndex        =   4
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir "
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
      Left            =   9840
      TabIndex        =   3
      Top             =   7320
      Width           =   1095
   End
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmConListadoCeo.frx":0004
      Left            =   1440
      List            =   "frmConListadoCeo.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   200
      Width           =   7335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5295
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   9340
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
      Left            =   11520
      Top             =   240
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
   Begin MSMask.MaskEdBox txtMes 
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtMoroso 
      Height          =   285
      Left            =   3720
      TabIndex        =   13
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Mes Corte"
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
      TabIndex        =   15
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Mes Moroso"
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
      Left            =   2520
      TabIndex        =   14
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Cantidad Socios Activos"
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
      Left            =   9000
      TabIndex        =   11
      Top             =   6840
      Width           =   2175
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11280
      TabIndex        =   10
      Top             =   6840
      Width           =   1095
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
      Left            =   240
      TabIndex        =   8
      Top             =   920
      Width           =   1140
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
      Left            =   360
      TabIndex        =   6
      Top             =   7320
      Width           =   7575
   End
   Begin VB.Label Label38 
      Alignment       =   1  'Right Justify
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
      Left            =   480
      TabIndex        =   1
      Top             =   200
      Width           =   855
   End
End
Attribute VB_Name = "frmConListadoCeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
   LlenaCab
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(8) As String, wreg As Integer, wtot As Integer
   Dim wNom As String, wNum1 As Long, wNum2 As Integer, wSoc As Integer, wFecIng As Date, wFecNac As Date
   Heading(0) = "NRO."
   Heading(1) = "GRADO"
   Heading(2) = "TIPO"
   Heading(3) = "APELLIDOS Y NOMBRES"
   Heading(4) = "FEC.ING."
   Heading(5) = "C.APORTE"
   Heading(6) = "RENOVAC"
   Heading(7) = "DEU.APORTE"
   Heading(8) = "DEU.RENOVAC"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 9)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 9)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "LISTADO ALFABETICO DE SOCIOS AOPIP CON SUS FAMILIARES DEPENDIENTES"
        For I = 1 To 9 Step 1
            .Cells(3, I) = Heading(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 5
        objExcel.Columns("B").ColumnWidth = 15
        objExcel.Columns("C").ColumnWidth = 10
        objExcel.Columns("D").ColumnWidth = 60
        objExcel.Columns("E").ColumnWidth = 11
        objExcel.Columns("F").ColumnWidth = 11
        objExcel.Columns("G").ColumnWidth = 11
        objExcel.Columns("H").ColumnWidth = 13
        objExcel.Columns("I").ColumnWidth = 13
   End With
   
   aa = Leerado3("SELECT * FROM TMP_REPCEO WHERE USU = '" + wcodusu + "' ORDER BY NOMBRE, TIPOPARIENTE, LIN ")
   If aa > 0 Then
      wtot = aa
      V = 4
      H = 1
      wNum1 = 1
      Do While Not ADO3.EOF
         wSoc = ADO3!codsocio
         wNom = ADO3!nombre
         
         objExcel.Range(objExcel.Cells(V, H + 7), objExcel.Cells(V, H + 8)).NumberFormat = "####,##0.00;;\ "
         objExcel.Range(objExcel.Cells(V, H + 0), objExcel.Cells(V, H + 8)).Font.Color = RGB(0, 0, 255)
         objExcel.Range(objExcel.Cells(V, H + 0), objExcel.Cells(V, H + 8)).Font.Bold = True
         objExcel.Range(objExcel.Cells(V, H + 0), objExcel.Cells(V, H + 8)).Borders.LineStyle = xlContinuous
         objExcel.Range(objExcel.Cells(V, H + 0), objExcel.Cells(V, H + 8)).Borders.Color = RGB(0, 0, 255)
         
         objExcel.Cells(V, H + 0) = wNum1
         objExcel.Cells(V, H + 1) = ADO3!nomgra
         objExcel.Cells(V, H + 2) = ADO3!e_socio
         objExcel.Cells(V, H + 3) = ADO3!nombre
         If Format(ADO3!fecing, "dd/mm/yyyy") > Format("01/01/1900", "dd/mm/yyyy") Then
            wFecIng = Format(ADO3!fecing, "dd/mm/yyyy")
            objExcel.Cells(V, H + 4) = wFecIng
         End If
         objExcel.Cells(V, H + 5) = IIf(ADO3!moneda = "S", "S/.", "US$") + Format(ADO3!aporte, "###0.00")
         objExcel.Cells(V, H + 6) = IIf(ADO3!moneda = "S", "S/.", "US$") + Format(ADO3!renova, "###0.00")
         objExcel.Cells(V, H + 7) = Format(ADO3!deuapo, "###0.00")
         objExcel.Cells(V, H + 8) = Format(ADO3!deuren, "###0.00")
         V = V + 1
         wNum2 = 1
         Do While wSoc = ADO3!codsocio
            DoEvents
            lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wreg, "####0") + " / " + Format(wtot, "####0")
            lblMensaje.Refresh
         
            If Len(Trim(ADO3!tipopariente)) > 0 Then
               objExcel.Cells(V, H + 1) = Format(wNum2, "##0")
               Select Case ADO3!tipopariente
               Case "E"
                    objExcel.Cells(V, H + 2) = "ESPOSO(A)"
               Case "P"
                    objExcel.Cells(V, H + 2) = "PADRE"
               Case "M"
                    objExcel.Cells(V, H + 2) = "MADRE"
               Case "H"
                    objExcel.Cells(V, H + 2) = "HIJO(A)"
               Case "N"
                    objExcel.Cells(V, H + 2) = "NIETO(A)"
               Case Else
                    objExcel.Cells(V, H + 2) = "DESCONOCIDO"
               End Select
               objExcel.Cells(V, H + 3) = ADO3!nomparien
               If Format(ADO3!fecnac, "dd/mm/yyyy") > Format("01/01/1900", "dd/mm/yyyy") Then
                  wFecNac = Format(ADO3!fecnac, "dd/mm/yyyy")
                  objExcel.Cells(V, H + 4) = wFecNac
               End If
               wNum2 = wNum2 + 1
               V = V + 1
            End If
            
            wreg = wreg + 1
            ADO3.MoveNext
            If ADO3.EOF Then
               Exit Do
            End If
         Loop
         V = V + 1
         wNum1 = wNum1 + 1
         
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
   Crys1.ReportFileName = xraiz + "ReportCtaCte\RepCeo.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= 'AL MES DE " + wMes + "' "
   Crys1.SelectionFormula = " {TMP_REPCEO.USU}='" + wcodusu + "' "
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
   
   cmbE_socio.Clear
   cmbE_socio.AddItem "Todos Los Estados de Socio"
   a = Leerado8("SELECT * FROM MAEE_SOCIO ORDER BY E_SOCIO ")
   If a > 0 Then
      ADO8.MoveFirst
      I = 0
      Do While Not ADO8.EOF
         cmbE_socio.AddItem ADO8!nombre
         I = I + 1
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   cmbE_socio.ListIndex = 0
   
   txtMes.Text = Left(zMesTope, 4) + "/" + Right(zMesTope, 2)
   
   
'   If Right(zMesTope, 2) = "01" Then
'      txtMoroso.Text = Format(Val(Left(zMesTope, 4)) - 1, "0000") + "12"
'   Else
'      txtMoroso.Text = Left(zMesTope, 4) + "/" + Format(Val(Right(zMesTope, 2)) - 1, "00")
'   End If
   
   txtMoroso.Text = Left(zMesTope, 4) + "/" + Right(zMesTope, 2)
   
   txtMes.Enabled = False
   txtMoroso.Enabled = False

   Set DataGrid1.DataSource = Nothing
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_REPCEO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   cmbE_socio.SetFocus
End Sub

Private Sub Form_Load()
   frmConListadoCeo.Left = (Screen.Width - Width) \ 2
   frmConListadoCeo.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_REPCEO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim aa As Long, w As String, wMes As String, wFec As Date, wE_S As String, wSoc As Integer, wNom As String, wSdo As Currency

   wMes = txtMoroso.Text
   wFec = Format(fundiames(Right(wMes, 2)) + "/" + Right(wMes, 2) + "/" + Left(wMes, 4), "dd/mm/yyyy")
   wE_S = BuscaCodEsocio(cmbE_socio.List(cmbE_socio.ListIndex))

   w = ""
   If wE_S <> "" Then
      w = " AND S.E_SOCIO = '" + wE_S + "' "
   End If

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_REPCEO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_REPCEO " _
   & " (USU, CODSOCIO, TIPOPARIENTE, LIN, CODIGO, INS, NOMBRE, NOMGRA, E_SOCIO, NOME_S, FECING, MONEDA, " _
   & "  APORTE, RENOVA, DEUAPO, DEUREN, NOMPARIEN, FECNAC ) " _
   & " SELECT '" + wcodusu + "', S.CODSOCIO, F.TIPOPARIENTE, F.LIN, S.CODIGO, S.INS, S.NOMBRE, G.NOMBRE AS NOMGRA, " _
   & "                 S.E_SOCIO, E.NOMBRE AS NOME_S, S.FECING, E.MONEDA, E.APORTE, 0 AS RENOVA, " _
   & "                 S.DEUDA_PT2, 0 AS DEUREN, F.NOMBRE AS NOMPARIENTE, F.FECNAC " _
   & " FROM MAESOCIO AS S LEFT  JOIN MAEFAMILIA AS F ON S.CODSOCIO = F.CODSOCIO " _
   & "                    INNER JOIN MAEGRADO AS G   ON S.GRADO = G.GRADO " _
   & "                    LEFT  JOIN MAETIPOPARIENTE AS T ON F.TIPOPARIENTE = T.TIPOPARIENTE " _
   & "                    INNER JOIN MAEE_SOCIO AS E ON S.E_SOCIO = E.E_SOCIO " _
   & " Where E.APORTE > 0 " + w + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_REPCEO " _
   & " SET NOMPARIEN = '' " _
   & " WHERE NOMPARIEN IS NULL ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_REPCEO " _
   & " SET TIPOPARIENTE = '' " _
   & " WHERE TIPOPARIENTE IS NULL ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_REPCEO " _
   & " SET LIN = '' " _
   & " WHERE LIN IS NULL ")
   Db.CommitTrans

   aa = Leerado8("SELECT CODSOCIO, NOMBRE " _
                & " FROM TMP_REPCEO " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " GROUP BY CODSOCIO, NOMBRE " _
                & " ORDER BY NOMBRE ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wSoc = ADO8!codsocio
         wNom = Trim(ADO8!nombre)
         wSdo = SaldoFoto(wSoc, wMes)
         
         DoEvents
         lblMensaje.Caption = "Calculando Socio " + wNom
         lblMensaje.Refresh
   
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_REPCEO " _
         & " SET DEUAPO = " + Str(wSdo) + " - APORTE " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "            USU = '" + wcodusu + "' ")
         Db.CommitTrans
   
         ADO8.MoveNext
      Loop
   End If
   
   DoEvents
   lblMensaje.Caption = ""
   lblMensaje.Refresh
   
   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, NOMGRA, NOME_S, MIN(DEUAPO) AS DEUDA, MIN(DEUREN) AS DEUREN " _
            & " FROM TMP_REPCEO " _
            & " WHERE USU = '" + wcodusu + "' GROUP BY CODSOCIO, CODIGO, INS, NOMBRE, NOMGRA, NOME_S " _
            & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO2
 
   lblTotal.Caption = Format(aa, "##,##0") + " "

   DataGrid1.Columns(0).Width = 700   ' CODSOCIO
   DataGrid1.Columns(0).Alignment = dbgLeft
   DataGrid1.Columns(0).Caption = "COD.SOCIO"
    
   DataGrid1.Columns(1).Width = 1000  ' CODIGO
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Caption = "CODIGO"
    
   DataGrid1.Columns(2).Width = 400   ' INS
   DataGrid1.Columns(2).Alignment = dbgLeft
   DataGrid1.Columns(2).Caption = "INS"
    
   DataGrid1.Columns(3).Width = 4650  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE"
    
   DataGrid1.Columns(4).Width = 1500   ' NOMGRA
   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Caption = "GRADO"
    
   DataGrid1.Columns(5).Width = 1500  ' NOME_S
   DataGrid1.Columns(5).Alignment = dbgLeft
   DataGrid1.Columns(5).Caption = "ESTADO"
    
   DataGrid1.Columns(6).Width = 1000  ' DEUDA_PT2
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).Caption = "DEUDA APORTE"
   DataGrid1.Columns(6).NumberFormat = "#####0.00;;\ "
   
   DataGrid1.Columns(7).Width = 1000  ' DEUREN
   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).Caption = "DEUDA RENOV"
   DataGrid1.Columns(7).NumberFormat = "#####0.00;;\ "
End Sub

