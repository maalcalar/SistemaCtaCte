VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAporteEstadoCta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estado de Cuenta"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   14070
   Begin VB.OptionButton optSaldos 
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
      Height          =   195
      Left            =   8400
      TabIndex        =   27
      Top             =   720
      Width           =   2895
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
      Height          =   195
      Left            =   8400
      TabIndex        =   26
      Top             =   360
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.CommandButton cndOtro 
      Caption         =   "&Otra Consulta"
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
      TabIndex        =   23
      Top             =   6600
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
      Left            =   12360
      TabIndex        =   22
      Top             =   6600
      Width           =   1095
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
      Left            =   11040
      TabIndex        =   21
      Top             =   6600
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
      Left            =   9720
      TabIndex        =   20
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox txtGrado 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      MaxLength       =   3
      TabIndex        =   14
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtTipCob 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      MaxLength       =   3
      TabIndex        =   13
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtNumdoc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7080
      MaxLength       =   8
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtIns 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      MaxLength       =   1
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      MaxLength       =   8
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtE_socio 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtCodSocio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      MaxLength       =   9
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4695
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Width           =   13815
      _ExtentX        =   24368
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
      Caption         =   "ESTADO DE CUENTA DE APORTACIONES ASOCIADOS"
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
   Begin MSMask.MaskEdBox txtMesCierre 
      Height          =   285
      Left            =   1080
      TabIndex        =   28
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
      Left            =   360
      Top             =   6720
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
   Begin VB.Label Label9 
      Caption         =   "Mes Cierre"
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Saldo Total"
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
      Height          =   255
      Left            =   7680
      TabIndex        =   25
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9240
      TabIndex        =   24
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lblGrado 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Grado"
      Height          =   195
      Left            =   360
      TabIndex        =   17
      Top             =   900
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo de Cobro"
      Height          =   195
      Left            =   2520
      TabIndex        =   16
      Top             =   900
      Width           =   1335
   End
   Begin VB.Label lblTipCob 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "D.N.I."
      Height          =   195
      Left            =   7080
      TabIndex        =   12
      Top             =   420
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Ins"
      Height          =   195
      Left            =   6720
      TabIndex        =   11
      Top             =   420
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Codofin"
      Height          =   195
      Left            =   5760
      TabIndex        =   10
      Top             =   420
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Estado del Socio"
      Height          =   195
      Left            =   5040
      TabIndex        =   9
      Top             =   900
      Width           =   1335
   End
   Begin VB.Label lblE_socio 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de Socio"
      Height          =   195
      Left            =   1200
      TabIndex        =   7
      Top             =   420
      Width           =   1695
   End
   Begin VB.Label lblCodSocio 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "Cod.Socio"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   420
      Width           =   975
   End
End
Attribute VB_Name = "frmAporteEstadoCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Limpiar()
   txtCodSocio.Text = ""
   txtCodigo.Text = ""
   txtIns.Text = ""
   txtNumdoc.Text = ""
   txtE_socio.Text = ""
   txtGrado.Text = ""
   txtTipCob.Text = ""

   Set DataGrid1.DataSource = Nothing
   txtCodSocio.SetFocus
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(10) As String, wRegAct As Integer, wRegTot As Integer
   Dim wSoc As Integer, wNom As String, wTot As Currency
   wSoc = Val(txtCodSocio.Text)
   wNom = Trim(lblCodSocio.Caption)
   
   Heading(0) = "CONCEPTO"
   Heading(1) = "NOMBRE CONCEPTO"
   Heading(2) = "MES"
   Heading(3) = "MONEDA"
   Heading(4) = "PROVIS."
   Heading(5) = "COBROS"
   Heading(6) = "SALDO"
   Heading(8) = "DIECO"
   Heading(9) = "CAJA MP"
   Heading(10) = "TESORERIA"
   aa = Leerado3("SELECT * FROM TMP_CTASXCAB " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "       CODSOCIO = " + Str(wSoc) + " " _
                & " ORDER BY CONCEPTO, MES ")
   If aa > 0 Then
      wRegAct = 1
      wRegTot = aa
      Set objExcel = New Excel.Application
      objExcel.SheetsInNewWorkbook = 1
      objExcel.Workbooks.Add
      With objExcel
           .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
           .Range(.Cells(3, 1), .Cells(3, 11)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 11)).Font.Bold = True
           .Cells(1, 1) = wnomcia
           .Cells(2, 1) = "ESTADO DE CUENTA POR SOCIO " + wNom
           For I = 1 To 11 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           objExcel.Columns("A").ColumnWidth = 10
           objExcel.Columns("B").ColumnWidth = 50
           objExcel.Columns("C").ColumnWidth = 11
           objExcel.Columns("D").ColumnWidth = 7
           
           objExcel.Columns("E").ColumnWidth = 11
           objExcel.Columns("F").ColumnWidth = 11
           objExcel.Columns("G").ColumnWidth = 11
           objExcel.Columns("H").ColumnWidth = 3
           objExcel.Columns("I").ColumnWidth = 11
           objExcel.Columns("J").ColumnWidth = 11
           objExcel.Columns("K").ColumnWidth = 11
      End With
      V = 4
      H = 1
      wTot = 0
      Do While Not ADO3.EOF
'         DoEvents
'         lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wRegAct, "####0") + " / " + Format(wRegTot, "####0")
'         lblMensaje.Refresh
         
         objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 6)).NumberFormat = "######0.00"
         objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 10)).NumberFormat = "######0.00;;\ "
         
         objExcel.Cells(V, H + 0) = ADO3!concepto
         objExcel.Cells(V, H + 1) = IIf(IsNull(ADO3!nomcon), "", ADO3!nomcon)
         objExcel.Cells(V, H + 2) = ADO3!mes
         objExcel.Cells(V, H + 3) = IIf(ADO3!moneda = "S", "S/.", "US$")
         
         objExcel.Cells(V, H + 4) = ADO3!cargos
         objExcel.Cells(V, H + 5) = ADO3!abonos
         objExcel.Cells(V, H + 6) = ADO3!sdonew
         objExcel.Cells(V, H + 7) = ""
         objExcel.Cells(V, H + 8) = ADO3!imppag1
         objExcel.Cells(V, H + 9) = ADO3!imppag2
         objExcel.Cells(V, H + 10) = ADO3!imppag3
         
         wTot = wTot + ADO3!sdonew
         wRegAct = wRegAct + 1
         V = V + 1
         ADO3.MoveNext
      Loop
      V = V + 1
      
      objExcel.Range(objExcel.Cells(V, H + 6), objExcel.Cells(V, H + 6)).NumberFormat = "######0.00 "
      objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 6)).Font.Color = RGB(255, 0, 0)
      
      objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 6)).Borders.Color = RGB(255, 0, 0)
      objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 6)).Borders.LineStyle = xlContinuous
      
      
      objExcel.Cells(V, H + 5) = "SALDO TOTAL"
      objExcel.Cells(V, H + 6) = wTot
      V = V + 1
      
      Set ADO3 = Nothing
      objExcel.Visible = True
      Set objExcel = Nothing
   End If
'   lblMensaje.Caption = ""
'   lblMensaje.Refresh
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdImprimir_Click()
   Dim wSoc As Integer, wMes As String, wmmm As String, waaa As String
   wSoc = Val(txtCodSocio.Text)
   wMes = txtMesCierre.Text
   wmmm = Right(wMes, 2)
   waaa = Left(wMes, 4)
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\NuevoEstado.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= 'AL MES DE " + Trim(funnommes(wmmm)) + " DEL " + waaa + "' "
   Crys1.SelectionFormula = " {TMP_CTASXCAB.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cndOtro_Click()
   Limpiar
End Sub

Private Sub Form_Activate()
   frmAporteEstadoCta.Left = (Screen.Width - Width) \ 2
   frmAporteEstadoCta.Top = 0
   
   txtMesCierre.Text = Left(zMesTope, 4) + "/" + Right(zMesTope, 2)
   txtMesCierre.Enabled = False
   
   
   Call Limpiar
   
   txtCodSocio.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_CTASXCAB WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim aa As Integer, wSoc As Long, WE_S As String, _
       wMesCierre As String, wFec As Date, _
       wMes As String, wAno As String, _
       wCargos As Currency, wAbonos As Currency, _
       zCargos As Currency, zAbonos As Currency, _
       wMon As String, _
       wFecPag1 As Date, wImpPag1 As Currency, wTipPag1 As String, _
       wFecPag2 As Date, wImpPag2 As Currency, wTipPag2 As String, _
       wFecPag3 As Date, wImpPag3 As Currency, wTipPag3 As String, _
       wFecPag4 As Date, wImpPag4 As Currency, wTipPag4 As String

   wSoc = Val(txtCodSocio.Text)
   wMesCierre = txtMesCierre.Text
   wAno = Left(wMesCierre, 4)
   wMes = Right(wMesCierre, 2)
   wFec = Format(fundiames(wMes) + "/" + wMes + "/" + wAno, "dd/mm/yyyy")
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_CTASXCAB WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_CTASXCAB " _
   & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, " _
   & "  CARGOS, ABONOS, SDONEW, " _
   & "  FECPAG1, IMPPAG1, " _
   & "  FECPAG2, IMPPAG2, " _
   & "  FECPAG3, IMPPAG3, USU ) " _
   & " SELECT " _
   & " CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, " _
   & " 0, 0, 0, " _
   & " NULL, 0, NULL, 0, NULL, 0, '" + wcodusu + "' " _
   & " FROM CTASXCAB " _
   & " WHERE CODSOCIO = " + Str(wSoc) + " ")
   Db.CommitTrans
      
   aa = Leerado8("SELECT * FROM CTASXDET " _
                & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                & "       FECHA <= '" + Format(wFec, "dd/mm/yyyy") + "' " _
                & " ORDER BY MES, FECHA ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wMes = ADO8!mes
         wCargos = 0: wAbonos = 0
         wImpPag1 = 0: wImpPag2 = 0: wImpPag3 = 0: wImpPag4 = 0
         
         Select Case ADO8!tipcob
         Case "00"
              If ADO8!cargos > 0 Then
                 wCargos = ADO8!cargos
              Else
                 wAbonos = ADO8!abonos
              End If
         Case "01"
              wImpPag1 = ADO8!cargos
              wAbonos = ADO8!abonos
         Case "02"
              wImpPag2 = ADO8!abonos
              wAbonos = ADO8!abonos
         Case "03"
              wImpPag3 = ADO8!abonos
              wAbonos = ADO8!abonos
         Case "04"
              If ADO8!cargos > 0 Then
                 wImpPag4 = -ADO8!cargos
                 wCargos = ADO8!cargos
              Else
                 wImpPag4 = ADO8!abonos
                 wAbonos = ADO8!abonos
              End If
         End Select
         
         WE_S = "": wMon = ""
         aa = Leerado7("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
         If aa > 0 Then
            WE_S = ADO7!e_socio
         End If
         Set ADO7 = Nothing
         
         aa = Leerado7("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + WE_S + "' ")
         If aa > 0 Then
            wMon = ADO7!moneda
         End If
         Set ADO7 = Nothing
         
         aa = Leerado7("SELECT * FROM TMP_CTASXCAB " _
                    & " WHERE      USU = '" + wcodusu + "' AND " _
                    & "       CODSOCIO = " + Str(wSoc) + " AND " _
                    & "            MES = '" + wMes + "' AND " _
                    & "       CONCEPTO = '01'  ")
         If aa = 0 Then
'            Db.BeginTrans
'            Db.Execute ("INSERT INTO TMP_CTASXCAB " _
'            & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, " _
'            & "  CARGOS, ABONOS, SDONEW, " _
'            & "  IMPPAG1, IMPPAG2, IMPPAG3, IMPPAG4, " _
'            & "   USU ) " _
'            & " VALUES " _
'            & " (" + Str(wSoc) + ", '" + wMes + "', '01', " _
'            & "  '" + WE_S + "', '" + wMon + "', " + Str(wCargo) + ", " + Str(wAbono) + ", " _
'            & "  " + Str(wCargo - wAbono) + ", " _
'            & "  " + Str(wImpPag1) + ", " + Str(wImpPag2) + ", " _
'            & "  " + Str(wImpPag3) + ", " + Str(wImpPag4) + ", " _
'            & "  '" + wcodusu + "') ")
'            Db.CommitTrans
         Else
            Db.BeginTrans
            Db.Execute ("UPDATE TMP_CTASXCAB " _
            & " SET CARGOS = CARGOS + " + Str(wCargos) + ", " _
            & "     ABONOS = ABONOS + " + Str(wAbonos) + ", " _
            & "     SDONEW = SDONEW + " + Str(wCargos) + " - " + Str(wAbonos) + ", " _
            & "     IMPPAG1 = IMPPAG1 + " + Str(wImpPag1) + ", " _
            & "     IMPPAG2 = IMPPAG2 + " + Str(wImpPag2) + ", " _
            & "     IMPPAG3 = IMPPAG3 + " + Str(wImpPag3) + ", " _
            & "     IMPPAG4 = IMPPAG4 + " + Str(wImpPag4) + " " _
            & " WHERE      USU = '" + wcodusu + "' AND " _
            & "       CODSOCIO = " + Str(wSoc) + " AND " _
            & "            MES = '" + wMes + "' AND " _
            & "       CONCEPTO = '01' ")
            Db.CommitTrans
         End If
         
         If IsDate(ADO8!fecha) Then
            Select Case ADO8!tipcob
            Case "01"
                 If IsDate(ADO8!fecha) Then
                    Db.BeginTrans
                    Db.Execute ("UPDATE TMP_CTASXCAB " _
                    & " SET FECPAG1 = '" + Format(ADO8!fecha, "dd/mm/yyyy") + "' " _
                    & " WHERE      USU = '" + wcodusu + "' AND " _
                    & "       CODSOCIO = " + Str(wSoc) + " AND " _
                    & "            MES = '" + wMes + "' AND " _
                    & "       CONCEPTO = '01' ")
                    Db.CommitTrans
                 End If
            Case "02"
                 If IsDate(ADO8!fecha) Then
                    Db.BeginTrans
                    Db.Execute ("UPDATE TMP_CTASXCAB " _
                    & " SET FECPAG2 = '" + Format(ADO8!fecha, "dd/mm/yyyy") + "' " _
                    & " WHERE      USU = '" + wcodusu + "' AND " _
                    & "       CODSOCIO = " + Str(wSoc) + " AND " _
                    & "            MES = '" + wMes + "' AND " _
                    & "       CONCEPTO = '01' ")
                    Db.CommitTrans
                 End If
            Case "03"
                 If IsDate(ADO8!fecha) Then
                    Db.BeginTrans
                    Db.Execute ("UPDATE TMP_CTASXCAB " _
                    & " SET FECPAG3 = '" + Format(ADO8!fecha, "dd/mm/yyyy") + "' " _
                    & " WHERE      USU = '" + wcodusu + "' AND " _
                    & "       CODSOCIO = " + Str(wSoc) + " AND " _
                    & "            MES = '" + wMes + "' AND " _
                    & "       CONCEPTO = '01' ")
                    Db.CommitTrans
                 End If
            Case "04"
                 If IsDate(ADO8!fecha) Then
                    Db.BeginTrans
                    Db.Execute ("UPDATE TMP_CTASXCAB " _
                    & " SET FECPAG4 = '" + Format(ADO8!fecha, "dd/mm/yyyy") + "' " _
                    & " WHERE      USU = '" + wcodusu + "' AND " _
                    & "       CODSOCIO = " + Str(wSoc) + " AND " _
                    & "            MES = '" + wMes + "' AND " _
                    & "       CONCEPTO = '01' ")
                    Db.CommitTrans
                 End If
            End Select
         End If
         
         ADO8.MoveNext
      Loop
   
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_CTASXCAB " _
   & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
   & "       CONCEPTO = '01' AND " _
   & "            USU = '" + wcodusu + "' AND " _
   & "         CARGOS = 0 AND " _
   & "         ABONOS = 0 AND " _
   & "         SDONEW = 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_CTASXCAB " _
   & " SET NOMCON = M.NOMBRE, " _
   & "     SDONEW = CARGOS - ABONOS " _
   & " FROM TMP_CTASXCAB AS T INNER JOIN MAECONCEPTO AS M " _
   & "   ON T.CONCEPTO = M.CONCEPTO " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   aa = Leerado2("SELECT CONCEPTO, NOMCON, MES, MONEDA, CARGOS, ABONOS, SDONEW, " _
                & "      IMPPAG1, IMPPAG2, IMPPAG3, FECPAG3, IMPPAG4, FECPAG4  " _
                & " FROM TMP_CTASXCAB " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY MES ")
   Set DataGrid1.DataSource = ADO2

End Sub

Private Sub LlenaCab1()
   DataGrid1.Columns(0).Width = 400
   DataGrid1.Columns(0).Alignment = dbgCenter
   DataGrid1.Columns(0).Caption = "CONCEPTO"

   DataGrid1.Columns(1).Width = 2600
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Caption = "NOMBRE CONCEPTO"

   DataGrid1.Columns(2).Width = 800
   DataGrid1.Columns(2).Alignment = dbgCenter
   DataGrid1.Columns(2).Caption = "MES"

   DataGrid1.Columns(3).Width = 500
   DataGrid1.Columns(3).Alignment = dbgCenter
   DataGrid1.Columns(3).Caption = "MON"

   DataGrid1.Columns(4).Width = 900
   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).Caption = "PROVIS"
   DataGrid1.Columns(4).NumberFormat = "#####0.00"

   DataGrid1.Columns(5).Width = 900
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).Caption = "COBROS"
   DataGrid1.Columns(5).NumberFormat = "#####0.00"

   DataGrid1.Columns(6).Width = 900
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).Caption = "SALDOS"
   DataGrid1.Columns(6).NumberFormat = "#####0.00"

   DataGrid1.Columns(7).Width = 900
   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).Caption = "DIECO"
   DataGrid1.Columns(7).NumberFormat = "###0.00;;\ "

   DataGrid1.Columns(8).Width = 900
   DataGrid1.Columns(8).Alignment = dbgRight
   DataGrid1.Columns(8).Caption = "CAJA MP"
   DataGrid1.Columns(8).NumberFormat = "###0.00;;\ "

   DataGrid1.Columns(9).Width = 900
   DataGrid1.Columns(9).Alignment = dbgRight
   DataGrid1.Columns(9).Caption = "TESORERIA"
   DataGrid1.Columns(9).NumberFormat = "###0.00;;\ "
   
   DataGrid1.Columns(10).Width = 1030
   DataGrid1.Columns(10).Alignment = dbgCenter
   DataGrid1.Columns(10).Caption = "FEC.TESOR"
   DataGrid1.Columns(10).NumberFormat = "dd/mm/yyyy"
   
   DataGrid1.Columns(11).Width = 900
   DataGrid1.Columns(11).Alignment = dbgRight
   DataGrid1.Columns(11).Caption = "OTROS"
   DataGrid1.Columns(11).NumberFormat = "###0.00;;\ "

   DataGrid1.Columns(12).Width = 1030
   DataGrid1.Columns(12).Alignment = dbgCenter
   DataGrid1.Columns(12).Caption = "FEC.OTROS"
   DataGrid1.Columns(12).NumberFormat = "dd/mm/yyyy"
End Sub

Private Sub TotalCab()
   Dim aa As Integer, wSoc As Integer, wTot As Currency
   wSoc = Val(txtCodSocio.Text)
   wTot = 0
   aa = Leerado8("SELECT SUM(SDONEW) AS SDONEW " _
                & " FROM TMP_CTASXCAB " _
                & " WHERE USU = '" + wcodusu + "' ")
   If aa > 0 Then
      wTot = IIf(IsNull(ADO8!sdonew), 0, ADO8!sdonew)
   End If
   Set ADO8 = Nothing

   lblTotal.Caption = Format(wTot, "#####0,00;;\  ")
End Sub

Private Sub txtCodSocio_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
   If aa > 0 Then
      lblCodSocio.Caption = ADO6a!nombre
   Else
      lblCodSocio.Caption = ""
      Limpiar
   End If
   Set ADO6a = Nothing
End Sub

Private Sub txtCodSocio_GotFocus()
   txtCodSocio.SelStart = 0
   If Len(Trim(txtCodSocio.Text)) > 0 Then
      txtCodSocio.SelLength = Len(Trim(txtCodSocio.Text))
   Else
      txtCodSocio.SelLength = 8
   End If
End Sub

Private Sub txtCodSocio_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim KeyNew As Integer
   KeyNew = Asc(UCase(Chr(KeyCode)))
   
   Select Case KeyCode
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
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtCodSocio.Text)) = 0 Then
         MsgBox "Codigo Socio En Blanco", vbExclamation
         txtCodSocio.Text = ""
         Exit Sub
      End If
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
      If aa = 0 Then
         MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
         txtCodSocio.Text = ""
         Exit Sub
      End If
      txtCodigo.Text = ADO8!codigo
      txtIns.Text = ADO8!ins
      txtNumdoc.Text = ADO8!numdoc
      txtE_socio.Text = ADO8!e_socio
      txtGrado.Text = ADO8!grado
      txtTipCob.Text = ADO8!tipcob
   
      LlenaCab
      LlenaCab1
   
      DataGrid1.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
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

Private Sub txtMesCierre_GotFocus()
   txtMesCierre.SelStart = 0
   txtMesCierre.SelLength = 7
End Sub

Private Sub txtMesCierre_KeyPress(KeyAscii As Integer)
   Dim wAno As String, wMes As String
   If KeyAscii = 13 Then
      If txtMesCierre.Text = "____/__" Then
         MsgBox "Mes de Cierre En Blanco", vbExclamation
         txtMesCierre.Text = "____/__"
         Exit Sub
      End If
      wAno = Mid(txtMesCierre.Text, 1, 4)
      wMes = Mid(txtMesCierre.Text, 6, 2)
      
      If wMes <> "01" And wMes <> "02" And wMes <> "03" And wMes <> "04" And _
         wMes <> "05" And wMes <> "06" And wMes <> "07" And wMes <> "08" And _
         wMes <> "09" And wMes <> "10" And wMes <> "11" And wMes <> "12" Then
         MsgBox "Mes Digitado es Invalido", vbExclamation
         txtMesCierre.Text = "____/__"
         Exit Sub
      End If
      If wAno < "2016" Or wAno > "2030" Then
         MsgBox "Año Digitado Fuera de Rango", vbExclamation
         txtMesCierre.Text = "____/__"
         Exit Sub
      End If
      
      txtCodSocio.SetFocus
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

