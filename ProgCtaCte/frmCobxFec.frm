VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCobxFec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Cobranzas Por Rango de Fechas "
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   13845
   Begin VB.TextBox txtAno 
      Enabled         =   0   'False
      Height          =   305
      Left            =   8520
      TabIndex        =   13
      Top             =   200
      Width           =   615
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
      Left            =   12480
      TabIndex        =   12
      Top             =   6960
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
      Left            =   9840
      TabIndex        =   11
      Top             =   6960
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
      Left            =   11160
      TabIndex        =   10
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox txtSocio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmCobxFec.frx":0000
      Left            =   1080
      List            =   "frmCobxFec.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   200
      Width           =   6615
   End
   Begin MSMask.MaskEdBox txtDesde 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   540
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtHasta 
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Top             =   540
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5535
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   9763
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
      Caption         =   "RELACION DE COBROS POR RANGO DE FECHA"
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
      Left            =   11280
      Top             =   120
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
      TabIndex        =   25
      Top             =   7560
      Width           =   8295
   End
   Begin VB.Label Label8 
      Caption         =   "Vendedor"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label lblVendedorDet 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1320
      TabIndex        =   23
      Top             =   7200
      Width           =   5535
   End
   Begin VB.Label Label7 
      Caption         =   "Tipo Cobro"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label lblTipCobDet 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1320
      TabIndex        =   21
      Top             =   6960
      Width           =   5535
   End
   Begin VB.Label Label11 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label lblClienteDet 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1320
      TabIndex        =   19
      Top             =   6720
      Width           =   5535
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
      Left            =   7080
      TabIndex        =   18
      Top             =   6720
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
      Left            =   8280
      TabIndex        =   17
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Caption         =   "Total US$"
      Height          =   255
      Left            =   7080
      TabIndex        =   16
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Total S/."
      Height          =   255
      Left            =   8280
      TabIndex        =   15
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label4 
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
      Left            =   7920
      TabIndex        =   14
      Top             =   200
      Width           =   495
   End
   Begin VB.Label lblSocio 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   840
      Width           =   6735
   End
   Begin VB.Label Label2 
      Caption         =   "Socio"
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
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Hasta"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   540
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "Desde"
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
      Top             =   540
      Width           =   975
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
      Top             =   200
      Width           =   855
   End
End
Attribute VB_Name = "frmCobxFec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(14) As String, _
       wreg As Integer, wTot As Integer, _
       wdes As String, whas As String, wFec As Date, _
       wdol As Currency, wsol As Currency, _
       zdol As Currency, zsol As Currency, wCob As String

   wdes = Format(txtDesde.Text, "dd/mm/yyyy")
   whas = Format(txtHasta.Text, "dd/mm/yyyy")
   
   Heading(0) = "SERIE"
   Heading(1) = "NUMCOB"
   Heading(2) = "FECHA"
   Heading(3) = "LIN"
   Heading(4) = "SOCIO"
   Heading(5) = "NOMBRE SOCIO"
   Heading(6) = "MES"
   Heading(7) = "CONC"
   Heading(8) = "NOMBRE CONCEPTO"
   Heading(9) = "US$"
   Heading(10) = "S/."
   Heading(11) = "MONEDA"
   Heading(12) = "SDO.ANTER"
   Heading(13) = "COBRANZA"
   Heading(14) = "SDO.NUEVO"
   wdol = 0: wsol = 0
   zdol = 0: zsol = 0
   
   aa = Leerado3("SELECT SERCOB, NUMCOB, FECHA, LINCOB, CODSOCIO, NOMBRE, MESCOB, CONPAGO, NOMCON, " _
                & "      DOLARE, SOLESS, MONDOC, SDOOLD, ABONOS, SDONEW " _
                & " FROM TMP_COBXFEC " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY SERCOB, NUMCOB, LINCOB ")
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
           .Cells(2, 1) = "DETALLE DE COBRANZAS POR FECHA - DEL " + wdes + " AL " + whas
           For I = 1 To 15 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           objExcel.Columns("A").ColumnWidth = 5
           objExcel.Columns("B").ColumnWidth = 11
           objExcel.Columns("C").ColumnWidth = 11
           objExcel.Columns("D").ColumnWidth = 5
           objExcel.Columns("E").ColumnWidth = 7
           objExcel.Columns("F").ColumnWidth = 40
           objExcel.Columns("G").ColumnWidth = 8
           objExcel.Columns("H").ColumnWidth = 4
           objExcel.Columns("I").ColumnWidth = 20
           objExcel.Columns("J").ColumnWidth = 10
           objExcel.Columns("K").ColumnWidth = 10
           objExcel.Columns("L").ColumnWidth = 5
           objExcel.Columns("M").ColumnWidth = 11
           objExcel.Columns("N").ColumnWidth = 11
           objExcel.Columns("O").ColumnWidth = 11
      End With
      V = 4
      H = 1
      wreg = 1
      Do While Not ADO3.EOF
         wCob = ADO3!sercob + "-" + ADO3!numcob
         wdol = 0: wsol = 0
         Do While ADO3!sercob + "-" + ADO3!numcob = wCob
            DoEvents
            lblMensaje.Caption = "Traslando Cobranzas a EXCEL - Registro " + _
                                 Format(wreg, "####0") + " / " + _
                                 Format(wTot, "####0")
            lblMensaje.Refresh
         
            objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 10)).NumberFormat = "######0.00;;\ "
            objExcel.Range(objExcel.Cells(V, H + 12), objExcel.Cells(V, H + 14)).NumberFormat = "######0.00"
            
            wFec = Format(ADO3!fecha, "dd/mm/yyyy")
         
            objExcel.Cells(V, H + 0) = IIf(IsNull(ADO3!sercob), "", ADO3!sercob)
            objExcel.Cells(V, H + 1) = IIf(IsNull(ADO3!numcob), "", ADO3!numcob)
            objExcel.Cells(V, H + 2) = wFec
            objExcel.Cells(V, H + 3) = IIf(IsNull(ADO3!lincob), "", ADO3!lincob)
            objExcel.Cells(V, H + 4) = ADO3!codsocio
            objExcel.Cells(V, H + 5) = ADO3!nombre
            objExcel.Cells(V, H + 6) = ADO3!mescob
            objExcel.Cells(V, H + 7) = ADO3!conpago
            objExcel.Cells(V, H + 8) = ADO3!nomcon
            objExcel.Cells(V, H + 9) = ADO3!dolare
            objExcel.Cells(V, H + 10) = ADO3!soless
            objExcel.Cells(V, H + 11) = ADO3!mondoc
            
            objExcel.Cells(V, H + 12) = ADO3!sdoold
            objExcel.Cells(V, H + 13) = ADO3!abonos
            objExcel.Cells(V, H + 14) = ADO3!sdonew
         
            wdol = wdol + ADO3!dolare
            wsol = wsol + ADO3!soless
         
            zdol = zdol + ADO3!dolare
            zsol = zsol + ADO3!soless
         
            wreg = wreg + 1
            V = V + 1
            ADO3.MoveNext
            If ADO3.EOF Then
               Exit Do
            End If
         Loop
         
         objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 10)).NumberFormat = "######0.00"
         objExcel.Cells(V, H + 8) = "TOTALES X NUMCOB " + wCob
         objExcel.Cells(V, H + 9) = wdol
         objExcel.Cells(V, H + 10) = wsol
         V = V + 2
      Loop
      objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 10)).NumberFormat = "######0.00"
      objExcel.Cells(V, H + 8) = "TOTALES FINALES"
      objExcel.Cells(V, H + 9) = zdol
      objExcel.Cells(V, H + 10) = zsol
      V = V + 1
      
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
   Dim wdes As String, whas As String
   wdes = Format(txtDesde.Text, "dd/mm/yyyy")
   whas = Format(txtHasta.Text, "dd/mm/yyyy")
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\CobxFec.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= 'DEL " + wdes + " AL " + whas + "' "
   Crys1.SelectionFormula = " {TMP_COBXFEC.USU}='" + wcodusu + "' "
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
   txtAno.Text = wanocia
   
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
   
   txtDesde.SetFocus
End Sub

Private Sub Form_Load()
   frmCobxFec.Left = (Screen.Width - Width) \ 2
   frmCobxFec.Top = 0
End Sub

Private Sub LlenaCab()
   Dim aa As Long, wdes As String, whas As String, _
       wSoc As String
   wdes = Format(txtDesde.Text, "yyyy/mm/dd")
   whas = Format(txtHasta.Text, "yyyy/mm/dd")
   wSoc = txtSocio.Text

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_COBXFEC WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_COBXFEC " _
   & " ( BLOQUE, TIPCOB, SERCOB, NUMCOB, LINCOB, FECHA, CODSOCIO, NOMBRE, DOLARE, SOLESS, " _
   & "   MESCOB, CONCEPTO, CONPAGO, NOMCON, MONDOC, FORPAG, " _
   & "   SDOOLD, ABONOS, SDONEW, IMPORTE, NOMPAR, GLOSA, USU ) " _
   & " SELECT " _
   & "   '1', D.TIPCOB, D.SERCOB, D.NUMCOB, D.LINCOB, C.FECHA, C.CODSOCIO, '', " _
   & "   D.DOLARE, D.SOLESS, D.MESCOB, D.CONCEPTO, D.CONPAGO, '', MONDOC, " _
   & "   C.FORPAG, " _
   & "   D.SDOOLD, D.ABONOS, D.SDONEW, D.IMPORTE, D.NOMBRE, C.GLOSA, '" + wcodusu + "' " _
   & " From COBRODET AS D INNER JOIN COBROCAB AS C " _
   & "   ON D.TIPCOB = C.TIPCOB AND " _
   & "      D.SERCOB = C.SERCOB AND " _
   & "      D.NUMCOB = C.NUMCOB " _
   & " WHERE (D.ANO='" + wanocia + "') AND " _
   & "       (C.FECHA >= '" + Format(wdes, "dd/mm/yyyy") + "') AND " _
   & "       (C.FECHA <= '" + Format(whas, "dd/mm/yyyy") + "') ")
   Db.CommitTrans

   If wSoc <> "" Then
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_COBXFEC " _
      & " WHERE      USU = '" + wcodusu + "' AND " _
      & "       CODSOCIO <> '" + wSoc + "' ")
      Db.CommitTrans
   End If
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_COBXFEC " _
   & " SET BLOQUE = '1' " _
   & " WHERE USU = '" + wcodusu + "' AND " _
   & "       SERCOB = '005' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_COBXFEC " _
   & " SET BLOQUE = '2' " _
   & " WHERE USU = '" + wcodusu + "' AND " _
   & "       SERCOB = '004' AND (MONDOC = 'S' OR MONDOC = '') AND FORPAG <> '08' AND FORPAG <> '09' AND FORPAG <> '10' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_COBXFEC " _
   & " SET BLOQUE = '3' " _
   & " WHERE USU = '" + wcodusu + "' AND " _
   & "       SERCOB = '004' AND MONDOC = 'D' AND FORPAG <> '08' AND FORPAG <> '09' AND FORPAG <> '10' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_COBXFEC " _
   & " SET BLOQUE = '4' " _
   & " WHERE USU = '" + wcodusu + "' AND " _
   & "       SERCOB = '004' AND FORPAG = '08' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_COBXFEC " _
   & " SET BLOQUE = '5' " _
   & " WHERE USU = '" + wcodusu + "' AND " _
   & "       SERCOB = '004' AND FORPAG = '09' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_COBXFEC " _
   & " SET BLOQUE = '6' " _
   & " WHERE USU = '" + wcodusu + "' AND " _
   & "       SERCOB = '004' AND FORPAG = '10' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_COBXFEC " _
   & " SET NOMBRE = M.NOMBRE, CODIGO = M.CODIGO, INS = M.INS " _
   & " FROM TMP_COBXFEC AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODSOCIO = M.CODSOCIO " _
   & " WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
       
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_COBXFEC " _
   & " SET NOMCON = M.DESCONCE " _
   & " FROM TMP_COBXFEC AS T INNER JOIN ZZZ_CONCEPTO AS M " _
   & "   ON T.CONPAGO = M.CONCEPTO " _
   & " WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
       
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_COBXFEC " _
   & " SET NOMBRE = 'DOCUMENTO ANULADO' " _
   & " WHERE USU = '" + wcodusu + "' AND CODSOCIO = 0 ")
   Db.CommitTrans
       
   aa = Leerado2("SELECT TIPCOB, SERCOB, NUMCOB, LINCOB, FECHA, CODSOCIO, NOMBRE, MESCOB, " _
                & "      CONPAGO, DOLARE, SOLESS, MONDOC, " _
                & "      SDOOLD, ABONOS, SDONEW, FORPAG, GLOSA " _
                & " FROM TMP_COBXFEC " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY SERCOB, NUMCOB, LINCOB ")
   Set DataGrid1.DataSource = ADO2
End Sub

Private Sub LlenaCab1()
   DataGrid1.Columns(0).Width = 350   ' TIPCOB
   DataGrid1.Columns(0).Alignment = dbgCenter
   DataGrid1.Columns(0).Caption = "TIP"
    
   DataGrid1.Columns(1).Width = 550   ' SERCOB
   DataGrid1.Columns(1).Alignment = dbgCenter
   DataGrid1.Columns(1).Caption = "SERIE"
    
   DataGrid1.Columns(2).Width = 1050  ' NUMCOB
   DataGrid1.Columns(2).Alignment = dbgCenter
   DataGrid1.Columns(2).Caption = "NUM.COB."
   
   DataGrid1.Columns(3).Width = 350   ' LINCOB
   DataGrid1.Columns(3).Alignment = dbgCenter
   DataGrid1.Columns(3).Caption = "LIN"
   
   DataGrid1.Columns(4).Width = 1050   ' FECHA
   DataGrid1.Columns(4).Alignment = dbgCenter
   DataGrid1.Columns(4).NumberFormat = "dd/mm/yyyy"
   DataGrid1.Columns(4).Caption = "FECHA"
    
   DataGrid1.Columns(5).Width = 650   ' CODSOCIO
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).Caption = "SOCIO"
       
   DataGrid1.Columns(6).Width = 3000  ' NOMBRE
   DataGrid1.Columns(6).Alignment = dbgLeft
   DataGrid1.Columns(6).Caption = "NOMBRE"
    
   DataGrid1.Columns(7).Width = 750   ' MESCOB
   DataGrid1.Columns(7).Alignment = dbgLeft
   DataGrid1.Columns(7).Caption = "MES"
    
   DataGrid1.Columns(8).Width = 540   ' CONCEPTO
   DataGrid1.Columns(8).Alignment = dbgLeft
   DataGrid1.Columns(8).Caption = "CONC"
    
   DataGrid1.Columns(9).Width = 900     ' DOLARE'
   DataGrid1.Columns(9).Alignment = dbgRight
   DataGrid1.Columns(9).Caption = "  US$  "
   DataGrid1.Columns(9).NumberFormat = "####0.00"
   
   DataGrid1.Columns(10).Width = 900     ' SOLESS'
   DataGrid1.Columns(10).Alignment = dbgRight
   DataGrid1.Columns(10).Caption = "  S/.  "
   DataGrid1.Columns(10).NumberFormat = "####0.00"
   
   DataGrid1.Columns(11).Width = 450   ' MONDOC '
   DataGrid1.Columns(11).Alignment = dbgLeft
   DataGrid1.Columns(11).Caption = "MON"
       
   DataGrid1.Columns(12).Width = 850     ' SDOOLD'
   DataGrid1.Columns(12).Alignment = dbgRight
   DataGrid1.Columns(12).Caption = "SDO.OLD"
   DataGrid1.Columns(12).NumberFormat = "####0.00"

   DataGrid1.Columns(13).Width = 850     ' ABONOS'
   DataGrid1.Columns(13).Alignment = dbgRight
   DataGrid1.Columns(13).Caption = "COBROS"
   DataGrid1.Columns(13).NumberFormat = "####0.00"
       
   DataGrid1.Columns(14).Width = 850     ' SDONEW'
   DataGrid1.Columns(14).Alignment = dbgRight
   DataGrid1.Columns(14).Caption = "NUEVO SDO"
   DataGrid1.Columns(14).NumberFormat = "####0.00"

   DataGrid1.Columns(15).Visible = False
   DataGrid1.Columns(16).Visible = False
End Sub

Private Sub LabelCab()
   If ADO2.RecordCount > 0 Then
'      lblClienteDet.Caption = IIf(IsNull(ADO2!nomcli), "", ADO2!nomcli)
   End If
End Sub

Private Sub TotalCab()
   Dim aa As Long, wdol As Currency, wsol As Currency
   aa = Leerado8("SELECT SUM(DOLARE) AS DOLARE, SUM(SOLESS) AS SOLESS " _
                & " FROM TMP_COBXFEC " _
                & " WHERE USU = '" + wcodusu + "' ")
   If aa > 0 Then
      wdol = IIf(IsNull(ADO8!dolare), 0, ADO8!dolare)
      wsol = IIf(IsNull(ADO8!soless), 0, ADO8!soless)
   End If
   Set ADO8 = Nothing
   lblTotDol.Caption = Format(wdol, "###,##0.00")
   lblTotSol.Caption = Format(wsol, "###,##0.00")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_COBXFEC WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub txtDesde_GotFocus()
   txtDesde.SelStart = 0
   txtDesde.SelLength = 10
End Sub

Private Sub txtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 40
        txtHasta.SetFocus
   End Select
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtDesde.Text = "__/__/____" Then
         MsgBox "Fecha Inicial En Blanco", vbExclamation
         txtDesde.Text = "__/__/____"
         Exit Sub
      End If
      If Not IsDate(txtDesde.Text) Then
         MsgBox "Fecha Inicial Digitado Es Invalido", vbExclamation
         txtDesde.Text = "__/__/____"
         Exit Sub
      End If
      txtHasta.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii), 1) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtHasta_GotFocus()
   txtHasta.SelStart = 0
   txtHasta.SelLength = 10
End Sub

Private Sub txtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDesde.SetFocus
   Case 40
        txtSocio.SetFocus
   End Select
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
   Dim wdes As String, whas As String
   If KeyAscii = 13 Then
      If txtHasta.Text = "__/__/____" Then
         MsgBox "Fecha Final En Blanco", vbExclamation
         txtHasta.Text = "__/__/____"
         Exit Sub
      End If
      If Not IsDate(txtHasta.Text) Then
         MsgBox "Fecha Final Digitado Es Invalido", vbExclamation
         txtHasta.Text = "__/__/____"
         Exit Sub
      End If
      wdes = Format(txtDesde.Text, "yyyy/mm/dd")
      whas = Format(txtHasta.Text, "yyyy/mm/dd")
      If wdes > whas Then
         MsgBox "Rango de Fechas Digitado Es Invalido", vbExclamation
         txtHasta.Text = "__/__/____"
         Exit Sub
      End If
      txtSocio.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii), 1) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtSocio_Change()
   Dim aa As Integer
   aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio.Text)) + " ")
   If aa > 0 Then
      lblSocio.Caption = ADO8!nombre
   Else
      lblSocio.Caption = ""
   End If
   Set ADO8 = Nothing
End Sub

Private Sub txtSocio_GotFocus()
   txtSocio.SelStart = 0
   txtSocio.SelLength = Len(Trim(txtSocio.Text))
End Sub

Private Sub txtSocio_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtHasta.SetFocus
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtSocio.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtSocio_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtSocio.Text)) <> 0 Then
         aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio.Text)) + " ")
         If aa = 0 Then
            MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
            txtSocio.Text = ""
            Exit Sub
         End If
      End If
      LlenaCab
      LlenaCab1
      LabelCab
      TotalCab
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub
