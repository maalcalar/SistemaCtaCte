VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmResumenxAno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuadro Resumen x Año"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   8985
   Begin VB.TextBox txtAnoCab 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   480
      Width           =   615
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
      Left            =   6120
      TabIndex        =   6
      Top             =   7920
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
      Left            =   4800
      TabIndex        =   5
      Top             =   7920
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
      Left            =   7440
      TabIndex        =   4
      Top             =   7920
      Width           =   1095
   End
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmConResumenxAno.frx":0000
      Left            =   1320
      List            =   "frmConResumenxAno.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   7335
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
      Left            =   3600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   1095
   End
   Begin Crystal.CrystalReport Crys1 
      Left            =   6480
      Top             =   600
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   11033
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
      Caption         =   "CUADRO RESUMEN POR AÑO"
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
   Begin VB.Label Label25 
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
      Left            =   720
      TabIndex        =   9
      Top             =   480
      Width           =   495
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
      Left            =   240
      TabIndex        =   7
      Top             =   8040
      Width           =   4215
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
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmResumenxAno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
   LlenaCab
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(6) As String, wReg As Integer, wTot As Integer
   Dim wDeuda0 As Integer, wDeuda3 As Integer, wDeuda6 As Integer, wDeuda7 As Integer, wTotDeu As Integer, _
       wMes As String, wAno As String
       
   wAno = Left(txtMoroso.Text, 4)
   wMes = Right(txtMoroso.Text, 2)
       
   Heading(0) = "TIPO"
   Heading(1) = "NOMBRE"
   Heading(2) = "SIN DEUDA"
   Heading(3) = "DEUDA 3 MESES"
   Heading(4) = "DEUDA 6 MESES"
   Heading(5) = "MAYOR 6 MESES"
   Heading(6) = "TOTAL"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 7)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 7)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "CUADRO RESUMEN DE APORTACIONES POR TIPO SOCIO - MES " + Trim(funnommes(wMes)) + " " + wAno
        For I = 1 To 7 Step 1
            .Cells(3, I) = Heading(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 10
        objExcel.Columns("B").ColumnWidth = 20
        objExcel.Columns("C").ColumnWidth = 12
        objExcel.Columns("D").ColumnWidth = 12
        objExcel.Columns("E").ColumnWidth = 12
        objExcel.Columns("F").ColumnWidth = 12
        objExcel.Columns("G").ColumnWidth = 12
   End With
   
   aa = Leerado3("SELECT * FROM TMP_RESXANO WHERE USU = '" + wcodusu + "' ORDER BY ORDEN ")
   If aa > 0 Then
      wTot = aa
      V = 4
      H = 1
      wNum1 = 1
      wDeuda0 = 0: wDeuda3 = 0: wDeuda6 = 0: wDeuda7 = 0: wTotDeu = 0
      Do While Not ADO3.EOF
         
         objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 6)).NumberFormat = "##,##0;;\ "
         
         objExcel.Cells(V, H + 0) = ADO3!e_socio
         objExcel.Cells(V, H + 1) = ADO3!e_socio
         objExcel.Cells(V, H + 2) = ADO3!deuda0
         objExcel.Cells(V, H + 3) = ADO3!deuda3
         objExcel.Cells(V, H + 4) = ADO3!deuda6
         objExcel.Cells(V, H + 5) = ADO3!deuda7
         objExcel.Cells(V, H + 6) = ADO3!totdeu
            
         wDeuda0 = wDeuda0 + ADO3!deuda0
         wDeuda3 = wDeuda3 + ADO3!deuda3
         wDeuda6 = wDeuda6 + ADO3!deuda6
         wDeuda7 = wDeuda7 + ADO3!deuda7
         wTotDeu = wTotDeu + ADO3!totdeu
         
         V = V + 1
         ADO3.MoveNext
      Loop
      V = V + 1
      
      objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 6)).NumberFormat = "##,##0;;\ "
      objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 6)).Font.Bold = True
      objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 6)).Font.Color = RGB(255, 0, 0)
      objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 6)).Borders.LineStyle = xlContinuous
      objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 6)).Borders.Color = RGB(255, 0, 0)
      
      objExcel.Cells(V, H + 1) = "TOTALES FINALES"
      objExcel.Cells(V, H + 2) = wDeuda0
      objExcel.Cells(V, H + 3) = wDeuda3
      objExcel.Cells(V, H + 4) = wDeuda6
      objExcel.Cells(V, H + 5) = wDeuda7
      objExcel.Cells(V, H + 6) = wTotDeu
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
   Dim wFec As Date, wMes As String, wAno As String
   wAno = Left(txtMoroso.Text, 4)
   wMes = Trim(funnommes(Right(txtMoroso.Text, 2))) + " DEL " + wAno
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\CuadroResumen.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= 'AL MES DE " + wMes + "' "
   Crys1.SelectionFormula = " {TMP_RESXANO.USU}='" + wcodusu + "' "
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
   
   txtAnoCab.Text = wanocia
   
   Set DataGrid1.DataSource = Nothing
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESXANO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   txtAnoCab.SetFocus
End Sub

Private Sub Form_Load()
   frmResumenxAno.Left = (Screen.Width - Width) \ 2
   frmResumenxAno.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESXANO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim aa As Long, wReg As Long, wTot As Long, _
       wMes As String, wAno As String, _
       wSoc As Integer, wSdo As Currency, wNom As String, wApo As Currency, wFac As Integer, _
       wDeu0 As Integer, wDeu3 As Integer, wDeu6 As Integer, wDeu7 As Integer

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESXANO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESXANO " _
   & " (E_SOCIO, ORDEN, NOMBRE, USU) " _
   & " SELECT " _
   & "  E_SOCIO, ORDEN, NOMBRE, '" + wcodusu + "' " _
   & " FROM MAEE_SOCIO " _
   & " WHERE ORDEN > 0 ")
   Db.CommitTrans
   
'   Db.BeginTrans
'   Db.Execute ("")
'   Db.CommitTrans
   
   
   aa = Leerado8a("SELECT S.CODSOCIO, S.CODIGO, S.INS, S.NOMBRE, S.E_SOCIO, E.APORTE " _
                & " FROM MAESOCIO AS S INNER JOIN MAEE_SOCIO AS E " _
                & "   ON S.E_SOCIO = E.E_SOCIO " _
                & " WHERE E.APORTE > 0 " _
                & " ORDER BY NOMBRE ")
   If aa > 0 Then
      ADO8a.MoveFirst
      wReg = 1
      wTot = aa
      Do While Not ADO8a.EOF
         DoEvents
         lblMensaje.Caption = Trim(Format(wReg, "###,##0")) + " / " + _
                              Trim(Format(wTot, "###,##0"))
         lblMensaje.Refresh
         
         wSoc = ADO8a!codsocio
         wNom = Trim(ADO8a!nombre)
         WE_S = ADO8a!e_socio
         wSdo = SaldoFoto(wSoc, wMes)
         wApo = ADO8a!aporte
         wDeu0 = 0: wDeu3 = 0: wDeu6 = 0: wDeu7 = 0
         wFac = Round(wSdo / wApo, 0)
          
         Select Case wFac
         Case 0
              wDeu0 = 1
         Case 1, 2, 3
              wDeu3 = 1
         Case 4, 5, 6
              wDeu6 = 1
         Case Is > 6
              wDeu7 = 1
         End Select
          
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_RESXANO " _
         & " SET DEUDA0 = DEUDA0 + " + Str(wDeu0) + ", " _
         & "     DEUDA3 = DEUDA3 + " + Str(wDeu3) + ", " _
         & "     DEUDA6 = DEUDA6 + " + Str(wDeu6) + ", " _
         & "     DEUDA7 = DEUDA7 + " + Str(wDeu7) + " " _
         & " WHERE E_SOCIO = '" + WE_S + "' AND " _
         & "           USU = '" + wcodusu + "' ")
         Db.CommitTrans
   
         wReg = wReg + 1
         ADO8a.MoveNext
      Loop
   End If
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_RESXANO " _
   & " SET TOTDEU = DEUDA0 + DEUDA3 + DEUDA6 + DEUDA7 " _
   & " WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   DoEvents
   lblMensaje.Caption = ""
   lblMensaje.Refresh
   
   aa = Leerado2("SELECT E_SOCIO, NOMBRE, DEUDA0, DEUDA3, DEUDA6, DEUDA7, TOTDEU " _
            & " FROM TMP_RESXANO " _
            & " WHERE USU = '" + wcodusu + "' " _
            & " ORDER BY ORDEN ")
   Set DataGrid1.DataSource = ADO2
 
   DataGrid1.Columns(0).Width = 600   ' CODSOCIO
   DataGrid1.Columns(0).Alignment = dbgLeft
   DataGrid1.Columns(0).Caption = "E.SOCIO"
    
   DataGrid1.Columns(1).Width = 2500   ' NOMBRE
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Caption = "NOMBRE"
    
   DataGrid1.Columns(2).Width = 1000  ' DEUDA0
   DataGrid1.Columns(2).Alignment = dbgRight
   DataGrid1.Columns(2).Caption = "SIN DEUDA"
   DataGrid1.Columns(2).NumberFormat = "#####0;;\ "

   DataGrid1.Columns(3).Width = 1000  ' DEUDA3
   DataGrid1.Columns(3).Alignment = dbgRight
   DataGrid1.Columns(3).Caption = "3 MESES"
   DataGrid1.Columns(3).NumberFormat = "#####0;;\ "

   DataGrid1.Columns(4).Width = 1000  ' DEUDA6
   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).Caption = "6 MESES"
   DataGrid1.Columns(4).NumberFormat = "#####0;;\ "

   DataGrid1.Columns(5).Width = 1000  ' DEUDA7
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).Caption = "> 6 MESES"
   DataGrid1.Columns(5).NumberFormat = "#####0;;\ "

   DataGrid1.Columns(6).Width = 1000  ' TOTDEU
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).Caption = "TOTALES"
   DataGrid1.Columns(6).NumberFormat = "#####0;;\ "
End Sub

