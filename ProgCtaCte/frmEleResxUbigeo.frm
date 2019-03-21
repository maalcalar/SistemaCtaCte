VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmEleResxUbigeo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen x UBIGEO"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   13965
   Begin VB.OptionButton optTres 
      Caption         =   "Socios con Tres Meses de Deuda"
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
      Left            =   1200
      TabIndex        =   10
      Top             =   840
      Width           =   4335
   End
   Begin VB.OptionButton optDos 
      Caption         =   "Socios con Dos Meses de Deuda"
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
      Left            =   1200
      TabIndex        =   9
      Top             =   600
      Width           =   4335
   End
   Begin VB.OptionButton optUno 
      Caption         =   "Socios con Un Mes de Deuda"
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
      Left            =   1200
      TabIndex        =   8
      Top             =   360
      Width           =   4335
   End
   Begin VB.OptionButton optTodos 
      Caption         =   "Todos Los Socios"
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
      Left            =   1200
      TabIndex        =   7
      Top             =   120
      Value           =   -1  'True
      Width           =   4335
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
      Left            =   5760
      TabIndex        =   4
      Top             =   240
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
      Left            =   7560
      TabIndex        =   2
      Top             =   6480
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
      Left            =   8760
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6480
      Width           =   975
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
      Left            =   9960
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6480
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4815
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   8493
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
      Caption         =   "RESUMEN POR TIPO DE SOCIO"
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
      Left            =   12840
      Top             =   1080
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
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   11400
      TabIndex        =   6
      Top             =   6000
      Width           =   735
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
      TabIndex        =   5
      Top             =   6480
      Width           =   6615
   End
End
Attribute VB_Name = "frmEleResxUbigeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
   LlenaCab
   LlenaCab1
End Sub

Private Sub cmdExporta_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(15) As String, wreg As Integer
   Dim wTit As Integer, wHij As Integer, wHer As Integer, wNie As Integer, wCiv As Integer, _
       wCi1 As Integer, wTra As Integer, wAdh As Integer, wViu As Integer, wTot As Integer, wPnp As Integer, wHon As Integer
   Dim zTit As Integer, zHij As Integer, zHer As Integer, zNie As Integer, zCiv As Integer, _
       zCi1 As Integer, zTra As Integer, zAdh As Integer, zViu As Integer, zTot As Integer, zPnp As Integer, zHon As Integer, _
       wDist As String, wNomDist As String, _
       wProv As String, wNomProv As String

   Dim wNom As String
   Heading(0) = "PROV"
   Heading(1) = "NOMBRE PROVINCIA"
   Heading(2) = "DIST"
   Heading(3) = "NOMBRE DISTRITO"
   Heading(4) = "TIT"
   Heading(5) = "HIJ"
   Heading(6) = "HER"
   Heading(7) = "NIE"
   Heading(8) = "CIV"
   Heading(9) = "CI1"
   Heading(10) = "TRA"
   Heading(11) = "ADH"
   Heading(12) = "VIU"
   Heading(13) = "PNP"
   Heading(14) = "HON"
   Heading(15) = "TOT"
   aa = Leerado3("SELECT * FROM TMP_CUAXUBIGEO WHERE USU = '" + wcodusu + "' ORDER BY PROV, DIST ")
   If aa > 0 Then
      wTot = aa
      Set objExcel = New Excel.Application
      objExcel.SheetsInNewWorkbook = 1
      objExcel.Workbooks.Add
      With objExcel
           .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
           .Range(.Cells(3, 1), .Cells(3, 16)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 16)).Font.Bold = True
           .Cells(1, 1) = wnomcia
           .Cells(2, 1) = "MAESTRO DE REGIONES"
           For I = 1 To 16 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           objExcel.Columns("A").ColumnWidth = 8
           objExcel.Columns("B").ColumnWidth = 30
           objExcel.Columns("C").ColumnWidth = 8
           objExcel.Columns("D").ColumnWidth = 30
           objExcel.Columns("E").ColumnWidth = 9
           objExcel.Columns("F").ColumnWidth = 9
           objExcel.Columns("G").ColumnWidth = 9
           objExcel.Columns("H").ColumnWidth = 9
           objExcel.Columns("I").ColumnWidth = 9
           objExcel.Columns("J").ColumnWidth = 9
           objExcel.Columns("K").ColumnWidth = 9
           objExcel.Columns("L").ColumnWidth = 9
           objExcel.Columns("M").ColumnWidth = 9
           objExcel.Columns("N").ColumnWidth = 9
           objExcel.Columns("O").ColumnWidth = 9
           objExcel.Columns("P").ColumnWidth = 9
      End With
      V = 4
      H = 1
      wreg = 1
      Do While Not ADO3.EOF
         wDist = ADO3!dist
         wNomDist = ADO3!nomdist
         wProv = ADO3!prov
         wNomProv = ADO3!nomprov
         wTit = 0: wHij = 0: wHer = 0: wNie = 0: wCiv = 0: wCi1 = 0: wTra = 0: wAdh = 0: wViu = 0: wPnp = 0: wHon = 0: wTot = 0
         
         Do While ADO3!prov = wProv
            lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wreg, "####0") + " / " + Format(wTot, "####0")
            lblMensaje.Refresh
         
            objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 15)).NumberFormat = "####0;;\ "
            
            objExcel.Cells(V, H + 0) = ADO3!prov
            objExcel.Cells(V, H + 1) = ADO3!nomprov
            objExcel.Cells(V, H + 2) = ADO3!dist
            objExcel.Cells(V, H + 3) = ADO3!nomdist
            objExcel.Cells(V, H + 4) = ADO3!tit
            objExcel.Cells(V, H + 5) = ADO3!hij
            objExcel.Cells(V, H + 6) = ADO3!her
            objExcel.Cells(V, H + 7) = ADO3!nie
            objExcel.Cells(V, H + 8) = ADO3!civ
            objExcel.Cells(V, H + 9) = ADO3!ci1
            objExcel.Cells(V, H + 10) = ADO3!tra
            objExcel.Cells(V, H + 11) = ADO3!adh
            objExcel.Cells(V, H + 12) = ADO3!viu
            objExcel.Cells(V, H + 13) = ADO3!pnp
            objExcel.Cells(V, H + 14) = ADO3!hon
            objExcel.Cells(V, H + 15) = ADO3!tot
         
            wTit = wTit + ADO3!tit
            wHij = wHij + ADO3!hij
            wHer = wHer + ADO3!her
            wNie = wNie + ADO3!nie
            wCiv = wCiv + ADO3!civ
            wCi1 = wCi1 + ADO3!ci1
            wTra = wTra + ADO3!tra
            wAdh = wAdh + ADO3!adh
            wViu = wViu + ADO3!viu
            wPnp = wPnp + ADO3!pnp
            wHon = wHon + ADO3!hon
            wTot = wTot + ADO3!tot
         
            zTit = zTit + ADO3!tit
            zHij = zHij + ADO3!hij
            zHer = zHer + ADO3!her
            zNie = zNie + ADO3!nie
            zCiv = zCiv + ADO3!civ
            zCi1 = zCi1 + ADO3!ci1
            zTra = zTra + ADO3!tra
            zAdh = zAdh + ADO3!adh
            zViu = zViu + ADO3!viu
            zPnp = zPnp + ADO3!pnp
            zHon = zHon + ADO3!hon
            zTot = zTot + ADO3!tot
         
            wreg = wreg + 1
            V = V + 1
         
            ADO3.MoveNext
            If ADO3.EOF Then
               Exit Do
            End If
         Loop
         V = V + 1
         
         objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 15)).NumberFormat = "####0;;\ "
          
         objExcel.Cells(V, H + 0) = wProv
         objExcel.Cells(V, H + 1) = wNomProv
         objExcel.Cells(V, H + 2) = ""
         objExcel.Cells(V, H + 3) = ""
         objExcel.Cells(V, H + 4) = wTit
         objExcel.Cells(V, H + 5) = wHij
         objExcel.Cells(V, H + 6) = wHer
         objExcel.Cells(V, H + 7) = wNie
         objExcel.Cells(V, H + 8) = wCiv
         objExcel.Cells(V, H + 9) = wCi1
         objExcel.Cells(V, H + 10) = wTra
         objExcel.Cells(V, H + 11) = wAdh
         objExcel.Cells(V, H + 12) = wViu
         objExcel.Cells(V, H + 13) = wPnp
         objExcel.Cells(V, H + 14) = wHon
         objExcel.Cells(V, H + 15) = wTot
         
         V = V + 2
         
      Loop
      
      objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 15)).NumberFormat = "####0;;\ "
          
      objExcel.Cells(V, H + 0) = ""
      objExcel.Cells(V, H + 1) = "TOTALES FINALES"
      objExcel.Cells(V, H + 2) = ""
      objExcel.Cells(V, H + 3) = ""
      objExcel.Cells(V, H + 4) = zTit
      objExcel.Cells(V, H + 5) = zHij
      objExcel.Cells(V, H + 6) = zHer
      objExcel.Cells(V, H + 7) = zNie
      objExcel.Cells(V, H + 8) = zCiv
      objExcel.Cells(V, H + 9) = zCi1
      objExcel.Cells(V, H + 10) = zTra
      objExcel.Cells(V, H + 11) = zAdh
      objExcel.Cells(V, H + 12) = zViu
      objExcel.Cells(V, H + 13) = zPnp
      objExcel.Cells(V, H + 14) = zHon
      objExcel.Cells(V, H + 15) = zTot
      
      V = V + 2
      
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
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\ResumenxUbigeo.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.SelectionFormula = " {TMP_CUAXUBIGEO.USU}='" + wcodusu + "' "
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
   frmEleResxUbigeo.Left = (Screen.Width - Width) \ 2
   frmEleResxUbigeo.Top = 0
   
   cmdBuscar.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_CUAXUBIGEO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

End Sub

Private Sub LlenaCab()
   Dim aa As Long, wRegAct As Long, wRegTot As Long, _
       wProv As String, wDist As String, wNomProv As String, wNomDist As String, wE_S As String, _
       wCan As Integer, _
       wTot As Integer, wTit As Integer, wHij As Integer, wHer As Integer, wNie As Integer, _
       wCiv As Integer, wCi1 As Integer, wTra As Integer, wAdh As Integer, wViu As Integer, _
       wSoc As Integer, wCod As Long, wIns As Integer, wSdo As Currency, wApo As Currency, wsw As Boolean

   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_CUAXUBIGEO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   If optTodos.Value = True Then
   
   aa = Leerado8a("SELECT M.E_SOCIO, M.UBIGEO, U.NOMBRE, LEFT(M.UBIGEO,4)+'00' AS prov, U2.NOMBRE as nomprov, COUNT(M.E_SOCIO) AS CANT " _
                & " FROM MAESOCIO AS M INNER JOIN MAEE_SOCIO AS E  ON M.E_SOCIO         = E.E_SOCIO " _
                & "                    LEFT JOIN MAEUBIGEO   AS U  ON M.UBIGEO          = U.CODIGO " _
                & "                    LEFT JOIN MAEUBIGEO   AS U2 ON left(M.UBIGEO,4) + '00' = U2.CODIGO " _
                & " Where E.APORTE > 0 OR E.E_SOCIO = 'HON' " _
                & " GROUP BY M.E_SOCIO, M.UBIGEO, U.NOMBRE, LEFT(M.UBIGEO,4), U2.NOMBRE " _
                & " ORDER BY M.E_SOCIO")
   If aa > 0 Then
      ADO8a.MoveFirst
      wRegAct = 1
      wRegTot = aa
      Do While Not ADO8a.EOF
         DoEvents
         lblMensaje.Caption = "Registro " + _
                              Trim(Format(wRegAct, "##,##0")) + " / " + _
                              Trim(Format(wRegTot, "##,##0"))
         lblMensaje.Refresh
         
         wE_S = ADO8a!e_socio
         wProv = IIf(ADO8a!prov = "    00", "9999", ADO8a!prov)
         wDist = IIf(ADO8a!ubigeo = "", "999999", ADO8a!ubigeo)
         wNomProv = IIf(IsNull(ADO8a!nomprov), "SIN UBIGEO", ADO8a!nomprov)
         wNomDist = IIf(IsNull(ADO8a!nombre), "SIN UBIGEO", ADO8a!nombre)
         wCan = ADO8a!cant
         
         
         aa = Leerado7a("SELECT * FROM TMP_CUAXUBIGEO " _
                        & " WHERE DIST = '" + wDist + "' AND " _
                        & "       PROV = '" + wProv + "' AND " _
                        & "        USU = '" + wcodusu + "' ")
         If aa = 0 Then
            Db.BeginTrans
            Db.Execute ("INSERT INTO TMP_CUAXUBIGEO " _
            & " (DIST, NOMDIST, PROV, NOMPROV, USU) " _
            & " VALUES " _
            & " ('" + wDist + "', '" + wNomDist + "', " _
            & "  '" + wProv + "', '" + wNomProv + "', " _
            & "  '" + wcodusu + "') ")
            Db.CommitTrans
         End If
          
         Select Case wE_S
         Case "TIT"
              Db.BeginTrans
              Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
              & " SET TIT = " + Str(wCan) + " " _
              & " WHERE  USU = '" + wcodusu + "' AND " _
              & "       DIST = '" + wDist + "' AND " _
              & "       PROV = '" + wProv + "' ")
              Db.CommitTrans
         Case "HIJ"
              Db.BeginTrans
              Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
              & " SET HIJ = " + Str(wCan) + " " _
              & " WHERE  USU = '" + wcodusu + "' AND " _
              & "       DIST = '" + wDist + "' AND " _
              & "       PROV = '" + wProv + "' ")
              Db.CommitTrans
         Case "HER"
              Db.BeginTrans
              Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
              & " SET HER = " + Str(wCan) + " " _
              & " WHERE  USU = '" + wcodusu + "' AND " _
              & "       DIST = '" + wDist + "' AND " _
              & "       PROV = '" + wProv + "' ")
              Db.CommitTrans
         Case "NIE"
              Db.BeginTrans
              Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
              & " SET NIE = " + Str(wCan) + " " _
              & " WHERE  USU = '" + wcodusu + "' AND " _
              & "       DIST = '" + wDist + "' AND " _
              & "       PROV = '" + wProv + "' ")
              Db.CommitTrans
         Case "CIV"
              Db.BeginTrans
              Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
              & " SET CIV = " + Str(wCan) + " " _
              & " WHERE  USU = '" + wcodusu + "' AND " _
              & "       DIST = '" + wDist + "' AND " _
              & "       PROV = '" + wProv + "' ")
              Db.CommitTrans
         Case "CI1"
              Db.BeginTrans
              Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
              & " SET CI1 = " + Str(wCan) + " " _
              & " WHERE  USU = '" + wcodusu + "' AND " _
              & "       DIST = '" + wDist + "' AND " _
              & "       PROV = '" + wProv + "' ")
              Db.CommitTrans
         Case "TRA"
              Db.BeginTrans
              Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
              & " SET TRA = " + Str(wCan) + " " _
              & " WHERE  USU = '" + wcodusu + "' AND " _
              & "       DIST = '" + wDist + "' AND " _
              & "       PROV = '" + wProv + "' ")
              Db.CommitTrans
         Case "ADH"
              Db.BeginTrans
              Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
              & " SET ADH = " + Str(wCan) + " " _
              & " WHERE  USU = '" + wcodusu + "' AND " _
              & "       DIST = '" + wDist + "' AND " _
              & "       PROV = '" + wProv + "' ")
              Db.CommitTrans
         Case "VIU"
              Db.BeginTrans
              Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
              & " SET VIU = " + Str(wCan) + " " _
              & " WHERE  USU = '" + wcodusu + "' AND " _
              & "       DIST = '" + wDist + "' AND " _
              & "       PROV = '" + wProv + "' ")
              Db.CommitTrans
         Case "PNP"
              Db.BeginTrans
              Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
              & " SET PNP = " + Str(wCan) + " " _
              & " WHERE  USU = '" + wcodusu + "' AND " _
              & "       DIST = '" + wDist + "' AND " _
              & "       PROV = '" + wProv + "' ")
              Db.CommitTrans
         Case "HON"
              Db.BeginTrans
              Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
              & " SET HON = " + Str(wCan) + " " _
              & " WHERE  USU = '" + wcodusu + "' AND " _
              & "       DIST = '" + wDist + "' AND " _
              & "       PROV = '" + wProv + "' ")
              Db.CommitTrans
         End Select
         
         wRegAct = wRegAct + 1
         ADO8a.MoveNext
      Loop
   
   End If
   
   Else
      wTit = 0: wHij = 0: wHer = 0: wNie = 0: wCiv = 0: wCi1 = 0: wTra = 0: wAdh = 0: wViu = 0: wPnp = 0: wHon = 0: wTot = 0
      
      aa = Leerado8a("select m.CODSOCIO, m.CODIGO, m.INS, m.nombre, m.E_SOCIO, e.nombre, M.UBIGEO, U.NOMBRE AS NOMUBI, LEFT(M.UBIGEO,4)+'00' AS prov, U2.NOMBRE as nomprov, e.aporte " _
                & " from maesocio as m inner join MAEE_SOCIO as e on m.E_SOCIO = e.e_socio " _
                & "                    LEFT  JOIN MAEUBIGEO AS  U ON M.UBIGEO  = U.CODIGO " _
                & "                    LEFT  JOIN MAEUBIGEO AS U2 ON left(M.UBIGEO,4) + '00' = U2.CODIGO " _
                & " where e.APORTE > 0 or m.E_SOCIO = 'HON'")
      If aa > 0 Then
         ADO8a.MoveFirst
         wRegAct = 1
         wRegTot = aa
         Do While Not ADO8a.EOF
            DoEvents
            lblMensaje.Caption = "Registro " + Trim(Format(wRegAct, "####0")) + " / " + Trim(Format(wRegTot, "####0"))
            lblMensaje.Refresh
            
            wSoc = ADO8a!codsocio
            wCod = ADO8a!codigo
            wIns = ADO8a!ins
            wE_S = ADO8a!e_socio
            wProv = IIf(ADO8a!prov = "    00", "9999", ADO8a!prov)
            wDist = IIf(ADO8a!ubigeo = "", "999999", ADO8a!ubigeo)
            wNomProv = IIf(IsNull(ADO8a!nomprov), "SIN UBIGEO", ADO8a!nomprov)
            wNomDist = IIf(IsNull(ADO8a!nomubi), "SIN UBIGEO", ADO8a!nomubi)
            wApo = ADO8a!aporte
            wSdo = SaldoFoto(wSoc, zMesTope)
            wsw = True
            Select Case True
            Case optUno.Value = True
                 If wSdo > wApo Then
                    wsw = False
                 End If
            Case optDos.Value = True
                 If wSdo > Round(wApo * 2, 2) Then
                    wsw = False
                 End If
            Case optTres.Value = True
                 If wSdo > Round(wApo * 3, 2) Then
                    wsw = False
                 End If
            End Select
      
            If wsw = True Then
               aa = Leerado7a("SELECT * FROM TMP_CUAXUBIGEO " _
                           & " WHERE DIST = '" + wDist + "' AND " _
                           & "       PROV = '" + wProv + "' AND " _
                           & "        USU = '" + wcodusu + "' ")
               If aa = 0 Then
                  Db.BeginTrans
                  Db.Execute ("INSERT INTO TMP_CUAXUBIGEO " _
                  & " (DIST, NOMDIST, PROV, NOMPROV, USU) " _
                  & " VALUES " _
                  & " ('" + wDist + "', '" + wNomDist + "', " _
                  & "  '" + wProv + "', '" + wNomProv + "', " _
                  & "  '" + wcodusu + "') ")
                  Db.CommitTrans
               End If
            
               Select Case wE_S
               Case "TIT"
                    Db.BeginTrans
                    Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
                    & " SET TIT = TIT + 1 " _
                    & " WHERE  USU = '" + wcodusu + "' AND " _
                    & "       DIST = '" + wDist + "' AND " _
                    & "       PROV = '" + wProv + "' ")
                    Db.CommitTrans
               Case "HIJ"
                    Db.BeginTrans
                    Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
                    & " SET HIJ = HIJ + 1 " _
                    & " WHERE  USU = '" + wcodusu + "' AND " _
                    & "       DIST = '" + wDist + "' AND " _
                    & "       PROV = '" + wProv + "' ")
                    Db.CommitTrans
               Case "HER"
                    Db.BeginTrans
                    Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
                    & " SET HER = HER + 1 " _
                    & " WHERE  USU = '" + wcodusu + "' AND " _
                    & "       DIST = '" + wDist + "' AND " _
                    & "       PROV = '" + wProv + "' ")
                    Db.CommitTrans
               Case "NIE"
                    Db.BeginTrans
                    Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
                    & " SET NIE = NIE + 1 " _
                    & " WHERE  USU = '" + wcodusu + "' AND " _
                    & "       DIST = '" + wDist + "' AND " _
                    & "       PROV = '" + wProv + "' ")
                    Db.CommitTrans
               Case "CIV"
                    Db.BeginTrans
                    Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
                    & " SET CIV = CIV + 1 " _
                    & " WHERE  USU = '" + wcodusu + "' AND " _
                    & "       DIST = '" + wDist + "' AND " _
                    & "       PROV = '" + wProv + "' ")
                    Db.CommitTrans
               Case "CI1"
                    Db.BeginTrans
                    Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
                    & " SET CI1 = CI1 + 1 " _
                    & " WHERE  USU = '" + wcodusu + "' AND " _
                    & "       DIST = '" + wDist + "' AND " _
                    & "       PROV = '" + wProv + "' ")
                    Db.CommitTrans
               Case "TRA"
                    Db.BeginTrans
                    Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
                    & " SET TRA = TRA + 1 " _
                    & " WHERE  USU = '" + wcodusu + "' AND " _
                    & "       DIST = '" + wDist + "' AND " _
                    & "       PROV = '" + wProv + "' ")
                    Db.CommitTrans
               Case "ADH"
                    Db.BeginTrans
                    Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
                    & " SET ADH = ADH + 1 " _
                    & " WHERE  USU = '" + wcodusu + "' AND " _
                    & "       DIST = '" + wDist + "' AND " _
                    & "       PROV = '" + wProv + "' ")
                    Db.CommitTrans
               Case "VIU"
                    Db.BeginTrans
                    Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
                    & " SET VIU = VIU + 1 " _
                    & " WHERE  USU = '" + wcodusu + "' AND " _
                    & "       DIST = '" + wDist + "' AND " _
                    & "       PROV = '" + wProv + "' ")
                    Db.CommitTrans
               Case "PNP"
                    Db.BeginTrans
                    Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
                    & " SET PNP = PNP + 1 " _
                    & " WHERE  USU = '" + wcodusu + "' AND " _
                    & "       DIST = '" + wDist + "' AND " _
                    & "       PROV = '" + wProv + "' ")
                    Db.CommitTrans
               Case "HON"
                    Db.BeginTrans
                    Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
                    & " SET HON = HON + 1 " _
                    & " WHERE  USU = '" + wcodusu + "' AND " _
                    & "       DIST = '" + wDist + "' AND " _
                    & "       PROV = '" + wProv + "' ")
                    Db.CommitTrans
               End Select
            End If
            
            wRegAct = wRegAct + 1
            ADO8a.MoveNext
         Loop
      End If
   End If
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_CUAXUBIGEO " _
   & " SET TOT = TIT + HIJ + HER + NIE + CIV + CI1 + TRA + ADH + VIU + PNP + HON " _
   & " WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   wTot = 0: wTit = 0: wHij = 0: wHer = 0: wNie = 0: wCiv = 0: wCi1 = 0: wTra = 0: wAdh = 0: wViu = 0
   aa = Leerado2("SELECT SUM(TIT) AS TIT, SUM(HIJ) AS HIJ, SUM(HER) AS HER, " _
                & "      SUM(NIE) AS NIE, SUM(CIV) AS CIV, SUM(CI1) AS CI1, " _
                & "      SUM(TRA) AS TRA, SUM(ADH) AS ADH, SUM(VIU) AS VIU, " _
                & "      SUM(PNP) AS PNP, SUM(HON) AS HON, SUM(TOT) AS TOT " _
                & " FROM TMP_CUAXUBIGEO " _
                & " WHERE USU = '" + wcodusu + "' ")
   If aa > 0 Then
      wTit = IIf(IsNull(ADO2!tit), 0, ADO2!tit)
      wHij = IIf(IsNull(ADO2!hij), 0, ADO2!hij)
      wHer = IIf(IsNull(ADO2!her), 0, ADO2!her)
      wNie = IIf(IsNull(ADO2!nie), 0, ADO2!nie)
      wCiv = IIf(IsNull(ADO2!civ), 0, ADO2!civ)
      wCi1 = IIf(IsNull(ADO2!ci1), 0, ADO2!ci1)
      wTra = IIf(IsNull(ADO2!tra), 0, ADO2!tra)
      wAdh = IIf(IsNull(ADO2!adh), 0, ADO2!adh)
      wViu = IIf(IsNull(ADO2!viu), 0, ADO2!viu)
      wPnp = IIf(IsNull(ADO2!pnp), 0, ADO2!pnp)
      wHon = IIf(IsNull(ADO2!hon), 0, ADO2!hon)
      wTot = IIf(IsNull(ADO2!tot), 0, ADO2!tot)
   End If
   
   DoEvents
   lblMensaje.Caption = ""
   lblMensaje.Refresh
   
   lblTot.Caption = Format(wTot, "###0;;\ ")
   
   aa = Leerado2("SELECT NOMPROV, NOMDIST, TIT, HIJ, HER, NIE, CIV, CI1, TRA, ADH, VIU, PNP, HON, TOT " _
                & " FROM TMP_CUAXUBIGEO " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY PROV, DIST ")
   Set DataGrid1.DataSource = ADO2
End Sub

Private Sub LlenaCab1()
   DataGrid1.Columns(0).Alignment = dbgLeft
   DataGrid1.Columns(0).Width = 2000
   DataGrid1.Columns(0).Caption = "PROVINCIA"

   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Width = 2000
   DataGrid1.Columns(1).Caption = "DISTRITO"

   DataGrid1.Columns(2).Alignment = dbgRight
   DataGrid1.Columns(2).Width = 650
   DataGrid1.Columns(2).Caption = "TIT"
   DataGrid1.Columns(2).NumberFormat = "##,##0;;\ "

   DataGrid1.Columns(3).Alignment = dbgRight
   DataGrid1.Columns(3).Width = 650
   DataGrid1.Columns(3).Caption = "HIJ"
   DataGrid1.Columns(3).NumberFormat = "##,##0;;\ "

   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).Width = 650
   DataGrid1.Columns(4).Caption = "HER"
   DataGrid1.Columns(4).NumberFormat = "##,##0;;\ "

   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).Width = 650
   DataGrid1.Columns(5).Caption = "NIE"
   DataGrid1.Columns(5).NumberFormat = "##,##0;;\ "

   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).Width = 650
   DataGrid1.Columns(6).Caption = "CIV"
   DataGrid1.Columns(6).NumberFormat = "##,##0;;\ "

   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).Width = 650
   DataGrid1.Columns(7).Caption = "CI1"
   DataGrid1.Columns(7).NumberFormat = "##,##0;;\ "

   DataGrid1.Columns(8).Alignment = dbgRight
   DataGrid1.Columns(8).Width = 650
   DataGrid1.Columns(8).Caption = "TRA"
   DataGrid1.Columns(8).NumberFormat = "##,##0;;\ "

   DataGrid1.Columns(9).Alignment = dbgRight
   DataGrid1.Columns(9).Width = 650
   DataGrid1.Columns(9).Caption = "ADH"
   DataGrid1.Columns(9).NumberFormat = "##,##0;;\ "

   DataGrid1.Columns(10).Alignment = dbgRight
   DataGrid1.Columns(10).Width = 650
   DataGrid1.Columns(10).Caption = "VIU"
   DataGrid1.Columns(10).NumberFormat = "##,##0;;\ "

   DataGrid1.Columns(11).Alignment = dbgRight
   DataGrid1.Columns(11).Width = 650
   DataGrid1.Columns(11).Caption = "PNP"
   DataGrid1.Columns(11).NumberFormat = "##,##0;;\ "

   DataGrid1.Columns(12).Alignment = dbgRight
   DataGrid1.Columns(12).Width = 650
   DataGrid1.Columns(12).Caption = "HON"
   DataGrid1.Columns(12).NumberFormat = "##,##0;;\ "

   DataGrid1.Columns(13).Alignment = dbgRight
   DataGrid1.Columns(13).Width = 650
   DataGrid1.Columns(13).Caption = "TOT"
   DataGrid1.Columns(13).NumberFormat = "##,##0;;\ "
End Sub


