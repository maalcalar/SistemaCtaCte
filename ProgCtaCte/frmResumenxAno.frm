VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmResumenxAno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Cobros x Año"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   10890
   Begin VB.CheckBox chkCAJMP 
      Caption         =   "Caja MP"
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
      Left            =   5280
      TabIndex        =   12
      Top             =   600
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkDIECO 
      Caption         =   "DIECO"
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
      Left            =   3720
      TabIndex        =   11
      Top             =   600
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkTesoreria 
      Caption         =   "Tesoreria"
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
      Left            =   2160
      TabIndex        =   10
      Top             =   600
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox txtAnoCab 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   600
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
      Left            =   8160
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
      Left            =   6840
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
      Left            =   9480
      TabIndex        =   4
      Top             =   7920
      Width           =   1095
   End
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmResumenxAno.frx":0000
      Left            =   1320
      List            =   "frmResumenxAno.frx":0002
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
      Left            =   7560
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   1095
   End
   Begin Crystal.CrystalReport Crys1 
      Left            =   9360
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
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   10215
      _ExtentX        =   18018
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
      Caption         =   "RESUMEN DE COBROS POR AÑO"
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
      Top             =   600
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
Private Sub chkCAJMP_Click()
   LlenaCab
End Sub

Private Sub chkDIECO_Click()
   LlenaCab
End Sub

Private Sub chkTesoreria_Click()
   LlenaCab
End Sub

Private Sub cmdBuscar_Click()
   LlenaCab
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(7) As String, wreg As Integer, wTot As Integer
   Dim wApoSol As Currency, wApoDol As Currency, _
       wInsSol As Currency, wInsDol As Currency, _
       wRenSol As Currency, wRenDol As Currency, _
       wAno As String
       
   wAno = txtAnoCab.Text
       
   Heading(0) = "TIPO"
   Heading(1) = "NOMBRE"
   Heading(2) = "APORTE S/."
   Heading(3) = "APORTE US$"
   Heading(4) = "INSCRIP S/."
   Heading(5) = "INSCRIP US$"
   Heading(6) = "RENOV S/."
   Heading(7) = "RENOV US$"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 8)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 8)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "RESUMEN DE COBROS X AÑO - EJERCICIO " + wAno
        For I = 1 To 8 Step 1
            .Cells(3, I) = Heading(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 6
        objExcel.Columns("B").ColumnWidth = 20
        objExcel.Columns("C").ColumnWidth = 12
        objExcel.Columns("D").ColumnWidth = 12
        objExcel.Columns("E").ColumnWidth = 12
        objExcel.Columns("F").ColumnWidth = 12
        objExcel.Columns("G").ColumnWidth = 12
        objExcel.Columns("H").ColumnWidth = 12
   End With
   
   aa = Leerado3("SELECT * FROM TMP_RESXANO WHERE USU = '" + wcodusu + "' ORDER BY ORDEN ")
   If aa > 0 Then
      wTot = aa
      V = 4
      H = 1
      wNum1 = 1
      wApoSol = 0: wApoDol = 0: wInsSol = 0: wInsDol = 0: wRenSol = 0: wRenDol = 0
      Do While Not ADO3.EOF
         
         objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 7)).NumberFormat = "#####,##0.00;;\ "
         
         objExcel.Cells(V, H + 0) = ADO3!e_socio
         objExcel.Cells(V, H + 1) = ADO3!nombre
         objExcel.Cells(V, H + 2) = ADO3!soless
         objExcel.Cells(V, H + 3) = ADO3!dolare
         objExcel.Cells(V, H + 4) = ADO3!inssol
         objExcel.Cells(V, H + 5) = ADO3!insdol
         objExcel.Cells(V, H + 6) = ADO3!rensol
         objExcel.Cells(V, H + 7) = ADO3!rendol
            
         wApoSol = wApoSol + ADO3!soless
         wApoDol = wApoDol + ADO3!dolare
         wInsSol = wInsSol + ADO3!inssol
         wInsDol = wInsDol + ADO3!insdol
         wRenSol = wRenSol + ADO3!rensol
         wRenDol = wRenDol + ADO3!rendol
         V = V + 1
         ADO3.MoveNext
      Loop
      V = V + 1
      
      objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 7)).NumberFormat = "##,##0;;\ "
      objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 7)).Font.Bold = True
      objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 7)).Font.Color = RGB(255, 0, 0)
      objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 7)).Borders.LineStyle = xlContinuous
      objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 7)).Borders.Color = RGB(255, 0, 0)
      
      objExcel.Cells(V, H + 1) = "TOTALES FINALES"
      objExcel.Cells(V, H + 2) = wApoSol
      objExcel.Cells(V, H + 3) = wApoDol
      objExcel.Cells(V, H + 4) = wInsSol
      objExcel.Cells(V, H + 5) = wInsDol
      objExcel.Cells(V, H + 6) = wRenSol
      objExcel.Cells(V, H + 7) = wRenDol
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
   Dim wAno As String
   wAno = txtAnoCab.Text
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\ResumenCobroxAno.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= 'EJERCICIO " + wAno + "' "
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
   
   txtAnoCab.Text = Format(Val(wanocia) - 1, "0000")
   
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
   Dim aa As Long, wreg As Long, wTot As Long, _
       wE_S As String, wAno As String, _
       wCan As Integer, _
       wApoSol As Currency, wApoDol As Currency, _
       wRenSol As Currency, wRenDol As Currency, _
       wInsSol As Currency, wInsDol As Currency

   wAno = txtAnoCab.Text
   
   Set DataGrid1.DataSource = Nothing
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESXANO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   DoEvents
   lblMensaje.Caption = "Preparando Archivo......."
   lblMensaje.Refresh
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESXANO " _
   & " (E_SOCIO, NOMBRE, ORDEN, USU) " _
   & " SELECT " _
   & "  E_SOCIO, NOMBRE, ORDEN, '" + wcodusu + "' " _
   & " FROM MAEE_SOCIO " _
   & " WHERE APORTE > 0 ")
   Db.CommitTrans
   
   aa = Leerado8("SELECT * FROM TMP_RESXANO ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wE_S = ADO8!e_socio
         wCan = 0: wApoSol = 0: wApoDol = 0: wRenSol = 0: wRenDol = 0: wInsSol = 0: wInsDol = 0
   
         aa = Leerado7("select distinct codsocio From v_cobros_tesor_aporte where ano = '" + wAno + "' and e_socio = '" + wE_S + "' " _
                    & " Union " _
                    & " select distinct codsocio From v_cobros_tesor_inscri where ano = '" + wAno + "' and e_socio = '" + wE_S + "' " _
                    & " Union " _
                    & " select distinct codsocio From v_cobros_tesor_RENOVA where ano = '" + wAno + "' and e_socio = '" + wE_S + "' " _
                    & " Union " _
                    & " select distinct codsocio From v_cobros_DIECO        where ano = '" + wAno + "' and e_socio = '" + wE_S + "' " _
                    & " Union " _
                    & " select distinct codsocio From v_cobros_CAJMP        where ano = '" + wAno + "' and e_socio = '" + wE_S + "' ")
         If aa > 0 Then
            wCan = aa
         End If
         Set ADO7 = Nothing
   
         If chkTesoreria.Value = vbChecked Then
            aa = Leerado7("SELECT COUNT(*) AS CANT, SUM(SOLES) AS SOLES, SUM(DOLARES) AS DOLARES " _
                & " From V_COBROS_TESOR_APORTE " _
                & " WHERE (ANO = '" + wAno + "') AND " _
                & "       (E_SOCIO = '" + wE_S + "')")
            If aa > 0 Then
               wApoSol = IIf(IsNull(ADO7!soles), 0, ADO7!soles)
               wApoDol = IIf(IsNull(ADO7!dolares), 0, ADO7!dolares)
            End If
            Set ADO7 = Nothing
   
            aa = Leerado7("SELECT COUNT(*) AS CANT, SUM(SOLES) AS SOLES, SUM(DOLARES) AS DOLARES " _
                & " From V_COBROS_TESOR_RENOVA " _
                & " WHERE (ANO = '" + wAno + "') AND " _
                & "       (E_SOCIO = '" + wE_S + "')")
            If aa > 0 Then
               wRenSol = IIf(IsNull(ADO7!soles), 0, ADO7!soles)
               wRenDol = IIf(IsNull(ADO7!dolares), 0, ADO7!dolares)
            End If
            Set ADO7 = Nothing
   
            aa = Leerado7("SELECT COUNT(*) AS CANT, SUM(SOLES) AS SOLES, SUM(DOLARES) AS DOLARES " _
                & " From V_COBROS_TESOR_INSCRI " _
                & " WHERE (ANO = '" + wAno + "') AND " _
                & "       (E_SOCIO = '" + wE_S + "')")
            If aa > 0 Then
               wInsSol = IIf(IsNull(ADO7!soles), 0, ADO7!soles)
               wInsDol = IIf(IsNull(ADO7!dolares), 0, ADO7!dolares)
            End If
            Set ADO7 = Nothing
         End If
   
         If chkDIECO.Value = vbChecked Then
            aa = Leerado7("SELECT COUNT(*) AS CANT, SUM(SOLES) AS SOLES, SUM(DOLARES) AS DOLARES " _
                & " From V_COBROS_DIECO " _
                & " WHERE (ANO = '" + wAno + "') AND " _
                & "       (E_SOCIO = '" + wE_S + "')")
            If aa > 0 Then
'              wCan = wCan + IIf(IsNull(ADO7!cant), 0, ADO7!cant)
               wApoSol = wApoSol + IIf(IsNull(ADO7!soles), 0, ADO7!soles)
               wApoDol = wApoDol + IIf(IsNull(ADO7!dolares), 0, ADO7!dolares)
            End If
            Set ADO7 = Nothing
         End If
   
         If chkCAJMP.Value = vbChecked Then
            aa = Leerado7("SELECT COUNT(*) AS CANT, SUM(SOLES) AS SOLES, SUM(DOLARES) AS DOLARES " _
                & " From V_COBROS_CAJMP " _
                & " WHERE (ANO = '" + wAno + "') AND " _
                & "       (E_SOCIO = '" + wE_S + "')")
            If aa > 0 Then
'              wCan = wCan + IIf(IsNull(ADO7!cant), 0, ADO7!cant)
               wApoSol = wApoSol + IIf(IsNull(ADO7!soles), 0, ADO7!soles)
               wApoDol = wApoDol + IIf(IsNull(ADO7!dolares), 0, ADO7!dolares)
            End If
            Set ADO7 = Nothing
         End If
   
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_RESXANO " _
         & " SET CANTIDAD = " + Str(wCan) + ", " _
         & "       SOLESS = " + Str(wApoSol) + ", " _
         & "       DOLARE = " + Str(wApoDol) + ", " _
         & "       RENSOL = " + Str(wRenSol) + ", " _
         & "       RENDOL = " + Str(wRenDol) + ", " _
         & "       INSSOL = " + Str(wInsSol) + ", " _
         & "       INSDOL = " + Str(wInsDol) + "  " _
         & " WHERE     USU = '" + wcodusu + "' AND " _
         & "       E_SOCIO = '" + wE_S + "' ")
         Db.CommitTrans
    
         ADO8.MoveNext
      Loop
   End If
   
   DoEvents
   lblMensaje.Caption = ""
   lblMensaje.Refresh
   
   aa = Leerado2("SELECT E_SOCIO, NOMBRE, CANTIDAD, SOLESS, DOLARE, INSSOL, INSDOL, RENSOL, RENDOL " _
            & " FROM TMP_RESXANO " _
            & " WHERE USU = '" + wcodusu + "' " _
            & " ORDER BY ORDEN ")
   Set DataGrid1.DataSource = ADO2
 
   DataGrid1.Columns(0).Width = 700   ' E_SOCIO
   DataGrid1.Columns(0).Alignment = dbgCenter
   DataGrid1.Columns(0).Caption = "E.SOCIO"
    
   DataGrid1.Columns(1).Width = 1500   ' NOMBRE
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Caption = "NOMBRE"
    
   DataGrid1.Columns(2).Width = 900   ' CANTIDAD
   DataGrid1.Columns(2).Alignment = dbgRight
   DataGrid1.Columns(2).Caption = "CANTIDAD"
   DataGrid1.Columns(2).NumberFormat = "#####,##0;;\ "

   DataGrid1.Columns(3).Width = 1100  ' SOLESS
   DataGrid1.Columns(3).Alignment = dbgRight
   DataGrid1.Columns(3).Caption = "APORT S/."
   DataGrid1.Columns(3).NumberFormat = "#####,##0.00;;\ "

   DataGrid1.Columns(4).Width = 1100  ' DOLARE
   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).Caption = "APORT US$"
   DataGrid1.Columns(4).NumberFormat = "#####,##0.00;;\ "

   DataGrid1.Columns(5).Width = 1100  ' INSSOL
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).Caption = "INSCR S/."
   DataGrid1.Columns(5).NumberFormat = "#####,##0.00;;\ "

   DataGrid1.Columns(6).Width = 1100  ' INSDOL
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).Caption = "INSCR US$"
   DataGrid1.Columns(6).NumberFormat = "#####,##0.00;;\ "

   DataGrid1.Columns(7).Width = 1100  ' RENSOL
   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).Caption = "RENOV S/."
   DataGrid1.Columns(7).NumberFormat = "#####,##0.00;;\ "

   DataGrid1.Columns(8).Width = 1100  ' RENDOL
   DataGrid1.Columns(8).Alignment = dbgRight
   DataGrid1.Columns(8).Caption = "RENOV US$"
   DataGrid1.Columns(8).NumberFormat = "#####,##0.00;;\ "
End Sub

Private Sub txtAnoCab_GotFocus()
   txtAnoCab.SelStart = 0
   txtAnoCab.SelLength = 4
End Sub

Private Sub txtAnoCab_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txtAnoCab.Text)) = 0 Then
         MsgBox "Año En Blanco", vbExclamation
         txtAnoCab.Text = Format((wanocia) - 1, "0000")
         Exit Sub
      End If
      If txtAnoCab.Text < "2015" And txtAnoCab.Text > "2030" Then
         MsgBox "Año Digitado Es Invalido", vbExclamation
         txtAnoCab.Text = Format((wanocia) - 1, "0000")
         Exit Sub
      End If
      cmdBuscar.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub
