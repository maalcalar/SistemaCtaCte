VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmDIECORelacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relación Socios Cobro x DIECO"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   10500
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
      Left            =   7440
      TabIndex        =   7
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
      Left            =   6120
      TabIndex        =   5
      Top             =   6960
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
      Left            =   8760
      TabIndex        =   4
      Top             =   6960
      Width           =   1095
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
      Left            =   8760
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmDIECORelacion.frx":0000
      Left            =   1080
      List            =   "frmDIECORelacion.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   10095
      _ExtentX        =   17806
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
      Left            =   360
      Top             =   6840
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
      Left            =   8760
      TabIndex        =   6
      Top             =   6360
      Width           =   1095
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmDIECORelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
   LlenaCab
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(5) As String, wreg As Integer, wTot As Integer
   Dim wNom As String, wNum1 As Long, wSoc As Integer, wFecIng As Date, wFecNac As Date, _
       wMes As String, wAno As String, wE_S As String, _
       wSdoSol As Currency, zSdoSol As Currency, _
       wSdoDol As Currency, zSdoDol As Currency
       
   Heading(0) = "NRO."
   Heading(1) = "CODIGO"
   Heading(2) = "CODOFIN"
   Heading(3) = "D.N.I."
   Heading(4) = "APELLIDOS Y NOMBRES"
   Heading(5) = "ESTADO"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 6)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 6)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "RELACION DE SOCIOS CON DESCUENTO DIECO"
        For I = 1 To 6 Step 1
            .Cells(3, I) = Heading(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 6
        objExcel.Columns("B").ColumnWidth = 10
        objExcel.Columns("C").ColumnWidth = 11
        objExcel.Columns("D").ColumnWidth = 10
        objExcel.Columns("E").ColumnWidth = 80
        objExcel.Columns("F").ColumnWidth = 12
   End With
   
   aa = Leerado3("SELECT * FROM TMP_DIECO_RELACION WHERE USU = '" + wcodusu + "' ORDER BY NOMBRE ")
   If aa > 0 Then
      wTot = aa
      V = 4
      H = 1
      wNum = 1
      Do While Not ADO3.EOF
         
         objExcel.Cells(V, H + 0) = Format(wNum, "###0")
         objExcel.Cells(V, H + 1) = ADO3!codsocio
         objExcel.Cells(V, H + 2) = Format(ADO3!codigo, "########") + "-" + Format(ADO3!ins, "#")
         objExcel.Cells(V, H + 3) = ADO3!numdoc
         objExcel.Cells(V, H + 4) = ADO3!nombre
         objExcel.Cells(V, H + 5) = ADO3!e_socio
         
         wNum = wNum + 1
         wreg = wreg + 1
         V = V + 1
         ADO3.MoveNext
      Loop
      Set ADO3 = Nothing
      objExcel.Visible = True
      Set objExcel = Nothing
   End If
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdImprimir_Click()
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\DiecoRelacion.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.SelectionFormula = " {TMP_DIECO_RELACION.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   frmDIECORelacion.Left = (Screen.Width - Width) \ 2
   frmDIECORelacion.Top = 0
   
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
   
   Set DataGrid1.DataSource = Nothing
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_DIECO_RELACION WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   cmdBuscar.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_DIECO_RELACION WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim aa As Long

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_DIECO_RELACION WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECO_RELACION " _
   & " (USU, CODSOCIO, CODIGO, INS, NUMDOC, NOMBRE, E_SOCIO ) " _
   & " SELECT '" + wcodusu + "', M.CODSOCIO, M.CODIGO, M.INS, M.NUMDOC, M.NOMBRE, M.E_SOCIO " _
   & " FROM MAESOCIO AS M INNER JOIN MAEE_SOCIO AS E ON M.E_SOCIO = E.E_SOCIO " _
   & " WHERE E.APORTE > 0 AND " _
   & "       M.FECRENU IS NULL AND " _
   & "       M.FECEXCLU IS NULL AND " _
   & "       M.FECEXPUL IS NULL AND " _
   & "       TIPCOB = '01'")
   Db.CommitTrans

   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NUMDOC, NOMBRE, E_SOCIO " _
            & " FROM TMP_DIECO_RELACION " _
            & " WHERE USU = '" + wcodusu + "' " _
            & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO2
 
   lblTotal.Caption = Format(aa, "##,##0") + " "

   DataGrid1.Columns(0).Width = 700   ' CODSOCIO
   DataGrid1.Columns(0).Alignment = dbgRight
   DataGrid1.Columns(0).Caption = "COD.SOCIO"
    
   DataGrid1.Columns(1).Width = 1000  ' CODIGO
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Caption = "CODIGO"
    
   DataGrid1.Columns(2).Width = 400   ' INS
   DataGrid1.Columns(2).Alignment = dbgLeft
   DataGrid1.Columns(2).Caption = "INS"
    
   DataGrid1.Columns(3).Width = 1000  ' DNI
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "DNI"
    
   DataGrid1.Columns(4).Width = 5400  ' NOMBRE
   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Caption = "NOMBRE"
    
   DataGrid1.Columns(5).Width = 700    ' E_SOCIO
   DataGrid1.Columns(5).Alignment = dbgCenter
   DataGrid1.Columns(5).Caption = "E_SOCIO"
End Sub



