VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmConAsignaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Socios y Asignaciones"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   12540
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
      Left            =   5160
      TabIndex        =   5
      Top             =   120
      Width           =   3015
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
      Left            =   5160
      TabIndex        =   4
      Top             =   480
      Value           =   -1  'True
      Width           =   3015
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
      TabIndex        =   3
      Top             =   7800
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
      Left            =   11160
      TabIndex        =   2
      Top             =   7800
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
      Left            =   3240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   12255
      _ExtentX        =   21616
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
End
Attribute VB_Name = "frmConAsignaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
   LlenaCab
   LlenaCab1
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, _
       Heading(25) As String, Headin2(25) As String, _
       wRegAct As Integer, wRegTot As Integer

   Heading(0) = "DATOS GENERALES DEL ASOCIADO"

   Heading(13) = "ESTADO"
   Heading(14) = "FECHA"
   Heading(15) = "TIPO"
   
   Heading(16) = "DATOS GENERALES DEL PADRE QUE ASIGNA"
   Heading(25) = "TIPO COBRO"


   Headin2(0) = "SOCIO"
   Headin2(1) = "CODIGO"
   Headin2(2) = "INS"
   Headin2(3) = "NOMBRE SOCIO"
   Heading(4) = "D.N.I."
   Headin2(5) = "GRADO"
   
   Headin2(6) = "DIRECCION"
   Headin2(7) = "UBICACION GEOGRAFICA"
   Headin2(8) = "TELEFONO"
   Headin2(9) = "TELF2"
   Headin2(10) = "CELULAR"
   Headin2(11) = "EMAIL"
   Headin2(12) = "EMAIL2"
   
   Headin2(13) = "SOCIO"
   Headin2(14) = "INGRESO"
   Headin2(15) = "COBRO"
   
   Headin2(16) = "SOCIO"
   Headin2(17) = "CODIGO"
   Headin2(18) = "INS"
   Headin2(19) = "NOMBRE"
   Headin2(20) = "TIP.COB"
   Headin2(21) = "LIN"
   Headin2(22) = "ESTADO"
   Headin2(23) = "OBSERV"
   Heading(24) = "FECTOP"
   
   Heading(25) = "FINAL"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 26)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(4, 26)).Font.Bold = True
        
        .Range(objExcel.Cells(2, 1), .Cells(2, 26)).Merge
        .Range(objExcel.Cells(2, 1), .Cells(2, 26)).HorizontalAlignment = xlCenter
        
        .Range(objExcel.Cells(3, 1), .Cells(3, 6)).Merge
        .Range(objExcel.Cells(3, 7), .Cells(3, 13)).Merge
        .Range(objExcel.Cells(3, 16), .Cells(3, 23)).Merge
        
        .Range(objExcel.Cells(3, 1), .Cells(3, 24)).HorizontalAlignment = xlCenter
        
        
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "RELACION DE ASOCIADOS Y ASIGNACIONES "
        For I = 1 To 26 Step 1
            .Cells(3, I) = Heading(I - 1)
            .Cells(4, I) = Headin2(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 7
        objExcel.Columns("B").ColumnWidth = 9
        objExcel.Columns("C").ColumnWidth = 3
        objExcel.Columns("D").ColumnWidth = 60
        objExcel.Columns("E").ColumnWidth = 11
        objExcel.Columns("F").ColumnWidth = 15
        
        objExcel.Columns("G").ColumnWidth = 60
        objExcel.Columns("H").ColumnWidth = 40
        objExcel.Columns("I").ColumnWidth = 18
        objExcel.Columns("J").ColumnWidth = 18
        objExcel.Columns("K").ColumnWidth = 12
        
        objExcel.Columns("L").ColumnWidth = 40
        objExcel.Columns("M").ColumnWidth = 40
        
        objExcel.Columns("N").ColumnWidth = 9
        objExcel.Columns("O").ColumnWidth = 12
        objExcel.Columns("P").ColumnWidth = 16
        
        objExcel.Columns("Q").ColumnWidth = 7
        objExcel.Columns("R").ColumnWidth = 9
        objExcel.Columns("S").ColumnWidth = 3
        objExcel.Columns("T").ColumnWidth = 60
        objExcel.Columns("U").ColumnWidth = 16
   
        objExcel.Columns("V").ColumnWidth = 4
        objExcel.Columns("W").ColumnWidth = 6
        objExcel.Columns("X").ColumnWidth = 18
        objExcel.Columns("Y").ColumnWidth = 12
   End With
   
   aa = Leerado3("SELECT * " _
                & " FROM TMP_SOCIOASIG " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY NOMBRE ")
   If aa > 0 Then
      wRegTot = aa
      V = 5
      H = 1
      wRegAct = 1
      Do While Not ADO3.EOF
         DoEvents
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + _
                              Format(wRegAct, "####0") + " / " + _
                              Format(wRegTot, "####0")
         lblMensaje.Refresh
         
         objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 16)).NumberFormat = "######0.00;;\ "
            
         objExcel.Cells(V, H + 0) = ADO3!codsocio
         objExcel.Cells(V, H + 1) = ADO3!codigo
         objExcel.Cells(V, H + 2) = ADO3!ins
         objExcel.Cells(V, H + 3) = ADO3!nombre
         objExcel.Cells(V, H + 4) = ADO3!numdoc
         objExcel.Cells(V, H + 5) = ADO3!nomgrado
         objExcel.Cells(V, H + 6) = ADO3!direc
         objExcel.Cells(V, H + 7) = ADO3!nomubigeo
         objExcel.Cells(V, H + 8) = ADO3!telefono
         objExcel.Cells(V, H + 9) = ADO3!telefon2
         objExcel.Cells(V, H + 10) = ADO3!celular
         objExcel.Cells(V, H + 11) = ADO3!email
         objExcel.Cells(V, H + 12) = ADO3!email2
         objExcel.Cells(V, H + 13) = ADO3!e_socio
         objExcel.Cells(V, H + 14) = ADO3!fecing
         objExcel.Cells(V, H + 15) = ADO3!nomcob
         
         objExcel.Cells(V, H + 16) = ADO3!socpadre
         objExcel.Cells(V, H + 17) = ADO3!codpadre
         objExcel.Cells(V, H + 18) = ADO3!inspadre
         objExcel.Cells(V, H + 19) = ADO3!nompadre
         objExcel.Cells(V, H + 20) = ADO3!nomcobpadre
         objExcel.Cells(V, H + 21) = ADO3!lin
         objExcel.Cells(V, H + 22) = ADO3!estado
         objExcel.Cells(V, H + 23) = ADO3!observ
         objExcel.Cells(V, H + 24) = ADO3!fectop
         objExcel.Cells(V, H + 25) = ADO3!nomcobdet
         
         wRegAct = wRegAct + 1
         V = V + 1
         ADO3.MoveNext
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
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\ResxTipo.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.SelectionFormula = " {TMP_RESXTIPO.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Form_Load
   
   Set DataGrid1.DataSource = Nothing
   
   cmdBuscar.SetFocus
End Sub

Private Sub Form_Load()
   frmConAsignaciones.Left = (Screen.Width - Width) \ 2
   frmConAsignaciones.Top = 0
End Sub

Private Sub LlenaCab()

   Dim aa As Long

   Set DataGrid1.DataSource = Nothing
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_SOCIOASIG WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_SOCIOASIG " _
   & " (CODSOCIO, CODIGO, INS, NOMBRE, GRADO, NOMGRADO, E_SOCIO, NUMDOC, FECING, " _
   & "  DIREC, UBIGEO, NOMUBIGEO, TELEFONO, TELEFON2, CELULAR, EMAIL, EMAIL2, TIPCOB, " _
   & "  NOMCOB, SOCPADRE, CODPADRE, INSPADRE, NOMPADRE, TIPCOBPADRE, NOMCOBPADRE, " _
   & "  LIN, ESTADO, OBSERV, FECTOP, TIPCOBDET, NOMCOBDET, USU ) " _
   & " SELECT " _
   & "  CODSOCIO, CODIGO, INS, NOMBRE, GRADO, NOMGRADO, E_SOCIO, NUMDOC, FECING, " _
   & "  DIREC, UBIGEO, NOMUBIGEO, TELEFONO, TELEFON2, CELULAR, EMAIL, EMAIL2, TIPCOB, " _
   & "  NOMCOB, SOCPADRE, CODPADRE, INSPADRE, NOMPADRE, TIPCOBPADRE, NOMCOBPADRE, " _
   & "  LIN, ESTADO, OBSERV, FECTOP, TIPCOBDET, NOMCOBDET, '" + wcodusu + "' " _
   & " FROM V_TOTALSOCIOS ")
   Db.CommitTrans
  
   If optActivos.Value = True Then
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_SOCIOASIG " _
      & " WHERE (USU = '" + wcodusu + "') AND " _
      & "       (E_SOCIO = 'FAL' OR " _
      & "        E_SOCIO = 'RET' OR " _
      & "        E_SOCIO = 'REN' OR " _
      & "        E_SOCIO = 'SEP' OR " _
      & "        E_SOCIO = 'EXP' OR " _
      & "        E_SOCIO = '998' OR " _
      & "        E_SOCIO = 'EXC' ) ")
      Db.CommitTrans
   End If
  
   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, TIPCOBDET, NOMCOBDET FROM TMP_SOCIOASIG " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO2
End Sub

Private Sub LlenaCab1()
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
    
   DataGrid1.Columns(3).Width = 5550  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE"
    
   DataGrid1.Columns(4).Width = 500  ' E_SOCIO
   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Caption = "ESTADO"
    
   DataGrid1.Columns(5).Width = 500  ' TIPCOB
   DataGrid1.Columns(5).Alignment = dbgLeft
   DataGrid1.Columns(5).Caption = "T.COB"
    
   DataGrid1.Columns(6).Width = 2000 ' NOMBRE COBRO
   DataGrid1.Columns(6).Alignment = dbgLeft
   DataGrid1.Columns(6).Caption = "NOMBRE COBRO"
End Sub

Private Sub optActivos_Click()
   LlenaCab
   LlenaCab1
End Sub

Private Sub optTodos_Click()
   LlenaCab
   LlenaCab1
End Sub
