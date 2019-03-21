VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmElePadronUnipag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Padron Activos Habiles x Unidad de Pago"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   13095
   Begin VB.CommandButton cmdExportar2 
      Caption         =   "&Exportar Formato 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   16
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdFormato2 
      Caption         =   "&Imprimir Formato 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7680
      Width           =   975
   End
   Begin VB.ComboBox cmbRegion 
      Height          =   315
      ItemData        =   "frmElePadronUnipag.frx":0000
      Left            =   1200
      List            =   "frmElePadronUnipag.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   840
      Width           =   3615
   End
   Begin VB.ComboBox cmbGrado 
      Height          =   315
      ItemData        =   "frmElePadronUnipag.frx":0004
      Left            =   1230
      List            =   "frmElePadronUnipag.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   120
      Width           =   3615
   End
   Begin VB.ComboBox cmbE_socio 
      Height          =   315
      ItemData        =   "frmElePadronUnipag.frx":0008
      Left            =   1230
      List            =   "frmElePadronUnipag.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   480
      Width           =   3615
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
      Left            =   5310
      TabIndex        =   8
      Top             =   720
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
      Height          =   615
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
      Height          =   615
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
      Height          =   615
      Left            =   10320
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6960
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   9128
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
      Left            =   12600
      Top             =   6960
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
      Caption         =   "Región"
      Height          =   195
      Index           =   3
      Left            =   525
      TabIndex        =   14
      Top             =   900
      Width           =   510
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Grado"
      Height          =   195
      Index           =   2
      Left            =   630
      TabIndex        =   12
      Top             =   180
      Width           =   435
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Estado Socio"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   945
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOTA: Se denomina Socio ACTIVO HABIL aquellos Socios Que No Tengan Deuda y que Sean Clase de Socio TIT, VIU, NIE, HIJ, HER"
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
      Left            =   6720
      TabIndex        =   7
      Top             =   120
      Width           =   6015
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
      Left            =   10920
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
      Left            =   9150
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
      Left            =   360
      TabIndex        =   4
      Top             =   6840
      Width           =   7575
   End
End
Attribute VB_Name = "frmElePadronUnipag"
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

Private Sub cmbRegion_KeyPress(KeyAscii As Integer)
   cmdBuscar.SetFocus
End Sub

Private Sub cmdBuscar_Click()
   LlenaCab
   LlenaCab1
   TotalCab
End Sub

Private Sub cmdExporta_Click()
   On Error GoTo err
   
   Dim aa As Integer, wRegAct As Integer, wRegTot As Integer, wRegReg As Integer, _
       I As Integer, Heading(9) As String, _
       wNum As Integer, wFec As Date, wreg As String, wNom As String, _
       wDni As String, wGra As Integer, wNomGra As String
       
   Heading(0) = "REGION"
   Heading(1) = "GRUPO"
   Heading(2) = "NUM"
   Heading(3) = "GRADO"
   Heading(4) = "NOMBRE ASOCIADO"
   Heading(5) = "FEC.ING"
   Heading(6) = "D.N.I."
   Heading(7) = "DEUDA"
   Heading(8) = "FIRMA"
   Heading(9) = "IMPRESION DIGITAL"
   aa = Leerado3("SELECT * FROM TMP_PADRON WHERE USU = '" + wcodusu + "' ORDER BY REGIONGRUPO, GRADOGRUPO, NOMBRE ")
   If aa > 0 Then
      wRegAct = 1: wRegReg = 1
      wRegTot = aa
      Set objExcel = New Excel.Application
      objExcel.SheetsInNewWorkbook = 1
      objExcel.Workbooks.Add
      With objExcel
           .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
           .Range(.Cells(3, 1), .Cells(3, 10)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 10)).Font.Bold = True
           .Cells(1, 1) = wnomcia
           .Cells(2, 1) = "PADRON GENERAL DE ASOCIADOS ACTIVOS HABILES"
           For I = 1 To 10 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           objExcel.Columns("A").ColumnWidth = 14
           objExcel.Columns("B").ColumnWidth = 15
           objExcel.Columns("C").ColumnWidth = 6
           objExcel.Columns("D").ColumnWidth = 15
           objExcel.Columns("E").ColumnWidth = 50
           objExcel.Columns("F").ColumnWidth = 11
           objExcel.Columns("G").ColumnWidth = 10
           objExcel.Columns("H").ColumnWidth = 11
           objExcel.Columns("I").ColumnWidth = 18
           objExcel.Columns("J").ColumnWidth = 18
      End With
      V = 4
      H = 1
      wNum = 1
      Do While Not ADO3.EOF
         wreg = IIf(IsNull(ADO3!regiongrupo), "", ADO3!regiongrupo)
         wNom = IIf(IsNull(ADO3!nomregiongrupo), "", ADO3!nomregiongrupo)
         wGra = IIf(IsNull(ADO3!gradogrupo), 0, ADO3!gradogrupo)
         wNomGra = IIf(IsNull(ADO3!nomgradogrupo), "", ADO3!nomgradogrupo)
         wNum = 1
         Do While IIf(IsNull(ADO3!regiongrupo), "", ADO3!regiongrupo) = wreg And IIf(IsNull(ADO3!gradogrupo), 0, ADO3!gradogrupo) = wGra
            DoEvents
            lblMensaje.Caption = "Traslando a EXCEL - Registro " + _
                                 Trim(Format(wRegAct, "####0")) + " / " + _
                                 Trim(Format(wRegTot, "####0"))
            lblMensaje.Refresh
            objExcel.Range(objExcel.Cells(V, H + 7), objExcel.Cells(V, H + 7)).NumberFormat = "####,##0.00"
         
            objExcel.Cells(V, H + 0) = ADO3!nomregiongrupo
            objExcel.Cells(V, H + 1) = ADO3!nomgradogrupo
            objExcel.Cells(V, H + 2) = wNum
            objExcel.Cells(V, H + 3) = Trim(ADO3!nomgra)
            objExcel.Cells(V, H + 4) = Trim(ADO3!nombre)
            If IsDate(ADO3!fecing) Then
               wFec = Format(ADO3!fecing, "dd/mm/yyyy")
               objExcel.Cells(V, H + 5) = wFec
            End If
            objExcel.Cells(V, H + 6) = ADO3!numdoc
            objExcel.Cells(V, H + 7) = ADO3!deuda_pt2
            objExcel.Cells(V, H + 8) = "_______________"
            objExcel.Cells(V, H + 9) = "_______________"
         
            wRegAct = wRegAct + 1
            wRegReg = wRegReg + 1
            V = V + 1
            wNum = wNum + 1
         
            ADO3.MoveNext
            If ADO3.EOF Then
               Exit Do
            End If
         Loop
         V = V + 1
         
         If ADO3.EOF Then
            objExcel.Cells(V, H + 4) = "TOTAL DE VOTANTES REGION " + wNom + " " + Format(wRegReg, "##,##0")
            V = V + 2
         Else
            If wreg <> ADO3!regiongrupo Then
               objExcel.Cells(V, H + 4) = "TOTAL DE VOTANTES REGION " + wNom + " " + Format(wRegReg, "##,##0")
               V = V + 2
               wRegReg = 0
            End If
         End If
      Loop
      V = V + 1
      objExcel.Cells(V, H + 4) = "TOTAL GENERAL DE VOTANTES " + Format(wRegTot, "##,##0")
      
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

Private Sub cmdExportar2_Click()
   On Error GoTo err
   
   Dim aa As Integer, wRegAct As Integer, wRegTot As Integer, wRegReg As Integer, _
       I As Integer, Heading(10) As String, _
       wNum As Integer, wFec As Date, wreg As String, wNom As String, _
       wDni As String, wGra As Integer, wNomGra As String, wDir As String, _
       wTel As String, wCel As String
       
   Heading(0) = "REGION"
   Heading(1) = "GRUPO"
   Heading(2) = "NUM"
   Heading(3) = "GRADO"
   Heading(4) = "TIPO SOCIO"
   Heading(5) = "NOMBRE ASOCIADO"
   Heading(6) = "D.N.I."
   Heading(7) = "DEUDA"
   Heading(8) = "DIRECCION"
   Heading(9) = "TELEFONO"
   Heading(10) = "CELULAR"
   aa = Leerado3("SELECT * FROM TMP_PADRON WHERE USU = '" + wcodusu + "' ORDER BY REGIONGRUPO, GRADOGRUPO, NOMBRE ")
   If aa > 0 Then
      wRegAct = 1: wRegReg = 1
      wRegTot = aa
      Set objExcel = New Excel.Application
      objExcel.SheetsInNewWorkbook = 1
      objExcel.Workbooks.Add
      With objExcel
           .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
           .Range(.Cells(3, 1), .Cells(3, 11)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 11)).Font.Bold = True
           .Cells(1, 1) = wnomcia
           .Cells(2, 1) = "PADRON GENERAL DE ASOCIADOS ACTIVOS HABILES"
           For I = 1 To 11 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           objExcel.Columns("A").ColumnWidth = 14
           objExcel.Columns("B").ColumnWidth = 15
           objExcel.Columns("C").ColumnWidth = 6
           objExcel.Columns("D").ColumnWidth = 15
           objExcel.Columns("E").ColumnWidth = 50
           objExcel.Columns("F").ColumnWidth = 11
           objExcel.Columns("G").ColumnWidth = 10
           objExcel.Columns("H").ColumnWidth = 11
           objExcel.Columns("I").ColumnWidth = 50
           objExcel.Columns("J").ColumnWidth = 18
           objExcel.Columns("K").ColumnWidth = 18
      End With
      V = 4
      H = 1
      wNum = 1
      Do While Not ADO3.EOF
         wreg = IIf(IsNull(ADO3!regiongrupo), "", ADO3!regiongrupo)
         wNom = IIf(IsNull(ADO3!nomregiongrupo), "", ADO3!nomregiongrupo)
         wGra = IIf(IsNull(ADO3!gradogrupo), 0, ADO3!gradogrupo)
         wNomGra = IIf(IsNull(ADO3!nomgradogrupo), "", ADO3!nomgradogrupo)
         wNum = 1
         Do While IIf(IsNull(ADO3!regiongrupo), "", ADO3!regiongrupo) = wreg And IIf(IsNull(ADO3!gradogrupo), 0, ADO3!gradogrupo) = wGra
            DoEvents
            lblMensaje.Caption = "Traslando a EXCEL - Registro " + _
                                 Trim(Format(wRegAct, "####0")) + " / " + _
                                 Trim(Format(wRegTot, "####0"))
            lblMensaje.Refresh
            objExcel.Range(objExcel.Cells(V, H + 7), objExcel.Cells(V, H + 7)).NumberFormat = "####,##0.00"
         
            wDir = "": wTel = "": wCel = 0
            aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(ADO3!codsocio) + " ")
            If aa > 0 Then
               wDir = ADO8!direc
               wTel = ADO8!telefono
               wCel = ADO8!celular
            End If
            Set ADO8 = Nothing
         
            objExcel.Cells(V, H + 0) = ADO3!nomregiongrupo
            objExcel.Cells(V, H + 1) = ADO3!nomgradogrupo
            objExcel.Cells(V, H + 2) = wNum
            objExcel.Cells(V, H + 3) = Trim(ADO3!nomgra)
            objExcel.Cells(V, H + 4) = ADO3!nome_socio
            objExcel.Cells(V, H + 5) = Trim(ADO3!nombre)
            objExcel.Cells(V, H + 6) = ADO3!numdoc
            objExcel.Cells(V, H + 7) = ADO3!deuda_pt2
            objExcel.Cells(V, H + 8) = wDir
            objExcel.Cells(V, H + 9) = wTel
            objExcel.Cells(V, H + 10) = wCel
         
            wRegAct = wRegAct + 1
            wRegReg = wRegReg + 1
            V = V + 1
            wNum = wNum + 1
         
            ADO3.MoveNext
            If ADO3.EOF Then
               Exit Do
            End If
         Loop
         V = V + 1
         
         If ADO3.EOF Then
            objExcel.Cells(V, H + 5) = "TOTAL DE VOTANTES REGION " + wNom + " " + Format(wRegReg, "##,##0")
            V = V + 2
         Else
            If wreg <> ADO3!regiongrupo Then
               objExcel.Cells(V, H + 5) = "TOTAL DE VOTANTES REGION " + wNom + " " + Format(wRegReg, "##,##0")
               V = V + 2
               wRegReg = 0
            End If
         End If
      Loop
      V = V + 1
      objExcel.Cells(V, H + 5) = "TOTAL GENERAL DE VOTANTES " + Format(wRegTot, "##,##0")
      
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

Private Sub cmdFormato2_Click()
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\PadronUniPag2.RPT"
   Crys1.SelectionFormula = " {TMP_PADRON.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdImprimir_Click()
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\PadronUniPag.RPT"
   Crys1.SelectionFormula = " {TMP_PADRON.USU}='" + wcodusu + "' "
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
   frmElePadronUnipag.Left = (Screen.Width - Width) \ 2
   frmElePadronUnipag.Top = 0
   
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
   aa = Leerado8("SELECT * FROM MAEE_SOCIO WHERE ELECCION = 1 ORDER BY E_SOCIO ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
   
         cmbE_Socio.AddItem ADO8!e_socio + " " + ADO8!nombre
   
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   cmbE_Socio.ListIndex = 0
   
   cmbRegion.Clear
   cmbRegion.AddItem "000 TODAS LAS REGIONES"
   aa = Leerado8("SELECT * FROM MAEREGIONGRUPO ORDER BY REGIONGRUPO ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
   
         cmbRegion.AddItem Str(ADO8!regiongrupo) + " " + ADO8!nombre
   
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   cmbRegion.ListIndex = 0
   
   cmdBuscar.SetFocus
End Sub

Private Sub LlenaCab()
   Dim aa As Long, wGra As Integer, wEso As String, wreg As Integer

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
   
   If cmbRegion.ListIndex = 0 Then
      wreg = 0
   Else
      wreg = Val(Left(cmbRegion.Text, 3))
   End If

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_PADRON WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_PADRON " _
   & " (CODSOCIO, CODIGO, INS, TIPDOC, NUMDOC, CARNETPNP, CARNETPIP, GRADO, SITU, " _
   & "  NOMBRE, E_SOCIO, FECRET, UNIPAG, REGION, FECING, DEUDA_PT2, NOMGRA, " _
   & "  NOME_SOCIO, NOMUNIPAG, NOMREGION, USU ) " _
   & " SELECT " _
   & "  CODSOCIO, CODIGO, INS, '01', NUMDOC, CARNETPNP, CARNETPIP, GRADO, SITU, " _
   & "  M.NOMBRE, M.E_SOCIO, FECRET, UNIPAG, REGION, FECING, DEUDA_PT2, " _
   & "  '', '', '', '', '" + wcodusu + "' " _
   & " FROM MAESOCIO AS M INNER JOIN MAEE_SOCIO AS E " _
   & "   ON M.E_SOCIO = E.E_SOCIO " _
   & " WHERE E.ELECCION = 1 ")
   Db.CommitTrans

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

   If wreg <> 0 Then
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_PADRON WHERE USU = '" + wcodusu + "' AND REGIONGRUPO <> " + Str(wreg) + " ")
      Db.CommitTrans
   End If

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_PADRON " _
   & " SET NOMREGIONGRUPO = R.NOMBRE " _
   & " FROM TMP_PADRON AS T INNER JOIN MAEREGIONGRUPO AS R " _
   & "   ON T.REGIONGRUPO = R.REGIONGRUPO " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans

   lblMensaje.Caption = ""
   lblMensaje.Refresh

   aa = Leerado("SELECT NOMREGION, NOMGRA, CODIGO, INS, NOMBRE, " _
                & "     FECING, NUMDOC, DEUDA_PT2, NOME_SOCIO, " _
                & "     NOMUNIPAG, CODIGO, GRADOGRUPO, REGIONGRUPO " _
                & " FROM TMP_PADRON " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY REGIONGRUPO, GRADOGRUPO, NOMBRE ")
   Set DataGrid1.DataSource = ADO1
End Sub

Private Sub LlenaCab1()
   DataGrid1.Columns(0).Width = 1500
   DataGrid1.Columns(0).Alignment = dbgLeft
   DataGrid1.Columns(0).Caption = "REGION"
   
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
    
   DataGrid1.Columns(10).Visible = False
   DataGrid1.Columns(11).Visible = False
   DataGrid1.Columns(12).Visible = False
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
