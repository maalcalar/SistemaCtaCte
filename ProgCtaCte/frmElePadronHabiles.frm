VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmElePadronHabiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Padron Activos Habiles"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   13095
   Begin VB.CommandButton cmdMatriz 
      Caption         =   "Matricial"
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7560
      Width           =   975
   End
   Begin VB.ComboBox cmbGrado 
      Height          =   315
      ItemData        =   "frmElePadronHabiles.frx":0000
      Left            =   1230
      List            =   "frmElePadronHabiles.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   120
      Width           =   3615
   End
   Begin VB.ComboBox cmbE_socio 
      Height          =   315
      ItemData        =   "frmElePadronHabiles.frx":0004
      Left            =   1230
      List            =   "frmElePadronHabiles.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   8
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
      Left            =   5400
      TabIndex        =   7
      Top             =   300
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
      Left            =   11520
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdExporta 
      Caption         =   "&Exportar Activos"
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
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   9763
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
      Left            =   600
      Top             =   7320
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Grado"
      Height          =   195
      Index           =   2
      Left            =   630
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   480
      Width           =   945
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
Attribute VB_Name = "frmElePadronHabiles"
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
       I As Integer, Heading(13) As String, _
       wNum As Integer, wFec As Date, wEso As String, wNomEso As String, _
       wDni As String, wDir As String, wDis As String, wTel As String, _
       wCor As String, wSoc As Integer, wFor As String
       
   Heading(1) = "NUM"
   Heading(2) = "GRADO"
   Heading(3) = "NOMBRE ASOCIADO"
   Heading(4) = "FEC.ING"
   Heading(5) = "D.N.I."
   Heading(6) = "DEUDA"
   Heading(7) = "DIRECCION"
   Heading(8) = "DISTRITO"
   Heading(9) = "TELEFONO"
   Heading(10) = "CORREO"
   Heading(11) = "FORMA PAGO"
   Heading(12) = "FIRMA"
   Heading(13) = "IMPRESION DIGITAL"
   
   aa = Leerado3("SELECT * FROM TMP_PADRON where usu = '" + wcodusu + "' ORDER BY NOMBRE ")
   If aa > 0 Then
      wRegAct = 1
      wRegTot = aa
      Set objExcel = New Excel.Application
      objExcel.SheetsInNewWorkbook = 1
      objExcel.Workbooks.Add
      With objExcel
           .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
           .Range(.Cells(3, 1), .Cells(3, 14)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 14)).Font.Bold = True
           .Cells(1, 1) = wnomcia
           .Cells(2, 1) = "PADRON GENERAL DE ASOCIADOS ACTIVOS HABILES"
           For I = 1 To 14 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           objExcel.Columns("A").ColumnWidth = 12
           objExcel.Columns("B").ColumnWidth = 6
           objExcel.Columns("C").ColumnWidth = 15
           objExcel.Columns("D").ColumnWidth = 50
           objExcel.Columns("E").ColumnWidth = 11
           objExcel.Columns("F").ColumnWidth = 10
           objExcel.Columns("G").ColumnWidth = 11
           objExcel.Columns("H").ColumnWidth = 70
           objExcel.Columns("I").ColumnWidth = 30
           objExcel.Columns("J").ColumnWidth = 25
           objExcel.Columns("K").ColumnWidth = 35
           objExcel.Columns("L").ColumnWidth = 30
           objExcel.Columns("M").ColumnWidth = 18
           objExcel.Columns("N").ColumnWidth = 18
      End With
      V = 4
      H = 1
      wNum = 1
      Do While Not ADO3.EOF
         DoEvents
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + _
                              Trim(Format(wRegAct, "####0")) + " / " + _
                              Trim(Format(wRegTot, "####0"))
         lblMensaje.Refresh
         
         wSoc = ADO3!codsocio
         wDir = "": wDis = "": wTel = "": wCor = "": wFor = ""
         aa = Leerado7a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
         If aa > 0 Then
            wDir = IIf(IsNull(ADO7a!direc), "", ADO7a!direc)
            wDis = IIf(IsNull(ADO7a!ubigeo), "", ADO7a!ubigeo)
            wTel = IIf(IsNull(ADO7a!telefono), "", ADO7a!telefono) + " " + _
                   IIf(IsNull(ADO7a!telefon2), "", ADO7a!telefon2) + " " + _
                   IIf(IsNull(ADO7a!celular), "", ADO7a!celular)
            wCor = IIf(IsNull(ADO7a!email), "", ADO7a!email)
            Select Case ADO7a!tipcob
            Case "01"
                 wFor = "DIECO"
            Case "02"
                 wFor = "CAJA MILITAR POLICIAL"
            Case "03"
                 wFor = "TESORERIA AOPIP"
            End Select
         End If
         Set ADO7 = Nothing
         If Len(Trim(wDis)) > 0 Then
            aa = Leerado7a("SELECT * FROM MAEUBIGEO WHERE CODIGO = '" + wDis + "' ")
            If aa > 0 Then
               wDis = ADO7a!nombre
            End If
            Set ADO7a = Nothing
         End If
         
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
         objExcel.Cells(V, H + 7) = wDir
         objExcel.Cells(V, H + 8) = wDis
         objExcel.Cells(V, H + 9) = wTel
         objExcel.Cells(V, H + 10) = wCor
         objExcel.Cells(V, H + 11) = wFor
         objExcel.Cells(V, H + 12) = "_______________"
         objExcel.Cells(V, H + 13) = "_______________"
         
         wRegAct = wRegAct + 1
         V = V + 1
         wNum = wNum + 1
         ADO3.MoveNext
      Loop
      
      V = V + 1
      objExcel.Cells(V, H + 3) = "TOTAL GENERAL DE VOTANTES " + Format(wRegTot, "##,##0")
      
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
   Crys1.ReportFileName = xraiz + "ReportCtaCte_31102018\PadronActivos.RPT"
   Crys1.SelectionFormula = " {TMP_PADRON.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdMatriz_Click()
   On Error GoTo err
   Dim filepdt1 As String, lin As Integer, pag As Long, wreg As Integer, resp As Integer, wTip As String
   filepdt1 = xraiz + "PDT\Padron.TXT"
   filebloc = "Padron"
   If Len(Dir$(filepdt1)) Then
      Kill filepdt1
   End If
   
   Dim numreg As Integer
    
   numreg = Leerado7("SELECT * FROM TMP_PADRON WHERE USU = '" + wcodusu + "' ORDER BY NOMBRE ")
   If numreg = 0 Then
      MsgBox "Archivo No Tiene Registros", vbInformation
      Exit Sub
   End If
   
   Open filepdt1 For Output As #1
   ADO7.MoveFirst
   pag = 0
   lin = 66
   wreg = 1
   Dim wCod As String, wNom As String, wRuc As String
   wTip = ADO7!sorteo
   Do While Not ADO7.EOF
      If lin > 58 Then
         pag = pag + 1
         MatrizTit (pag)
         lin = 8
      End If
      If wTip <> ADO7!sorteo Then
         pag = pag + 1
         MatrizTit (pag)
         lin = 8
      End If
      wTip = ADO7!sorteo
      Print #1, ""
      lin = lin + 1
      Print #1, ""
      lin = lin + 1
      Print #1, ""
      lin = lin + 1
      Print #1, ""
      lin = lin + 1
      Print #1, Format(wreg, "@@@@") + "  " + LlenaDat(ADO7!nombre, 40) + " " + _
                Format(ADO7!fecing, "dd/mm/yyyy") + " " + ADO7!numdoc + "  " + _
                String(20, "_") + "  " + _
                String(20, "_") + "  " + _
                String(20, "_")
      lin = lin + 1
      
      wreg = wreg + 1
      ADO7.MoveNext
   Loop
   Close #1
   Set ADO7 = Nothing
   FileReport = filepdt1
   Shell "notepad " & filepdt1, vbNormalFocus
   Kill filepdt1

   Do While True
      apli = FindWindow(vbNullString, filebloc & " - Bloc de Notas")
      If apli = 0 Then
         Exit Do
      End If
   Loop

   If MsgBox("Desea Imprimirlo???..", vbYesNo, "Imprimir Padrón") = vbYes Then
      printmatriz
   End If
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub MatrizTit(pag As Integer)
   Print #1, wnomcia
   Print #1, "RUC " + wruccia + Space(107) + Format(Date, "dd/mm/yyyy")
   Print #1, "PADRON ELECTORAL DE ASOCIADOS ACTIVOS HABILES ELECCIONES GENERALES DICIEMBRE " + wanocia + _
             " DIRECTIVOS AOPIP PERIODO ENE" + _
             Format(Val(wanocia) + 1, "0000") + _
             "-ENE" + Format(Val(wanocia) + 3, "0000") + _
             "(02 AÑOS)"
   Print #1, "MESA NRO " + ADO7!sorteo + String(100, " ") + "PAGINA : " + Format(pag, "###0")
   Print #1, String(132, "=")
   Print #1, "  CODIGO   " + Space(10) + String(22, " ") + "NOMBRE" + String(22, " ")
   Print #1, String(132, "=")
End Sub

Private Sub printmatriz()
    Dim a As String, numreg As Integer, wreg As Long, wTip As String
    numreg = Leerado7("SELECT * FROM TMP_PADRON WHERE USU = '" + wcodusu + "' ORDER BY NOMBRE ")
    If numreg = 0 Then
       MsgBox "Archivo No Tiene Registros", vbInformation
       Exit Sub
    End If
    
    CommonDialog1.CancelError = True
    On Error GoTo ErrManejo
    CommonDialog1.ShowPrinter
       
    Set fimpresora2 = New frmImpresora2
    fimpresora2.Show vbModal
    If Not fimpresora2.OK Then
       Exit Sub
    End If
    Unload fimpresora2
       
    Printer.FontName = "Draft 12cpi"
'    Printer.PrintQuality = -1
    Dim fila As Double, pag As Integer
    
    Dim wCod As String, wNom As String
    fila = 66
    pag = 1
    wreg = 1
    
    
   ADO7.MoveFirst
   wTip = ADO7!sorteo
   Do While Not ADO7.EOF
      If fila > 58 Then
         printmatriztit pag, desdeprint, hastaprint, todoprint
         pag = pag + 1
         fila = 8
      End If
      If wTip <> ADO7!sorteo Then
         printmatriztit pag, desdeprint, hastaprint, todoprint
         pag = pag + 1
         lin = 8
      End If
      wTip = ADO7!sorteo
      If todoprint = True Or (pag - 1 >= desdeprint And pag - 1 <= hastaprint) Then
         a = IXY(0, fila, String(1, " "))
      End If
      fila = fila + 1
      If todoprint = True Or (pag - 1 >= desdeprint And pag - 1 <= hastaprint) Then
         a = IXY(0, fila, String(1, " "))
      End If
      fila = fila + 1
      If todoprint = True Or (pag - 1 >= desdeprint And pag - 1 <= hastaprint) Then
         a = IXY(0, fila, String(1, " "))
      End If
      fila = fila + 1
      If todoprint = True Or (pag - 1 >= desdeprint And pag - 1 <= hastaprint) Then
         a = IXY(0, fila, String(1, " "))
      End If
      fila = fila + 1
      If todoprint = True Or (pag - 1 >= desdeprint And pag - 1 <= hastaprint) Then
         a = IXY(0, fila, Format(wreg, "@@@@") + "  " + LlenaDat(ADO7!nombre, 40) + " " + _
                          Format(ADO7!fecing, "dd/mm/yyyy") + " " + ADO7!numdoc + "  " + _
                          String(20, "_") + "  " + _
                          String(20, "_") + "  " + _
                          String(20, "_"))
      End If
      fila = fila + 1
      
      wreg = wreg + 1
      ADO7.MoveNext
   Loop
    Set ADO7 = Nothing
    Exit Sub
ErrManejo:
    MsgBox "Se Cancela La Opción de Impresión", vbExclamation
    Exit Sub
End Sub

Private Sub printmatriztit(pag As Integer, desde As Integer, hasta As Integer, todo As Boolean)
   Dim a As String
   If todo = True Or (pag >= desde And pag <= hasta) Then
      If pag > 0 Then Printer.EndDoc
      a = IXY(0, 1, wnomcia)
      a = IXY(0, 2, "RUC " + wruccia + Space(107) + Format(Date, "dd/mm/yyyy"))
      a = IXY(0, 3, "PADRON ELECTORAL DE ASOCIADOS ACTIVOS HABILES ELECCIONES GENERALES DICIEMBRE " + wanocia + _
                    " DIRECTIVOS AOPIP PERIODO ENE" + _
                    Format(Val(wanocia) + 1, "0000") + _
                    "-ENE" + Format(Val(wanocia) + 3, "0000") + _
                    "(02 AÑOS)")
      a = IXY(0, 4, "MESA NRO " + ADO7!sorteo + String(106, " ") + "PAGINA : " + Format(pag, "###0"))
      a = IXY(0, 5, String(132, "="))
      a = IXY(0, 6, "NRO." + "  " + String(10, " ") + "APELLIDOS Y NOMBRES " + String(10, " ") + " " + _
                    " FEC.ING. " + " " + " D.N.I. " + "  " + _
                    " F   I   R   M   A  " + "  " + _
                    "  IMPRESION DIGITAL " + "  " + _
                    "    OBSERVACION     ")
      a = IXY(0, 7, String(132, "="))
   End If
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Form_Initialize
End Sub

Private Sub Form_Initialize()
   frmElePadronHabiles.Left = (Screen.Width - Width) \ 2
   frmElePadronHabiles.Top = 0
   
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
   
   cmdBuscar.SetFocus
End Sub

Private Sub LlenaCab()
   Dim aa As Long, wRegAct As Long, wRegTot As Long, wGra As Integer, wEso As String

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

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_PADRON " _
   & " SET NOMREGIONGRUPO = R.NOMBRE " _
   & " FROM TMP_PADRON AS T INNER JOIN MAEREGIONGRUPO AS R " _
   & "   ON T.REGIONGRUPO = R.REGIONGRUPO " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans

   lblMensaje.Caption = ""
   lblMensaje.Refresh
   
   Dim wSdo As Currency, wSoc As Integer, wMesTope As String
   
   aa = Leerado("SELECT CODSOCIO, NOMGRA, CODIGO, INS, NOMBRE, " _
                & "     FECING, NUMDOC, DEUDA_PT2, NOME_SOCIO, " _
                & "     NOMUNIPAG, NOMREGION " _
                & " FROM TMP_PADRON " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY NOMBRE ")
   If aa > 0 Then
      ADO1.MoveFirst
      wRegAct = 1
      wRegTot = aa
      Do While Not ADO1.EOF
         DoEvents
         lblMensaje.Caption = "Registro - " + _
                              Trim(Format(wRegAct, "####0")) + " / " + _
                              Trim(Format(wRegTot, "####0"))
         lblMensaje.Refresh
         
         
         wSoc = ADO1!codsocio
         wSdo = 0
         wSdo = SaldoFoto(wSoc, zMesTope)
     
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_PADRON " _
         & " SET DEUDA_PT2 = " + Str(wSdo) + " " _
         & " WHERE      USU = '" + wcodusu + "' AND " _
         & "       CODSOCIO = " + Str(wSoc) + " ")
         Db.CommitTrans
     
         wRegAct = wRegAct + 1
         ADO1.MoveNext
      Loop
   End If
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_PADRON " _
   & " WHERE USU = '" + wcodusu + "' AND " _
   & "       DEUDA_PT2 >= 50 ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_PADRON " _
   & " SET SORTEO = '1' + LEFT(NOMBRE,1) " _
   & " WHERE LEFT(NOMBRE,1) = 'A' OR " _
   & "       LEFT(NOMBRE,1) = 'B' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_PADRON " _
   & " SET SORTEO = '2' + LEFT(NOMBRE,1) " _
   & " WHERE LEFT(NOMBRE,1) = 'C' OR " _
   & "       LEFT(NOMBRE,1) = 'D' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_PADRON " _
   & " SET SORTEO = '3' + LEFT(NOMBRE,1) " _
   & " WHERE LEFT(NOMBRE,1) = 'E' OR " _
   & "       LEFT(NOMBRE,1) = 'F' OR " _
   & "       LEFT(NOMBRE,1) = 'G' OR " _
   & "       LEFT(NOMBRE,1) = 'H' OR " _
   & "       LEFT(NOMBRE,1) = 'I' OR " _
   & "       LEFT(NOMBRE,1) = 'J' OR " _
   & "       LEFT(NOMBRE,1) = 'K' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_PADRON " _
   & " SET SORTEO = '4' + LEFT(NOMBRE,1) " _
   & " WHERE LEFT(NOMBRE,1) = 'L' OR " _
   & "       LEFT(NOMBRE,1) = 'M' OR " _
   & "       LEFT(NOMBRE,1) = 'N' OR " _
   & "       LEFT(NOMBRE,1) = 'O' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_PADRON " _
   & " SET SORTEO = '5' + LEFT(NOMBRE,1) " _
   & " WHERE LEFT(NOMBRE,1) = 'P' OR " _
   & "       LEFT(NOMBRE,1) = 'Q' OR " _
   & "       LEFT(NOMBRE,1) = 'R' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_PADRON " _
   & " SET SORTEO = '6' + LEFT(NOMBRE,1) " _
   & " WHERE LEFT(NOMBRE,1) = 'S' OR " _
   & "       LEFT(NOMBRE,1) = 'T' OR " _
   & "       LEFT(NOMBRE,1) = 'U' OR " _
   & "       LEFT(NOMBRE,1) = 'V' OR " _
   & "       LEFT(NOMBRE,1) = 'W' OR " _
   & "       LEFT(NOMBRE,1) = 'X' OR " _
   & "       LEFT(NOMBRE,1) = 'Y' OR " _
   & "       LEFT(NOMBRE,1) = 'Z' ")
   Db.CommitTrans
   
   aa = Leerado("SELECT CODSOCIO, NOMGRA, CODIGO, INS, NOMBRE, " _
                & "     FECING, NUMDOC, DEUDA_PT2, NOME_SOCIO, " _
                & "     NOMUNIPAG, NOMREGION " _
                & " FROM TMP_PADRON " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY NOMBRE ")
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
