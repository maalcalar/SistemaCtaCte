VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConSocNoMorosos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relación de Socios Activos Habiles"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12930
   Begin VB.CommandButton cmdExpAlfa 
      Caption         =   "&Exportar Alfabetico"
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
      Left            =   8640
      TabIndex        =   15
      Top             =   8400
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
      Left            =   9960
      TabIndex        =   11
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar x Estado"
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
      Left            =   8640
      TabIndex        =   10
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
      Left            =   11280
      TabIndex        =   9
      Top             =   7800
      Width           =   1095
   End
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmConSocNoMorosos.frx":0000
      Left            =   1320
      List            =   "frmConSocNoMorosos.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   7335
   End
   Begin VB.ComboBox cmbE_Socio 
      Height          =   315
      ItemData        =   "frmConSocNoMorosos.frx":0004
      Left            =   1320
      List            =   "frmConSocNoMorosos.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   3375
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
      Left            =   5160
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   645
      Width           =   1095
   End
   Begin MSMask.MaskEdBox txtMes 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin Crystal.CrystalReport Crys1 
      Left            =   11760
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6135
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   10821
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
      Caption         =   "RELACION DE SOCIOS ACTIVOS HABILES"
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
   Begin MSMask.MaskEdBox txtMoroso 
      Height          =   285
      Left            =   3600
      TabIndex        =   13
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
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
      Left            =   2400
      TabIndex        =   14
      Top             =   480
      Width           =   1095
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
      TabIndex        =   12
      Top             =   8040
      Width           =   7575
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
      Left            =   11160
      TabIndex        =   8
      Top             =   7440
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
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   855
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
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   975
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
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1140
   End
End
Attribute VB_Name = "frmConSocNoMorosos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
   LlenaCab
End Sub

Private Sub cmdExpAlfa_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(18) As String, wreg As Integer, wTot As Integer
   Dim wNom As String, wNum1 As Long, wSoc As Integer, wFecIng As Date, wFecNac As Date, _
       wMes As String, wAno As String, wE_S As String, _
       wSdoSol As Currency, zSdoSol As Currency, _
       wSdoDol As Currency, zSdoDol As Currency
       
   wAno = Left(txtMoroso.Text, 4)
   wMes = Right(txtMoroso.Text, 2)
       
   Heading(0) = "TIPO"
   Heading(1) = "NRO."
   Heading(2) = "CODIGO"
   Heading(3) = "CODOFIN"
   Heading(4) = "D.N.I."
   Heading(5) = "APELLIDOS Y NOMBRES"
   Heading(6) = "GRADO"
   Heading(7) = "TELEFONOS"
   Heading(8) = "TELEFONOS2"
   Heading(9) = "CELULAR"
   Heading(10) = "CORREO ELECTRONICO"
   Heading(11) = "CORREO ELECTRONICO 2"
   Heading(12) = "DIRECCION"
   Heading(13) = "REFERENCIA"
   Heading(14) = "UBIGEO"
   Heading(15) = "MONEDA"
   Heading(16) = "S/. MOROSOS"
   Heading(17) = "US$ MOROSOS"
   Heading(18) = "OBSERVAC"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 19)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 19)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "RELACION DE SOCIOS ACTIVOS HABILES - MES " + Trim(funnommes(wMes)) + " " + wAno
        For I = 1 To 19 Step 1
            .Cells(3, I) = Heading(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 14
        objExcel.Columns("B").ColumnWidth = 6
        objExcel.Columns("C").ColumnWidth = 10
        objExcel.Columns("D").ColumnWidth = 10
        objExcel.Columns("E").ColumnWidth = 12
        objExcel.Columns("F").ColumnWidth = 50
        objExcel.Columns("G").ColumnWidth = 16
        objExcel.Columns("H").ColumnWidth = 10
        objExcel.Columns("I").ColumnWidth = 10
        objExcel.Columns("J").ColumnWidth = 10
        objExcel.Columns("K").ColumnWidth = 30
        objExcel.Columns("L").ColumnWidth = 30
        objExcel.Columns("M").ColumnWidth = 50
        objExcel.Columns("N").ColumnWidth = 50
        objExcel.Columns("O").ColumnWidth = 50
        objExcel.Columns("P").ColumnWidth = 6
        objExcel.Columns("Q").ColumnWidth = 12
        objExcel.Columns("R").ColumnWidth = 12
        objExcel.Columns("S").ColumnWidth = 20
   End With
   
   aa = Leerado3("SELECT * FROM TMP_SOCMOR WHERE USU = '" + wcodusu + "' ORDER BY NOMBRE ")
   If aa > 0 Then
      wTot = aa
      V = 4
      H = 1
      wNum1 = 1
      zSdoSol = 0: zSdoDol = 0
      Do While Not ADO3.EOF
         DoEvents
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wreg, "####0") + " / " + Format(wTot, "####0")
         lblMensaje.Refresh
            
         objExcel.Range(objExcel.Cells(V, H + 14), objExcel.Cells(V, H + 15)).NumberFormat = "####,##0.00;-####,##0.00;\ "
         
         objExcel.Cells(V, H + 0) = ADO3!nome_s
         objExcel.Cells(V, H + 1) = Format(wNum2, "##0")
         objExcel.Cells(V, H + 2) = ADO3!codsocio
         objExcel.Cells(V, H + 3) = Format(ADO3!codigo, "########") + "-" + Format(ADO3!ins, "#")
         objExcel.Cells(V, H + 4) = ADO3!numdoc
         objExcel.Cells(V, H + 5) = ADO3!nombre
         objExcel.Cells(V, H + 6) = ADO3!nomgra
         objExcel.Cells(V, H + 7) = Trim(ADO3!telefono)
         objExcel.Cells(V, H + 8) = Trim(ADO3!telefon2)
         objExcel.Cells(V, H + 9) = Trim(ADO3!celular)
         objExcel.Cells(V, H + 10) = ADO3!email
         objExcel.Cells(V, H + 11) = ADO3!email2
         objExcel.Cells(V, H + 12) = ADO3!direc
         objExcel.Cells(V, H + 13) = ADO3!refer
         objExcel.Cells(V, H + 14) = ADO3!nomgeo
         objExcel.Cells(V, H + 15) = IIf(ADO3!moneda = "S", "S/.", "US$")
         If ADO3!moneda = "S" Then
            objExcel.Cells(V, H + 16) = ADO3!sdonew
            wSdoSol = wSdoSol + ADO3!sdonew
            zSdoSol = zSdoSol + ADO3!sdonew
         Else
            objExcel.Cells(V, H + 17) = ADO3!sdonew
            wSdoDol = wSdoDol + ADO3!sdonew
            zSdoDol = zSdoDol + ADO3!sdonew
         End If
         objExcel.Cells(V, H + 18) = ADO3!observac
         V = V + 1
         wreg = wreg + 1
         ADO3.MoveNext
      Loop
      V = V + 1
      
      objExcel.Range(objExcel.Cells(V, H + 16), objExcel.Cells(V, H + 17)).NumberFormat = "####,##0.00;-####,##0.00;\ "
      objExcel.Range(objExcel.Cells(V, H + 14), objExcel.Cells(V, H + 17)).Font.Bold = True
      objExcel.Range(objExcel.Cells(V, H + 14), objExcel.Cells(V, H + 17)).Font.Color = RGB(255, 0, 0)
      objExcel.Range(objExcel.Cells(V, H + 14), objExcel.Cells(V, H + 17)).Borders.LineStyle = xlContinuous
      objExcel.Range(objExcel.Cells(V, H + 14), objExcel.Cells(V, H + 17)).Borders.Color = RGB(255, 0, 0)
      
      objExcel.Cells(V, H + 14) = "TOTALES FINALES"
      objExcel.Cells(V, H + 16) = zSdoSol
      objExcel.Cells(V, H + 17) = zSdoDol
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

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(18) As String, wreg As Integer, wTot As Integer
   Dim wNom As String, wNum1 As Long, wSoc As Integer, wFecIng As Date, wFecNac As Date, _
       wMes As String, wAno As String, wE_S As String, _
       wSdoSol As Currency, zSdoSol As Currency, _
       wSdoDol As Currency, zSdoDol As Currency
       
   wAno = Left(txtMoroso.Text, 4)
   wMes = Right(txtMoroso.Text, 2)
       
   Heading(0) = "TIPO"
   Heading(1) = "NRO."
   Heading(2) = "CODIGO"
   Heading(3) = "CODOFIN"
   Heading(4) = "D.N.I."
   Heading(5) = "APELLIDOS Y NOMBRES"
   Heading(6) = "GRADO"
   Heading(7) = "TELEFONOS"
   Heading(8) = "TELEFONOS2"
   Heading(9) = "CELULAR"
   Heading(10) = "CORREO ELECTRONICO"
   Heading(11) = "CORREO ELECTRONICO 2"
   Heading(12) = "DIRECCION"
   Heading(13) = "REFERENCIA"
   Heading(14) = "UBIGEO"
   Heading(15) = "MONEDA"
   Heading(16) = "S/. MOROSOS"
   Heading(17) = "US$ MOROSOS"
   Heading(18) = "OBSERVAC"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 19)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 19)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "RELACION DE SOCIOS ACTIVOS HABILES - MES " + Trim(funnommes(wMes)) + " " + wAno
        For I = 1 To 19 Step 1
            .Cells(3, I) = Heading(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 14
        objExcel.Columns("B").ColumnWidth = 6
        objExcel.Columns("C").ColumnWidth = 10
        objExcel.Columns("D").ColumnWidth = 11
        objExcel.Columns("E").ColumnWidth = 10
        objExcel.Columns("F").ColumnWidth = 50
        objExcel.Columns("G").ColumnWidth = 16
        objExcel.Columns("H").ColumnWidth = 10
        objExcel.Columns("I").ColumnWidth = 10
        objExcel.Columns("J").ColumnWidth = 10
        objExcel.Columns("K").ColumnWidth = 30
        objExcel.Columns("L").ColumnWidth = 30
        objExcel.Columns("M").ColumnWidth = 50
        objExcel.Columns("N").ColumnWidth = 50
        objExcel.Columns("O").ColumnWidth = 50
        objExcel.Columns("P").ColumnWidth = 6
        objExcel.Columns("Q").ColumnWidth = 12
        objExcel.Columns("R").ColumnWidth = 12
        objExcel.Columns("S").ColumnWidth = 20
   End With
   
   aa = Leerado3("SELECT * FROM TMP_SOCMOR WHERE USU = '" + wcodusu + "' ORDER BY NOME_S, NOMBRE ")
   If aa > 0 Then
      wTot = aa
      V = 4
      H = 1
      wNum1 = 1
      zSdoSol = 0: zSdoDol = 0
      Do While Not ADO3.EOF
         wE_S = ADO3!e_socio
         wNom = ADO3!nome_s
         wSdoSol = 0: wSdoDol = 0
         wNum2 = 1
         Do While wE_S = ADO3!e_socio
            DoEvents
            lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wreg, "####0") + " / " + Format(wTot, "####0")
            lblMensaje.Refresh
            
            objExcel.Range(objExcel.Cells(V, H + 15), objExcel.Cells(V, H + 16)).NumberFormat = "####,##0.00;-####,##0.00;\ "
         
            objExcel.Cells(V, H + 0) = ADO3!nome_s
            objExcel.Cells(V, H + 1) = Format(wNum2, "##0")
            objExcel.Cells(V, H + 2) = ADO3!codsocio
            objExcel.Cells(V, H + 3) = Format(ADO3!codigo, "########") + "-" + Format(ADO3!ins, "#")
            objExcel.Cells(V, H + 4) = ADO3!numdoc
            objExcel.Cells(V, H + 5) = ADO3!nombre
            objExcel.Cells(V, H + 6) = ADO3!nomgra
            objExcel.Cells(V, H + 7) = Trim(ADO3!telefono)
            objExcel.Cells(V, H + 8) = Trim(ADO3!telefon2)
            objExcel.Cells(V, H + 9) = Trim(ADO3!celular)
            objExcel.Cells(V, H + 10) = ADO3!email
            objExcel.Cells(V, H + 11) = ADO3!email2
            objExcel.Cells(V, H + 12) = ADO3!direc
            objExcel.Cells(V, H + 13) = ADO3!refer
            objExcel.Cells(V, H + 14) = ADO3!nomgeo
            objExcel.Cells(V, H + 15) = IIf(ADO3!moneda = "S", "S/.", "US$")
            If ADO3!moneda = "S" Then
               objExcel.Cells(V, H + 16) = ADO3!sdonew
               wSdoSol = wSdoSol + ADO3!sdonew
               zSdoSol = zSdoSol + ADO3!sdonew
            Else
               objExcel.Cells(V, H + 17) = ADO3!sdonew
               wSdoDol = wSdoDol + ADO3!sdonew
               zSdoDol = zSdoDol + ADO3!sdonew
            End If
            objExcel.Cells(V, H + 18) = ADO3!observac
            
            wNum2 = wNum2 + 1
            wreg = wreg + 1
            V = V + 1
            ADO3.MoveNext
            If ADO3.EOF Then
               Exit Do
            End If
         Loop
         objExcel.Range(objExcel.Cells(V, H + 16), objExcel.Cells(V, H + 17)).NumberFormat = "####,##0.00;-####,##0.00;\ "
         objExcel.Range(objExcel.Cells(V, H + 14), objExcel.Cells(V, H + 17)).Font.Bold = True
         objExcel.Range(objExcel.Cells(V, H + 14), objExcel.Cells(V, H + 17)).Font.Color = RGB(0, 0, 255)
         objExcel.Range(objExcel.Cells(V, H + 14), objExcel.Cells(V, H + 17)).Borders.LineStyle = xlContinuous
         objExcel.Range(objExcel.Cells(V, H + 14), objExcel.Cells(V, H + 17)).Borders.Color = RGB(0, 0, 255)
      
         objExcel.Cells(V, H + 14) = "TOTALES POR TIPO " + wE_S
         objExcel.Cells(V, H + 16) = wSdoSol
         objExcel.Cells(V, H + 17) = wSdoDol
         V = V + 2
         
         wNum1 = wNum1 + 1
      Loop
      V = V + 1
      
      objExcel.Range(objExcel.Cells(V, H + 15), objExcel.Cells(V, H + 16)).NumberFormat = "####,##0.00;-####,##0.00;\ "
      objExcel.Range(objExcel.Cells(V, H + 13), objExcel.Cells(V, H + 16)).Font.Bold = True
      objExcel.Range(objExcel.Cells(V, H + 13), objExcel.Cells(V, H + 16)).Font.Color = RGB(255, 0, 0)
      objExcel.Range(objExcel.Cells(V, H + 13), objExcel.Cells(V, H + 16)).Borders.LineStyle = xlContinuous
      objExcel.Range(objExcel.Cells(V, H + 13), objExcel.Cells(V, H + 16)).Borders.Color = RGB(255, 0, 0)
      
      objExcel.Cells(V, H + 13) = "TOTALES FINALES"
      objExcel.Cells(V, H + 15) = zSdoSol
      objExcel.Cells(V, H + 16) = zSdoDol
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
   Crys1.ReportFileName = xraiz + "ReportCtaCte\SociosHabiles.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= 'MOROSIDAD AL MES DE " + wMes + "' "
   Crys1.SelectionFormula = " {TMP_SOCMOR.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   Select Case ColIndex
   Case 0
        ADO2.Sort = "CODSOCIO"
   Case 1
        ADO2.Sort = "CODIGO"
   Case 3
        ADO2.Sort = "NOMBRE"
   Case 5
        ADO2.Sort = "NOMGRA"
   End Select
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
   
   cmbE_Socio.Clear
   cmbE_Socio.AddItem "Todos Los Estados de Socio"
   a = Leerado8("SELECT * FROM MAEE_SOCIO WHERE APORTE > 0 ORDER BY E_SOCIO ")
   If a > 0 Then
      ADO8.MoveFirst
      I = 0
      Do While Not ADO8.EOF
         cmbE_Socio.AddItem ADO8!nombre
         I = I + 1
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   cmbE_Socio.ListIndex = 0
   
   txtMes.Text = Left(zMesTope, 4) + "/" + Right(zMesTope, 2)
   
   txtMoroso.Text = Left(zMesTope, 4) + "/" + Right(zMesTope, 2)
   
'   If Right(zMesTope, 2) = "01" Then
'      txtMoroso.Text = Format(Val(Left(zMesTope, 4)) - 1, "0000") + "/" + "12"
'   Else
'      txtMoroso.Text = Left(zMesTope, 4) + "/" + Format(Val(Right(zMesTope, 2)) - 1, "00")
'   End If
   
   txtMes.Enabled = False
   txtMoroso.Enabled = False

   Set DataGrid1.DataSource = Nothing
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_SOCMOR WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   cmbE_Socio.SetFocus
End Sub

Private Sub Form_Load()
   frmConSocNoMorosos.Left = (Screen.Width - Width) \ 2
   frmConSocNoMorosos.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_SOCMOR WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim aa As Long, wreg As Long, wTot As Long, _
       w As String, wMes As String, wFec As Date, wE_S As String, _
       wSoc As Integer, wSdo As Currency, wNom As String

   wMes = Left(txtMoroso.Text, 4) + Right(txtMoroso.Text, 2)
   wFec = Format(fundiames(Right(wMes, 2)) + "/" + Right(wMes, 2) + "/" + Left(wMes, 4), "dd/mm/yyyy")
   wE_S = BuscaCodEsocio(cmbE_Socio.List(cmbE_Socio.ListIndex))

   w = ""
   If wE_S <> "" Then
      w = " AND S.E_SOCIO = '" + wE_S + "' "
   End If

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_SOCMOR WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_SOCMOR " _
   & " (USU, CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, NOME_S, GRADO, NOMGRA, " _
   & "  DIREC, TELEFONO, TELEFON2, CELULAR, EMAIL, EMAIL2, REFER, MONEDA, SDONEW, " _
   & "  UBIGEO, NOMGEO, NUMDOC, APORTE ) " _
   & " SELECT '" + wcodusu + "', S.CODSOCIO, S.CODIGO, S.INS, S.NOMBRE, S.E_SOCIO, " _
   & "        E.NOMBRE, S.GRADO, G.NOMBRE, S.DIREC, S.TELEFONO, S.TELEFON2, " _
   & "        S.CELULAR, S.EMAIL, S.EMAIL2, S.REFER, E.MONEDA, 0, S.UBIGEO, '', " _
   & "        S.NUMDOC, E.APORTE " _
   & " FROM MAESOCIO AS S INNER JOIN MAEGRADO   AS G ON S.GRADO = G.GRADO " _
   & "                    INNER JOIN MAEE_SOCIO AS E ON S.E_SOCIO = E.E_SOCIO " _
   & " WHERE E.APORTE > 0 " + w + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_SOCMOR " _
   & " SET NOMGEO = U.NOMBRE " _
   & " FROM TMP_SOCMOR AS T INNER JOIN MAEUBIGEO AS U " _
   & "   ON T.UBIGEO = U.CODIGO " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans

   aa = Leerado8a("SELECT * FROM TMP_SOCMOR WHERE USU = '" + wcodusu + "' ORDER BY NOMBRE ")
   If aa > 0 Then
      ADO8a.MoveFirst
      wreg = 1
      wTot = aa
      Do While Not ADO8a.EOF
         DoEvents
         lblMensaje.Caption = Trim(Format(wreg, "###,##0")) + " / " + _
                              Trim(Format(wTot, "###,##0"))
         lblMensaje.Refresh
         
         wSoc = ADO8a!codsocio
         wNom = Trim(ADO8a!nombre)
         wSdo = SaldoFoto(wSoc, wMes)
         
'         DoEvents
'         lblMensaje.Caption = "Calculando Socio " + wNom
'         lblMensaje.Refresh
   
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_SOCMOR " _
         & " SET SDONEW = " + Str(wSdo) + " " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "            USU = '" + wcodusu + "' ")
         Db.CommitTrans
   
         wreg = wreg + 1
         ADO8a.MoveNext
      Loop
   End If
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_SOCMOR WHERE USU = '" + wcodusu + "' AND SDONEW > 300")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_SOCMOR " _
   & " SET OBSERVAC = 'SIN DEUDAS' " _
   & " WHERE USU = '" + wcodusu + "' AND SDONEW = 0")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_SOCMOR " _
   & " SET OBSERVAC = 'DEUDAS HASTA 3 MESES' " _
   & " WHERE USU = '" + wcodusu + "' AND SDONEW > 0 AND SDONEW <= ROUND(APORTE * 3, 2) ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_SOCMOR " _
   & " SET OBSERVAC = 'DEUDAS HASTA 6 MESES' " _
   & " WHERE USU = '" + wcodusu + "' AND SDONEW > ROUND(APORTE * 3,2) AND SDONEW <= ROUND(APORTE * 6,2) ")
   Db.CommitTrans
   
   DoEvents
   lblMensaje.Caption = ""
   lblMensaje.Refresh
   
   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, NOMGRA, NOME_S, MONEDA, SDONEW, UBIGEO, NOMGEO " _
            & " FROM TMP_SOCMOR " _
            & " WHERE USU = '" + wcodusu + "' " _
            & " ORDER BY E_SOCIO, NOMBRE ")
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
    
   DataGrid1.Columns(6).Width = 450   ' MONEDA
   DataGrid1.Columns(6).Alignment = dbgCenter
   DataGrid1.Columns(6).Caption = "MON"
   
   DataGrid1.Columns(7).Width = 1000  ' SDONEW
   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).Caption = "SALDOS"
   DataGrid1.Columns(7).NumberFormat = "#####0.00;;\ "

   DataGrid1.Columns(8).Visible = False
   DataGrid1.Columns(9).Visible = False
End Sub

