VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmConAporteAnual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen Aportes Anuales"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   14370
   Begin VB.CheckBox chkCiviles 
      Caption         =   "CIVILES"
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
      Left            =   3600
      TabIndex        =   12
      Top             =   960
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox chkCajaMP 
      Caption         =   "CAJA MILITAR"
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
      Left            =   1680
      TabIndex        =   11
      Top             =   960
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkDieco 
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
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Value           =   1  'Checked
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
      Left            =   12720
      TabIndex        =   8
      Top             =   7320
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
      Left            =   10080
      TabIndex        =   7
      Top             =   7320
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
      Left            =   11400
      TabIndex        =   6
      Top             =   7320
      Width           =   1095
   End
   Begin VB.ComboBox cmbE_Socio 
      Height          =   315
      ItemData        =   "frmConAporteAnual.frx":0000
      Left            =   1440
      List            =   "frmConAporteAnual.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtAno 
      Height          =   305
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin Crystal.CrystalReport Crys1 
      Left            =   10200
      Top             =   240
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
      Height          =   5535
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   13815
      _ExtentX        =   24368
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
      TabIndex        =   9
      Top             =   7560
      Width           =   9135
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
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   1140
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
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmConAporteAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
   LlenaCab
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(23) As String, Headin2(23) As String, _
       wRegAct As Integer, wRegTot As Integer, wAno As String, _
       wPago01 As Currency, wPago02 As Currency, wPago03 As Currency, wPago04 As Currency, _
       wPago05 As Currency, wPago06 As Currency, wPago07 As Currency, wPago08 As Currency, _
       wPago09 As Currency, wPago10 As Currency, wPago11 As Currency, wPago12 As Currency, _
       wTotpag As Currency, wTotDeu As Currency, wTotAno As Currency, _
       zPago01 As Currency, zPago02 As Currency, zPago03 As Currency, zPago04 As Currency, _
       zPago05 As Currency, zPago06 As Currency, zPago07 As Currency, zPago08 As Currency, _
       zPago09 As Currency, zPago10 As Currency, zPago11 As Currency, zPago12 As Currency, _
       zTotpag As Currency, zTotDeu As Currency, zTotAno As Currency, _
       dPago01 As Currency, dPago02 As Currency, dPago03 As Currency, dPago04 As Currency, _
       dPago05 As Currency, dPago06 As Currency, dPago07 As Currency, dPago08 As Currency, _
       dPago09 As Currency, dPago10 As Currency, dPago11 As Currency, dPago12 As Currency, _
       dTotpag As Currency, dTotDeu As Currency, dTotAno As Currency, _
       WE_S As String, wNomE_S As String, wFec As Date

   wAno = txtAno.Text
   
   Heading(0) = "E S T A D O"
   Heading(2) = "CODIGO"
   Heading(3) = "CODIGO"
   Heading(4) = "INS"
   Heading(5) = "APELLIDOS Y NOMBRES DEL ASOCIADO"
   Heading(6) = "FECHA"
   Heading(7) = "MON"
   Heading(8) = "APORTE"
   Heading(9) = "APORTE"
   
   Heading(10) = "COBRANZAS DEL EJERCICIO"
   Heading(22) = "TOTAL"
   Heading(23) = "TOTAL"
   
   Headin2(0) = "CODIGO"
   Headin2(1) = "NOMBRE"
   Headin2(2) = "SOCIO"
   Headin2(3) = "CODOFIN"
   Headin2(6) = "INGRESO"
   Headin2(7) = "EDA"
   Headin2(8) = "MENSUAL"
   Headin2(9) = "ANUAL"
   Headin2(10) = "ENE"
   Headin2(11) = "FEB"
   Headin2(12) = "MAR"
   Headin2(13) = "ABR"
   Headin2(14) = "MAY"
   Headin2(15) = "JUN"
   Headin2(16) = "JUL"
   Headin2(17) = "AGO"
   Headin2(18) = "SET"
   Headin2(19) = "OCT"
   Headin2(20) = "NOV"
   Headin2(21) = "DIC"
   Headin2(22) = "COBROS"
   Headin2(23) = "DEUDA"
   
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(4, 24)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(4, 24)).Font.Bold = True
        
        
        .Range(objExcel.Cells(2, 1), .Cells(2, 24)).Merge
        .Range(objExcel.Cells(2, 1), .Cells(2, 24)).HorizontalAlignment = xlCenter
        
        .Range(objExcel.Cells(3, 1), .Cells(3, 2)).Merge
        .Range(objExcel.Cells(3, 1), .Cells(3, 24)).HorizontalAlignment = xlCenter
        
        .Range(objExcel.Cells(3, 11), .Cells(3, 22)).Merge
        .Range(objExcel.Cells(3, 11), .Cells(3, 22)).HorizontalAlignment = xlCenter
        
        .Range(objExcel.Cells(4, 1), .Cells(4, 24)).HorizontalAlignment = xlCenter
        
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "RESUMEN DE APORTACIONES CIVILES - EJERCICIO " + wAno
        For I = 1 To 24 Step 1
            .Cells(3, I) = Heading(I - 1)
            .Cells(4, I) = Headin2(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 7
        objExcel.Columns("B").ColumnWidth = 14
        objExcel.Columns("C").ColumnWidth = 7
        objExcel.Columns("D").ColumnWidth = 9
        objExcel.Columns("E").ColumnWidth = 3
        objExcel.Columns("F").ColumnWidth = 60
        objExcel.Columns("G").ColumnWidth = 11
        objExcel.Columns("H").ColumnWidth = 6
        objExcel.Columns("I").ColumnWidth = 10
        objExcel.Columns("J").ColumnWidth = 10
        objExcel.Columns("K").ColumnWidth = 11
        objExcel.Columns("L").ColumnWidth = 11
        objExcel.Columns("M").ColumnWidth = 11
        objExcel.Columns("N").ColumnWidth = 11
        objExcel.Columns("O").ColumnWidth = 11
        objExcel.Columns("P").ColumnWidth = 11
        objExcel.Columns("Q").ColumnWidth = 11
        objExcel.Columns("R").ColumnWidth = 11
        objExcel.Columns("S").ColumnWidth = 11
        objExcel.Columns("T").ColumnWidth = 11
        objExcel.Columns("U").ColumnWidth = 11
        objExcel.Columns("V").ColumnWidth = 11
        objExcel.Columns("W").ColumnWidth = 11
        objExcel.Columns("X").ColumnWidth = 11
        objExcel.Columns("Y").ColumnWidth = 11
        objExcel.Columns("Z").ColumnWidth = 11
   End With
   
   aa = Leerado3("SELECT * " _
                & " FROM TMP_ANUAL_APORTE " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY E_SOCIO, NOMBRE ")
   If aa > 0 Then
      wRegTot = aa
      V = 5
      H = 1
      wRegAct = 1
      wPago01 = 0: wPago02 = 0: wPago03 = 0: wPago04 = 0: wPago05 = 0: wPago06 = 0
      wPago07 = 0: wPago08 = 0: wPago09 = 0: wPago10 = 0: wPago11 = 0: wPago12 = 0
      dPago01 = 0: dPago02 = 0: dPago03 = 0: dPago04 = 0: dPago05 = 0: dPago06 = 0
      dPago07 = 0: dPago08 = 0: dPago09 = 0: dPago10 = 0: dPago11 = 0: dPago12 = 0
      
      Do While Not ADO3.EOF
         WE_S = ADO3!e_socio
         wNomE_S = ""
         
         aa = Leerado8("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + WE_S + "' ")
         If aa > 0 Then
            wNomE_S = ADO8!nombre
         End If
         Set ADO8 = Nothing
         objExcel.Cells(V, H + 0) = WE_S
         objExcel.Cells(V, H + 1) = wNomE_S
         V = V + 1
         zPago01 = 0: zPago02 = 0: zPago03 = 0: zPago04 = 0: zPago05 = 0: zPago06 = 0
         zPago07 = 0: zPago08 = 0: zPago09 = 0: zPago10 = 0: zPago11 = 0: zPago12 = 0
         zTotAno = 0: zTotpag = 0: zTotDeu = 0
         Do While ADO3!e_socio = WE_S
            DoEvents
            lblMensaje.Caption = "Traslando Resumen a EXCEL - Registro " + _
                                 Format(wRegAct, "####0") + " / " + _
                                 Format(wRegTot, "####0")
            lblMensaje.Refresh
         
            objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 23)).NumberFormat = "######0.00;-######0.00;\ "
         
            objExcel.Cells(V, H + 0) = WE_S
            objExcel.Cells(V, H + 1) = wNomE_S
            objExcel.Cells(V, H + 2) = ADO3!codsocio
            objExcel.Cells(V, H + 3) = ADO3!codigo
            objExcel.Cells(V, H + 4) = ADO3!ins
            objExcel.Cells(V, H + 5) = ADO3!nombre
            If IsDate(ADO3!fecing) Then
               wFec = Format(ADO3!fecing, "dd/mm/yyyy")
               objExcel.Cells(V, H + 6) = wFec
            End If
            objExcel.Cells(V, H + 7) = ADO3!moneda
            objExcel.Cells(V, H + 8) = ADO3!aporte
            objExcel.Cells(V, H + 9) = ADO3!totano
            objExcel.Cells(V, H + 10) = ADO3!pago01
            objExcel.Cells(V, H + 11) = ADO3!pago02
            objExcel.Cells(V, H + 12) = ADO3!pago03
            objExcel.Cells(V, H + 13) = ADO3!pago04
            objExcel.Cells(V, H + 14) = ADO3!pago05
            objExcel.Cells(V, H + 15) = ADO3!pago06
            objExcel.Cells(V, H + 16) = ADO3!pago07
            objExcel.Cells(V, H + 17) = ADO3!pago08
            objExcel.Cells(V, H + 18) = ADO3!pago09
            objExcel.Cells(V, H + 19) = ADO3!pago10
            objExcel.Cells(V, H + 20) = ADO3!pago11
            objExcel.Cells(V, H + 21) = ADO3!pago12
            objExcel.Cells(V, H + 22) = ADO3!totpag
            objExcel.Cells(V, H + 23) = ADO3!totdeu
         
            If ADO3!moneda = "S" Then
               wPago01 = wPago01 + ADO3!pago01
               wPago02 = wPago02 + ADO3!pago02
               wPago03 = wPago03 + ADO3!pago03
               wPago04 = wPago04 + ADO3!pago04
               wPago05 = wPago05 + ADO3!pago05
               wPago06 = wPago06 + ADO3!pago06
               wPago07 = wPago07 + ADO3!pago07
               wPago08 = wPago08 + ADO3!pago08
               wPago09 = wPago09 + ADO3!pago09
               wPago10 = wPago10 + ADO3!pago10
               wPago11 = wPago11 + ADO3!pago11
               wPago12 = wPago12 + ADO3!pago12
               wTotAno = wTotAno + ADO3!totano
               wTotpag = wTotpag + ADO3!totpag
               wTotDeu = wTotDeu + ADO3!totdeu
            Else
               dPago01 = dPago01 + ADO3!pago01
               dPago02 = dPago02 + ADO3!pago02
               dPago03 = dPago03 + ADO3!pago03
               dPago04 = dPago04 + ADO3!pago04
               dPago05 = dPago05 + ADO3!pago05
               dPago06 = dPago06 + ADO3!pago06
               dPago07 = dPago07 + ADO3!pago07
               dPago08 = dPago08 + ADO3!pago08
               dPago09 = dPago09 + ADO3!pago09
               dPago10 = dPago10 + ADO3!pago10
               dPago11 = dPago11 + ADO3!pago11
               dPago12 = dPago12 + ADO3!pago12
               dTotAno = dTotAno + ADO3!totano
               dTotpag = dTotpag + ADO3!totpag
               dTotDeu = dTotDeu + ADO3!totdeu
            End If
         
            zPago01 = zPago01 + ADO3!pago01
            zPago02 = zPago02 + ADO3!pago02
            zPago03 = zPago03 + ADO3!pago03
            zPago04 = zPago04 + ADO3!pago04
            zPago05 = zPago05 + ADO3!pago05
            zPago06 = zPago06 + ADO3!pago06
            zPago07 = zPago07 + ADO3!pago07
            zPago08 = zPago08 + ADO3!pago08
            zPago09 = zPago09 + ADO3!pago09
            zPago10 = zPago10 + ADO3!pago10
            zPago11 = zPago11 + ADO3!pago11
            zPago12 = zPago12 + ADO3!pago12
            zTotAno = zTotAno + ADO3!totano
            zTotpag = zTotpag + ADO3!totpag
            zTotDeu = zTotDeu + ADO3!totdeu
         
            wRegAct = wRegAct + 1
            V = V + 1
            ADO3.MoveNext
            If ADO3.EOF Then
               Exit Do
            End If
         Loop
         V = V + 1
         objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 23)).NumberFormat = "######0.00;-######0.00;\ "
         objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 23)).Font.Bold = True
         objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 23)).Font.Color = RGB(0, 0, 255)
         objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 23)).Borders.LineStyle = xlContinuous
         objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 23)).Borders.Color = RGB(0, 0, 255)
      
         objExcel.Cells(V, H + 5) = "TOTALES " + WE_S
         objExcel.Cells(V, H + 9) = zTotAno
         objExcel.Cells(V, H + 10) = zPago01
         objExcel.Cells(V, H + 11) = zPago02
         objExcel.Cells(V, H + 12) = zPago03
         objExcel.Cells(V, H + 13) = zPago04
         objExcel.Cells(V, H + 14) = zPago05
         objExcel.Cells(V, H + 15) = zPago06
         objExcel.Cells(V, H + 16) = zPago07
         objExcel.Cells(V, H + 17) = zPago08
         objExcel.Cells(V, H + 18) = zPago09
         objExcel.Cells(V, H + 19) = zPago10
         objExcel.Cells(V, H + 20) = zPago11
         objExcel.Cells(V, H + 21) = zPago12
         objExcel.Cells(V, H + 22) = zTotpag
         objExcel.Cells(V, H + 23) = zTotDeu
         V = V + 1
      Loop
      V = V + 1
         
      objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 23)).NumberFormat = "######0.00;-######0.00;\ "
      objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 23)).Font.Bold = True
      objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 23)).Font.Color = RGB(255, 0, 0)
      objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 23)).Borders.LineStyle = xlContinuous
      objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 23)).Borders.Color = RGB(255, 0, 0)
      
      objExcel.Cells(V, H + 5) = "TOTALES FINALES"
      objExcel.Cells(V, H + 8) = "S/."
      objExcel.Cells(V, H + 9) = wTotAno
      objExcel.Cells(V, H + 10) = wPago01
      objExcel.Cells(V, H + 11) = wPago02
      objExcel.Cells(V, H + 12) = wPago03
      objExcel.Cells(V, H + 13) = wPago04
      objExcel.Cells(V, H + 14) = wPago05
      objExcel.Cells(V, H + 15) = wPago06
      objExcel.Cells(V, H + 16) = wPago07
      objExcel.Cells(V, H + 17) = wPago08
      objExcel.Cells(V, H + 18) = wPago09
      objExcel.Cells(V, H + 19) = wPago10
      objExcel.Cells(V, H + 20) = wPago11
      objExcel.Cells(V, H + 21) = wPago12
      objExcel.Cells(V, H + 22) = wTotpag
      objExcel.Cells(V, H + 23) = wTotDeu
      V = V + 1
      
      objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 23)).NumberFormat = "######0.00;-######0.00;\ "
      objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 23)).Font.Bold = True
      objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 23)).Font.Color = RGB(255, 0, 0)
      objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 23)).Borders.LineStyle = xlContinuous
      objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 23)).Borders.Color = RGB(255, 0, 0)
      
      objExcel.Cells(V, H + 5) = "TOTALES FINALES"
      objExcel.Cells(V, H + 8) = "US$"
      objExcel.Cells(V, H + 9) = dTotAno
      objExcel.Cells(V, H + 10) = dPago01
      objExcel.Cells(V, H + 11) = dPago02
      objExcel.Cells(V, H + 12) = dPago03
      objExcel.Cells(V, H + 13) = dPago04
      objExcel.Cells(V, H + 14) = dPago05
      objExcel.Cells(V, H + 15) = dPago06
      objExcel.Cells(V, H + 16) = dPago07
      objExcel.Cells(V, H + 17) = dPago08
      objExcel.Cells(V, H + 18) = dPago09
      objExcel.Cells(V, H + 19) = dPago10
      objExcel.Cells(V, H + 20) = dPago11
      objExcel.Cells(V, H + 21) = dPago12
      objExcel.Cells(V, H + 22) = dTotpag
      objExcel.Cells(V, H + 23) = dTotDeu
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
   wAno = txtAno.Text
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\AportesxAno.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= 'EJERCICIO " + wAno + "'"
   Crys1.SelectionFormula = " {TMP_ANUAL_APORTE.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   txtAno.Text = wanocia
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_ANUAL_APORTE WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   cmbE_Socio.Clear
   cmbE_Socio.AddItem "Todos Los Estados de Socio"
   a = Leerado8("SELECT * FROM MAEE_SOCIO ORDER BY E_SOCIO ")
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
   
   txtAno.SetFocus
End Sub

Private Sub Form_Load()
   frmConAporteAnual.Left = (Screen.Width - Width) \ 2
   frmConAporteAnual.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_ANUAL_APORTE WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim aa As Long, w As String, WE_S As String, wAno As String, II As Integer, wmmm As String

   wAno = txtAno.Text
   WE_S = BuscaCodEsocio(cmbE_Socio.List(cmbE_Socio.ListIndex))

   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_ANUAL_APORTE WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   w = ""
   If WE_S <> "" Then
      w = " WHERE M.E_SOCIO = '" + WE_S + "' "
   End If

   If Not (chkDieco.Value = True And _
           chkCajaMP.Value = True And _
           chkCiviles.Value = True) Then
   
       If chkDieco.Value = vbChecked Then
          If w = "" Then
             w = "WHERE M.TIPCOB = '01' "
          Else
             w = w + " AND M.TIPCOB = '01' "
          End If
       End If
   
       If chkCajaMP.Value = vbChecked Then
          If w = "" Then
             w = "WHERE M.TIPCOB = '02' "
          Else
             w = w + " OR M.TIPCOB = '02' "
          End If
       End If
   
       If chkCiviles.Value = vbChecked Then
          If w = "" Then
             w = "WHERE M.TIPCOB = '03' "
          Else
             w = w + " OR M.TIPCOB = '03' "
          End If
       End If
   
   End If
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_ANUAL_APORTE " _
   & " (CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, FECING, MONEDA, APORTE, TOTANO, USU) " _
   & " SELECT " _
   & "  M.CODSOCIO, M.CODIGO, M.INS, M.NOMBRE, M.E_SOCIO, M.FECING, E.MONEDA, E.APORTE, ROUND(E.APORTE * 12,2),  '" + wcodusu + "' " _
   & " FROM MAESOCIO AS M INNER JOIN MAEE_SOCIO AS E " _
   & "   ON M.E_SOCIO = E.E_SOCIO " _
   & " " + w + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ANUAL_APORTE " _
   & " SET TOTANO = ROUND(APORTE * 11,2) " _
   & " WHERE FECING >= '" + Format("16/01/" + wAno, "dd/mm/yyyy") + "' AND FECING <= '" + Format("15/02/" + wAno, "dd/mm/yyyy") + "'  ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ANUAL_APORTE " _
   & " SET TOTANO = ROUND(APORTE * 10,2) " _
   & " WHERE FECING >= '" + Format("16/02/" + wAno, "dd/mm/yyyy") + "' AND FECING <= '" + Format("15/03/" + wAno, "dd/mm/yyyy") + "'  ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ANUAL_APORTE " _
   & " SET TOTANO = ROUND(APORTE * 9,2) " _
   & " WHERE FECING >= '" + Format("16/03/" + wAno, "dd/mm/yyyy") + "' AND FECING <= '" + Format("15/04/" + wAno, "dd/mm/yyyy") + "'  ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ANUAL_APORTE " _
   & " SET TOTANO = ROUND(APORTE * 8,2) " _
   & " WHERE FECING >= '" + Format("16/04/" + wAno, "dd/mm/yyyy") + "' AND FECING <= '" + Format("15/05/" + wAno, "dd/mm/yyyy") + "'  ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ANUAL_APORTE " _
   & " SET TOTANO = ROUND(APORTE * 7,2) " _
   & " WHERE FECING >= '" + Format("16/05/" + wAno, "dd/mm/yyyy") + "' AND FECING <= '" + Format("15/06/" + wAno, "dd/mm/yyyy") + "'  ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ANUAL_APORTE " _
   & " SET TOTANO = ROUND(APORTE * 6,2) " _
   & " WHERE FECING >= '" + Format("16/06/" + wAno, "dd/mm/yyyy") + "' AND FECING <= '" + Format("15/07/" + wAno, "dd/mm/yyyy") + "'  ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ANUAL_APORTE " _
   & " SET TOTANO = ROUND(APORTE * 5,2) " _
   & " WHERE FECING >= '" + Format("16/07/" + wAno, "dd/mm/yyyy") + "' AND FECING <= '" + Format("15/08/" + wAno, "dd/mm/yyyy") + "'  ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ANUAL_APORTE " _
   & " SET TOTANO = ROUND(APORTE * 4,2) " _
   & " WHERE FECING >= '" + Format("16/08/" + wAno, "dd/mm/yyyy") + "' AND FECING <= '" + Format("15/09/" + wAno, "dd/mm/yyyy") + "'  ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ANUAL_APORTE " _
   & " SET TOTANO = ROUND(APORTE * 3,2) " _
   & " WHERE FECING >= '" + Format("16/09/" + wAno, "dd/mm/yyyy") + "' AND FECING <= '" + Format("15/10/" + wAno, "dd/mm/yyyy") + "'  ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ANUAL_APORTE " _
   & " SET TOTANO = ROUND(APORTE * 2,2) " _
   & " WHERE FECING >= '" + Format("16/10/" + wAno, "dd/mm/yyyy") + "' AND FECING <= '" + Format("15/11/" + wAno, "dd/mm/yyyy") + "'  ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ANUAL_APORTE " _
   & " SET TOTANO = ROUND(APORTE * 1,2) " _
   & " WHERE FECING >= '" + Format("16/11/" + wAno, "dd/mm/yyyy") + "' AND FECING <= '" + Format("15/12/" + wAno, "dd/mm/yyyy") + "'  ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ANUAL_APORTE " _
   & " SET TOTANO = ROUND(APORTE * 0,2) " _
   & " WHERE FECING >= '" + Format("16/12/" + wAno, "dd/mm/yyyy") + "' ")
   Db.CommitTrans

   For II = 1 To 12
       wmmm = Format(II, "00")
   
       Db.BeginTrans
       Db.Execute ("UPDATE TMP_ANUAL_APORTE " _
       & " SET PAGO" + wmmm + " = PAGO" + wmmm + " + V.IMPORTE " _
       & " FROM TMP_ANUAL_APORTE AS T INNER JOIN V_COBRO_TESOR AS V " _
       & "   ON T.CODIGO = V.CODIGO AND T.INS = V.INS " _
       & " WHERE V.ANO = '" + wAno + "' AND V.MES = '" + wmmm + "' ")
       Db.CommitTrans
   
   Next

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ANUAL_APORTE " _
   & " SET TOTPAG = PAGO01 + PAGO02 + PAGO03 + PAGO04 + PAGO05 + PAGO06 + PAGO07 + PAGO08 + PAGO09 + PAGO10 + PAGO11 + PAGO12 " _
   & " WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ANUAL_APORTE " _
   & " SET TOTDEU = TOTANO - TOTPAG " _
   & " WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_ANUAL_APORTE " _
   & " WHERE USU = '" + wcodusu + "' AND " _
   & "       APORTE = 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_ANUAL_APORTE " _
   & " WHERE USU = '" + wcodusu + "' AND " _
   & "       TOTANO = 0 AND " _
   & "       TOTPAG = 0 ")
   Db.CommitTrans


   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, FECING, MONEDA, " _
                & "      APORTE, TOTANO, TOTPAG, TOTDEU, PAGO01, PAGO02, PAGO03, " _
                & "      PAGO04, PAGO05, PAGO06, PAGO07, PAGO08, PAGO09, PAGO10, " _
                & "      PAGO11, PAGO12, USU " _
                & " FROM TMP_ANUAL_APORTE " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY E_SOCIO, NOMBRE ")
   Set DataGrid1.DataSource = ADO2

   DataGrid1.Columns(0).Width = 700   ' CODSOCIO
   DataGrid1.Columns(0).Alignment = dbgLeft
   DataGrid1.Columns(0).Caption = "COD.SOCIO"
    
   DataGrid1.Columns(1).Width = 1000  ' CODIGO
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Caption = "CODIGO"
    
   DataGrid1.Columns(2).Width = 400   ' INS
   DataGrid1.Columns(2).Alignment = dbgLeft
   DataGrid1.Columns(2).Caption = "INS"
    
   DataGrid1.Columns(3).Width = 5250  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE"
    
   DataGrid1.Columns(4).Width = 380   ' NOME_S
   DataGrid1.Columns(4).Alignment = dbgCenter
   DataGrid1.Columns(4).Caption = "ESTADO"
    
   DataGrid1.Columns(5).Width = 1050  ' FECING
   DataGrid1.Columns(5).Alignment = dbgCenter
   DataGrid1.Columns(5).Caption = "FEC.ING"
   DataGrid1.Columns(5).NumberFormat = "dd/mm/yyyy"
   
   DataGrid1.Columns(6).Width = 350   ' MONEDA
   DataGrid1.Columns(6).Alignment = dbgCenter
   DataGrid1.Columns(6).Caption = "MON"
    
   DataGrid1.Columns(7).Width = 850   ' APORTE
   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).Caption = "APORTE"
   DataGrid1.Columns(7).NumberFormat = "###0.00"
   
   DataGrid1.Columns(8).Width = 950   ' ANUAL
   DataGrid1.Columns(8).Alignment = dbgRight
   DataGrid1.Columns(8).Caption = "APORTE"
   DataGrid1.Columns(8).NumberFormat = "####0.00"
   
   DataGrid1.Columns(9).Width = 950   ' TOTPAG
   DataGrid1.Columns(9).Alignment = dbgRight
   DataGrid1.Columns(9).Caption = "TOT.COB"
   DataGrid1.Columns(9).NumberFormat = "####0.00"
   
   DataGrid1.Columns(10).Width = 950   ' SALDOS
   DataGrid1.Columns(10).Alignment = dbgRight
   DataGrid1.Columns(10).Caption = "DEUDAS"
   DataGrid1.Columns(10).NumberFormat = "####0.00"
   
   DataGrid1.Columns(11).Visible = False
   DataGrid1.Columns(12).Visible = False
   DataGrid1.Columns(13).Visible = False
   DataGrid1.Columns(14).Visible = False
   DataGrid1.Columns(15).Visible = False
   DataGrid1.Columns(16).Visible = False
   DataGrid1.Columns(17).Visible = False
   DataGrid1.Columns(18).Visible = False
   DataGrid1.Columns(19).Visible = False
   DataGrid1.Columns(20).Visible = False
   DataGrid1.Columns(21).Visible = False
   DataGrid1.Columns(22).Visible = False
   DataGrid1.Columns(23).Visible = False
   
End Sub

Private Sub txtAno_GotFocus()
   txtAno.SelStart = 0
   txtAno.SelLength = 4
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txtAno.Text)) = 0 Then
         MsgBox "Año de Proceso En Cero", vbExclamation
         txtAno.Text = wanocia
         Exit Sub
      End If
      If txtAno.Text < "2014" Or txtAno.Text > "2040" Then
         MsgBox "Año de Proceso Fuera de Rango", vbExclamation
         txtAno.Text = wanocia
         Exit Sub
      End If
      cmbE_Socio.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii), 1) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub


