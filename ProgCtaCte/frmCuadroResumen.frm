VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCuadroResumen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuadro Resumen de Aportaciones"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   11700
   Begin VB.CommandButton cmdImpxSol 
      Caption         =   "&Imprimir con S/"
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
      TabIndex        =   14
      Top             =   7680
      Width           =   1095
   End
   Begin VB.ComboBox cmbE_Socio 
      Height          =   315
      ItemData        =   "frmCuadroResumen.frx":0000
      Left            =   1320
      List            =   "frmCuadroResumen.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   800
      Width           =   3375
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
      Left            =   8760
      TabIndex        =   6
      Top             =   7080
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
      Left            =   7440
      TabIndex        =   5
      Top             =   7080
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
      Left            =   10080
      TabIndex        =   4
      Top             =   7080
      Width           =   1095
   End
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmCuadroResumen.frx":0004
      Left            =   1320
      List            =   "frmCuadroResumen.frx":0006
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
      Left            =   10080
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   1095
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5295
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9340
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
      Caption         =   "CUADRO RESUMEN DE APORTACIONES"
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
   Begin MSMask.MaskEdBox txtMes 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtMoroso 
      Height          =   285
      Left            =   3600
      TabIndex        =   9
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
   Begin VB.Label lblCajMP 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5880
      TabIndex        =   28
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "CAJA MP"
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
      Left            =   4680
      TabIndex        =   27
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Left            =   4680
      TabIndex        =   26
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblDieco 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5880
      TabIndex        =   25
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label lblImpDe9 
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9960
      TabIndex        =   24
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label lblDeuda9 
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9240
      TabIndex        =   23
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblImpDe7 
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8040
      TabIndex        =   22
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label lblImpDe6 
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   21
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label lblImpDe3 
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label lblImpDe0 
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   19
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label lblDeuda7 
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7320
      TabIndex        =   18
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblDeuda6 
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5520
      TabIndex        =   17
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblDeuda3 
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3720
      TabIndex        =   16
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblDeuda0 
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   6480
      Width           =   735
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
      TabIndex        =   13
      Top             =   800
      Width           =   1140
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
      TabIndex        =   11
      Top             =   480
      Width           =   975
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
      TabIndex        =   10
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
      TabIndex        =   7
      Top             =   7200
      Width           =   6615
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
Attribute VB_Name = "frmCuadroResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbE_Socio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdBuscar.SetFocus
   End If
End Sub

Private Sub cmdBuscar_Click()
   LlenaCab
   TotalCab
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(11) As String, Headin2(11) As String, wreg As Integer, wTot As Integer
   Dim wDeuda0 As Integer, wDeuda3 As Integer, wDeuda6 As Integer, wDeuda7 As Integer, wTotDeu As Integer, _
       wImpDe0 As Integer, wImpDe3 As Integer, wImpDe6 As Integer, wImpDe7 As Integer, wImpTot As Integer, _
       wMes As String, wAno As String
       
   wAno = Left(txtMoroso.Text, 4)
   wMes = Right(txtMoroso.Text, 2)
       
   Heading(0) = "TIPO"
   Heading(1) = "NOMBRE"
   Heading(2) = "SIN DEUDA"
   Heading(4) = "DEUDA 3 MESES"
   Heading(6) = "DEUDA 6 MESES"
   Heading(8) = "MAYOR 6 MESES"
   Heading(10) = "TOTAL"
   
   Heading(0) = "SOCIO"
   Heading(1) = "NOMBRE"
   Heading(2) = "CANT"
   Heading(3) = "S/"
   Heading(4) = "CANT"
   Heading(5) = "S/"
   Heading(6) = "CANT"
   Heading(7) = "S/"
   Heading(8) = "CANT"
   Heading(9) = "S/"
   Heading(10) = "CANT"
   Heading(11) = "S/"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 12)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 12)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "CUADRO RESUMEN DE APORTACIONES POR TIPO SOCIO - MES " + Trim(funnommes(wMes)) + " " + wAno
        For I = 1 To 12 Step 1
            .Cells(3, I) = Heading(I - 1)
            .Cells(4, I) = Headin2(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 10
        objExcel.Columns("B").ColumnWidth = 20
        objExcel.Columns("C").ColumnWidth = 8
        objExcel.Columns("D").ColumnWidth = 12
        objExcel.Columns("E").ColumnWidth = 8
        objExcel.Columns("F").ColumnWidth = 12
        objExcel.Columns("G").ColumnWidth = 8
        objExcel.Columns("H").ColumnWidth = 12
        objExcel.Columns("I").ColumnWidth = 8
        objExcel.Columns("J").ColumnWidth = 12
        objExcel.Columns("K").ColumnWidth = 8
        objExcel.Columns("L").ColumnWidth = 12
   End With
   
   aa = Leerado3("SELECT * FROM TMP_RESDEU WHERE USU = '" + wcodusu + "' ORDER BY ORDEN ")
   If aa > 0 Then
      wTot = aa
      V = 4
      H = 1
      wNum1 = 1
      wDeuda0 = 0: wDeuda3 = 0: wDeuda6 = 0: wDeuda7 = 0: wTotDeu = 0
      wImpDe0 = 0: wImpDe3 = 0: wImpDe6 = 0: wImpDe7 = 0: wImpTot = 0
      Do While Not ADO3.EOF
         
         objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 2)).NumberFormat = "##,##0;;\ "
         objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 4)).NumberFormat = "##,##0;;\ "
         objExcel.Range(objExcel.Cells(V, H + 6), objExcel.Cells(V, H + 6)).NumberFormat = "##,##0;;\ "
         objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 8)).NumberFormat = "##,##0;;\ "
         objExcel.Range(objExcel.Cells(V, H + 10), objExcel.Cells(V, H + 10)).NumberFormat = "##,##0;;\ "
         
         objExcel.Range(objExcel.Cells(V, H + 3), objExcel.Cells(V, H + 3)).NumberFormat = "#####,##0.00;;\ "
         objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 5)).NumberFormat = "#####,##0.00;;\ "
         objExcel.Range(objExcel.Cells(V, H + 7), objExcel.Cells(V, H + 7)).NumberFormat = "#####,##0.00;;\ "
         objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 9)).NumberFormat = "#####,##0.00;;\ "
         objExcel.Range(objExcel.Cells(V, H + 11), objExcel.Cells(V, H + 11)).NumberFormat = "#####,##0.00;;\ "
         
         
         objExcel.Cells(V, H + 0) = ADO3!e_socio
         objExcel.Cells(V, H + 1) = ADO3!nombre
         objExcel.Cells(V, H + 2) = ADO3!deuda0
         objExcel.Cells(V, H + 3) = ADO3!impde0
         objExcel.Cells(V, H + 4) = ADO3!deuda3
         objExcel.Cells(V, H + 5) = ADO3!impde3
         objExcel.Cells(V, H + 6) = ADO3!deuda6
         objExcel.Cells(V, H + 7) = ADO3!impde6
         objExcel.Cells(V, H + 8) = ADO3!deuda7
         objExcel.Cells(V, H + 9) = ADO3!impde7
         objExcel.Cells(V, H + 10) = ADO3!totdeu
         objExcel.Cells(V, H + 11) = ADO3!imptot
            
         wDeuda0 = wDeuda0 + ADO3!deuda0
         wDeuda3 = wDeuda3 + ADO3!deuda3
         wDeuda6 = wDeuda6 + ADO3!deuda6
         wDeuda7 = wDeuda7 + ADO3!deuda7
         wTotDeu = wTotDeu + ADO3!totdeu
         
         wImpDe0 = wImpDe0 + ADO3!impde0
         wImpDe3 = wImpDe3 + ADO3!impde3
         wImpDe6 = wImpDe6 + ADO3!impde6
         wImpDe7 = wImpDe7 + ADO3!impde7
         wimpdeu = wimpdeu + ADO3!imptot
         
         V = V + 1
         ADO3.MoveNext
      Loop
      V = V + 1
      
      objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 2)).NumberFormat = "##,##0;;\ "
      objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 4)).NumberFormat = "##,##0;;\ "
      objExcel.Range(objExcel.Cells(V, H + 6), objExcel.Cells(V, H + 6)).NumberFormat = "##,##0;;\ "
      objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 8)).NumberFormat = "##,##0;;\ "
      objExcel.Range(objExcel.Cells(V, H + 10), objExcel.Cells(V, H + 10)).NumberFormat = "##,##0;;\ "
      
      objExcel.Range(objExcel.Cells(V, H + 3), objExcel.Cells(V, H + 3)).NumberFormat = "#####,##0.00;;\ "
      objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 5)).NumberFormat = "#####,##0.00;;\ "
      objExcel.Range(objExcel.Cells(V, H + 7), objExcel.Cells(V, H + 7)).NumberFormat = "#####,##0.00;;\ "
      objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 9)).NumberFormat = "#####,##0.00;;\ "
      objExcel.Range(objExcel.Cells(V, H + 11), objExcel.Cells(V, H + 11)).NumberFormat = "#####,##0.00;;\ "
      
      objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 6)).Font.Bold = True
      objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 6)).Font.Color = RGB(255, 0, 0)
      objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 6)).Borders.LineStyle = xlContinuous
      objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 6)).Borders.Color = RGB(255, 0, 0)
      
      objExcel.Cells(V, H + 1) = "TOTALES FINALES"
      objExcel.Cells(V, H + 2) = wDeuda0
      objExcel.Cells(V, H + 3) = wImpDe0
      objExcel.Cells(V, H + 4) = wDeuda3
      objExcel.Cells(V, H + 5) = wImpDe3
      objExcel.Cells(V, H + 6) = wDeuda6
      objExcel.Cells(V, H + 7) = wImpDe6
      objExcel.Cells(V, H + 8) = wDeuda7
      objExcel.Cells(V, H + 9) = wImpDe7
      objExcel.Cells(V, H + 10) = wTotDeu
      objExcel.Cells(V, H + 11) = wImpTot
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
   Crys1.SelectionFormula = " {TMP_RESDEU.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdImpxSol_Click()
   Dim wFec As Date, wMes As String, wAno As String
   wAno = Left(txtMoroso.Text, 4)
   wMes = Trim(funnommes(Right(txtMoroso.Text, 2))) + " DEL " + wAno
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\CuadroResumenImportes.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= 'AL MES DE " + wMes + "' "
   Crys1.SelectionFormula = " {TMP_RESDEU.USU}='" + wcodusu + "' "
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
   
   Set DataGrid1.DataSource = Nothing
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESDEU WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   cmbE_Socio.SetFocus
End Sub

Private Sub Form_Load()
   frmCuadroResumen.Left = (Screen.Width - Width) \ 2
   frmCuadroResumen.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESDEU WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim aa As Long, wreg As Long, wTot As Long, _
       wMes As String, wAno As String, _
       wSoc As Integer, wSdo As Currency, wNom As String, wApo As Currency, wFac As Integer, _
       wDeu0 As Integer, wDeu3 As Integer, wDeu6 As Integer, wDeu7 As Integer, _
       wImp0 As Currency, wImp3 As Currency, wImp6 As Currency, wImp7 As Currency, _
       wDieco As Integer, wCajMP As Integer

   Dim w As String, wFec As Date, WE_S As String

   wMes = Left(txtMoroso.Text, 4) + Right(txtMoroso.Text, 2)
   wFec = Format(fundiames(Right(wMes, 2)) + "/" + Right(wMes, 2) + "/" + Left(wMes, 4), "dd/mm/yyyy")
   WE_S = BuscaCodEsocio(cmbE_Socio.List(cmbE_Socio.ListIndex))

   wDieco = 1
   aa = Leerado8("SELECT SUM(TOTENVIO) AS TOTENVIO " _
                & " FROM DIECOCAB " _
                & " WHERE MES = '" + wMes + "' ")
   If aa > 0 Then
      wDieco = 2
   End If
   Set ADO8 = Nothing

   aa = Leerado8("SELECT SUM(DSCDIECO) AS DSCDIECO " _
                & " FROM DIECOCAB " _
                & " WHERE MES = '" + wMes + "' ")
   If aa > 0 Then
      If ADO8!dscdieco > 0 Then
         wDieco = 3
      End If
   End If
   lblDieco.FontBold = True
   Select Case wDieco
   Case 1
        lblDieco.ForeColor = RGB(0, 255, 0)
        lblDieco.Caption = "DIECO - Mes " + Left(wMes, 4) + "-" + Right(wMes, 2) + " Sin Envio"
   Case 2
        lblDieco.ForeColor = RGB(0, 0, 255)
        lblDieco.Caption = "DIECO - Mes " + Left(wMes, 4) + "-" + Right(wMes, 2) + " Enviado Descuento"
   Case 3
        lblDieco.ForeColor = RGB(255, 0, 0)
        lblDieco.Caption = "DIECO - Mes " + Left(wMes, 4) + "-" + Right(wMes, 2) + " Descto Efectuado"
   End Select


   wCajMP = 1
   aa = Leerado8("SELECT SUM(TOTENVIO) AS TOTENVIO " _
                & " FROM CAJMPCAB " _
                & " WHERE MES = '" + wMes + "' ")
   If aa > 0 Then
      wCajMP = 2
   End If
   Set ADO8 = Nothing

   aa = Leerado8("SELECT SUM(DSCCAJMP) AS DSCCAJMP " _
                & " FROM CAJMPCAB " _
                & " WHERE MES = '" + wMes + "' ")
   If aa > 0 Then
      If ADO8!dsccajmp > 10000 Then
         wCajMP = 3
      End If
   End If
   lblCajMP.FontBold = True
   Select Case wCajMP
   Case 1
        lblCajMP.ForeColor = RGB(0, 255, 0)
        lblCajMP.Caption = "CAJA MP - Mes " + Left(wMes, 4) + "-" + Right(wMes, 2) + " Sin Envio"
   Case 2
        lblCajMP.ForeColor = RGB(0, 0, 255)
        lblCajMP.Caption = "CAJA MP - Mes " + Left(wMes, 4) + "-" + Right(wMes, 2) + " Enviado Descuento"
   Case 3
        lblCajMP.ForeColor = RGB(255, 0, 0)
        lblCajMP.Caption = "CAJA MP - Mes " + Left(wMes, 4) + "-" + Right(wMes, 2) + " Descto Efectuado"
   End Select

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESDEU WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESDEU " _
   & " (E_SOCIO, NOMBRE, ORDEN, USU) " _
   & " SELECT " _
   & "  E_SOCIO, NOMBRE, ORDEN, '" + wcodusu + "' " _
   & " FROM MAEE_SOCIO " _
   & " WHERE ORDEN <> 0 " + w + " ")
   Db.CommitTrans

   w = ""
   If WE_S <> "" Then
      w = " AND S.E_SOCIO = '" + WE_S + "' "
   End If

   aa = Leerado8a("SELECT S.CODSOCIO, S.CODIGO, S.INS, S.NOMBRE, S.E_SOCIO, E.APORTE, S.TIPCOB " _
                & " FROM MAESOCIO AS S INNER JOIN MAEE_SOCIO AS E " _
                & "   ON S.E_SOCIO = E.E_SOCIO " _
                & " WHERE E.APORTE > 0 " + w + " " _
                & " ORDER BY NOMBRE ")
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
         WE_S = ADO8a!e_socio
         wSdo = SaldoFoto(wSoc, wMes)
         If wSdo < 0 Then
            wSdo = 0
         End If
         wApo = ADO8a!aporte
         
         If ADO8a!tipcob = "01" Then
            If wDieco = "1" Or wDieco = "2" Then
               If wSdo >= wApo Then
                  wSdo = wSdo - wApo
               End If
            End If
         End If
         
         If ADO8a!tipcob = "02" Then
            If wCajMP = "1" Or wCajMP = "2" Then
               If wSdo >= wApo Then
                  wSdo = wSdo - wApo
               End If
            End If
         End If
         
         wDeu0 = 0: wDeu3 = 0: wDeu6 = 0: wDeu7 = 0
         wImp0 = 0: wImp3 = 0: wImp6 = 0: wImp7 = 0
         wFac = Round(wSdo / wApo, 0)
          
         If wSdo <= 0 Then
            wDeu0 = 1
            wImp0 = wSdo
         Else
            If wSdo > 0 And wSdo <= Round(wApo * 3, 2) Then
               wDeu3 = 1
               wImp3 = wSdo
            Else
               If wSdo > Round(wApo * 3, 2) And wSdo <= Round(wApo * 6, 2) Then
                  wDeu6 = 1
                  wImp6 = wSdo
               Else
                  wDeu7 = 1
                  wImp7 = wSdo
               End If
            End If
         End If
         
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_RESDEU " _
         & " SET DEUDA0 = DEUDA0 + " + Str(wDeu0) + ", " _
         & "     DEUDA3 = DEUDA3 + " + Str(wDeu3) + ", " _
         & "     DEUDA6 = DEUDA6 + " + Str(wDeu6) + ", " _
         & "     DEUDA7 = DEUDA7 + " + Str(wDeu7) + ", " _
         & "     IMPDE0 = IMPDE0 + " + Str(wImp0) + ", " _
         & "     IMPDE3 = IMPDE3 + " + Str(wImp3) + ", " _
         & "     IMPDE6 = IMPDE6 + " + Str(wImp6) + ", " _
         & "     IMPDE7 = IMPDE7 + " + Str(wImp7) + " " _
         & " WHERE E_SOCIO = '" + WE_S + "' AND " _
         & "           USU = '" + wcodusu + "' ")
         Db.CommitTrans
   
         wreg = wreg + 1
         ADO8a.MoveNext
      Loop
   End If
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_RESDEU " _
   & " SET TOTDEU = DEUDA0 + DEUDA3 + DEUDA6 + DEUDA7, " _
   & "     IMPTOT = IMPDE0 + IMPDE3 + IMPDE6 + IMPDE7 " _
   & " WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   DoEvents
   lblMensaje.Caption = ""
   lblMensaje.Refresh
   
   aa = Leerado2("SELECT E_SOCIO, NOMBRE, DEUDA0, IMPDE0, DEUDA3, IMPDE3, DEUDA6, IMPDE6, DEUDA7, IMPDE7, TOTDEU, IMPTOT " _
            & " FROM TMP_RESDEU " _
            & " WHERE USU = '" + wcodusu + "' " _
            & " ORDER BY ORDEN ")
   Set DataGrid1.DataSource = ADO2
 
   DataGrid1.Columns(0).Width = 600   ' CODSOCIO
   DataGrid1.Columns(0).Alignment = dbgLeft
   DataGrid1.Columns(0).Caption = "E.SOCIO"
    
   DataGrid1.Columns(1).Width = 2000   ' NOMBRE
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Caption = "NOMBRE"
    
   DataGrid1.Columns(2).Width = 500   ' DEUDA0
   DataGrid1.Columns(2).Alignment = dbgRight
   DataGrid1.Columns(2).Caption = "SIN DEUDA"
   DataGrid1.Columns(2).NumberFormat = "#####0;;\ "

   DataGrid1.Columns(3).Width = 1100  ' IMPDA0
   DataGrid1.Columns(3).Alignment = dbgRight
   DataGrid1.Columns(3).Caption = "S/ DEUDA0"
   DataGrid1.Columns(3).NumberFormat = "####,##0.00;;\ "

   DataGrid1.Columns(4).Width = 500   ' DEUDA3
   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).Caption = "3 MESES"
   DataGrid1.Columns(4).NumberFormat = "#####0;;\ "

   DataGrid1.Columns(5).Width = 1100  ' IMPDA3
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).Caption = "S/ DEUDA3"
   DataGrid1.Columns(5).NumberFormat = "####,##0.00;;\ "

   DataGrid1.Columns(6).Width = 500   ' DEUDA6
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).Caption = "6 MESES"
   DataGrid1.Columns(6).NumberFormat = "#####0;;\ "

   DataGrid1.Columns(7).Width = 1100  ' IMPDA6
   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).Caption = "S/ DEUDA6"
   DataGrid1.Columns(7).NumberFormat = "####,##0.00;;\ "

   DataGrid1.Columns(8).Width = 500   ' DEUDA7
   DataGrid1.Columns(8).Alignment = dbgRight
   DataGrid1.Columns(8).Caption = "> 6 MESES"
   DataGrid1.Columns(8).NumberFormat = "#####0;;\ "

   DataGrid1.Columns(9).Width = 1100  ' IMPDA7
   DataGrid1.Columns(9).Alignment = dbgRight
   DataGrid1.Columns(9).Caption = "S/ DEUDA7"
   DataGrid1.Columns(9).NumberFormat = "####,##0.00;;\ "

   DataGrid1.Columns(10).Width = 500   ' TOTDEU
   DataGrid1.Columns(10).Alignment = dbgRight
   DataGrid1.Columns(10).Caption = "TOTALES"
   DataGrid1.Columns(10).NumberFormat = "#####0;;\ "

   DataGrid1.Columns(11).Width = 1100  ' IMPTOT
   DataGrid1.Columns(11).Alignment = dbgRight
   DataGrid1.Columns(11).Caption = "S/ TOTAL"
   DataGrid1.Columns(11).NumberFormat = "####,##0.00;;\ "

End Sub

Private Sub TotalCab()
   Dim wDeu0 As Integer, wDeu3 As Integer, wDeu6 As Integer, wDeu7 As Integer
   Dim wImp0 As Currency, wImp3 As Currency, wImp6 As Currency, wImp7 As Currency
   Dim wDeu9 As Integer, wImp9 As Currency
   Dim aa As Integer

   aa = Leerado8("SELECT SUM(DEUDA0) AS DEUDA0, SUM(IMPDE0) AS IMPDE0, " _
                & "      SUM(DEUDA3) AS DEUDA3, SUM(IMPDE3) AS IMPDE3, " _
                & "      SUM(DEUDA6) AS DEUDA6, SUM(IMPDE6) AS IMPDE6, " _
                & "      SUM(DEUDA7) AS DEUDA7, SUM(IMPDE7) AS IMPDE7, " _
                & "      SUM(totdeu) AS DEUTOT, SUM(IMPTOT) AS IMPTOT " _
                & " FROM TMP_RESDEU " _
                & " WHERE USU = '" + wcodusu + "' ")
   If aa > 0 Then
      wDeu0 = IIf(IsNull(ADO8!deuda0), 0, ADO8!deuda0)
      wImp0 = IIf(IsNull(ADO8!impde0), 0, ADO8!impde0)
      wDeu3 = IIf(IsNull(ADO8!deuda3), 0, ADO8!deuda3)
      wImp3 = IIf(IsNull(ADO8!impde3), 0, ADO8!impde3)
      wDeu6 = IIf(IsNull(ADO8!deuda6), 0, ADO8!deuda6)
      wImp6 = IIf(IsNull(ADO8!impde6), 0, ADO8!impde6)
      wDeu7 = IIf(IsNull(ADO8!deuda7), 0, ADO8!deuda7)
      wImp7 = IIf(IsNull(ADO8!impde7), 0, ADO8!impde7)
      wDeu9 = IIf(IsNull(ADO8!deutot), 0, ADO8!deutot)
      wImp9 = IIf(IsNull(ADO8!imptot), 0, ADO8!imptot)
   End If
   Set ADO8 = Nothing
  
   lblDeuda0.Caption = Format(wDeu0, "###,##0;;\ ")
   lblDeuda3.Caption = Format(wDeu3, "###,##0;;\ ")
   lblDeuda6.Caption = Format(wDeu6, "###,##0;;\ ")
   lblDeuda7.Caption = Format(wDeu7, "###,##0;;\ ")
   lblDeuda9.Caption = Format(wDeu9, "###,##0;;\ ")

   lblImpDe0.Caption = Format(wImp0, "###,##0.00;;\ ")
   lblImpDe3.Caption = Format(wImp3, "###,##0.00;;\ ")
   lblImpDe6.Caption = Format(wImp6, "###,##0.00;;\ ")
   lblImpDe7.Caption = Format(wImp7, "###,##0.00;;\ ")
   lblImpDe9.Caption = Format(wImp9, "###,##0.00;;\ ")
End Sub

Private Sub txtMes_GotFocus()
   txtMes.SelStart = 0
   txtMes.SelLength = 7
End Sub

Private Sub txtMes_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 40
        cmbE_Socio.SetFocus
   End Select
End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)
   Dim wMes As String, wAno As String
   If KeyAscii = 13 Then
      If txtMes.Text = "___/__" Then
         MsgBox "Mes En Blanco", vbExclamation
         txtMes.Text = Left(zMesTope, 4) + "/" + Right(zMesTope, 2)
         Exit Sub
      End If
      wAno = Left(txtMes.Text, 4)
      wMes = Right(txtMes.Text, 2)
      If wMes <> "01" And wMes <> "02" And wMes <> "03" And wMes <> "04" And _
         wMes <> "05" And wMes <> "06" And wMes <> "07" And wMes <> "08" And _
         wMes <> "09" And wMes <> "10" And wMes <> "11" And wMes <> "12" Then
         MsgBox "Mes Digitado Es Invalido", vbExclamation
         txtMes.Text = Left(zMesTope, 4) + "/" + Right(zMesTope, 2)
         Exit Sub
      End If
      If wAno < "2010" And wAno > "2030" Then
         MsgBox "Año Digitado Fuera De Rango", vbExclamation
         txtMes.Text = Left(zMesTope, 4) + "/" + Right(zMesTope, 2)
         Exit Sub
      End If
      txtMoroso.Text = txtMes.Text
      
      cmbE_Socio.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub
