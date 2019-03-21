VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConCobxMes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Cobros x Mes"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   16950
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
      Left            =   9480
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   480
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
      Left            =   13680
      TabIndex        =   14
      Top             =   7800
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
      Left            =   12360
      TabIndex        =   13
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
      Left            =   15000
      TabIndex        =   12
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox txtSocio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   780
      Width           =   855
   End
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmConCobxMes.frx":0000
      Left            =   1080
      List            =   "frmConCobxMes.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   6615
   End
   Begin VB.TextBox txtAno 
      Height          =   305
      Left            =   8520
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin MSMask.MaskEdBox txtDesde 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtHasta 
      Height          =   285
      Left            =   3360
      TabIndex        =   6
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5895
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   10398
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
      Caption         =   "RELACION DE COBROS POR SOCIO"
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
            LCID            =   3082
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
            LCID            =   3082
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
      Left            =   14880
      Top             =   0
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
   Begin VB.Label lblSocioDet 
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
      Left            =   120
      TabIndex        =   42
      Top             =   7680
      Width           =   8415
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Caption         =   "Total"
      Height          =   255
      Left            =   15480
      TabIndex        =   41
      Top             =   7320
      Width           =   975
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   15480
      TabIndex        =   40
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Caption         =   "Dic"
      Height          =   255
      Left            =   14520
      TabIndex        =   39
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label lblMes12 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   14520
      TabIndex        =   38
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Caption         =   "Nov"
      Height          =   255
      Left            =   13560
      TabIndex        =   37
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label lblMes11 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   13560
      TabIndex        =   36
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Caption         =   "Oct"
      Height          =   255
      Left            =   12600
      TabIndex        =   35
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label lblMes10 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   12600
      TabIndex        =   34
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Caption         =   "Set"
      Height          =   255
      Left            =   11640
      TabIndex        =   33
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label lblMes09 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   11640
      TabIndex        =   32
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "Ago"
      Height          =   255
      Left            =   10680
      TabIndex        =   31
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label lblMes08 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   10680
      TabIndex        =   30
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "Jul"
      Height          =   255
      Left            =   9720
      TabIndex        =   29
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label lblMes07 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   9720
      TabIndex        =   28
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Jun"
      Height          =   255
      Left            =   8760
      TabIndex        =   27
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label lblMes06 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   8760
      TabIndex        =   26
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "May"
      Height          =   255
      Left            =   7800
      TabIndex        =   25
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label lblMes05 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   7800
      TabIndex        =   24
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Abr"
      Height          =   255
      Left            =   6840
      TabIndex        =   23
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label lblMes04 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6840
      TabIndex        =   22
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Mar"
      Height          =   255
      Left            =   5880
      TabIndex        =   21
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label lblMes03 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5880
      TabIndex        =   20
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Feb"
      Height          =   255
      Left            =   4920
      TabIndex        =   19
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label lblMes02 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Ene"
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label lblMes01 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   7080
      Width           =   975
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
      Left            =   1320
      TabIndex        =   15
      Top             =   8160
      Width           =   9135
   End
   Begin VB.Label Label14 
      Caption         =   "Desde"
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
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Hasta"
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
      TabIndex        =   9
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Socio"
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
      TabIndex        =   8
      Top             =   780
      Width           =   855
   End
   Begin VB.Label lblSocio 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   780
      Width           =   6735
   End
   Begin VB.Label Label38 
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
      TabIndex        =   3
      Top             =   120
      Width           =   855
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
      Left            =   7920
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmConCobxMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
   LlenaCab
   LabelCab
   TotalCab
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(16) As String, _
       wRegAct As Integer, wRegTot As Integer, wAno As String, _
       wMes01 As Currency, wMes02 As Currency, wMes03 As Currency, wMes04 As Currency, _
       wMes05 As Currency, wMes06 As Currency, wMes07 As Currency, wMes08 As Currency, _
       wMes09 As Currency, wMes10 As Currency, wMes11 As Currency, wMes12 As Currency, _
       wTotal As Currency

   wAno = txtAno.Text
   
   Heading(0) = "SOCIO"
   Heading(1) = "CODIGO"
   Heading(2) = "INS"
   Heading(3) = "NOMBRE SOCIO"
   Heading(4) = "ENE"
   Heading(5) = "FEB"
   Heading(6) = "MAR"
   Heading(7) = "ABR"
   Heading(8) = "MAY"
   Heading(9) = "JUN"
   Heading(10) = "JUL"
   Heading(11) = "AGO"
   Heading(12) = "SET"
   Heading(13) = "OCT"
   Heading(14) = "NOV"
   Heading(15) = "DIC"
   Heading(16) = "TOTAL"
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 17)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 17)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "DETALLE DE COBRANZAS POR MES - EJERCICIO " + wAno
        For I = 1 To 15 Step 1
            .Cells(3, I) = Heading(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 7
        objExcel.Columns("B").ColumnWidth = 9
        objExcel.Columns("C").ColumnWidth = 3
        objExcel.Columns("D").ColumnWidth = 60
        objExcel.Columns("E").ColumnWidth = 11
        objExcel.Columns("F").ColumnWidth = 11
        objExcel.Columns("G").ColumnWidth = 11
        objExcel.Columns("H").ColumnWidth = 11
        objExcel.Columns("I").ColumnWidth = 11
        objExcel.Columns("J").ColumnWidth = 11
        objExcel.Columns("K").ColumnWidth = 11
        objExcel.Columns("L").ColumnWidth = 11
        objExcel.Columns("M").ColumnWidth = 11
        objExcel.Columns("N").ColumnWidth = 11
        objExcel.Columns("O").ColumnWidth = 11
        objExcel.Columns("P").ColumnWidth = 11
        objExcel.Columns("Q").ColumnWidth = 11
   End With
   
   aa = Leerado3("SELECT * " _
                & " FROM TMP_COBXMES " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY NOMBRE ")
   If aa > 0 Then
      wRegTot = aa
      V = 4
      H = 1
      wRegAct = 1
      Do While Not ADO3.EOF
         DoEvents
         lblMensaje.Caption = "Traslando Cobranzas a EXCEL - Registro " + _
                              Format(wRegAct, "####0") + " / " + _
                              Format(wRegTot, "####0")
         lblMensaje.Refresh
         
         objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 16)).NumberFormat = "######0.00;;\ "
            
         objExcel.Cells(V, H + 0) = ADO3!codsocio
         objExcel.Cells(V, H + 1) = ADO3!codigo
         objExcel.Cells(V, H + 2) = ADO3!ins
         objExcel.Cells(V, H + 3) = ADO3!nombre
         objExcel.Cells(V, H + 4) = ADO3!mes01
         objExcel.Cells(V, H + 5) = ADO3!mes02
         objExcel.Cells(V, H + 6) = ADO3!mes03
         objExcel.Cells(V, H + 7) = ADO3!mes04
         objExcel.Cells(V, H + 8) = ADO3!mes05
         objExcel.Cells(V, H + 9) = ADO3!mes06
         objExcel.Cells(V, H + 10) = ADO3!mes07
         objExcel.Cells(V, H + 11) = ADO3!mes08
         objExcel.Cells(V, H + 12) = ADO3!mes09
         objExcel.Cells(V, H + 13) = ADO3!mes10
         objExcel.Cells(V, H + 14) = ADO3!mes11
         objExcel.Cells(V, H + 15) = ADO3!mes12
         objExcel.Cells(V, H + 16) = ADO3!Total
         
         wMes01 = wMes01 + ADO3!mes01
         wMes02 = wMes02 + ADO3!mes02
         wMes03 = wMes03 + ADO3!mes03
         wMes04 = wMes04 + ADO3!mes04
         wMes05 = wMes05 + ADO3!mes05
         wMes06 = wMes06 + ADO3!mes06
         wMes07 = wMes07 + ADO3!mes07
         wMes08 = wMes08 + ADO3!mes08
         wMes09 = wMes09 + ADO3!mes09
         wMes10 = wMes10 + ADO3!mes10
         wMes11 = wMes11 + ADO3!mes11
         wMes12 = wMes12 + ADO3!mes12
         wTotal = wTotal + ADO3!Total
         
         wRegAct = wRegAct + 1
         V = V + 1
         ADO3.MoveNext
      Loop
      V = V + 1
         
      objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 16)).NumberFormat = "######0.00;;\  "
      objExcel.Range(objExcel.Cells(V, H + 3), objExcel.Cells(V, H + 16)).Font.Bold = True
      objExcel.Range(objExcel.Cells(V, H + 3), objExcel.Cells(V, H + 16)).Font.Color = RGB(255, 0, 0)
      objExcel.Range(objExcel.Cells(V, H + 3), objExcel.Cells(V, H + 16)).Borders.LineStyle = xlContinuous
      objExcel.Range(objExcel.Cells(V, H + 3), objExcel.Cells(V, H + 16)).Borders.Color = RGB(255, 0, 0)
      
      objExcel.Cells(V, H + 3) = "TOTALES FINALES"
      objExcel.Cells(V, H + 4) = wMes01
      objExcel.Cells(V, H + 5) = wMes02
      objExcel.Cells(V, H + 6) = wMes03
      objExcel.Cells(V, H + 7) = wMes04
      objExcel.Cells(V, H + 8) = wMes05
      objExcel.Cells(V, H + 9) = wMes06
      objExcel.Cells(V, H + 10) = wMes07
      objExcel.Cells(V, H + 11) = wMes08
      objExcel.Cells(V, H + 12) = wMes09
      objExcel.Cells(V, H + 13) = wMes10
      objExcel.Cells(V, H + 14) = wMes11
      objExcel.Cells(V, H + 15) = wMes12
      objExcel.Cells(V, H + 16) = wTotal
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
   Crys1.ReportFileName = xraiz + "ReportCtaCte\CobxMes.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= 'EJERCICIO " + wAno + "'"
   Crys1.SelectionFormula = " {TMP_COBXMES.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   LabelCab
End Sub

Private Sub Form_Activate()
   txtAno.Text = wanocia
   
   Dim a As Integer
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
   
   txtAno.Text = wanocia
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_COBXMES WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   txtAno.SetFocus
End Sub

Private Sub Form_Load()
   frmConCobxMes.Left = (Screen.Width - Width) \ 2
   frmConCobxMes.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_COBXMES WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim aa As Long, II As Integer, wAno As String, wMes As String, wFec As Date, wSoc As Integer
   Dim wdes As String, whas As String
   wAno = txtAno.Text
   wdes = Format(txtDesde.Text, "dd/mm/yyyy")
   whas = Format(txtHasta.Text, "dd/mm/yyyy")
   wSoc = Val(txtSocio.Text)

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_COBXMES WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_COBXMES " _
   & " (CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, USU) " _
   & " SELECT " _
   & "  CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, '" + wcodusu + "' " _
   & " FROM MAESOCIO ")
   Db.CommitTrans
   
   For II = Month(wdes) To Month(whas)
       wMes = Format(II, "00")
   
       Db.BeginTrans
       Db.Execute ("UPDATE TMP_COBXMES " _
       & " SET MES" + wMes + " = MES" + wMes + " + D.DSCSOCIO " _
       & " FROM TMP_COBXMES AS T INNER JOIN DIECOCAB AS D " _
       & "   ON T.CODSOCIO = D.CODSOCIO " _
       & " WHERE T.USU = '" + wcodusu + "' AND " _
       & "       D.MES = '" + wAno + wMes + "' ")
       Db.CommitTrans
   
       Db.BeginTrans
       Db.Execute ("UPDATE TMP_COBXMES " _
       & " SET MES" + wMes + " = MES" + wMes + " + D.DSCASIG1 " _
       & " FROM TMP_COBXMES AS T INNER JOIN DIECOCAB AS D " _
       & "   ON T.CODSOCIO = D.CODASIG1 " _
       & " WHERE T.USU = '" + wcodusu + "' AND " _
       & "       D.MES = '" + wAno + wMes + "' AND " _
       & "       D.CODASIG1 > 0 ")
       Db.CommitTrans
   
       Db.BeginTrans
       Db.Execute ("UPDATE TMP_COBXMES " _
       & " SET MES" + wMes + " = MES" + wMes + " + D.DSCASIG2 " _
       & " FROM TMP_COBXMES AS T INNER JOIN DIECOCAB AS D " _
       & "   ON T.CODSOCIO = D.CODASIG2 " _
       & " WHERE T.USU = '" + wcodusu + "' AND " _
       & "       D.MES = '" + wAno + wMes + "' AND " _
       & "       D.CODASIG2 > 0 ")
       Db.CommitTrans
   
       Db.BeginTrans
       Db.Execute ("UPDATE TMP_COBXMES " _
       & " SET MES" + wMes + " = MES" + wMes + " + D.DSCASIG3 " _
       & " FROM TMP_COBXMES AS T INNER JOIN DIECOCAB AS D " _
       & "   ON T.CODSOCIO = D.CODASIG3 " _
       & " WHERE T.USU = '" + wcodusu + "' AND " _
       & "       D.MES = '" + wAno + wMes + "' AND " _
       & "       D.CODASIG3 > 0 ")
       Db.CommitTrans
   
       Db.BeginTrans
       Db.Execute ("UPDATE TMP_COBXMES " _
       & " SET MES" + wMes + " = MES" + wMes + " + D.DSCASIG4 " _
       & " FROM TMP_COBXMES AS T INNER JOIN DIECOCAB AS D " _
       & "   ON T.CODSOCIO = D.CODASIG4 " _
       & " WHERE T.USU = '" + wcodusu + "' AND " _
       & "       D.MES = '" + wAno + wMes + "' AND " _
       & "       D.CODASIG4 > 0 ")
       Db.CommitTrans
   
       Db.BeginTrans
       Db.Execute ("UPDATE TMP_COBXMES " _
       & " SET MES" + wMes + " = MES" + wMes + " + D.DSCASIG5 " _
       & " FROM TMP_COBXMES AS T INNER JOIN DIECOCAB AS D " _
       & "   ON T.CODSOCIO = D.CODASIG5 " _
       & " WHERE T.USU = '" + wcodusu + "' AND " _
       & "       D.MES = '" + wAno + wMes + "' AND " _
       & "       D.CODASIG5 > 0 ")
       Db.CommitTrans
   
       Db.BeginTrans
       Db.Execute ("UPDATE TMP_COBXMES " _
       & " SET MES" + wMes + " = MES" + wMes + " + D.DSCSOCIO " _
       & " FROM TMP_COBXMES AS T INNER JOIN CAJMPCAB AS D " _
       & "   ON T.CODSOCIO = D.CODSOCIO " _
       & " WHERE T.USU = '" + wcodusu + "' AND " _
       & "       D.MES = '" + wAno + wMes + "' ")
       Db.CommitTrans
   
       Db.BeginTrans
       Db.Execute ("UPDATE TMP_COBXMES " _
       & " SET MES" + wMes + " = MES" + wMes + " + D.DSCASIG1 " _
       & " FROM TMP_COBXMES AS T INNER JOIN CAJMPCAB AS D " _
       & "   ON T.CODSOCIO = D.CODASIG1 " _
       & " WHERE T.USU = '" + wcodusu + "' AND " _
       & "       D.MES = '" + wAno + wMes + "' AND " _
       & "       D.CODASIG1 > 0 ")
       Db.CommitTrans
   
       Db.BeginTrans
       Db.Execute ("UPDATE TMP_COBXMES " _
       & " SET MES" + wMes + " = MES" + wMes + " + D.DSCASIG2 " _
       & " FROM TMP_COBXMES AS T INNER JOIN CAJMPCAB AS D " _
       & "   ON T.CODSOCIO = D.CODASIG2 " _
       & " WHERE T.USU = '" + wcodusu + "' AND " _
       & "       D.MES = '" + wAno + wMes + "' AND " _
       & "       D.CODASIG2 > 0 ")
       Db.CommitTrans
   
       Db.BeginTrans
       Db.Execute ("UPDATE TMP_COBXMES " _
       & " SET MES" + wMes + " = MES" + wMes + " + D.DSCASIG3 " _
       & " FROM TMP_COBXMES AS T INNER JOIN CAJMPCAB AS D " _
       & "   ON T.CODSOCIO = D.CODASIG3 " _
       & " WHERE T.USU = '" + wcodusu + "' AND " _
       & "       D.MES = '" + wAno + wMes + "' AND " _
       & "       D.CODASIG3 > 0 ")
       Db.CommitTrans
   
       Db.BeginTrans
       Db.Execute ("UPDATE TMP_COBXMES " _
       & " SET MES" + wMes + " = MES" + wMes + " + D.DSCASIG4 " _
       & " FROM TMP_COBXMES AS T INNER JOIN CAJMPCAB AS D " _
       & "   ON T.CODSOCIO = D.CODASIG4 " _
       & " WHERE T.USU = '" + wcodusu + "' AND " _
       & "       D.MES = '" + wAno + wMes + "' AND " _
       & "       D.CODASIG4 > 0 ")
       Db.CommitTrans
   
       Db.BeginTrans
       Db.Execute ("UPDATE TMP_COBXMES " _
       & " SET MES" + wMes + " = MES" + wMes + " + D.DSCASIG5 " _
       & " FROM TMP_COBXMES AS T INNER JOIN CAJMPCAB AS D " _
       & "   ON T.CODSOCIO = D.CODASIG5 " _
       & " WHERE T.USU = '" + wcodusu + "' AND " _
       & "       D.MES = '" + wAno + wMes + "' AND " _
       & "       D.CODASIG5 > 0 ")
       Db.CommitTrans
   
' mrecibos
       Db.BeginTrans
       Db.Execute ("UPDATE TMP_COBXMES " _
       & " SET MES" + wMes + " = MES" + wMes + " + D.MONTOSOLES " _
       & " FROM TMP_COBXMES AS T INNER JOIN V_COBROS_MRECIBOS AS D " _
       & "   ON T.CODIGO = D.CODIGO AND T.INS = D.INS " _
       & " WHERE T.USU = '" + wcodusu + "' AND " _
       & "       D.MES = '" + wMes + "' AND " _
       & "       D.ANO = '" + wAno + "' AND D.MONTOSOLES > 0 ")
       Db.CommitTrans
   
'' cobrodet
'       Db.BeginTrans
'       Db.Execute ("UPDATE TMP_COBXMES " _
'       & " SET MES" + wMes + " = MES" + wMes + " + D.MONTOSOLES " _
'       & " FROM TMP_COBXMES AS T INNER JOIN V_COBROS_COBRODET AS D " _
'       & "   ON T.CODIGO = D.CODIGO AND T.INS = D.INS " _
'       & " WHERE T.USU = '" + wcodusu + "' AND " _
'       & "       D.MES = '" + wMes + "' AND " _
'       & "       D.ANO = '" + wAno + "' ")
'       Db.CommitTrans
   
   Next
      
      
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_COBXMES " _
   & " SET TOTAL = MES01 + MES02 + MES03 + MES04 + MES05 + MES06 + " _
   & "             MES07 + MES08 + MES09 + MES10 + MES11 + MES12 " _
   & " WHERE USU = '" + wcodusu + "' AND TOTAL = 0 ")
   Db.CommitTrans
      
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_COBXMES " _
   & " WHERE USU = '" + wcodusu + "' AND TOTAL = 0 ")
   Db.CommitTrans

   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, " _
            & "          MES01, MES02, MES03, MES04, MES05, MES06, " _
            & "          MES07, MES08, MES09, MES10, MES11, MES12, TOTAL " _
            & " FROM TMP_COBXMES " _
            & " WHERE USU = '" + wcodusu + "' " _
            & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO2
 
   lblTotal.Caption = Format(aa, "##,##0") + " "

   DataGrid1.Columns(0).Width = 700   ' CODSOCIO
   DataGrid1.Columns(0).Alignment = dbgRight
   DataGrid1.Columns(0).Caption = "SOCIO"
    
   DataGrid1.Columns(1).Width = 900   ' CODIGO
   DataGrid1.Columns(1).Alignment = dbgRight
   DataGrid1.Columns(1).Caption = "CODIGO"
    
   DataGrid1.Columns(2).Width = 360   ' INS
   DataGrid1.Columns(2).Alignment = dbgLeft
   DataGrid1.Columns(2).Caption = "INS"
    
   DataGrid1.Columns(3).Width = 3600  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE"
    
   DataGrid1.Columns(4).Width = 800    ' mes01
   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).Caption = "ENE"
   DataGrid1.Columns(4).NumberFormat = "#####0.00;;\ "
    
   DataGrid1.Columns(5).Width = 800   ' mes02
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).Caption = "FEB"
   DataGrid1.Columns(5).NumberFormat = "#####0.00;;\ "
    
   DataGrid1.Columns(6).Width = 800   ' mes03
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).Caption = "MAR"
   DataGrid1.Columns(6).NumberFormat = "#####0.00;;\ "
   
   DataGrid1.Columns(7).Width = 800   ' mes04
   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).Caption = "ABR"
   DataGrid1.Columns(7).NumberFormat = "#####0.00;;\ "

   DataGrid1.Columns(8).Width = 800   ' mes05
   DataGrid1.Columns(8).Alignment = dbgRight
   DataGrid1.Columns(8).Caption = "MAY"
   DataGrid1.Columns(8).NumberFormat = "#####0.00;;\ "

   DataGrid1.Columns(9).Width = 800   ' mes06
   DataGrid1.Columns(9).Alignment = dbgRight
   DataGrid1.Columns(9).Caption = "JUN"
   DataGrid1.Columns(9).NumberFormat = "#####0.00;;\ "

   DataGrid1.Columns(10).Width = 800   ' mes07
   DataGrid1.Columns(10).Alignment = dbgRight
   DataGrid1.Columns(10).Caption = "JUL"
   DataGrid1.Columns(10).NumberFormat = "#####0.00;;\ "

   DataGrid1.Columns(11).Width = 800   ' mes08
   DataGrid1.Columns(11).Alignment = dbgRight
   DataGrid1.Columns(11).Caption = "AGO"
   DataGrid1.Columns(11).NumberFormat = "#####0.00;;\ "

   DataGrid1.Columns(12).Width = 800   ' mes09
   DataGrid1.Columns(12).Alignment = dbgRight
   DataGrid1.Columns(12).Caption = "SET"
   DataGrid1.Columns(12).NumberFormat = "#####0.00;;\ "

   DataGrid1.Columns(13).Width = 800   ' mes10
   DataGrid1.Columns(13).Alignment = dbgRight
   DataGrid1.Columns(13).Caption = "OCT"
   DataGrid1.Columns(13).NumberFormat = "#####0.00;;\ "

   DataGrid1.Columns(14).Width = 800   ' mes11
   DataGrid1.Columns(14).Alignment = dbgRight
   DataGrid1.Columns(14).Caption = "NOV"
   DataGrid1.Columns(14).NumberFormat = "#####0.00;;\ "

   DataGrid1.Columns(15).Width = 800   ' mes12
   DataGrid1.Columns(15).Alignment = dbgRight
   DataGrid1.Columns(15).Caption = "DIC"
   DataGrid1.Columns(15).NumberFormat = "#####0.00;;\ "

   DataGrid1.Columns(16).Width = 800   ' TOTAL
   DataGrid1.Columns(16).Alignment = dbgRight
   DataGrid1.Columns(16).Caption = "TOTAL"
   DataGrid1.Columns(16).NumberFormat = "#####0.00;;\ "

End Sub

Private Sub LabelCab()
   lblSocioDet.Caption = Str(ADO2!codsocio) + " " + ADO2!nombre
End Sub

Private Sub TotalCab()
   Dim wMes01 As Currency, wMes02 As Currency, wMes03 As Currency, wMes04 As Currency, _
       wMes05 As Currency, wMes06 As Currency, wMes07 As Currency, wMes08 As Currency, _
       wMes09 As Currency, wMes10 As Currency, wMes11 As Currency, wMes12 As Currency, _
       wTotal As Currency, aa As Integer
  
   aa = Leerado8("SELECT SUM(MES01) AS MES01, SUM(MES02) AS MES02, SUM(MES03) AS MES03, " _
            & "          SUM(MES04) AS MES04, SUM(MES05) AS MES05, SUM(MES06) AS MES06, " _
            & "          SUM(MES07) AS MES07, SUM(MES08) AS MES08, SUM(MES09) AS MES09, " _
            & "          SUM(MES10) AS MES10, SUM(MES11) AS MES11, SUM(MES12) AS MES12, " _
            & "          SUM(TOTAL) AS TOTAL " _
            & " FROM TMP_COBXMES " _
            & " WHERE USU = '" + wcodusu + "' ")
   If aa > 0 Then
      wMes01 = ADO8!mes01
      wMes02 = ADO8!mes02
      wMes03 = ADO8!mes03
      wMes04 = ADO8!mes04
      wMes05 = ADO8!mes05
      wMes06 = ADO8!mes06
      wMes07 = ADO8!mes07
      wMes08 = ADO8!mes08
      wMes09 = ADO8!mes09
      wMes10 = ADO8!mes10
      wMes11 = ADO8!mes11
      wMes12 = ADO8!mes12
      wTotal = ADO8!Total
   End If
   Set ADO8 = Nothing

   lblMes01.Caption = Format(wMes01, "###,##0.00")
   lblMes02.Caption = Format(wMes02, "###,##0.00")
   lblMes03.Caption = Format(wMes03, "###,##0.00")
   lblMes04.Caption = Format(wMes04, "###,##0.00")
   lblMes05.Caption = Format(wMes05, "###,##0.00")
   lblMes06.Caption = Format(wMes06, "###,##0.00")
   lblMes07.Caption = Format(wMes07, "###,##0.00")
   lblMes08.Caption = Format(wMes08, "###,##0.00")
   lblMes09.Caption = Format(wMes09, "###,##0.00")
   lblMes10.Caption = Format(wMes10, "###,##0.00")
   lblMes11.Caption = Format(wMes11, "###,##0.00")
   lblMes12.Caption = Format(wMes12, "###,##0.00")
   lblTotal.Caption = Format(wTotal, "###,##0.00")
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
      txtDesde.Text = Format("01/01/" + txtAno.Text, "dd/mm/yyyy")
      txtHasta.Text = Format("31/12/" + txtAno.Text, "dd/mm/yyyy")
      
      txtDesde.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii), 1) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDesde_GotFocus()
   txtDesde.SelStart = 0
   txtDesde.SelLength = 10
End Sub

Private Sub txtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 40
        txtHasta.SetFocus
   End Select
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtDesde.Text = "__/__/____" Then
         MsgBox "Fecha Inicial En Blanco", vbExclamation
         txtDesde.Text = "__/__/____"
         Exit Sub
      End If
      If Not IsDate(txtDesde.Text) Then
         MsgBox "Fecha Inicial Digitado Es Invalido", vbExclamation
         txtDesde.Text = "__/__/____"
         Exit Sub
      End If
      txtHasta.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii), 1) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtHasta_GotFocus()
   txtHasta.SelStart = 0
   txtHasta.SelLength = 10
End Sub

Private Sub txtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDesde.SetFocus
   Case 40
        txtSocio.SetFocus
   End Select
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
   Dim wdes As String, whas As String
   If KeyAscii = 13 Then
      If txtHasta.Text = "__/__/____" Then
         MsgBox "Fecha Final En Blanco", vbExclamation
         txtHasta.Text = "__/__/____"
         Exit Sub
      End If
      If Not IsDate(txtHasta.Text) Then
         MsgBox "Fecha Final Digitado Es Invalido", vbExclamation
         txtHasta.Text = "__/__/____"
         Exit Sub
      End If
      wdes = Format(txtDesde.Text, "yyyy/mm/dd")
      whas = Format(txtHasta.Text, "yyyy/mm/dd")
      If wdes > whas Then
         MsgBox "Rango de Fechas Digitado Es Invalido", vbExclamation
         txtHasta.Text = "__/__/____"
         Exit Sub
      End If
      txtSocio.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii), 1) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtSocio_Change()
   Dim aa As Integer
   aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio.Text)) + " ")
   If aa > 0 Then
      lblSocio.Caption = ADO8!nombre
   Else
      lblSocio.Caption = ""
   End If
   Set ADO8 = Nothing
End Sub

Private Sub txtSocio_GotFocus()
   txtSocio.SelStart = 0
   txtSocio.SelLength = Len(Trim(txtSocio.Text))
End Sub

Private Sub txtSocio_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtHasta.SetFocus
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtSocio.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtSocio_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtSocio.Text)) <> 0 Then
         aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio.Text)) + " ")
         If aa = 0 Then
            MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
            txtSocio.Text = ""
            Exit Sub
         End If
      End If
      cmdBuscar.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

