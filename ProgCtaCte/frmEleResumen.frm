VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEleResumen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Socios Por Tipo"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   12120
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
      Left            =   8640
      TabIndex        =   3
      Top             =   1320
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
      Left            =   8520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   8880
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
      Left            =   7320
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   8880
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
      Left            =   6120
      TabIndex        =   0
      Top             =   8880
      Width           =   975
   End
   Begin Crystal.CrystalReport Crys1 
      Left            =   8520
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5106
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
      Caption         =   "ASOCIADOS ACTIVOS"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4048
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
      Caption         =   "ASOCIADOS NO ACTIVOS"
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   6480
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3413
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
      Caption         =   "FORMA DE PAGO"
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
   Begin MSMask.MaskEdBox txtFecIng 
      Height          =   285
      Left            =   8640
      TabIndex        =   16
      Top             =   300
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblTotDeuda6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10080
      TabIndex        =   27
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label lblTotDeuda3 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10080
      TabIndex        =   26
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblCanDeuda6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8880
      TabIndex        =   25
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Socios que adeudan Mas de Meses"
      Height          =   210
      Left            =   8640
      TabIndex        =   24
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label lblCanDeuda3 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8880
      TabIndex        =   23
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Socios que adeudan de 3 a 6 Meses"
      Height          =   210
      Left            =   8640
      TabIndex        =   22
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label lblCanRei2 
      Alignment       =   2  'Center
      Caption         =   "Ingresos del 01/01/2017 al 31/01/2017"
      Height          =   210
      Left            =   8520
      TabIndex        =   21
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label lblCanRei 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8880
      TabIndex        =   20
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblCanIng2 
      Alignment       =   2  'Center
      Caption         =   "Ingresos del 01/01/2017 al 31/01/2017"
      Height          =   210
      Left            =   8520
      TabIndex        =   19
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Label lblCanIng 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8880
      TabIndex        =   18
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Fecha Ing.Tope"
      Height          =   210
      Left            =   8520
      TabIndex        =   17
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total x Forma Pago"
      Height          =   195
      Index           =   1
      Left            =   5445
      TabIndex        =   15
      Top             =   8400
      Width           =   1380
   End
   Begin VB.Label lblFormPag 
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
      Left            =   6885
      TabIndex        =   14
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Label lblTotales 
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
      Left            =   6840
      TabIndex        =   12
      Top             =   6105
      Width           =   1455
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total Asociados"
      Height          =   195
      Index           =   3
      Left            =   5400
      TabIndex        =   11
      Top             =   6120
      Width           =   1380
   End
   Begin VB.Label lblNoActiv 
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
      Left            =   6840
      TabIndex        =   10
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total No Activos"
      Height          =   195
      Index           =   2
      Left            =   5595
      TabIndex        =   9
      Top             =   5760
      Width           =   1185
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total Activos"
      Height          =   195
      Index           =   0
      Left            =   5760
      TabIndex        =   8
      Top             =   3000
      Width           =   930
   End
   Begin VB.Label lblActivos 
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
      Left            =   6840
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
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
      Top             =   9000
      Width           =   5655
   End
End
Attribute VB_Name = "frmEleResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wCanNuevos As Integer, wCanReingr As Integer, _
    wCanDeuda3 As Integer, wCanDeuda6 As Integer, _
    wTotDeuda3 As Currency, wTotDeuda6 As Currency

Private Sub cmdBuscar_Click()
   LlenaCab
   LlenaCab1
End Sub

Private Sub cmdExporta_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, _
       Heading(2) As String, _
       wreg As Integer, wTot As Integer
   Dim wNom As String, _
       wCan As Integer
   
   Heading(0) = "CODIGO"
   Heading(1) = "NOMBRE"
   Heading(2) = "CANTIDAD"
   
   wTot = aa
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 3)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 3)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "RESUMEN DE SOCIOS POR TIPO"
        For I = 1 To 3 Step 1
            .Cells(3, I) = Heading(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 9
        objExcel.Columns("B").ColumnWidth = 40
        objExcel.Columns("C").ColumnWidth = 12
   End With
   
   V = 4
   aa = Leerado3("SELECT * FROM TMP_RESUMEN WHERE USU = '" + wcodusu + "' ORDER BY CODIGO ")
   If aa > 0 Then
      H = 1
      wreg = 1
      wCan = 0
      Do While Not ADO3.EOF
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wreg, "####0") + " / " + Format(wTot, "####0")
         lblMensaje.Refresh
         
         objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 2)).NumberFormat = "####,##0;;\ "
         
         objExcel.Cells(V, H + 0) = ADO3!codigo
         objExcel.Cells(V, H + 1) = ADO3!nombre
         objExcel.Cells(V, H + 2) = ADO3!num
         
         wCan = wCan + ADO3!num
         wreg = wreg + 1
         V = V + 1
         ADO3.MoveNext
      Loop
      V = V + 1
      
      objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 2)).NumberFormat = "####,##0;;\ "
      objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 2)).Font.Bold = True
      objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 2)).Font.Color = RGB(255, 0, 0)
      objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 2)).Borders.LineStyle = xlContinuous
      objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 2)).Borders.Color = RGB(255, 0, 0)
                  
      objExcel.Cells(V, H + 1) = "TOTALES FINALES"
      objExcel.Cells(V, H + 2) = wCan
      V = V + 1
      
   End If
   
   V = V + 4
   
   aa = Leerado3("SELECT * FROM TMP_RESUMEN2 WHERE USU = '" + wcodusu + "' ORDER BY CODIGO ")
   If aa > 0 Then
      H = 1
      wreg = 1
      wCan = 0
      Do While Not ADO3.EOF
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wreg, "####0") + " / " + Format(wTot, "####0")
         lblMensaje.Refresh
         
         objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 2)).NumberFormat = "####,##0;;\ "
         
         objExcel.Cells(V, H + 0) = ADO3!codigo
         objExcel.Cells(V, H + 1) = ADO3!nombre
         objExcel.Cells(V, H + 2) = ADO3!num
         
         wCan = wCan + ADO3!num
         wreg = wreg + 1
         V = V + 1
         ADO3.MoveNext
      Loop
      V = V + 1
      
      objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 2)).NumberFormat = "####,##0;;\ "
      objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 2)).Font.Bold = True
      objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 2)).Font.Color = RGB(255, 0, 0)
      objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 2)).Borders.LineStyle = xlContinuous
      objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 2)).Borders.Color = RGB(255, 0, 0)
                  
      objExcel.Cells(V, H + 1) = "TOTALES FINALES"
      objExcel.Cells(V, H + 2) = wCan
      V = V + 1
      
   End If
   
   V = V + 4
   
   wCan = 0
   aa = Leerado3("SELECT * FROM TMP_RESGRADO WHERE USU = '" + wcodusu + "' ORDER BY CODIGO ")
   If aa > 0 Then
      ADO3.MoveFirst
      Do While Not ADO3.EOF
         objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 2)).NumberFormat = "####,##0;;\ "
   
         objExcel.Cells(V, H + 1) = ADO3!nombre
         objExcel.Cells(V, H + 2) = ADO3!num
         
         wCan = wCan + ADO3!num
         V = V + 1
         ADO3.MoveNext
      Loop
   End If
   V = V + 1
   
   objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 2)).NumberFormat = "####,##0;;\ "
   
   objExcel.Cells(V, H + 1) = "TOTALES"
   objExcel.Cells(V, H + 2) = wCan
   V = V + 1
         
   Set ADO3 = Nothing
   objExcel.Visible = True
   Set objExcel = Nothing
   
   lblMensaje.Caption = ""
   lblMensaje.Refresh
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdImprimir_Click()
   Dim wFecing As Date
   wFecing = Format(txtFecIng.Text, "dd/mm/yyyy")
   
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\ResumenSocios.RPT"
   Crys1.Formulas(0) = "GLOSA1= 'INGRESOS DE NUEVOS ASOCIADOS DEL 01/01/" + wanocia + " AL " + Format(wFecing, "") + "' "
   Crys1.Formulas(1) = "GLOSA2= 'REINGRESODE ASOCIADOS DEL 01/01/" + wanocia + " AL " + Format(wFecing, "") + "' "
   Crys1.Formulas(2) = "CANT1= '" + Format(wCanNuevos, "##,##0") + "' "
   Crys1.Formulas(3) = "CANT2= '" + Format(wCanReingr, "##,##0") + "' "
   Crys1.Formulas(4) = "CANT3= '" + Format(wCanDeuda3, "##,##0") + "' "
   Crys1.Formulas(5) = "CANT4= '" + Format(wCanDeuda6, "##,##0") + "' "
   Crys1.Formulas(6) = "TOTAL3= '" + Format(wTotDeuda3, "##,###,##0.00") + "' "
   Crys1.Formulas(7) = "TOTAL4= '" + Format(wTotDeuda6, "##,###,##0.00") + "' "
   Crys1.SelectionFormula = " {TMP_ELE.USU}='" + wcodusu + "' "
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
   frmEleResumen.Left = (Screen.Width - Width) \ 2
   frmEleResumen.Top = 0
   
   txtFecIng.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing
   Set DataGrid2.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESUMEN WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESGRADO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_ELE WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim aa As Long, _
       wTotSoc As Integer, wTotViu As Integer, wTotHij As Integer, wTotHer As Integer, wTotNie As Integer, _
       wTotTra As Integer, wTotCiv As Integer, wTotCi1 As Integer, wTotHon As Integer, wTotFal As Integer, _
       wTotRen As Integer, wTotSep As Integer, wTotExc As Integer, wTotEsp As Integer, wTotTit As Integer, _
       wTotDie As Integer, wTotCaj As Integer, wTotCMP As Integer, wtotNoa As Integer, wTotExp As Integer, _
       wTotAct As Integer, wTotAc2 As Integer, wTotSus As Integer, wTotPnp As Integer, wTotFor As Integer, _
       wFecing As Date

   If txtFecIng.Text = "__/__/____" Then
      MsgBox "Fecha Ingreso Tope En Blanco", vbExclamation
      txtFecIng.Text = "__/__/____"
      Exit Sub
   End If
   If Not IsDate(txtFecIng.Text) Then
      MsgBox "Fecha Ingreso Tope Digitada Es Invalida", vbExclamation
      txtFecIng.Text = "__/__/____"
      Exit Sub
   End If
   wFecing = Format(txtFecIng.Text, "dd/mm/yyyy")
   
   lblMensaje.Caption = "Preparando Archivo......Espere"
   lblMensaje.Refresh

   Set DataGrid1.DataSource = Nothing
   DataGrid1.AllowUpdate = False
   DataGrid2.AllowUpdate = False
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESUMEN WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESUMEN2 WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESGRADO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_ELE " _
   & " (COD, USU) VALUES ('00', '" + wcodusu + "') ")
   Db.CommitTrans

   wTotSoc = 0: wTotViu = 0: wTotHij = 0: wTotHer = 0: wTotFal = 0: wTotCiv = 0: wTotCi1 = 0
   wTotSus = 0: wTotPnp = 0
   
   
   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'TIT' ")
   If aa > 0 Then
      wTotTit = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('1', '01', 'ACTIVOS - PIP / PNP     ', " + Str(wTotTit) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans
   
   
   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'VIU' ")
   If aa > 0 Then
      wTotViu = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('1', '02', 'ACTIVO - VIUDAS  ', " + Str(wTotViu) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans


   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'HIJ' ")
   If aa > 0 Then
      wTotHij = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('1', '03', 'ACTIVOS - HIJOS', " + Str(wTotHij) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans


   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'HER' ")
   If aa > 0 Then
      wTotHer = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('1', '04', 'ACTIVOS - HERMANOS', " + Str(wTotHer) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans


   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'NIE' ")
   If aa > 0 Then
      wTotNie = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('1', '05', 'ACTIVOS - NIETOS', " + Str(wTotNie) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans
   
   
   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'TRA' ")
   If aa > 0 Then
      wTotTra = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('1', '06', 'ADHERENTES I', " + Str(wTotTra) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans


   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'CIV' ")
   If aa > 0 Then
      wTotCiv = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('1', '07', 'ADHERENTES II', " + Str(wTotCiv) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans
   
   
   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'CI1' ")
   If aa > 0 Then
      wTotCi1 = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('1', '08', 'ADHERENTES III', " + Str(wTotCi1) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans


   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'HON' ")
   If aa > 0 Then
      wTotHon = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('1', '09', 'HONORARIOS', " + Str(wTotHon) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans

   
   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'PNP' ")
   If aa > 0 Then
      wTotPnp = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('1', '10', 'P.N.P.', " + Str(wTotPnp) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans

   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'FAL' ")
   If aa > 0 Then
      wTotFal = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('2', '11', 'FALLECIDOS', " + Str(wTotFal) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN2 " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('2', '01', 'FALLECIDOS', " + Str(wTotFal) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans



   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'REN' ")
   If aa > 0 Then
      wTotRen = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('2', '12', 'RENUNCIANTES', " + Str(wTotRen) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN2 " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('2', '02', 'RENUNCIANTES', " + Str(wTotRen) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans


   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'SEP' ")
   If aa > 0 Then
      wTotSep = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('2', '13', 'SEPARADOS POR DEUDA', " + Str(wTotSep) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN2 " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('2', '03', 'SEPARADOS POR DEUDA', " + Str(wTotSep) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans


   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'EXC' ")
   If aa > 0 Then
      wTotExc = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('2', '14', 'EXCLUIDOS', " + Str(wTotExc) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN2 " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('2', '04', 'EXCLUIDOS', " + Str(wTotExc) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans


   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'ESP' ")
   If aa > 0 Then
      wTotEsp = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('2', '15', 'EN ESPERA', " + Str(wTotEsp) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN2 " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('2', '05', 'EN ESPERA', " + Str(wTotEsp) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans


   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'EXP' ")
   If aa > 0 Then
      wTotExp = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('2', '16', 'EXPULSADOS', " + Str(wTotExp) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN2 " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('2', '06', 'EXPULSADOS', " + Str(wTotExp) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans

   
   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE E_SOCIO = 'SUS' ")
   If aa > 0 Then
      wTotSus = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('2', '17', 'SUSPENDIDO', " + Str(wTotSus) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESUMEN2 " _
   & " (TIPO, CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('2', '07', 'SUSPENDIDO', " + Str(wTotSus) + ", '" + wcodusu + "' ) ")
   Db.CommitTrans

   wTotAct = wTotTit + wTotViu + wTotHij + wTotHer + wTotNie + wTotTra + wTotCiv + wTotCi1 + wTotHon + wTotPnp
   wTotAc2 = wTotFal + wTotRen + wTotSep + wTotExc + wTotEsp + wTotExp + wTotSus
   wTotSoc = wTotAct + wTotAc2
   
   lblTotales.Caption = Format(wTotSoc, "###,##0;;\ ")
   lblActivos.Caption = Format(wTotAct, "###,##0;;\ ")
   lblNoActiv.Caption = Format(wTotAc2, "###,##0;;\ ")
   
   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE (TIPCOB = '01') AND " _
                & "       (E_SOCIO='TIT' OR E_SOCIO='VIU' OR E_SOCIO='HIJ' OR E_SOCIO='HER' OR " _
                & "        E_SOCIO='NIE' OR E_SOCIO='TRA' OR E_SOCIO='CIV' OR E_SOCIO='CI1' OR " _
                & "        E_SOCIO='PNP') ")
   If aa > 0 Then
      wTotDie = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESGRADO " _
   & " (CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('01', 'DIECO', " _
   & "  " + Str(wTotDie) + ", '" + wcodusu + "'  )  ")
   Db.CommitTrans
  
   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE (TIPCOB = '02') AND " _
                & "       (E_SOCIO='TIT' OR E_SOCIO='VIU' OR E_SOCIO='HIJ' OR E_SOCIO='HER' OR " _
                & "        E_SOCIO='NIE' OR E_SOCIO='TRA' OR E_SOCIO='CIV' OR E_SOCIO='CI1' OR " _
                & "        E_SOCIO='PNP') ")
   If aa > 0 Then
      wTotCMP = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESGRADO " _
   & " (CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('02', 'CAJA MILITAR POLICIAL', " _
   & "  " + Str(wTotCMP) + ", '" + wcodusu + "'  )  ")
   Db.CommitTrans
   
   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE (TIPCOB = '03') AND " _
                & "       (E_SOCIO='TIT' OR E_SOCIO='VIU' OR E_SOCIO='HIJ' OR E_SOCIO='HER' OR " _
                & "        E_SOCIO='NIE' OR E_SOCIO='TRA' OR E_SOCIO='CIV' OR E_SOCIO='CI1' OR " _
                & "        E_SOCIO='PNP') ")
   If aa > 0 Then
      wTotCaj = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESGRADO " _
   & " (CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('03', 'TESORERIA AOPIP', " _
   & "  " + Str(wTotCaj) + ", '" + wcodusu + "'  )  ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESGRADO " _
   & " (CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('04', 'NO APORTAN HONORARIOS', " _
   & "  " + Str(wTotHon) + ", '" + wcodusu + "'  )  ")
   Db.CommitTrans
      
   wTotFor = wTotDie + wTotCMP + wTotCaj + wTotHon
   
   lblFormPag.Caption = Format(wTotFor, "###,##0;;\ ")

   wCanNuevos = 0: wCanReingr = 0
   wCanDeuda3 = 0: wCanDeuda6 = 0: wTotDeuda3 = 0: wTotDeuda6 = 0
   aa = Leerado8("SELECT * FROM MAESOCIO " _
                & " WHERE FECING >= '01/01/2017' AND " _
                & "       FECING <= '" + Format(wFecing, "dd/mm/yyyy") + "' ")
   If aa > 0 Then
      wCanNuevos = aa
   End If
   Set ADO8 = Nothing

   aa = Leerado8("SELECT * FROM MAESOCIO " _
                & " WHERE FECREIN >= '01/01/2017' AND " _
                & "       FECREIN <= '" + Format(wFecing, "dd/mm/yyyy") + "' ")
   If aa > 0 Then
      wCanReingr = aa
   End If
   Set ADO8 = Nothing


   aa = Leerado8("SELECT COUNT(*) AS TOT, SUM(DEUDA_PT2) DEUDA " _
            & " FROM MAESOCIO AS M INNER JOIN MAEE_SOCIO AS E " _
            & "   ON M.E_SOCIO = E.E_SOCIO " _
            & " WHERE (M.E_SOCIO='TIT' OR M.E_SOCIO='VIU' OR M.E_SOCIO='HIJ' OR M.E_SOCIO='HER' OR " _
            & "        M.E_SOCIO='NIE' OR M.E_SOCIO='TRA' OR M.E_SOCIO='CIV' OR M.E_SOCIO='CI1' OR " _
            & "        M.E_SOCIO='PNP') AND " _
            & "       (M.DEUDA_PT2 > ROUND(3 * E.APORTE,2) AND M.DEUDA_PT2 <= ROUND(6 * E.APORTE,2)) ")
   If aa > 0 Then
      wCanDeuda3 = ADO8!tot
      wTotDeuda3 = ADO8!deuda
   End If
   Set ADO8 = Nothing

   aa = Leerado8("SELECT COUNT(*) AS TOT, SUM(DEUDA_PT2) DEUDA " _
            & " FROM MAESOCIO AS M INNER JOIN MAEE_SOCIO AS E " _
            & "   ON M.E_SOCIO = E.E_SOCIO " _
            & " WHERE (M.E_SOCIO='TIT' OR M.E_SOCIO='VIU' OR M.E_SOCIO='HIJ' OR M.E_SOCIO='HER' OR " _
            & "        M.E_SOCIO='NIE' OR M.E_SOCIO='TRA' OR M.E_SOCIO='CIV' OR M.E_SOCIO='CI1' OR " _
            & "        M.E_SOCIO='PNP') AND " _
            & "       (M.DEUDA_PT2 > ROUND(6 * E.APORTE,2)) ")
   If aa > 0 Then
      wCanDeuda6 = ADO8!tot
      wTotDeuda6 = ADO8!deuda
   End If
   Set ADO8 = Nothing

   lblCanIng2.Caption = "Ingresos del 01/01/" + wanocia + " Al " + Format(wFecing, "dd/mm/yyyy")
   lblCanIng.Caption = Format(wCanNuevos, "##,##0")

   lblCanRei2.Caption = "Reingresos del 01/01/" + wanocia + " Al " + Format(wFecing, "dd/mm/yyyy")
   lblCanRei.Caption = Format(wCanReingr, "##,##0")

   lblCanDeuda3.Caption = Format(wCanDeuda3, "##,##0")
   lblCanDeuda6.Caption = Format(wCanDeuda6, "##,##0")
   lblTotDeuda3.Caption = Format(wTotDeuda3, "##,##0.00")
   lblTotDeuda6.Caption = Format(wTotDeuda6, "##,##0.00")

   lblMensaje.Caption = ""
   lblMensaje.Refresh

   aa = Leerado2("SELECT CODIGO, NOMBRE, NUM, USU " _
                & " FROM TMP_RESUMEN " _
                & " WHERE  USU = '" + wcodusu + "' AND " _
                & "       TIPO = '1' " _
                & " ORDER BY CODIGO ")
   Set DataGrid1.DataSource = ADO2

   aa = Leerado2("SELECT CODIGO, NOMBRE, NUM, USU " _
                & " FROM TMP_RESUMEN2 " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY CODIGO ")
   Set DataGrid2.DataSource = ADO2

   aa = Leerado3("SELECT CODIGO, NOMBRE, NUM, USU " _
                & " FROM TMP_RESGRADO " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY CODIGO ")
   Set DataGrid3.DataSource = ADO3

End Sub

Private Sub LlenaCab1()
   DataGrid1.Columns(0).Width = 600
   DataGrid1.Columns(0).Alignment = dbgCenter
   DataGrid1.Columns(0).Caption = "CODIG"
   
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Width = 6100
   DataGrid1.Columns(1).Caption = "DESCRIPCION"

   DataGrid1.Columns(2).Alignment = dbgRight
   DataGrid1.Columns(2).Width = 1000
   DataGrid1.Columns(2).Caption = "CANT."
   DataGrid1.Columns(2).NumberFormat = "##,##0"

   DataGrid1.Columns(3).Visible = False

   DataGrid2.Columns(0).Width = 600
   DataGrid2.Columns(0).Alignment = dbgCenter
   DataGrid2.Columns(0).Caption = "CODIG"
   
   DataGrid2.Columns(1).Alignment = dbgLeft
   DataGrid2.Columns(1).Width = 6100
   DataGrid2.Columns(1).Caption = "NOMBRE"

   DataGrid2.Columns(2).Alignment = dbgRight
   DataGrid2.Columns(2).Width = 1000
   DataGrid2.Columns(2).Caption = "CANT."
   DataGrid2.Columns(2).NumberFormat = "##,##0"

   DataGrid2.Columns(3).Visible = False

   DataGrid3.Columns(0).Width = 600
   DataGrid3.Columns(0).Alignment = dbgCenter
   DataGrid3.Columns(0).Caption = "CODIG"
   
   DataGrid3.Columns(1).Alignment = dbgLeft
   DataGrid3.Columns(1).Width = 6100
   DataGrid3.Columns(1).Caption = "NOMBRE"

   DataGrid3.Columns(2).Alignment = dbgRight
   DataGrid3.Columns(2).Width = 1000
   DataGrid3.Columns(2).Caption = "CANT."
   DataGrid3.Columns(2).NumberFormat = "##,##0"

   DataGrid3.Columns(3).Visible = False
End Sub

Private Sub txtFecIng_GotFocus()
   txtFecIng.SelStart = 0
   txtFecIng.SelLength = 10
End Sub

Private Sub txtFecIng_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFecIng.Text = "__/__/____" Then
         MsgBox "Fecha Ingreso Tope En Blanco", vbExclamation
         txtFecIng.Text = "__/__/____"
         Exit Sub
      End If
      If Not IsDate(txtFecIng.Text) Then
         MsgBox "Fecha Ingreso Tope Digitada Es Invalida", vbExclamation
         txtFecIng.Text = "__/__/____"
         Exit Sub
      End If
      cmdBuscar.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub


