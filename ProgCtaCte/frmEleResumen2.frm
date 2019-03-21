VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmEleResumen2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuadro Resumen"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   14985
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
      Left            =   13560
      TabIndex        =   3
      Top             =   5640
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
      Left            =   8640
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   8040
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
      Left            =   7440
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   8040
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
      Left            =   6240
      TabIndex        =   0
      Top             =   8040
      Width           =   975
   End
   Begin Crystal.CrystalReport Crys1 
      Left            =   10440
      Top             =   6360
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
      Height          =   4815
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   8493
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
      Caption         =   "RESUMEN POR TIPO DE SOCIO"
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
      Height          =   1815
      Left            =   720
      TabIndex        =   6
      Top             =   5640
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3201
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
      Caption         =   "RESUMEN POR FORMA DE PAGO"
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
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cant=0"
      Height          =   195
      Index           =   9
      Left            =   1440
      TabIndex        =   26
      Top             =   5280
      Width           =   510
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Tot.x Tipo"
      Height          =   195
      Index           =   8
      Left            =   13590
      TabIndex        =   25
      Top             =   5280
      Width           =   720
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Can.x Tipo"
      Height          =   195
      Index           =   7
      Left            =   12585
      TabIndex        =   24
      Top             =   5280
      Width           =   780
   End
   Begin VB.Label lblTot9 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   13320
      TabIndex        =   23
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label lblCan9 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   12600
      TabIndex        =   22
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblCan0 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1320
      TabIndex        =   21
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cant<3"
      Height          =   195
      Index           =   6
      Left            =   2160
      TabIndex        =   20
      Top             =   5280
      Width           =   510
   End
   Begin VB.Label lblCan3 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cant>3 <6"
      Height          =   195
      Index           =   5
      Left            =   4080
      TabIndex        =   18
      Top             =   5280
      Width           =   750
   End
   Begin VB.Label lblCan6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4080
      TabIndex        =   17
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cant>36"
      Height          =   195
      Index           =   4
      Left            =   10755
      TabIndex        =   16
      Top             =   5280
      Width           =   600
   End
   Begin VB.Label lblCan7 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   10680
      TabIndex        =   15
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblTot3 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Total < 3 Meses"
      Height          =   195
      Index           =   3
      Left            =   2880
      TabIndex        =   13
      Top             =   5280
      Width           =   1140
   End
   Begin VB.Label lblTot6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Total > 3  y < 6"
      Height          =   195
      Index           =   2
      Left            =   4800
      TabIndex        =   11
      Top             =   5280
      Width           =   1200
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total Asociados"
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
      Index           =   1
      Left            =   5910
      TabIndex        =   10
      Top             =   7440
      Width           =   1380
   End
   Begin VB.Label lblHabiles 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7410
      TabIndex        =   9
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Total > 36 Meses"
      Height          =   195
      Index           =   0
      Left            =   11355
      TabIndex        =   8
      Top             =   5280
      Width           =   1230
   End
   Begin VB.Label lblTot7 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   11370
      TabIndex        =   7
      Top             =   5040
      Width           =   1215
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
      TabIndex        =   5
      Top             =   7920
      Width           =   5415
   End
End
Attribute VB_Name = "frmEleResumen2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
   LlenaCab
   LlenaCab1
End Sub

Private Sub cmdExporta_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, _
       Heading(10) As String, Headin2(10) As String, _
       wreg As Integer, wTot As Integer
   Dim wNom As String, _
       wCan0 As Integer, wCan3 As Integer, wCan6 As Integer, wCan7 As Integer, wCan9 As Integer, _
       wTot0 As Currency, wTot3 As Currency, wTot6 As Currency, wTot7 As Currency, wTot9 As Currency
   
   Heading(0) = "TIPO"
   Heading(1) = "SOCIOS SIN SALDO"
   Heading(3) = "DEUDA HASTA 3 MESES"
   Heading(5) = "DEUDA HASTA 6 MESES"
   Heading(7) = "DEUDA > 6 MESES"
   Heading(9) = "TOTAL X TIPO"
   
   Headin2(0) = "SOCIO"
   Headin2(1) = "CANT"
   Headin2(2) = "TOTAL"
   Headin2(3) = "CANT"
   Headin2(4) = "TOTAL"
   Headin2(5) = "CANT"
   Headin2(6) = "TOTAL"
   Headin2(7) = "CANT"
   Headin2(8) = "TOTAL"
   Headin2(9) = "CANT"
   Headin2(10) = "TOTAL"
   
   aa = Leerado3("SELECT * FROM TMP_RESUMEN3 WHERE USU = '" + wcodusu + "' ORDER BY ORDEN ")
   If aa > 0 Then
      wTot = aa
      Set objExcel = New Excel.Application
      objExcel.SheetsInNewWorkbook = 1
      objExcel.Workbooks.Add
      With objExcel
           .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
           .Range(.Cells(3, 1), .Cells(3, 11)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 11)).Font.Bold = True
           .Cells(1, 1) = wnomcia
           .Cells(2, 1) = "CUADRO RESUMEN DE APORTACIONES"
           For I = 1 To 11 Step 1
               .Cells(3, I) = Heading(I - 1)
               .Cells(4, I) = Headin2(I - 1)
           Next
           objExcel.Columns("A").ColumnWidth = 25
           objExcel.Columns("B").ColumnWidth = 9
           objExcel.Columns("C").ColumnWidth = 12
           objExcel.Columns("D").ColumnWidth = 9
           objExcel.Columns("E").ColumnWidth = 12
           objExcel.Columns("F").ColumnWidth = 9
           objExcel.Columns("G").ColumnWidth = 12
           objExcel.Columns("H").ColumnWidth = 9
           objExcel.Columns("I").ColumnWidth = 12
           objExcel.Columns("J").ColumnWidth = 9
           objExcel.Columns("K").ColumnWidth = 12
      End With
      V = 5
      H = 1
      wreg = 1
      wCan0 = 0: wCan6 = 0: wCan7 = 0: wCan9 = 0
      wTot0 = 0: wTot6 = 0: wTot7 = 0: wTot9 = 0
      Do While Not ADO3.EOF
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wreg, "####0") + " / " + Format(wTot, "####0")
         lblMensaje.Refresh
         
         objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 1)).NumberFormat = "####,##0;;\ "
         objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 2)).NumberFormat = "####,##0.00;;\ "
         objExcel.Range(objExcel.Cells(V, H + 3), objExcel.Cells(V, H + 3)).NumberFormat = "####,##0;;\ "
         objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 4)).NumberFormat = "####,##0.00;;\ "
         objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 5)).NumberFormat = "####,##0;;\ "
         objExcel.Range(objExcel.Cells(V, H + 6), objExcel.Cells(V, H + 6)).NumberFormat = "####,##0.00;;\ "
         objExcel.Range(objExcel.Cells(V, H + 7), objExcel.Cells(V, H + 7)).NumberFormat = "####,##0;;\ "
         objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 8)).NumberFormat = "####,##0.00;;\ "
         objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 9)).NumberFormat = "####,##0;;\ "
         objExcel.Range(objExcel.Cells(V, H + 10), objExcel.Cells(V, H + 10)).NumberFormat = "####,##0.00;;\ "
         
         objExcel.Cells(V, H + 0) = ADO3!nombre
         objExcel.Cells(V, H + 1) = ADO3!can0
         objExcel.Cells(V, H + 2) = ADO3!tot0
         objExcel.Cells(V, H + 3) = ADO3!can3
         objExcel.Cells(V, H + 4) = ADO3!tot3
         objExcel.Cells(V, H + 5) = ADO3!can6
         objExcel.Cells(V, H + 6) = ADO3!tot6
         objExcel.Cells(V, H + 7) = ADO3!can7
         objExcel.Cells(V, H + 8) = ADO3!tot7
         objExcel.Cells(V, H + 9) = ADO3!can9
         objExcel.Cells(V, H + 10) = ADO3!tot9
         
         wCan0 = wCan0 + ADO3!can0
         wCan3 = wCan3 + ADO3!can3
         wCan6 = wCan6 + ADO3!can6
         wCan7 = wCan7 + ADO3!can7
         wCan9 = wCan9 + ADO3!can9
         
         wTot0 = wTot0 + ADO3!tot0
         wTot3 = wTot3 + ADO3!tot3
         wTot6 = wTot6 + ADO3!tot6
         wTot7 = wTot7 + ADO3!tot7
         wTot9 = wTot9 + ADO3!tot9
         
         wreg = wreg + 1
         V = V + 1
         ADO3.MoveNext
      Loop
      V = V + 1
      
      objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 1)).NumberFormat = "####,##0;;\ "
      objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 2)).NumberFormat = "####,##0.00;;\ "
      objExcel.Range(objExcel.Cells(V, H + 3), objExcel.Cells(V, H + 3)).NumberFormat = "####,##0;;\ "
      objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 4)).NumberFormat = "####,##0.00;;\ "
      objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 5)).NumberFormat = "####,##0;;\ "
      objExcel.Range(objExcel.Cells(V, H + 6), objExcel.Cells(V, H + 6)).NumberFormat = "####,##0.00;;\ "
      objExcel.Range(objExcel.Cells(V, H + 7), objExcel.Cells(V, H + 7)).NumberFormat = "####,##0;;\ "
      objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 8)).NumberFormat = "####,##0.00;;\ "
      objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 9)).NumberFormat = "####,##0;;\ "
      objExcel.Range(objExcel.Cells(V, H + 10), objExcel.Cells(V, H + 10)).NumberFormat = "####,##0.00;;\ "
      
      objExcel.Range(objExcel.Cells(V, H + 0), objExcel.Cells(V, H + 10)).Font.Bold = True
      objExcel.Range(objExcel.Cells(V, H + 0), objExcel.Cells(V, H + 10)).Font.Color = RGB(255, 0, 0)
      objExcel.Range(objExcel.Cells(V, H + 0), objExcel.Cells(V, H + 10)).Borders.LineStyle = xlContinuous
      objExcel.Range(objExcel.Cells(V, H + 0), objExcel.Cells(V, H + 10)).Borders.Color = RGB(255, 0, 0)
                  
      objExcel.Cells(V, H + 0) = "TOTALES FINALES"
      objExcel.Cells(V, H + 1) = wCan0
      objExcel.Cells(V, H + 2) = wTot0
      objExcel.Cells(V, H + 3) = wCan3
      objExcel.Cells(V, H + 4) = wTot3
      objExcel.Cells(V, H + 5) = wCan6
      objExcel.Cells(V, H + 6) = wTot6
      objExcel.Cells(V, H + 7) = wCan7
      objExcel.Cells(V, H + 8) = wTot7
      objExcel.Cells(V, H + 9) = wCan9
      objExcel.Cells(V, H + 10) = wTot9
         
      V = V + 1
      
   End If
   
   V = V + 4
   
   wCan0 = 0
   aa = Leerado3("SELECT * FROM TMP_RESGRADO WHERE USU = '" + wcodusu + "' ORDER BY CODIGO ")
   If aa > 0 Then
      ADO3.MoveFirst
      Do While Not ADO3.EOF
         objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 1)).NumberFormat = "####,##0;;\ "
   
         objExcel.Cells(V, H + 0) = ADO3!nombre
         objExcel.Cells(V, H + 1) = ADO3!num
         
         wCan0 = wCan0 + ADO3!num
         V = V + 1
         ADO3.MoveNext
      Loop
   End If
   V = V + 1
   
   objExcel.Range(objExcel.Cells(V, H + 1), objExcel.Cells(V, H + 1)).NumberFormat = "####,##0;;\ "
   
   objExcel.Range(objExcel.Cells(V, H + 0), objExcel.Cells(V, H + 1)).Font.Bold = True
   objExcel.Range(objExcel.Cells(V, H + 0), objExcel.Cells(V, H + 1)).Font.Color = RGB(255, 0, 0)
   objExcel.Range(objExcel.Cells(V, H + 0), objExcel.Cells(V, H + 1)).Borders.LineStyle = xlContinuous
   objExcel.Range(objExcel.Cells(V, H + 0), objExcel.Cells(V, H + 1)).Borders.Color = RGB(255, 0, 0)
   
   objExcel.Cells(V, H + 0) = TOTALES
   objExcel.Cells(V, H + 1) = wCan0
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
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\ResumenSocios2.RPT"
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
   frmEleResumen2.Left = (Screen.Width - Width) \ 2
   frmEleResumen2.Top = 0
   
   cmdBuscar.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing
   Set DataGrid2.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESUMEN3 WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESGRADO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_ELE ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim aa As Long, wRegAct As Long, wRegTot As Long, _
       wCan00 As Integer, wCan03 As Integer, wCan06 As Integer, wCan12 As Integer, wCan24 As Integer, wCan36 As Integer, wCan37 As Integer, wCan99 As Integer, _
       wTot00 As Currency, wTot03 As Currency, wTot06 As Currency, wTot12 As Currency, wTot24 As Currency, wTot36 As Currency, wTot37 As Currency, wTot99 As Currency, _
       wTip As String, wNom As String, wApo As Currency, _
       wTop00 As Currency, wTop03 As Currency, wTop06 As Currency, _
       wCaz00 As Integer, wCaz03 As Integer, wCaz06 As Integer, wCaz07 As Currency, wCaz09 As Integer, _
       wToz00 As Currency, wToz03 As Currency, wToz06 As Currency, wToz07 As Currency, wToz09 As Currency, _
       wSoc As Integer, wCod As Long, wIns As Integer, wSdo As Currency, _
       wDieco As Integer, wCajMP As Integer, wE_S As String, wMes As String

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

   wCajMP = 1
   aa = Leerado8("SELECT SUM(TOTENVIO) AS TOTENVIO " _
                & " FROM CAJMPCAB " _
                & " WHERE MES = '" + wMes + "' ")
   If aa > 0 Then
      wCajMP = 2
   End If
   Set ADO8 = Nothing

   lblMensaje.Caption = "Preparando Archivo......Espere"
   lblMensaje.Refresh

   Set DataGrid1.DataSource = Nothing
   DataGrid1.AllowUpdate = False
   DataGrid2.AllowUpdate = False
   
   wMes = Left(zMesTope, 4) + "/" + Right(zMesTope, 2)
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESUMEN3 WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESGRADO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_ELE " _
   & " (COD, USU) VALUES ('00', '" + wcodusu + "') ")
   Db.CommitTrans
   
   aa = Leerado8a("select S.CODSOCIO, S.CODIGO, S.INS, S.NOMBRE, " _
                & "      S.E_SOCIO, E.NOMBRE AS NOME_S, E.APORTE, E.MONEDA, S.TIPCOB, E.ORDEN " _
                & " from MAESOCIO AS S INNER JOIN MAEE_SOCIO AS E " _
                & "   ON S.E_SOCIO = E.E_SOCIO " _
                & " WHERE E.ORDEN >0")
   If aa > 0 Then
      ADO8a.MoveFirst
      wRegAct = 1
      wRegTot = aa
      Do While Not ADO8a.EOF
         DoEvents
         lblMensaje.Caption = "Registro " + _
                              Trim(Format(wRegAct, "##,##0")) + " / " + _
                              Trim(Format(wRegTot, "##,##0"))
         lblMensaje.Refresh
         
         wSoc = ADO8a!codsocio
         wCod = ADO8a!codigo
         wIns = ADO8a!ins
         wE_S = ADO8a!e_socio
         wApo = ADO8a!aporte
         wSdo = SaldoFoto(wSoc, wMes)
         If wSdo < 0 Then
            wSdo = 0
         End If
         
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
         
         wCan00 = 0: wCan03 = 0: wCan06 = 0: wCan12 = 0: wCan24 = 0: wCan36 = 0: wCan37 = 0
         wTot00 = 0: wTot03 = 0: wTot06 = 0: wTot12 = 0: wCan24 = 0: wCan36 = 0: wTot37 = 0
          
         Select Case True
         Case wSdo <= 0
              wCan00 = 1
              wTot00 = wSdo
         Case wSdo > 0 And wSdo <= Round(wApo * 3, 2)
              wCan03 = 1
              wTot03 = wSdo
         Case wSdo > Round(wApo * 3, 2) And wSdo <= Round(wApo * 6, 2)
              wCan06 = 1
              wTot06 = wSdo
         Case wSdo > Round(wApo * 6, 2) And wSdo <= Round(wApo * 12, 2)
              wCan12 = 1
              wTot12 = wSdo
         Case wSdo > Round(wApo * 12, 2) And wSdo <= Round(wApo * 24, 2)
              wCan24 = 1
              wTot24 = wSdo
         Case wSdo > Round(wApo * 24, 2) And wSdo <= Round(wApo * 36, 2)
              wCan36 = 1
              wTot36 = wSdo
         Case Else
              wCan37 = 1
              wTot37 = wSdo
         End Select
         
         wCan99 = wCan00 + wCan03 + wCan06 + wCan12 + wCan24 + wCan36 + wCan37
         wTot99 = wTot00 + wTot03 + wTot06 + wTot12 + wTot24 + wTot36 + wTot37
         
         wCaz00 = wCaz00 + wCan00
         wCaz03 = wCaz03 + wCan03
         wCaz06 = wCaz06 + wCan06
         wCaz12 = wCaz12 + wCan12
         wCaz24 = wCaz24 + wCan24
         wCaz36 = wCaz36 + wCan36
         wCaz37 = wCaz37 + wCan37
         wCaz99 = wCaz99 + wCan99
         
         wToz00 = wToz00 + wTot00
         wToz03 = wToz03 + wTot03
         wToz06 = wToz06 + wTot06
         wToz12 = wToz12 + wTot12
         wToz24 = wToz24 + wTot24
         wToz36 = wToz36 + wTot36
         wToz37 = wToz37 + wTot37
         wToz99 = wToz99 + wTot99
         
         aa = Leerado7a("SELECT * FROM TMP_RESUMEN3 WHERE E_SOCIO = '" + wE_S + "' AND USU = '" + wcodusu + "' ")
         If aa = 0 Then
            Db.BeginTrans
            Db.Execute ("INSERT INTO TMP_RESUMEN3 " _
            & " (E_SOCIO, NOMBRE, ORDEN, " _
            & "  CAN00, TOT00, CAN03, TOT03, CAN06, TOT06, CAN12, TOT12, CAN24, TOT24, CAN36, TOT36, CAN37, TOT37, CAN99, TOT99, USU) " _
            & " VALUES " _
            & " ('" + wE_S + "', '" + Trim(ADO8a!nome_s) + "', " + Str(ADO8a!orden) + ", " _
            & "  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '" + wcodusu + "') ")
            Db.CommitTrans
         End If
         
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_RESUMEN3 " _
         & " SET CAN00 = CAN00 + " + Str(wCan00) + ", TOT00 = TOT00 + " + Str(wTot00) + ", " _
         & "     CAN03 = CAN03 + " + Str(wCan03) + ", TOT03 = TOT03 + " + Str(wTot03) + ", " _
         & "     CAN06 = CAN06 + " + Str(wCan06) + ", TOT06 = TOT06 + " + Str(wTot06) + ", " _
         & "     CAN12 = CAN12 + " + Str(wCan12) + ", TOT12 = TOT12 + " + Str(wTot12) + ", " _
         & "     CAN24 = CAN24 + " + Str(wCan24) + ", TOT24 = TOT24 + " + Str(wTot24) + ", " _
         & "     CAN36 = CAN36 + " + Str(wCan36) + ", TOT36 = TOT36 + " + Str(wTot36) + ", " _
         & "     CAN37 = CAN37 + " + Str(wCan37) + ", TOT37 = TOT37 + " + Str(wTot37) + ", " _
         & "     CAN99 = CAN99 + " + Str(wCan99) + ", TOT99 = TOT99 + " + Str(wTot99) + "  " _
         & " WHERE E_SOCIO = '" + wE_S + "' AND " _
         & "           USU = '" + wcodusu + "' ")
         Db.CommitTrans
         
         wRegAct = wRegAct + 1
         ADO8a.MoveNext
      Loop
   End If
   
   Dim wTotDie As Integer, wTotCMP As Integer, wTotCaj As Integer, wTotHon As Integer, wTotFor As Integer
   
   wTotDie = 0: wTotCMP = 0: wTotCaj = 0: wTotHon = 0: wTotFor = 0
   
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
   
   aa = Leerado8("SELECT COUNT(*) AS TOT FROM MAESOCIO " _
                & " WHERE (E_SOCIO='HON') ")
   If aa > 0 Then
      wTotHon = IIf(IsNull(ADO8!tot), 0, ADO8!tot)
   End If
   Set ADO8 = Nothing
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESGRADO " _
   & " (CODIGO, NOMBRE, NUM, USU) " _
   & " VALUES " _
   & " ('04', 'NO APORTAN HONORARIOS', " _
   & "  " + Str(wTotHon) + ", '" + wcodusu + "'  )  ")
   Db.CommitTrans
      
   wTotFor = wTotDie + wTotCMP + wTotCaj + wTotHon
   
   lblCan0.Caption = Format(wCaz00, "##,##0;;\ ")
   lblCan3.Caption = Format(wCaz03, "##,##0;;\ ")
   lblCan6.Caption = Format(wCaz06, "##,##0;;\ ")
   lblCan7.Caption = Format(wCaz37, "##,##0;;\ ")
   lblCan9.Caption = Format(wCaz99, "##,##0;;\ ")
   
   lblTot3.Caption = Format(wToz03, "##,###,##0.00;;\ ")
   lblTot6.Caption = Format(wToz06, "##,###,##0.00;;\ ")
   lblTot7.Caption = Format(wToz37, "##,###,##0.00;;\ ")
   lblTot9.Caption = Format(wToz99, "##,###,##0.00;;\ ")
   lblHabiles.Caption = Format(wTotFor, "###,##0;;\ ")

   lblMensaje.Caption = ""
   lblMensaje.Refresh

   aa = Leerado2("SELECT NOMBRE, CAN00, CAN03, TOT03, CAN06, TOT06, CAN12, TOT12, CAN24, TOT24, CAN36, TOT36, CAN37, TOT37, CAN99, TOT99, USU, ORDEN " _
                & " FROM TMP_RESUMEN3 " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY ORDEN ")
   Set DataGrid1.DataSource = ADO2

   aa = Leerado3("SELECT CODIGO, NOMBRE, NUM, POR, USU " _
                & " FROM TMP_RESGRADO " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY CODIGO ")
   Set DataGrid2.DataSource = ADO3

End Sub

Private Sub LlenaCab1()
   DataGrid1.Columns(0).Alignment = dbgLeft
   DataGrid1.Columns(0).Width = 2000
   DataGrid1.Columns(0).Caption = "DESCRIPCION"

   DataGrid1.Columns(1).Alignment = dbgRight
   DataGrid1.Columns(1).Width = 650
   DataGrid1.Columns(1).Caption = "CANT.0"
   DataGrid1.Columns(1).NumberFormat = "##,##0"

   DataGrid1.Columns(2).Alignment = dbgRight
   DataGrid1.Columns(2).Width = 650
   DataGrid1.Columns(2).Caption = "CANT.3"
   DataGrid1.Columns(2).NumberFormat = "##,##0"

   DataGrid1.Columns(3).Alignment = dbgRight
   DataGrid1.Columns(3).Width = 1000
   DataGrid1.Columns(3).Caption = "TOTAL 3"
   DataGrid1.Columns(3).NumberFormat = "##,###,##0.00"

   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).Width = 650
   DataGrid1.Columns(4).Caption = "CANT.6"
   DataGrid1.Columns(4).NumberFormat = "##,##0"

   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).Width = 1000
   DataGrid1.Columns(5).Caption = "TOTAL 6"
   DataGrid1.Columns(5).NumberFormat = "##,###,##0.00"

   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).Width = 650
   DataGrid1.Columns(6).Caption = "CAN.12"
   DataGrid1.Columns(6).NumberFormat = "##,##0"

   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).Width = 1000
   DataGrid1.Columns(7).Caption = "TOT 12"
   DataGrid1.Columns(7).NumberFormat = "##,###,##0.00"

   DataGrid1.Columns(8).Alignment = dbgRight
   DataGrid1.Columns(8).Width = 650
   DataGrid1.Columns(8).Caption = "CAN.24"
   DataGrid1.Columns(8).NumberFormat = "##,##0"

   DataGrid1.Columns(9).Alignment = dbgRight
   DataGrid1.Columns(9).Width = 1000
   DataGrid1.Columns(9).Caption = "TOT 24"
   DataGrid1.Columns(9).NumberFormat = "##,###,##0.00"

   DataGrid1.Columns(10).Alignment = dbgRight
   DataGrid1.Columns(10).Width = 650
   DataGrid1.Columns(10).Caption = "CAN.36"
   DataGrid1.Columns(10).NumberFormat = "##,##0"

   DataGrid1.Columns(11).Alignment = dbgRight
   DataGrid1.Columns(11).Width = 1000
   DataGrid1.Columns(11).Caption = "TOT 36"
   DataGrid1.Columns(11).NumberFormat = "##,###,##0.00"

   DataGrid1.Columns(12).Alignment = dbgRight
   DataGrid1.Columns(12).Width = 650
   DataGrid1.Columns(12).Caption = "CANT.37"
   DataGrid1.Columns(12).NumberFormat = "##,##0"

   DataGrid1.Columns(13).Alignment = dbgRight
   DataGrid1.Columns(13).Width = 1000
   DataGrid1.Columns(13).Caption = "TOTAL 37"
   DataGrid1.Columns(13).NumberFormat = "##,###,##0.00"

   DataGrid1.Columns(14).Alignment = dbgRight
   DataGrid1.Columns(14).Width = 650
   DataGrid1.Columns(14).Caption = "TOT.CANT"
   DataGrid1.Columns(14).NumberFormat = "##,##0"

   DataGrid1.Columns(15).Alignment = dbgRight
   DataGrid1.Columns(15).Width = 1000
   DataGrid1.Columns(15).Caption = "TOTAL DEUDA"
   DataGrid1.Columns(15).NumberFormat = "##,###,##0.00"

   DataGrid1.Columns(16).Visible = False
   DataGrid1.Columns(17).Visible = False

   
   
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
   DataGrid2.Columns(4).Visible = False
End Sub

