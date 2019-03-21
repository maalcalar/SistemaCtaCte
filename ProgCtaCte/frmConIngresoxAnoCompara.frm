VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmConIngresoxAnoCompara 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comparativo de Ingresantes x Año"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   10635
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
      Left            =   7800
      TabIndex        =   8
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
      Left            =   6480
      TabIndex        =   7
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
      Left            =   9120
      TabIndex        =   6
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox txtHasta 
      Height          =   305
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   615
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
      Left            =   2400
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtDesde 
      Height          =   305
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
   Begin Crystal.CrystalReport Crys1 
      Left            =   4320
      Top             =   360
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
      Height          =   6015
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10610
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
      Caption         =   "RELACION DE ASOCIADOS INGRESANTES POR AÑO"
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
   Begin VB.Label lblAno6b 
      Alignment       =   2  'Center
      Caption         =   "Ano6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9120
      TabIndex        =   21
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lblAno5b 
      Alignment       =   2  'Center
      Caption         =   "Ano5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8280
      TabIndex        =   20
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lblAno4b 
      Alignment       =   2  'Center
      Caption         =   "Ano4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7440
      TabIndex        =   19
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lblAno3b 
      Alignment       =   2  'Center
      Caption         =   "Ano3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6600
      TabIndex        =   18
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lblAno2b 
      Alignment       =   2  'Center
      Caption         =   "Ano2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5760
      TabIndex        =   17
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lblAno1b 
      Alignment       =   2  'Center
      Caption         =   "Ano1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lblAno6 
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
      Left            =   9120
      TabIndex        =   15
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lblAno5 
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
      Left            =   8280
      TabIndex        =   14
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lblAno4 
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
      Left            =   7440
      TabIndex        =   13
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lblAno3 
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
      Left            =   6600
      TabIndex        =   12
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lblAno2 
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
      Left            =   5760
      TabIndex        =   11
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lblAno1 
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
      Left            =   4920
      TabIndex        =   10
      Top             =   7080
      Width           =   855
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
      Left            =   0
      TabIndex        =   9
      Top             =   7800
      Width           =   3855
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
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label4 
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
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmConIngresoxAnoCompara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBuscar_Click()
   LlenaCab
   TotalCab
   DataGrid1.SetFocus
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(7) As String, wreg As Integer, wTot As Integer, wFec As Date
   Dim wNom As String, _
       wAno1 As String, wAno2 As String, wAno3 As String, _
       wAno4 As String, wAno5 As String, wAno6 As String, _
       wxx1 As Integer, wxx2 As Integer, wxx3 As Integer, _
       wxx4 As Integer, wxx5 As Integer, wxx6 As Integer
   
   wAno1 = txtDesde.Text
   wAno2 = Format(Val(wAno1) + 1, "0000")
   wAno3 = Format(Val(wAno1) + 2, "0000")
   wAno4 = Format(Val(wAno1) + 3, "0000")
   wAno5 = Format(Val(wAno1) + 4, "0000")
   wAno6 = Format(Val(wAno1) + 5, "0000")
   
   Heading(0) = "ESTADO"
   Heading(1) = "NOMBRE"
   Heading(2) = wAno1
   Heading(3) = wAno2
   Heading(4) = wAno3
   Heading(5) = wAno4
   Heading(6) = wAno5
   Heading(7) = wAno6
   aa = Leerado3("SELECT * FROM TMP_COMPARAING WHERE USU = '" + wcodusu + "' ORDER BY E_SOCIO ")
   If aa > 0 Then
      wTot = aa
      Set objExcel = New Excel.Application
      objExcel.SheetsInNewWorkbook = 1
      objExcel.Workbooks.Add
      With objExcel
           .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
           .Range(.Cells(3, 1), .Cells(3, 8)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 8)).Font.Bold = True
           .Cells(1, 1) = wnomcia
           .Cells(2, 1) = "COMPARATIVO DE INGRESANTES - DEL " + wAno1 + " AL " + wAno6
           For I = 1 To 8 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           objExcel.Columns("A").ColumnWidth = 11
           objExcel.Columns("B").ColumnWidth = 40
           objExcel.Columns("C").ColumnWidth = 10
           objExcel.Columns("D").ColumnWidth = 10
           objExcel.Columns("E").ColumnWidth = 10
           objExcel.Columns("F").ColumnWidth = 10
           objExcel.Columns("G").ColumnWidth = 10
           objExcel.Columns("H").ColumnWidth = 10
      End With
      V = 4
      H = 1
      wreg = 1
      wxx1 = 0: wxx2 = 0: wxx3 = 0: wxx4 = 0: wxx5 = 0: wxx6 = 0
      Do While Not ADO3.EOF
         DoEvents
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wreg, "####0") + " / " + Format(wTot, "####0")
         lblMensaje.Refresh
         objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 7)).NumberFormat = "####,##0;;\ "
         
         objExcel.Cells(V, H + 0) = ADO3!e_socio
         objExcel.Cells(V, H + 1) = ADO3!nombre
         objExcel.Cells(V, H + 2) = ADO3!ano1
         objExcel.Cells(V, H + 3) = ADO3!ano2
         objExcel.Cells(V, H + 4) = ADO3!ano3
         objExcel.Cells(V, H + 5) = ADO3!ano4
         objExcel.Cells(V, H + 6) = ADO3!ano5
         objExcel.Cells(V, H + 7) = ADO3!ano6
         
         wxx1 = wxx1 + ADO3!ano1
         wxx2 = wxx2 + ADO3!ano2
         wxx3 = wxx3 + ADO3!ano3
         wxx4 = wxx4 + ADO3!ano4
         wxx5 = wxx5 + ADO3!ano5
         wxx6 = wxx6 + ADO3!ano6
         wreg = wreg + 1
         V = V + 1
         ADO3.MoveNext
      Loop
      V = V + 1
      objExcel.Range(objExcel.Cells(V, H + 2), objExcel.Cells(V, H + 7)).NumberFormat = "####,##0;;\ "
      
      objExcel.Cells(V, H + 1) = "TOTALES"
      objExcel.Cells(V, H + 2) = wxx1
      objExcel.Cells(V, H + 3) = wxx2
      objExcel.Cells(V, H + 4) = wxx3
      objExcel.Cells(V, H + 5) = wxx4
      objExcel.Cells(V, H + 6) = wxx5
      objExcel.Cells(V, H + 7) = wxx6
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
   Dim wAno1 As String, wAno2 As String, wAno3 As String, _
       wAno4 As String, wAno5 As String, wAno6 As String, wGlo As String
   
   wAno1 = txtDesde.Text
   wAno2 = Format(Val(wAno1) + 1, "0000")
   wAno3 = Format(Val(wAno1) + 2, "0000")
   wAno4 = Format(Val(wAno1) + 3, "0000")
   wAno5 = Format(Val(wAno1) + 4, "0000")
   wAno6 = Format(Val(wAno1) + 5, "0000")
   
   wGlo = "DEL " + wAno1 + " AL " + wAno6
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\RepIngxAnoCompara.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= '" + wGlo + "' "
   Crys1.Formulas(3) = "ANO1='" + wAno1 + "' "
   Crys1.Formulas(4) = "ANO2='" + wAno2 + "' "
   Crys1.Formulas(5) = "ANO3='" + wAno3 + "' "
   Crys1.Formulas(6) = "ANO4='" + wAno4 + "' "
   Crys1.Formulas(7) = "ANO5='" + wAno5 + "' "
   Crys1.Formulas(8) = "ANO6='" + wAno6 + "' "
   Crys1.SelectionFormula = " {TMP_COMPARAING.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   txtDesde.Text = Format(Val(wanocia) - 6, "0000")
   txtHasta.Text = Format(Val(wanocia) - 1, "0000")
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_COMPARAING WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   txtDesde.SetFocus
End Sub

Private Sub Form_Initialize()
   frmConIngresoxAnoCompara.Left = (Screen.Width - Width) \ 2
   frmConIngresoxAnoCompara.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_COMPARAING WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim aa As Integer, wDesde As String, wHasta As String, waaa As String, II As Integer, _
       wAno1 As String, wAno2 As String, wAno3 As String, _
       wAno4 As String, wAno5 As String, wAno6 As String, wVez As Integer
   
   wDesde = txtDesde.Text
   wHasta = txtHasta.Text

   wAno1 = wDesde
   wAno2 = Format(Val(wDesde) + 1, "###0")
   wAno3 = Format(Val(wDesde) + 2, "###0")
   wAno4 = Format(Val(wDesde) + 3, "###0")
   wAno5 = Format(Val(wDesde) + 4, "###0")
   wAno6 = Format(Val(wDesde) + 5, "###0")

   Set DataGrid1.DataSource = Nothing
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_COMPARAING WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_COMPARAING " _
   & " (E_SOCIO, NOMBRE, ANO1, ANO2, ANO3, ANO4, ANO5, ANO6, USU) " _
   & " SELECT " _
   & "  E_SOCIO, NOMBRE, 0, 0, 0, 0, 0, 0, '" + wcodusu + "' " _
   & " FROM MAEE_SOCIO ")
   Db.CommitTrans

   wVez = 1
   For II = Val(wDesde) To Val(wHasta)
       waaa = Format(II, "0000")
   
       Db.BeginTrans
       Db.Execute ("UPDATE TMP_COMPARAING " _
       & " SET ANO" + Format(wVez, "0") + " = V.NUM " _
       & " FROM TMP_COMPARAING AS T INNER JOIN V_ING_X_ANO AS V " _
       & "   ON T.E_SOCIO = V.E_SOCIO " _
       & " WHERE V.ANO = " + Str(Val(waaa)) + " ")
       Db.CommitTrans

       wVez = wVez + 1
   Next
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_COMPARAING " _
   & " WHERE USU = '" + wcodusu + "' AND " _
   & "       ANO1 = 0 AND ANO2 = 0 AND ANO3 = 0 AND " _
   & "       ANO4 = 0 AND ANO5 = 0 AND ANO6 = 0 ")
   Db.CommitTrans

   aa = Leerado2("SELECT E_SOCIO, NOMBRE, ANO1, ANO2, ANO3, ANO4, ANO5, ANO6, USU " _
                & " FROM TMP_COMPARAING " _
                & " WHERE usu = '" + wcodusu + "' " _
                & " ORDER BY E_SOCIO ")
   Set DataGrid1.DataSource = ADO2

   DataGrid1.Columns(0).Width = 500   ' E_SOCIO
   DataGrid1.Columns(0).Alignment = dbgLeft
   DataGrid1.Columns(0).Caption = "ESTADO"
    
   DataGrid1.Columns(1).Width = 4000  ' NOMBRE
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Caption = "NOMBRE"
    
   DataGrid1.Columns(2).Width = 800   ' ANO1
   DataGrid1.Columns(2).Alignment = dbgRight
   DataGrid1.Columns(2).Caption = wAno1
   DataGrid1.Columns(2).NumberFormat = "###,##0;;\ "
    
   DataGrid1.Columns(3).Width = 800   ' ANO2
   DataGrid1.Columns(3).Alignment = dbgRight
   DataGrid1.Columns(3).Caption = wAno2
   DataGrid1.Columns(3).NumberFormat = "###,##0;;\ "
    
   DataGrid1.Columns(4).Width = 800   ' ANO3
   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).Caption = wAno3
   DataGrid1.Columns(4).NumberFormat = "###,##0;;\ "
    
   DataGrid1.Columns(5).Width = 800   ' ANO4
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).Caption = wAno4
   DataGrid1.Columns(5).NumberFormat = "###,##0;;\ "
    
   DataGrid1.Columns(6).Width = 800   ' ANO5
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).Caption = wAno5
   DataGrid1.Columns(6).NumberFormat = "###,##0;;\ "
    
   DataGrid1.Columns(7).Width = 800   ' ANO6
   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).Caption = wAno6
   DataGrid1.Columns(7).NumberFormat = "###,##0;;\ "
    
   DataGrid1.Columns(8).Visible = False
End Sub

Private Sub TotalCab()
   Dim wAno1 As String, wAno2 As String, wAno3 As String, _
       wAno4 As String, wAno5 As String, wAno6 As String
   
   
   wAno1 = txtDesde.Text
   wAno2 = Format(Val(wAno1) + 1, "0000")
   wAno3 = Format(Val(wAno1) + 2, "0000")
   wAno4 = Format(Val(wAno1) + 3, "0000")
   wAno5 = Format(Val(wAno1) + 4, "0000")
   wAno6 = Format(Val(wAno1) + 5, "0000")

   zz = Leerado8("SELECT sum(ano1) AS NUM " _
                & " FROM TMP_COMPARAING " _
                & " WHERE USU = '" + wcodusu + "' ")
   If zz > 0 Then
      wTot = IIf(IsNull(ADO8!num), 0, ADO8!num)
   End If
   Set ADO8 = Nothing
   lblAno1.Caption = Format(wTot, "###,##0")
   lblAno1b.Caption = wAno1

   zz = Leerado8("SELECT sum(ano2) AS NUM " _
                & " FROM TMP_COMPARAING " _
                & " WHERE USU = '" + wcodusu + "' ")
   If zz > 0 Then
      wTot = IIf(IsNull(ADO8!num), 0, ADO8!num)
   End If
   Set ADO8 = Nothing
   lblAno2.Caption = Format(wTot, "###,##0")
   lblAno2b.Caption = wAno2

   zz = Leerado8("SELECT sum(ano3) AS NUM " _
                & " FROM TMP_COMPARAING " _
                & " WHERE USU = '" + wcodusu + "' ")
   If zz > 0 Then
      wTot = IIf(IsNull(ADO8!num), 0, ADO8!num)
   End If
   Set ADO8 = Nothing
   lblAno3.Caption = Format(wTot, "###,##0")
   lblAno3b.Caption = wAno3

   zz = Leerado8("SELECT sum(ano4) AS NUM " _
                & " FROM TMP_COMPARAING " _
                & " WHERE USU = '" + wcodusu + "' ")
   If zz > 0 Then
      wTot = IIf(IsNull(ADO8!num), 0, ADO8!num)
   End If
   Set ADO8 = Nothing
   lblAno4.Caption = Format(wTot, "###,##0")
   lblAno4b.Caption = wAno4

   zz = Leerado8("SELECT sum(ano5) AS NUM " _
                & " FROM TMP_COMPARAING " _
                & " WHERE USU = '" + wcodusu + "' ")
   If zz > 0 Then
      wTot = IIf(IsNull(ADO8!num), 0, ADO8!num)
   End If
   Set ADO8 = Nothing
   lblAno5.Caption = Format(wTot, "###,##0")
   lblAno5b.Caption = wAno5

   zz = Leerado8("SELECT sum(ano6) AS NUM " _
                & " FROM TMP_COMPARAING " _
                & " WHERE USU = '" + wcodusu + "' ")
   If zz > 0 Then
      wTot = IIf(IsNull(ADO8!num), 0, ADO8!num)
   End If
   Set ADO8 = Nothing
   lblAno6.Caption = Format(wTot, "###,##0")
   lblAno6b.Caption = wAno6
   
End Sub

Private Sub txtDesde_GotFocus()
   txtDesde.SelStart = 0
   txtDesde.SelLength = 4
End Sub

Private Sub txtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 40
        txtHasta.SetFocus
   End Select
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txtDesde.Text)) = 0 Then
         MsgBox "Año Inicial En Blanco", vbExclamation
         txtDesde.Text = Format(Val(wanocia) - 6, "0000")
      End If
      txtHasta.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtHasta_GotFocus()
   txtHasta.SelStart = 0
   txtHasta.SelLength = 4
End Sub

Private Sub txtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDesde.SetFocus
   End Select
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txtHasta.Text)) = 0 Then
         MsgBox "Año Final En Blanco", vbExclamation
         txtHasta.Text = Format(Val(wanocia) - 1, "0000")
      End If
      cmdBuscar.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub
