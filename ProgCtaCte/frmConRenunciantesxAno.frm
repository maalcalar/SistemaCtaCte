VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmConRenunciantesxAno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relacion de Renuncias x Año"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   10530
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
      Left            =   8760
      TabIndex        =   7
      Top             =   7200
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
      Left            =   6120
      TabIndex        =   6
      Top             =   7200
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
      Left            =   7440
      TabIndex        =   5
      Top             =   7200
      Width           =   1095
   End
   Begin VB.TextBox txtAno 
      Height          =   305
      Left            =   1200
      TabIndex        =   1
      Top             =   360
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5775
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   10186
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
      Caption         =   "RELACION DE ASOCIADOS RENUNCIANTES POR AÑO"
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
      Left            =   9600
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
      Left            =   8880
      TabIndex        =   9
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Renunciantes"
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
      TabIndex        =   8
      Top             =   6720
      Width           =   1815
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
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   495
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
      TabIndex        =   2
      Top             =   7200
      Width           =   5895
   End
End
Attribute VB_Name = "frmConRenunciantesxAno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
   LlenaCab
   TotalCab
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(6) As String, wreg As Integer, wTot As Integer, wFec As Date
   Dim wNom As String
   Heading(0) = "CODSOCIO"
   Heading(1) = "CODIGO"
   Heading(2) = "INS"
   Heading(3) = "NOMBRE"
   Heading(4) = "ESTADO"
   Heading(5) = "FEC.RENUNC"
   aa = Leerado3("SELECT * FROM TMP_SOCIOS WHERE USU = '" + wcodusu + "' ORDER BY FECRENU, NOMBRE ")
   If aa > 0 Then
      wTot = aa
      Set objExcel = New Excel.Application
      objExcel.SheetsInNewWorkbook = 1
      objExcel.Workbooks.Add
      With objExcel
           .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
           .Range(.Cells(3, 1), .Cells(3, 6)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 6)).Font.Bold = True
           .Cells(1, 1) = wnomcia
           .Cells(2, 1) = "RELACION DE ASOCIADOS RENUNCIANTES - EJERCICIO " + txtAno.Text
           For I = 1 To 6 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           objExcel.Columns("A").ColumnWidth = 11
           objExcel.Columns("B").ColumnWidth = 11
           objExcel.Columns("C").ColumnWidth = 5
           objExcel.Columns("D").ColumnWidth = 60
           objExcel.Columns("E").ColumnWidth = 8
           objExcel.Columns("F").ColumnWidth = 11
      End With
      V = 4
      H = 1
      wreg = 1
      Do While Not ADO3.EOF
         DoEvents
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wreg, "####0") + " / " + Format(wTot, "####0")
         lblMensaje.Refresh
         
         objExcel.Cells(V, H + 0) = ADO3!codsocio
         objExcel.Cells(V, H + 1) = ADO3!codigo
         objExcel.Cells(V, H + 2) = ADO3!ins
         objExcel.Cells(V, H + 3) = ADO3!nombre
         objExcel.Cells(V, H + 4) = ADO3!e_socio
         If IsDate(ADO3!fecrenu) Then
            wFec = Format(ADO3!fecrenu, "dd/mm/yyyy")
            objExcel.Cells(V, H + 5) = wFec
         End If
         wreg = wreg + 1
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
   Dim wAno As String
   wAno = txtAno.Text
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\RepRenunciaxAno.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= 'EJERCICIO " + wAno + "' "
   Crys1.SelectionFormula = " {TMP_SOCIOS.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   txtAno.Text = Format(Val(wanocia) - 1, "0000")
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_SOCIOS WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   txtAno.SetFocus
End Sub

Private Sub Form_Initialize()
   frmConRenunciantesxAno.Left = (Screen.Width - Width) \ 2
   frmConRenunciantesxAno.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_SOCIOS WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim aa As Integer, wAno As String
   wAno = txtAno.Text

   Set DataGrid1.DataSource = Nothing
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_SOCIOS WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_SOCIOS " _
   & " (CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, FECRENU, MONEDA, APORTE, " _
   & "  USU ) " _
   & " SELECT " _
   & "  S.CODSOCIO, S.CODIGO, S.INS, S.NOMBRE, S.E_SOCIO, S.FECRENU, " _
   & "  E.MONEDA, E.APORTE, '" + wcodusu + "' " _
   & " FROM MAESOCIO AS S INNER JOIN MAEE_SOCIO AS E " _
   & "   ON S.E_SOCIO = E.E_SOCIO " _
   & " WHERE YEAR(S.FECRENU) =  " + Str(Val(wAno)) + "  ")
   Db.CommitTrans

   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, FECRENU, " _
                & "      MONEDA, APORTE " _
                & " FROM TMP_SOCIOS " _
                & " WHERE usu = '" + wcodusu + "' " _
                & " ORDER BY FECRENU, NOMBRE ")
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
    
   DataGrid1.Columns(3).Width = 4650  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE"
    
   DataGrid1.Columns(4).Width = 700    ' E_SOCIO
   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Caption = "ESTADO"
    
   DataGrid1.Columns(5).Width = 1100   ' FECING
   DataGrid1.Columns(5).Alignment = dbgLeft
   DataGrid1.Columns(5).Caption = "FEC.RENUNC"
   DataGrid1.Columns(5).NumberFormat = "dd/mm/yyyy"
    
   DataGrid1.Columns(6).Visible = False
   DataGrid1.Columns(7).Visible = False
End Sub

Private Sub TotalCab()
   Dim zz As Integer, wTot As Integer

   zz = Leerado8("SELECT COUNT(*) AS NUM " _
                & " FROM TMP_SOCIOS " _
                & " WHERE USU = '" + wcodusu + "' ")
   If zz > 0 Then
      wTot = ADO8!num
   End If

   lblTotal.Caption = Format(wTot, "###,##0")
End Sub

Private Sub txtAno_GotFocus()
   txtAno.SelStart = 0
   txtAno.SelLength = 4
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txtAno.Text)) = 0 Then
         MsgBox "Año En Blanco", vbExclamation
         txtAno.Text = Format(Val(wanocia) - 1, "0000")
         Exit Sub
      End If
      If txtAno.Text < "2015" And txtAno.Text > "2030" Then
         MsgBox "Año Digitado esta fuera de Rango", vbExclamation
         txtAno.Text = Format(Val(wanocia) - 1, "0000")
         Exit Sub
      End If
      cmdBuscar.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub
