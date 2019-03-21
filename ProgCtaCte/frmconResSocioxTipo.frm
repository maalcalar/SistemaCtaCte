VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmconResSocioxTipo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Socios x Tipo"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   9570
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
      Left            =   1800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
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
      Left            =   6000
      TabIndex        =   2
      Top             =   6480
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
      Left            =   4680
      TabIndex        =   1
      Top             =   6480
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
      Left            =   7320
      TabIndex        =   0
      Top             =   6480
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5055
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8916
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
      Left            =   8760
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Socios"
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
      Left            =   3000
      TabIndex        =   6
      Top             =   6000
      Width           =   1815
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
      Left            =   4920
      TabIndex        =   5
      Top             =   6000
      Width           =   1095
   End
End
Attribute VB_Name = "frmconResSocioxTipo"
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
   
   Dim aa As Long, I As Integer, Heading(5) As String, wRegAct As Long, wRegTot As Long
   Dim wtot As Currency
   Heading(0) = "INS"
   Heading(1) = "E_SOCIO"
   Heading(2) = "NOMBRE"
   Heading(3) = "CANTIDAD"
   Heading(4) = "INSCRIP"
   Heading(5) = "APORTE"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 6)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 6)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "RESUMEN DE SOCIOS POR TIPO"
        For I = 1 To 6 Step 1
            .Cells(3, I) = Heading(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 8
        objExcel.Columns("B").ColumnWidth = 8
        objExcel.Columns("C").ColumnWidth = 30
        objExcel.Columns("D").ColumnWidth = 11
        objExcel.Columns("E").ColumnWidth = 12
        objExcel.Columns("F").ColumnWidth = 12
   End With
   aa = Leerado3("SELECT * FROM TMP_RESXTIPO WHERE USU = '" + wcodusu + "' ORDER BY LIN ")
   If aa > 0 Then
      V = 4
      H = 1
      wtot = 0
      wRegAct = 1
      wRegTot = aa
      Do While Not ADO3.EOF
         
         objExcel.Range(objExcel.Cells(V, H + 3), objExcel.Cells(V, H + 3)).NumberFormat = "####,##0;;\ "
         objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 5)).NumberFormat = "####,##0.00;;\ "
         
         objExcel.Cells(V, H + 0) = ADO3!ins
         objExcel.Cells(V, H + 1) = ADO3!e_socio
         objExcel.Cells(V, H + 2) = ADO3!nombre
         objExcel.Cells(V, H + 3) = ADO3!cant
         objExcel.Cells(V, H + 4) = ADO3!inscrip
         objExcel.Cells(V, H + 5) = ADO3!aporte
         
         If Len(Trim(ADO3!e_socio)) = 0 Then
            wtot = wtot + ADO3!cant
         End If
         
         V = V + 1
         ADO3.MoveNext
      Loop
      V = V + 1
      objExcel.Range(objExcel.Cells(V, H + 3), objExcel.Cells(V, H + 3)).NumberFormat = "####,##0;;\ "
      
      objExcel.Cells(V, H + 2) = "TOTALES "
      objExcel.Cells(V, H + 3) = wtot
      V = V + 1
      
      Set ADO3 = Nothing
      objExcel.Visible = True
      Set objExcel = Nothing
   End If
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdImprimir_Click()
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\ResxTipo.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.SelectionFormula = " {TMP_RESXTIPO.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESXTIPO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   cmdBuscar.SetFocus
End Sub

Private Sub Form_Initialize()
   frmconResSocioxTipo.Left = (Screen.Width - Width) \ 2
   frmconResSocioxTipo.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESXTIPO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim aa As Integer

   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_RESXTIPO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESXTIPO " _
   & " (LIN, INS, NOMBRE, E_SOCIO, CANT, INSCRIP, APORTE, USU ) " _
   & " select " _
   & "  CAST(M.INS AS VARCHAR), m.ins, i.nombre, '', COUNT(*) as cant, 500, 50, '" + wcodusu + "' " _
   & " from MAESOCIO as m inner join maeins     as i on m.INS = i.ins " _
   & "                    inner join maee_socio as e on m.e_socio = e.e_socio " _
   & " where e.aporte > 0 " _
   & " group by m.INS, i.nombre " _
   & " order by m.ins ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_RESXTIPO " _
   & " SET INSCRIP = 0 " _
   & " WHERE USU = '" + wcodusu + "' AND " _
   & "       INS = 2 ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_RESXTIPO " _
   & " SET INSCRIP = 0, APORTE = 0 " _
   & " WHERE USU = '" + wcodusu + "' AND " _
   & "       INS = 8 ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_RESXTIPO " _
   & " (LIN, INS, E_SOCIO, NOMBRE, CANT, INSCRIP, APORTE, USU ) " _
   & " select " _
   & "  CAST(m.INS as varchar) + CAST(E.ORDEN AS VARCHAR), " _
   & "  m.ins, M.E_SOCIO, E.NOMBRE , COUNT(*), 0, 0, '" + wcodusu + "' " _
   & " from MAESOCIO as m inner join maee_socio as e on m.E_SOCIO = e.e_socio " _
   & " where m.INS = '7' and e.APORTE > 0 " _
   & " group by m.ins, m.e_socio, e.orden, E.NOMBRE ")
   Db.CommitTrans
   
   aa = Leerado2("SELECT LIN, INS, E_SOCIO, NOMBRE, CANT, INSCRIP, APORTE, USU " _
                & " FROM TMP_RESXTIPO " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " order by lin ")
   Set DataGrid1.DataSource = ADO2
   
   DataGrid1.Columns(0).Width = 450   ' LIN
   DataGrid1.Columns(0).Alignment = dbgLeft
   DataGrid1.Columns(0).Caption = "LIN"
    
   DataGrid1.Columns(1).Width = 550   ' INS
   DataGrid1.Columns(1).Alignment = dbgCenter
   DataGrid1.Columns(1).Caption = "INS"
    
   DataGrid1.Columns(2).Width = 650   ' E_SOCIO
   DataGrid1.Columns(2).Alignment = dbgCenter
   DataGrid1.Columns(2).Caption = "E_SOC"
    
   DataGrid1.Columns(3).Width = 3200  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE"
    
   DataGrid1.Columns(4).Width = 800   ' CANT
   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).Caption = "CANT"
   DataGrid1.Columns(4).NumberFormat = "###,##0;;\ "
    
   DataGrid1.Columns(5).Width = 1250  ' INSCRIPCION
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).Caption = "INSCRIP"
   DataGrid1.Columns(5).NumberFormat = "###,##0.00;;\ "
    
   DataGrid1.Columns(6).Width = 1250  ' APORTE
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).Caption = "APORTE MES"
   DataGrid1.Columns(6).NumberFormat = "###,##0.00;;\ "
    
   DataGrid1.Columns(0).Visible = False
   DataGrid1.Columns(7).Visible = False
End Sub

Private Sub TotalCab()
   Dim zz As Integer, wtot As Currency

   wtot = 0
   zz = Leerado8("SELECT SUM(CANT) AS CANT " _
                & " FROM TMP_RESXTIPO " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "       E_SOCIO = '' ")
   If zz > 0 Then
      wtot = IIf(IsNull(ADO8!cant), 0, ADO8!cant)
   End If
   Set ADO8 = Nothing

   lblTotal.Caption = Format(wtot, "###,##0")
End Sub

