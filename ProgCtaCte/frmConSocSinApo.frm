VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmConSocSinApo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Socios Sin Aportes"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11205
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
      TabIndex        =   11
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
      Left            =   6600
      TabIndex        =   10
      Top             =   7080
      Width           =   1095
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
      Left            =   5040
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtHasta 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1305
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtDesde 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox cmbE_Socio 
      Height          =   315
      ItemData        =   "frmConSocSinApo.frx":0000
      Left            =   1320
      List            =   "frmConSocSinApo.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   3375
   End
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmConSocSinApo.frx":0004
      Left            =   1320
      List            =   "frmConSocSinApo.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   7335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4815
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   10815
      _ExtentX        =   19076
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
      TabIndex        =   13
      Top             =   6960
      Width           =   5295
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
      TabIndex        =   12
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Hasta Importe"
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
      Index           =   2
      Left            =   45
      TabIndex        =   6
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Desde Importe"
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
      Left            =   15
      TabIndex        =   4
      Top             =   960
      Width           =   1245
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
      TabIndex        =   3
      Top             =   600
      Width           =   1140
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
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmConSocSinApo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbE_Socio_Click()
   cmbE_Socio_KeyPress (13)
End Sub

Private Sub cmbE_Socio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtDesde.SetFocus
   End If
End Sub

Private Sub cmdBuscar_Click()
   LlenaCab
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(13) As String, wreg As Integer, wTot As Integer
   Dim wFecIng As Date, wDes As Currency, wHas As Currency

   wDes = Format(Val(txtDesde.Text), "#####0.00")
   wHas = Format(Val(txtHasta.Text), "#####0.00")
       
   Heading(0) = "NUM"
   Heading(1) = "SOCIO"
   Heading(2) = "CODOFIN"
   Heading(3) = "INS"
   Heading(4) = "APELLIDOS Y NOMBRES"
   Heading(5) = "ESTADO"
   Heading(6) = "DIRECCION"
   Heading(7) = "TELEFONOS"
   Heading(8) = "TELEFONOS2"
   Heading(9) = "CELULAR"
   Heading(10) = "CORREO ELECTRONICO"
   Heading(11) = "CORREO ELECTRONICO 2"
   Heading(12) = "FEC.ING"
   Heading(13) = "TOT.APORTES"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 14)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 14)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "RELACION DE SOCIOS CON APORTES DE " + Format(wDes, "#####0.00") + " A " + Format(wHas, "#####0.00")
        For I = 1 To 14 Step 1
            .Cells(3, I) = Heading(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 6
        objExcel.Columns("B").ColumnWidth = 10
        objExcel.Columns("C").ColumnWidth = 11
        objExcel.Columns("D").ColumnWidth = 5
        objExcel.Columns("E").ColumnWidth = 50
        objExcel.Columns("F").ColumnWidth = 10
        objExcel.Columns("G").ColumnWidth = 50
        objExcel.Columns("H").ColumnWidth = 18
        objExcel.Columns("I").ColumnWidth = 18
        objExcel.Columns("J").ColumnWidth = 12
        objExcel.Columns("K").ColumnWidth = 25
        objExcel.Columns("L").ColumnWidth = 25
        objExcel.Columns("M").ColumnWidth = 12
        objExcel.Columns("N").ColumnWidth = 11
   End With
   
   aa = Leerado3("SELECT * FROM TMP_SOCSINAPO WHERE USU = '" + wcodusu + "' ORDER BY NOMBRE ")
   If aa > 0 Then
      wreg = 1
      wTot = aa
      V = 4
      H = 1
      Do While Not ADO3.EOF
         DoEvents
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wreg, "####0") + " / " + Format(wTot, "####0")
         lblMensaje.Refresh
            
         objExcel.Range(objExcel.Cells(V, H + 13), objExcel.Cells(V, H + 13)).NumberFormat = "####,##0.00;;\ "
         
         wFecIng = ADO3!fecing
         
         objExcel.Cells(V, H + 0) = wreg
         objExcel.Cells(V, H + 1) = ADO3!codsocio
         objExcel.Cells(V, H + 2) = ADO3!codigo
         objExcel.Cells(V, H + 3) = ADO3!ins
         objExcel.Cells(V, H + 4) = ADO3!nombre
         objExcel.Cells(V, H + 5) = ADO3!e_socio
         objExcel.Cells(V, H + 6) = ADO3!direc
         objExcel.Cells(V, H + 7) = ADO3!telefono
         objExcel.Cells(V, H + 8) = ADO3!telefon2
         objExcel.Cells(V, H + 9) = Trim(ADO3!celular)
         objExcel.Cells(V, H + 10) = ADO3!email
         objExcel.Cells(V, H + 11) = ADO3!email2
         objExcel.Cells(V, H + 12) = wFecIng
         objExcel.Cells(V, H + 13) = ADO3!totapo
         
         wreg = wreg + 1
         V = V + 1
         
         ADO3.MoveNext
      Loop
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

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   frmConSocSinApo.Left = (Screen.Width - Width) \ 2
   frmConSocSinApo.Top = 0
   
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
   
   Set DataGrid1.DataSource = Nothing
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_SOCSINAPO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   cmbE_Socio.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_SOCSINAPO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim aa As Long, wreg As Long, wTot As Long, _
       w As String, wE_S As String, _
       wSoc As Integer, wApo As Currency, wDes As Currency, wHas As Currency

   wE_S = BuscaCodEsocio(cmbE_Socio.List(cmbE_Socio.ListIndex))
   wDes = Val(txtDesde.Text)
   wHas = Val(txtHasta.Text)

   w = ""
   If wE_S <> "" Then
      w = " WHERE S.E_SOCIO = '" + wE_S + "' "
   End If

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_SOCSINAPO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_SOCSINAPO " _
   & " (USU, CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, NOME_S, " _
   & "  DIREC, TELEFONO, TELEFON2, CELULAR, EMAIL, EMAIL2, FECING, TOTAPO ) " _
   & " SELECT '" + wcodusu + "', S.CODSOCIO, S.CODIGO, S.INS, S.NOMBRE, S.E_SOCIO, " _
   & "        E.NOMBRE, S.DIREC, S.TELEFONO, S.TELEFON2, " _
   & "        S.CELULAR, S.EMAIL, S.EMAIL2, S.FECING, 0 " _
   & " FROM MAESOCIO AS S INNER JOIN MAEE_SOCIO AS E ON S.E_SOCIO = E.E_SOCIO " _
   & " " + w + " AND E.APORTE > 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_SOCSINAPO " _
   & " SET TOTAPO = T.TOTAPO + V.MONTO " _
   & " FROM TMP_SOCSINAPO AS T INNER JOIN V_APO_X_TESO AS V " _
   & "   ON T.CODSOCIO = V.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_SOCSINAPO " _
   & " SET TOTAPO = T.TOTAPO + V.MONTO " _
   & " FROM TMP_SOCSINAPO AS T INNER JOIN V_APO_X_BCO AS V " _
   & "   ON T.CODSOCIO = V.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_SOCSINAPO " _
   & " SET TOTAPO = T.TOTAPO - V.MONTO " _
   & " FROM TMP_SOCSINAPO AS T INNER JOIN V_APO_X_DEV AS V " _
   & "   ON T.CODSOCIO = V.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_SOCSINAPO " _
   & " SET TOTAPO = T.TOTAPO + V.MONTO " _
   & " FROM TMP_SOCSINAPO AS T INNER JOIN V_APO_X_DIECO AS V " _
   & "   ON T.CODSOCIO = V.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_SOCSINAPO " _
   & " SET TOTAPO = T.TOTAPO + V.MONTO " _
   & " FROM TMP_SOCSINAPO AS T INNER JOIN V_APO_X_CAJMP AS V " _
   & "   ON T.CODSOCIO = V.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_SOCSINAPO WHERE USU = '" + wcodusu + "' AND TOTAPO > " + Str(wHas) + " ")
   Db.CommitTrans
   
   DoEvents
   lblMensaje.Caption = ""
   lblMensaje.Refresh
   
   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, FECING, TOTAPO " _
            & " FROM TMP_SOCSINAPO " _
            & " WHERE USU = '" + wcodusu + "' " _
            & " ORDER BY NOMBRE ")
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
    
   DataGrid1.Columns(3).Width = 5450  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE"
    
   DataGrid1.Columns(4).Width = 600  ' E_SOCIO
   DataGrid1.Columns(4).Alignment = dbgCenter
   DataGrid1.Columns(4).Caption = "ESTADO"
    
   DataGrid1.Columns(5).Width = 1000  ' FECING
   DataGrid1.Columns(5).Alignment = dbgCenter
   DataGrid1.Columns(5).Caption = "FEC.ING"
   DataGrid1.Columns(5).NumberFormat = "dd/mm/yyyy"

   DataGrid1.Columns(6).Width = 1000  ' TOTAPO
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).Caption = "TOT.APORTES"
   DataGrid1.Columns(6).NumberFormat = "#####0.00;;\ "

End Sub

Private Sub txtDesde_GotFocus()
   txtDesde.SelStart = 0
   txtDesde.SelLength = Len(Trim(txtDesde.Text))
End Sub

Private Sub txtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        cmbE_Socio.SetFocus
   Case 40
        txtHasta.SetFocus
   End Select
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtDesde.Text = Format(txtDesde.Text, "######0.00")
      txtHasta.SetFocus
   Else
      If InStr(1, "0123456789." + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtHasta_GotFocus()
   txtHasta.SelStart = 0
   txtHasta.SelLength = Len(Trim(txtHasta.Text))
End Sub

Private Sub txtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDesde.SetFocus
   End Select
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
   Dim wDes As Currency, wHas As Currency
   If KeyAscii = 13 Then
      txtHasta.Text = Format(txtHasta.Text, "######0.00")
      
      wDes = Val(txtDesde.Text)
      wHas = Val(txtHasta.Text)
      
      If wDes > wHas Then
         MsgBox "Rango de Importes Es Invalido", vbExclamation
         txtHasta.Text = ""
         Exit Sub
      End If
      cmdBuscar.SetFocus
   Else
      If InStr(1, "0123456789." + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub
