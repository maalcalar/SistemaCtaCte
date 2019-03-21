VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmServEstadoCta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estado de Cuenta desde OCT 2017"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   12750
   Begin VB.ComboBox cmbE_Socio 
      Height          =   315
      ItemData        =   "frmServEstadoCta.frx":0000
      Left            =   9960
      List            =   "frmServEstadoCta.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   900
      Width           =   2295
   End
   Begin VB.CommandButton cndOtro 
      Caption         =   "&Otra Consulta"
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
      Left            =   5640
      TabIndex        =   35
      Top             =   7440
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
      Left            =   10920
      TabIndex        =   34
      Top             =   7440
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
      Left            =   9600
      TabIndex        =   33
      Top             =   7440
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
      Left            =   8280
      TabIndex        =   32
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4815
      Left            =   360
      TabIndex        =   28
      Top             =   1920
      Width           =   11775
      _ExtentX        =   20770
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
   Begin VB.ComboBox cmbTipCob 
      Height          =   315
      ItemData        =   "frmServEstadoCta.frx":0004
      Left            =   7800
      List            =   "frmServEstadoCta.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   1380
      Width           =   2175
   End
   Begin VB.TextBox txtCarnetPIP 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   13
      Top             =   1380
      Width           =   930
   End
   Begin VB.TextBox txtCarnetPNP 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   360
      MaxLength       =   8
      TabIndex        =   12
      Top             =   1380
      Width           =   930
   End
   Begin VB.TextBox txtNumdoc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8880
      MaxLength       =   8
      TabIndex        =   9
      Top             =   880
      Width           =   975
   End
   Begin VB.ComboBox cmbGrado 
      Height          =   315
      ItemData        =   "frmServEstadoCta.frx":0008
      Left            =   10080
      List            =   "frmServEstadoCta.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1365
      Width           =   2175
   End
   Begin VB.TextBox txtCodSocio 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      MaxLength       =   9
      TabIndex        =   4
      Top             =   880
      Width           =   975
   End
   Begin VB.TextBox txtIns 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   1
      Top             =   880
      Width           =   375
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   360
      MaxLength       =   8
      TabIndex        =   0
      Top             =   880
      Width           =   975
   End
   Begin MSMask.MaskEdBox txtFecIng 
      Height          =   285
      Left            =   2280
      TabIndex        =   16
      Top             =   1380
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecRenu 
      Height          =   285
      Left            =   3360
      TabIndex        =   17
      Top             =   1380
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecExclu 
      Height          =   285
      Left            =   4440
      TabIndex        =   18
      Top             =   1380
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecExpul 
      Height          =   285
      Left            =   5520
      TabIndex        =   19
      Top             =   1380
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFecRein 
      Height          =   285
      Left            =   6600
      TabIndex        =   25
      Top             =   1380
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtTope 
      Height          =   285
      Left            =   480
      TabIndex        =   29
      Top             =   300
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Estado de Socio"
      Height          =   195
      Index           =   16
      Left            =   10230
      TabIndex        =   44
      Top             =   720
      Width           =   1170
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Provisiones"
      Height          =   255
      Left            =   8760
      TabIndex        =   42
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label9 
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8760
      TabIndex        =   41
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Cobros"
      Height          =   255
      Left            =   9840
      TabIndex        =   40
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9840
      TabIndex        =   39
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Saldos"
      Height          =   255
      Left            =   10920
      TabIndex        =   38
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label lblSaldos 
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   10920
      TabIndex        =   37
      Top             =   6720
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
      Left            =   840
      TabIndex        =   36
      Top             =   6840
      Width           =   5655
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Mes Tope"
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
      TabIndex        =   31
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label lblTope 
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
      Height          =   285
      Left            =   1320
      TabIndex        =   30
      Top             =   300
      Width           =   1935
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tipo Cobro"
      Height          =   195
      Index           =   18
      Left            =   8460
      TabIndex        =   27
      Top             =   1200
      Width           =   780
   End
   Begin VB.Label Label11 
      Caption         =   "Fec.Reingreso"
      Height          =   210
      Left            =   6600
      TabIndex        =   24
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Fec.Expulsión"
      Height          =   210
      Left            =   5520
      TabIndex        =   23
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "Fec.Exclusión"
      Height          =   210
      Left            =   4440
      TabIndex        =   22
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "Fec.Renuncia"
      Height          =   210
      Left            =   3360
      TabIndex        =   21
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "Fecha Ing."
      Height          =   210
      Left            =   2160
      TabIndex        =   20
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Carnet PIP"
      Height          =   195
      Index           =   7
      Left            =   1395
      TabIndex        =   15
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Carnet PNP"
      Height          =   195
      Index           =   6
      Left            =   360
      TabIndex        =   14
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "D.N.I."
      Height          =   195
      Left            =   8880
      TabIndex        =   11
      Top             =   705
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Grado"
      Height          =   195
      Index           =   2
      Left            =   10200
      TabIndex        =   10
      Top             =   1185
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Nombre de Socio"
      Height          =   195
      Left            =   2760
      TabIndex        =   7
      Top             =   705
      Width           =   3375
   End
   Begin VB.Label lblCodSocio 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   885
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "Cod.Socio"
      Height          =   195
      Left            =   1680
      TabIndex        =   5
      Top             =   700
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Ins"
      Height          =   195
      Left            =   1320
      TabIndex        =   3
      Top             =   700
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Codofin"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   700
      Width           =   975
   End
End
Attribute VB_Name = "frmServEstadoCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Limpiar()
   txtCodSocio.Text = ""
   lblCodSocio.Caption = ""
   txtCodigo.Text = ""
   txtIns.Text = ""
   txtNumdoc.Text = ""
   txtCarnetPNP.Text = ""
   txtCarnetPIP.Text = ""
   txtFecIng.Text = "__/__/____"
   txtFecRenu.Text = "__/__/____"
   txtFecExpul.Text = "__/__/____"
   txtFecExclu.Text = "__/__/____"
   txtFecRein.Text = "__/__/____"

   cmbGrado.ListIndex = 0
   cmbE_Socio.ListIndex = 0
   cmbTipCob.ListIndex = 0
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   frmServEstadoCta.Left = (Screen.Width - Width) \ 2
   frmServEstadoCta.Top = 0
   
   txtTope.Text = Left(zMesTope, 4) + "/" + Right(zMesTope, 2)
   txtTope.Enabled = False
   
   a = Leerado8("SELECT * FROM MAEGRADO ORDER BY GRADO ")
   If a > 0 Then
      ADO8.MoveFirst
      I = 0
      Do While Not ADO8.EOF
         cmbGrado.AddItem ADO8!nombre
         I = I + 1
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   
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
   
   a = Leerado8("SELECT * FROM MAETIPCOB ORDER BY TIPCOB ")
   If a > 0 Then
      ADO8.MoveFirst
      I = 0
      Do While Not ADO8.EOF
         cmbTipCob.AddItem ADO8!nombre
         I = I + 1
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   
   Call Limpiar
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_ESTCTA WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   txtCodigo.SetFocus
End Sub

Private Sub llenadet()
   Dim aa As Integer, wCod As Long, wDni As String
   wCod = Val(txtCodigo.Text)
   wDni = txtNumdoc.Text

   aa = Leerado7a("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(wCod) + " ")
   If aa > 0 Then
      txtCodSocio.Text = ADO7a!codsocio
      txtCodigo.Text = ADO7a!codigo
      txtIns.Text = ADO7a!ins
      txtNumdoc.Text = ADO7a!numdoc
      txtCarnetPNP.Text = ADO7a!carnetpnp
      txtCarnetPIP.Text = ADO7a!carnetpip
      
      If IsDate(ADO7a!fecing) Then
         txtFecIng.Text = Format(ADO7a!fecing, "dd/mm/yyyy")
      Else
         txtFecIng.Text = "__/__/____"
      End If
      If IsDate(ADO7a!fecrenu) Then
         txtFecRenu.Text = Format(ADO7a!fecrenu, "dd/mm/yyyy")
      Else
         txtFecRenu.Text = "__/__/____"
      End If
      If IsDate(ADO7a!fecexpul) Then
         txtFecExpul.Text = Format(ADO7a!fecexpul, "dd/mm/yyyy")
      Else
         txtFecExpul.Text = "__/__/____"
      End If
      If IsDate(ADO7a!fecexclu) Then
         txtFecExclu.Text = Format(ADO7a!fecexclu, "dd/mm/yyyy")
      Else
         txtFecExclu.Text = "__/__/____"
      End If
      If IsDate(ADO7a!fecrein) Then
         txtFecRein.Text = Format(ADO7a!fecrein, "dd/mm/yyyy")
      Else
         txtFecRein.Text = "__/__/____"
      End If
      cmbGrado.ListIndex = BuscaGrado(ADO7a!grado)
      cmbE_Socio.ListIndex = BuscaEsocio(ADO7a!e_socio)
      cmbTipCob.ListIndex = BuscaTipCob(ADO7a!tipcob)
   End If

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_ESTCTA WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_ESTCTA " _
   & " (codsocio, codigo, ins, nombre, mes, fecha, tipcob, sercob, numcob, " _
   & "  sdoold, cargos, abonos, sdonew, usu) " _
   & " select " _
   & "   " _
   & " from ctasxdet as d inner join maesocio as m " _
   & "   on d.codsocio = m.codsocio ")
   Db.CommitTrans

End Sub

Private Sub txtCodigo_Change()
   Dim aa As Integer
   If Val(txtCodigo.Text) <> 0 Then
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(Val(txtCodigo.Text)) + " ")
      If aa > 0 Then
         lblCodSocio.Caption = ADO8!nombre
      Else
         lblCodSocio.Caption = ""
      End If
      Set ADO8 = Nothing
   Else
      lblCodSocio.Caption = ""
   End If
End Sub

Private Sub txtCodigo_GotFocus()
   txtCodigo.SelStart = 0
   txtCodigo.SelLength = 8
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim KeyNew As Integer
   KeyNew = Asc(UCase(Chr(KeyCode)))
   
   Select Case KeyCode
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio2.Show 1
        If xseleccion <> "" Then
           txtCodigo.Text = xseleccion
        End If
   Case 65 To 90
        xlista = "1"
        xseleccion = Chr(KeyNew)
        frmSelecSocio2.Show 1
        If xseleccion <> "" Then
           txtCodigo.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtCodigo.Text)) = 0 Then
         MsgBox "Codofin En Blanco", vbExclamation
         txtCodigo.Text = ""
         Exit Sub
      End If
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(Val(txtCodigo.Text)) + " ")
      If aa = 0 Then
         MsgBox "Codofin Digitado NO Existe", vbExclamation
         txtCodigo.Text = ""
         Exit Sub
      End If
      lblCodSocio.Caption = ADO8!nombre
      
      llenadet
   
      cmdImprimir.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtTope_Change()
   Dim wMes As String, wAno As String
   If txtTope.Text <> "____-__" Then
      wAno = Left(txtTope.Text, 4)
      wMes = Right(txtTope.Text, 2)
               
      If wMes <> "01" And wMes <> "02" And wMes <> "03" And wMes <> "04" And wMes <> "05" And wMes <> "06" And _
         wMes <> "07" And wMes <> "08" And wMes <> "09" And wMes <> "10" And wMes <> "11" And wMes <> "12" Then
         lblTope.Caption = ""
      Else
         lblTope.Caption = Trim(funnommes(wMes)) + " " + wAno
      End If
   Else
      lblTope.Caption = ""
   End If
End Sub
