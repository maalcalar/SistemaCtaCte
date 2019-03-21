VERSION 5.00
Begin VB.Form frmServModCodofin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar CODOFIN de Asociados"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11505
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
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
      TabIndex        =   28
      Top             =   5760
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
      Left            =   9600
      TabIndex        =   27
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOtro 
      Caption         =   "&Otro Socio"
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
      Left            =   3840
      TabIndex        =   26
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "CODOFIN Nuevo"
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
      Height          =   1935
      Left            =   120
      TabIndex        =   21
      Top             =   2040
      Width           =   11295
      Begin VB.TextBox txtInsNew 
         Height          =   285
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   23
         Top             =   420
         Width           =   375
      End
      Begin VB.TextBox txtCodigoNew 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         MaxLength       =   8
         TabIndex        =   22
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Ins"
         Height          =   195
         Left            =   1200
         TabIndex        =   25
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Codofin"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Anteriores"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      Begin VB.ComboBox cmbTipCob 
         Height          =   315
         ItemData        =   "frmServModCodofin.frx":0000
         Left            =   8280
         List            =   "frmServModCodofin.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1020
         Width           =   2655
      End
      Begin VB.ComboBox cmbE_Socio 
         Height          =   315
         ItemData        =   "frmServModCodofin.frx":0004
         Left            =   5040
         List            =   "frmServModCodofin.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1020
         Width           =   3255
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
         Left            =   4080
         MaxLength       =   8
         TabIndex        =   13
         Top             =   1020
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
         Left            =   3120
         MaxLength       =   8
         TabIndex        =   12
         Top             =   1020
         Width           =   930
      End
      Begin VB.ComboBox cmbGrado 
         Height          =   315
         ItemData        =   "frmServModCodofin.frx":0008
         Left            =   240
         List            =   "frmServModCodofin.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1020
         Width           =   2895
      End
      Begin VB.TextBox txtNumdoc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9960
         MaxLength       =   8
         TabIndex        =   4
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox txtIns 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   3
         Top             =   420
         Width           =   375
      End
      Begin VB.TextBox txtCodSocio 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   9000
         MaxLength       =   9
         TabIndex        =   2
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         MaxLength       =   8
         TabIndex        =   1
         Top             =   420
         Width           =   975
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cobro"
         Height          =   195
         Index           =   18
         Left            =   8940
         TabIndex        =   20
         Top             =   840
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Estado de Socio"
         Height          =   195
         Index           =   16
         Left            =   5310
         TabIndex        =   18
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Carnet PIP"
         Height          =   195
         Index           =   7
         Left            =   4155
         TabIndex        =   16
         Top             =   840
         Width           =   765
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Carnet PNP"
         Height          =   195
         Index           =   6
         Left            =   3120
         TabIndex        =   15
         Top             =   840
         Width           =   840
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Grado"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "D.N.I."
         Height          =   195
         Left            =   9960
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Ins"
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Codofin"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre de Socio"
         Height          =   195
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblCodSocio 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   420
         Width           =   7335
      End
      Begin VB.Label Label3 
         Caption         =   "Cod.Socio"
         Height          =   195
         Left            =   9000
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmServModCodofin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Limpiar()
   txtCodigo.Text = ""
   txtIns.Text = ""
   txtCodSocio.Text = ""
   txtNumdoc.Text = ""
   txtCarnetPNP.Text = ""
   txtCarnetPIP.Text = ""

   cmbGrado.ListIndex = 0
   cmbE_Socio.ListIndex = 0
   cmbGrado.ListIndex = 0

   txtCodigoNew.Text = ""
   txtInsNew.Text = ""
End Sub

Private Sub cmdGrabar_Click()

   Dim wCodNew As Long, wInsNew As Integer, _
       wCodOld As Long, wINsOld As Integer

   wCodNew = Val(txtCodigoNew.Text)
   wInsNew = Val(txtInsNew.Text)


   wCodOld = Val(txtCodigo.Text)
   wINsOld = Val(txtIns.Text)

   cmdSalir.Enabled = False
   cmdGrabar.Enabled = False

   Db.BeginTrans
   Db.Execute ("UPDATE BCPCAB " _
   & " SET CODIGO = " + Str(wCodNew) + ", " _
   & "        INS = " + Str(wInsNew) + " " _
   & " WHERE CODIGO = " + Str(wCodOld) + " AND " _
   & "          INS = " + Str(wINsOld) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE CAJMPCAB " _
   & " SET CODIGO = " + Str(wCodNew) + ", " _
   & "        INS = " + Str(wInsNew) + " " _
   & " WHERE CODIGO = " + Str(wCodOld) + " AND " _
   & "          INS = " + Str(wINsOld) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE DIECOCAB " _
   & " SET CODIGO = " + Str(wCodNew) + ", " _
   & "        INS = " + Str(wInsNew) + " " _
   & " WHERE CODIGO = " + Str(wCodOld) + " AND " _
   & "          INS = " + Str(wINsOld) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE DEVOL " _
   & " SET CODIGO = " + Str(wCodNew) + ", " _
   & "        INS = " + Str(wInsNew) + " " _
   & " WHERE CODIGO = " + Str(wCodOld) + " AND " _
   & "          INS = " + Str(wINsOld) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE FRACCAB " _
   & " SET CODIGO = " + Str(wCodNew) + ", " _
   & "        INS = " + Str(wInsNew) + " " _
   & " WHERE CODIGO = " + Str(wCodOld) + " AND " _
   & "          INS = " + Str(wINsOld) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE MAESOCIO " _
   & " SET CODIGO = " + Str(wCodNew) + ", " _
   & "        INS = " + Str(wInsNew) + " " _
   & " WHERE CODIGO = " + Str(wCodOld) + " AND " _
   & "          INS = " + Str(wINsOld) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_APOR_PLA " _
   & " SET CODIGO = " + Str(wCodNew) + ", " _
   & "        INS = " + Str(wInsNew) + " " _
   & " WHERE CODIGO = " + Str(wCodOld) + " AND " _
   & "          INS = " + Str(wINsOld) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_BCORECAU " _
   & " SET CODIGO = " + Str(wCodNew) + ", " _
   & "        INS = " + Str(wInsNew) + " " _
   & " WHERE CODIGO = " + Str(wCodOld) + " AND " _
   & "          INS = " + Str(wINsOld) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_D_ASIGNA " _
   & " SET CODIGO = " + Str(wCodNew) + ", " _
   & "        INS = " + Str(wInsNew) + " " _
   & " WHERE CODIGO = " + Str(wCodOld) + " AND " _
   & "          INS = " + Str(wINsOld) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_DEVOL " _
   & " SET CODIGO = " + Str(wCodNew) + ", " _
   & "        INS = " + Str(wInsNew) + " " _
   & " WHERE CODIGO = " + Str(wCodOld) + " AND " _
   & "          INS = " + Str(wINsOld) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_FAMILIA " _
   & " SET CODOFIN = " + Str(wCodNew) + ", " _
   & "         INS = " + Str(wInsNew) + " " _
   & " WHERE CODOFIN = " + Str(wCodOld) + " AND " _
   & "           INS = " + Str(wINsOld) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_M_ASIGNA " _
   & " SET CODIGO = " + Str(wCodNew) + ", " _
   & "        INS = " + Str(wInsNew) + " " _
   & " WHERE CODIGO = " + Str(wCodOld) + " AND " _
   & "          INS = " + Str(wINsOld) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_M_ASIGNA " _
   & " SET COD_PADRE = " + Str(wCodNew) + ", " _
   & "     INS_PADRE = " + Str(wInsNew) + " " _
   & " WHERE COD_PADRE = " + Str(wCodOld) + " AND " _
   & "       INS_PADRE = " + Str(wINsOld) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_MAESTRO " _
   & " SET CODIGO = " + Str(wCodNew) + ", " _
   & "        INS = " + Str(wInsNew) + " " _
   & " WHERE CODIGO = " + Str(wCodOld) + " AND " _
   & "          INS = " + Str(wINsOld) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_MAESTRO_INICIAL " _
   & " SET CODIGO = " + Str(wCodNew) + ", " _
   & "        INS = " + Str(wInsNew) + " " _
   & " WHERE CODIGO = " + Str(wCodOld) + " AND " _
   & "          INS = " + Str(wINsOld) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_MRECIBOS " _
   & " SET CODIGO = " + Str(wCodNew) + ", " _
   & "        INS = " + Str(wInsNew) + " " _
   & " WHERE CODIGO = " + Str(wCodOld) + " AND " _
   & "          INS = " + Str(wINsOld) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_RENOVA " _
   & " SET CODIGO = " + Str(wCodNew) + ", " _
   & "        INS = " + Str(wInsNew) + " " _
   & " WHERE CODIGO = " + Str(wCodOld) + " AND " _
   & "          INS = " + Str(wINsOld) + " ")
   Db.CommitTrans

   cmdSalir.Enabled = True
   cmdGrabar.Enabled = True

   MsgBox "CODOFIN Cambiado OK", vbExclamation
End Sub

Private Sub cmdOtro_Click()
   Call Limpiar
   
   txtCodigo.SetFocus
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   frmServModCodofin.Left = (Screen.Width - Width) \ 2
   frmServModCodofin.Top = 0
   
   Dim a As Integer, I As Integer, wAno As String, wMes As String
   
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
   
   txtCodigo.SetFocus
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
      
      LlenaCab
   
      txtCodigoNew.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub LlenaCab()
   Dim aa As Integer, wCod As Long
   wCod = Val(txtCodigo.Text)

   aa = Leerado7a("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(wCod) + " ")
   If aa > 0 Then
      txtCodSocio.Text = ADO7a!codsocio
      txtCodigo.Text = ADO7a!codigo
      txtIns.Text = ADO7a!ins
      txtNumdoc.Text = ADO7a!numdoc
      txtCarnetPNP.Text = ADO7a!carnetpnp
      txtCarnetPIP.Text = ADO7a!carnetpip
      
      cmbGrado.ListIndex = BuscaGrado(ADO7a!grado)
      cmbE_Socio.ListIndex = BuscaEsocio(ADO7a!e_socio)
      cmbTipCob.ListIndex = BuscaTipCob(ADO7a!tipcob)
   End If
End Sub

Private Sub txtCodigoNew_GotFocus()
   txtCodigoNew.SelStart = 0
   txtCodigoNew.SelLength = 8
End Sub

Private Sub txtCodigoNew_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtCodigoNew.Text) = 0 Then
         MsgBox "Codigo Nuevo En Cero", vbExclamation
         txtCodigoNew.Text = ""
         Exit Sub
      End If
   
      txtInsNew.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtInsNew_GotFocus()
   txtInsNew.SelStart = 0
   txtInsNew.SelLength = 1
End Sub

Private Sub txtInsNew_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtCodigoNew.SetFocus
   End Select
End Sub

Private Sub txtInsNew_KeyPress(KeyAscii As Integer)
   Dim aa As Integer, _
       wCodNew As Long, wInsNew As Integer
   
   If KeyAscii = 13 Then
      If Val(txtInsNew.Text) = 0 Then
         MsgBox "Institucion En Cero", vbExclamation
         txtInsNew.Text = ""
         Exit Sub
      End If
      wCodNew = Val(txtCodigoNew.Text)
      wInsNew = Val(txtInsNew.Text)
      
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(wCodNew) + " AND INS = " + Str(wInsNew) + " ")
      If aa > 0 Then
         MsgBox "CODOFIN Nuevo Ya Existe", vbExclamation
         txtInsNew.Text = ""
         Exit Sub
      End If
      cmdGrabar.SetFocus
   Else
      If InStr(1, "123457" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If


End Sub
