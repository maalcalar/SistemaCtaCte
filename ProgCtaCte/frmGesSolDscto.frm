VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmGesSolDscto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Descuentos"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10170
   Begin VB.CommandButton cndOtro 
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
      Height          =   615
      Left            =   7080
      TabIndex        =   37
      Top             =   5520
      Width           =   1095
   End
   Begin VB.ComboBox cmbE_Socio 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmGesSolDscto.frx":0000
      Left            =   5520
      List            =   "frmGesSolDscto.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   900
      Width           =   2175
   End
   Begin VB.CommandButton cmdCAJMP 
      Caption         =   "Imprimir  Carta    CAJA M.P."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdDIECO 
      Caption         =   "  Imprimir      Carta     DIECO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cmbGrado 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmGesSolDscto.frx":0004
      Left            =   600
      List            =   "frmGesSolDscto.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   920
      Width           =   2775
   End
   Begin VB.TextBox txtCodSocio 
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
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      MaxLength       =   8
      TabIndex        =   3
      Top             =   420
      Width           =   690
   End
   Begin VB.TextBox txtCodigo 
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
      Left            =   600
      MaxLength       =   8
      TabIndex        =   2
      Top             =   420
      Width           =   930
   End
   Begin VB.TextBox txtIns 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   1
      Top             =   420
      Width           =   330
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
      Height          =   615
      Left            =   8640
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Modificables"
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
      Height          =   3975
      Left            =   240
      TabIndex        =   15
      Top             =   1320
      Width           =   8295
      Begin VB.ComboBox cmbSitu 
         Height          =   315
         ItemData        =   "frmGesSolDscto.frx":0008
         Left            =   6000
         List            =   "frmGesSolDscto.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   540
         Width           =   2055
      End
      Begin VB.TextBox txtNumDoc 
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
         Left            =   5040
         MaxLength       =   8
         TabIndex        =   38
         Top             =   540
         Width           =   930
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar Cambios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6360
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtDirec 
         Height          =   285
         Left            =   600
         MaxLength       =   50
         TabIndex        =   28
         Top             =   1140
         Width           =   7455
      End
      Begin VB.TextBox txtTelefono 
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
         Left            =   600
         MaxLength       =   20
         TabIndex        =   27
         Top             =   2085
         Width           =   2610
      End
      Begin VB.TextBox txtRefer 
         Height          =   285
         Left            =   600
         MaxLength       =   50
         TabIndex        =   26
         Top             =   1605
         Width           =   6855
      End
      Begin VB.TextBox txtTelefon2 
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
         Left            =   3360
         MaxLength       =   20
         TabIndex        =   25
         Top             =   2085
         Width           =   2370
      End
      Begin VB.TextBox txtCelular 
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
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   24
         Top             =   2085
         Width           =   1410
      End
      Begin VB.TextBox txtEMail2 
         Height          =   285
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   23
         Top             =   2565
         Width           =   3615
      End
      Begin VB.TextBox txteMail 
         Height          =   285
         Left            =   600
         MaxLength       =   50
         TabIndex        =   22
         Top             =   2565
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
         Left            =   1590
         MaxLength       =   8
         TabIndex        =   18
         Top             =   560
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
         Left            =   600
         MaxLength       =   8
         TabIndex        =   17
         Top             =   560
         Width           =   930
      End
      Begin VB.ComboBox cmbTipCob 
         Height          =   315
         ItemData        =   "frmGesSolDscto.frx":000C
         Left            =   2640
         List            =   "frmGesSolDscto.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   560
         Width           =   2415
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Situación Policial"
         Height          =   195
         Index           =   8
         Left            =   6360
         TabIndex        =   41
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "D.N.I."
         Height          =   195
         Index           =   5
         Left            =   5235
         TabIndex        =   39
         Top             =   360
         Width           =   420
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   195
         Index           =   11
         Left            =   1080
         TabIndex        =   35
         Top             =   960
         Width           =   675
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Teléfonos"
         Height          =   195
         Index           =   12
         Left            =   735
         TabIndex        =   34
         Top             =   1875
         Width           =   705
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Referencia"
         Height          =   195
         Index           =   15
         Left            =   600
         TabIndex        =   33
         Top             =   1425
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Celular"
         Height          =   195
         Index           =   20
         Left            =   3720
         TabIndex        =   32
         Top             =   1875
         Width           =   480
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Celular 2"
         Height          =   195
         Index           =   13
         Left            =   6105
         TabIndex        =   31
         Top             =   1875
         Width           =   615
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Correo Electrónico 2"
         Height          =   195
         Index           =   19
         Left            =   4080
         TabIndex        =   30
         Top             =   2355
         Width           =   1440
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Correo Electrónico"
         Height          =   195
         Index           =   14
         Left            =   840
         TabIndex        =   29
         Top             =   2355
         Width           =   1305
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Carnet PIP"
         Height          =   195
         Index           =   7
         Left            =   1665
         TabIndex        =   21
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Carnet PNP"
         Height          =   195
         Index           =   6
         Left            =   600
         TabIndex        =   20
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cobro"
         Height          =   195
         Index           =   18
         Left            =   3180
         TabIndex        =   19
         Top             =   360
         Width           =   780
      End
   End
   Begin Crystal.CrystalReport Crys1 
      Left            =   240
      Top             =   5760
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
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Estado de Socio"
      Height          =   195
      Index           =   16
      Left            =   5790
      TabIndex        =   14
      Top             =   720
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de Socio"
      Height          =   195
      Left            =   2760
      TabIndex        =   12
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblCodSocio 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   420
      Width           =   6375
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Grado"
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   8
      Top             =   720
      Width           =   915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Codigo"
      Height          =   195
      Index           =   0
      Left            =   2025
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Codofin"
      Height          =   195
      Index           =   1
      Left            =   675
      TabIndex        =   5
      Top             =   240
      Width           =   540
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Ins"
      Height          =   195
      Index           =   4
      Left            =   1560
      TabIndex        =   4
      Top             =   240
      Width           =   210
   End
End
Attribute VB_Name = "frmGesSolDscto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Limpiar()
   txtCodigo.Text = ""
   txtIns.Text = ""
   txtCodSocio.Text = ""
   txtNumDoc.Text = ""
   lblCodSocio.Caption = ""
   txtCarnetPNP.Text = ""
   txtCarnetPIP.Text = ""
   txtDirec.Text = ""
   txtRefer.Text = ""
   txtTelefono.Text = ""
   txtTelefon2.Text = ""
   txtCelular.Text = ""
   txteMail.Text = ""
   txtEMail2.Text = ""
   
   cmbE_Socio.ListIndex = 0
   cmbGrado.ListIndex = 0
   cmbSitu.ListIndex = 0
   cmbTipCob.ListIndex = 0
End Sub

Private Sub Llenar()
   Dim zz As Integer, wCod As Long
   
   wCod = Val(txtCodigo.Text)
   zz = Leerado8("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(wCod) + " ")
   If zz = 0 Then
      MsgBox "Codofin Digitado No Existe", vbExclamation
      Limpiar
      Exit Sub
   End If
   txtCodigo.Text = ADO8!codigo
   txtIns.Text = ADO8!ins
   txtCodSocio.Text = ADO8!codsocio
   txtNumDoc.Text = ADO8!numdoc
   txtCarnetPNP.Text = ADO8!carnetpnp
   txtCarnetPIP.Text = ADO8!carnetpip
   txtDirec.Text = ADO8!direc
   txtRefer.Text = ADO8!refer
   txtTelefono.Text = ADO8!telefono
   txtTelefon2.Text = ADO8!telefon2
   txtCelular.Text = ADO8!celular
   txteMail.Text = ADO8!email
   txtEMail2.Text = ADO8!email2

   cmbGrado.ListIndex = BuscaGrado(ADO8!grado)
   cmbE_Socio.ListIndex = BuscaEsocio(ADO8!e_socio)
   cmbTipCob.ListIndex = BuscaTipCob(ADO8!tipcob)
   cmbSitu.ListIndex = BuscaSitu(ADO8!situ)

   Select Case ADO8!tipcob
   Case "01"
        cmdDIECO.Enabled = True
        cmdCAJMP.Enabled = False
   Case "02"
        cmdDIECO.Enabled = False
        cmdCAJMP.Enabled = True
   Case "03"
        cmdDIECO.Enabled = False
        cmdCAJMP.Enabled = False
   End Select
   Set ADO8 = Nothing
End Sub

Private Sub cmbSitu_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtDirec.SetFocus
   End If
End Sub

Private Sub cmbTipCob_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtNumDoc.SetFocus
   End If
End Sub

Private Sub cmdCAJMP_Click()
   Dim wNombre As String, wDni As String, wCodofin As String, _
       wDirec As String, wDist As String, wCorreo As String, _
       wTelefono As String, wCelular As String, wSdoPen As String, _
       wCuoIni As String, wCanMes As String, wCuoMes As String, _
       wFecDia As String, wFecMes As String, wFecAno As String, _
       zz As Integer, wSoc As Integer, wCod As Long, wIns As Integer, _
       wNomAsig1 As String, wDniAsig1 As String, wCodAsig1 As Integer, _
       wNomAsig2 As String, wDniAsig2 As String, wCodAsig2 As Integer, _
       wNomAsig3 As String, wDniAsig3 As String, wCodAsig3 As Integer, _
       wNomAsig4 As String, wDniAsig4 As String, wCodAsig4 As Integer, _
       wNomAsig5 As String, wDniAsig5 As String, wCodAsig5 As Integer, _
       wCip As String
   
   wSoc = Val(txtCodSocio.Text)
   wNombre = Trim(lblCodSocio.Caption)
   wDni = txtNumDoc.Text
   wCodofin = Trim(txtCodigo.Text) + "-" + Trim(txtIns.Text)
   wDirec = "": wDist = "": wCorreo = "": wTelefono = "": wCelular = "": wCip = ""
   wNomAsig1 = "": wDniAsig1 = "": wCodAsig1 = 0
   wNomAsig2 = "": wDniAsig2 = "": wCodAsig2 = 0
   wNomAsig3 = "": wDniAsig3 = "": wCodAsig3 = 0
   wNomAsig4 = "": wDniAsig4 = "": wCodAsig4 = 0
   wNomAsig5 = "": wDniAsig5 = "": wCodAsig5 = 0
   
   zz = Leerado8("SELECT * FROM MAEASIGNADO WHERE CODSOCIO = " + Str(wSoc) + " ")
   If zz > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
   
         Select Case ADO8!lin
         Case "01"
              wCodAsig1 = ADO8!codhijo
              wNomAsig1 = BuscaDatosSocio(wCodAsig1, 3)
              wDniAsig1 = BuscaDatosSocio(wCodAsig1, 4)
         Case "02"
              wCodAsig2 = ADO8!codhijo
              wNomAsig2 = BuscaDatosSocio(wCodAsig2, 3)
              wDniAsig2 = BuscaDatosSocio(wCodAsig2, 4)
         Case "03"
              wCodAsig3 = ADO8!codhijo
              wNomAsig3 = BuscaDatosSocio(wCodAsig3, 3)
              wDniAsig3 = BuscaDatosSocio(wCodAsig3, 4)
         Case "04"
              wCodAsig4 = ADO8!codhijo
              wNomAsig4 = BuscaDatosSocio(wCodAsig4, 3)
              wDniAsig4 = BuscaDatosSocio(wCodAsig4, 4)
         Case "05"
              wCodAsig5 = ADO8!codhijo
              wNomAsig5 = BuscaDatosSocio(wCodAsig5, 3)
              wDniAsig5 = BuscaDatosSocio(wCodAsig5, 4)
         End Select
   
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
   If zz > 0 Then
      wDirec = ADO8!direc
      wDist = ADO8!ubigeo
      If Len(Trim(ADO8!email)) > 0 Then
         wCorreo = ADO8!email
      Else
         If Len(Trim(ADO8!email2)) > 0 Then
            wCorreo = ADO8!email2
         End If
      End If
      wTelefono = Trim(IIf(IsNull(ADO8!telefono), "", Trim(ADO8!telefono)) + " " + _
                       IIf(IsNull(ADO8!telefon2), "", Trim(ADO8!telefon2)))
      wCelular = IIf(IsNull(ADO8!celular), "", Trim(ADO8!celular))
      wCip = IIf(IsNull(ADO8!carnetpnp), "", ADO8!carnetpnp)
   
   
      zz = Leerado7("SELECT * FROM MAESITU WHERE SITU = " + Str(ADO8!situ) + " ")
      If zz > 0 Then
         wSituac = ADO7!nombre
      End If
      Set ADO7 = Nothing
   
   End If
   Set ADO8 = Nothing
   
   If Len(Trim(wDist)) > 0 Then
      zz = Leerado8("SELECT * FROM MAEUBIGEO WHERE CODIGO = '" + wDist + "' ")
      If zz > 0 Then
         wDist = ADO8!nombre
      End If
      Set ADO8 = Nothing
   End If
   wFecDia = Format(Day(Date), "00")
   wFecMes = Trim(funnommes(Format(Month(Date), "00")))
   wFecAno = Right(Format(Year(Date), "0000"), 2)
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\DesctoCAJMP.RPT"
   Crys1.Formulas(0) = "NOMBRE= '" + wNombre + "' "
   Crys1.Formulas(1) = "DNI= '" + wDni + "' "
   Crys1.Formulas(2) = "CODOFIN= '" + wCodofin + "' "
   Crys1.Formulas(3) = "DIREC= '" + wDirec + "' "
   Crys1.Formulas(4) = "DIST= '" + wDist + "' "
   Crys1.Formulas(5) = "CORREO= '" + wCorreo + "' "
   Crys1.Formulas(6) = "TELEFONO= '" + wTelefono + "' "
   Crys1.Formulas(7) = "CELULAR= '" + wCelular + "' "
   Crys1.Formulas(8) = "FECDIA= '" + wFecDia + "' "
   Crys1.Formulas(9) = "FECMES= '" + wFecMes + "' "
   Crys1.Formulas(10) = "FECANO= '" + wFecAno + "' "
   Crys1.Formulas(11) = "NOMASIG1= '" + wNomAsig1 + "' "
   Crys1.Formulas(12) = "NOMASIG2= '" + wNomAsig2 + "' "
   Crys1.Formulas(13) = "NOMASIG3= '" + wNomAsig3 + "' "
   Crys1.Formulas(14) = "NOMASIG4= '" + wNomAsig4 + "' "
   Crys1.Formulas(15) = "NOMASIG5= '" + wNomAsig5 + "' "
   Crys1.Formulas(16) = "DNIASIG1= '" + wDniAsig1 + "' "
   Crys1.Formulas(17) = "DNIASIG2= '" + wDniAsig2 + "' "
   Crys1.Formulas(18) = "DNIASIG3= '" + wDniAsig3 + "' "
   Crys1.Formulas(19) = "DNIASIG4= '" + wDniAsig4 + "' "
   Crys1.Formulas(20) = "DNIASIG5= '" + wDniAsig5 + "' "
   Crys1.Formulas(21) = "CIP= '" + wCip + "' "
   Crys1.Formulas(22) = "SITUAC= '" + wSituac + "' "
   Crys1.SelectionFormula = " {MAESOCIO.CODSOCIO}=" + Str(wSoc) + " "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdDIECO_Click()
   Dim wNombre As String, wDni As String, wCodofin As String, _
       wDirec As String, wDist As String, wCorreo As String, _
       wTelefono As String, wCelular As String, wSdoPen As String, _
       wCuoIni As String, wCanMes As String, wCuoMes As String, _
       wFecDia As String, wFecMes As String, wFecAno As String, _
       zz As Integer, wSoc As Integer, wCod As Long, wIns As Integer, _
       wNomAsig1 As String, wDniAsig1 As String, wCodAsig1 As Integer, _
       wNomAsig2 As String, wDniAsig2 As String, wCodAsig2 As Integer, _
       wNomAsig3 As String, wDniAsig3 As String, wCodAsig3 As Integer, _
       wNomAsig4 As String, wDniAsig4 As String, wCodAsig4 As Integer, _
       wNomAsig5 As String, wDniAsig5 As String, wCodAsig5 As Integer, _
       wCip As String, wSituac As String
   
   wSoc = Val(txtCodSocio.Text)
   wNombre = Trim(lblCodSocio.Caption)
   wDni = txtNumDoc.Text
   wCodofin = Trim(txtCodigo.Text) + "-" + Trim(txtIns.Text)
   wDirec = "": wDist = "": wCorreo = "": wTelefono = "": wCelular = "": wCip = ""
   
   wNomAsig1 = "": wDniAsig1 = "": wCodAsig1 = 0
   wNomAsig2 = "": wDniAsig2 = "": wCodAsig2 = 0
   wNomAsig3 = "": wDniAsig3 = "": wCodAsig3 = 0
   wNomAsig4 = "": wDniAsig4 = "": wCodAsig4 = 0
   wNomAsig5 = "": wDniAsig5 = "": wCodAsig5 = 0
   
   zz = Leerado8("SELECT * FROM MAEASIGNADO WHERE CODSOCIO = " + Str(wSoc) + " ")
   If zz > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
   
         Select Case ADO8!lin
         Case "01"
              wCodAsig1 = ADO8!codhijo
              wNomAsig1 = BuscaDatosSocio(wCodAsig1, 3)
              wDniAsig1 = BuscaDatosSocio(wCodAsig1, 4)
         Case "02"
              wCodAsig2 = ADO8!codhijo
              wNomAsig2 = BuscaDatosSocio(wCodAsig2, 3)
              wDniAsig2 = BuscaDatosSocio(wCodAsig2, 4)
         Case "03"
              wCodAsig3 = ADO8!codhijo
              wNomAsig3 = BuscaDatosSocio(wCodAsig3, 3)
              wDniAsig3 = BuscaDatosSocio(wCodAsig3, 4)
         Case "04"
              wCodAsig4 = ADO8!codhijo
              wNomAsig4 = BuscaDatosSocio(wCodAsig4, 3)
              wDniAsig4 = BuscaDatosSocio(wCodAsig4, 4)
         Case "05"
              wCodAsig5 = ADO8!codhijo
              wNomAsig5 = BuscaDatosSocio(wCodAsig5, 3)
              wDniAsig5 = BuscaDatosSocio(wCodAsig5, 4)
         End Select
   
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
   If zz > 0 Then
      wDirec = ADO8!direc
      wDist = ADO8!ubigeo
      If Len(Trim(ADO8!email)) > 0 Then
         wCorreo = ADO8!email
      Else
         If Len(Trim(ADO8!email2)) > 0 Then
            wCorreo = ADO8!email2
         End If
      End If
      wTelefono = Trim(IIf(IsNull(ADO8!telefono), "", Trim(ADO8!telefono)) + " " + _
                       IIf(IsNull(ADO8!telefon2), "", Trim(ADO8!telefon2)))
      wCelular = IIf(IsNull(ADO8!celular), "", Trim(ADO8!celular))
      wCip = IIf(IsNull(ADO8!carnetpnp), "", ADO8!carnetpnp)
   
      zz = Leerado7("SELECT * FROM MAESITU WHERE SITU = " + Str(ADO8!situ) + " ")
      If zz > 0 Then
         wSituac = ADO7!nombre
      End If
      Set ADO7 = Nothing
   
   End If
   Set ADO8 = Nothing
   
   If Len(Trim(wDist)) > 0 Then
      zz = Leerado8("SELECT * FROM MAEUBIGEO WHERE CODIGO = '" + wDist + "' ")
      If zz > 0 Then
         wDist = ADO8!nombre
      End If
      Set ADO8 = Nothing
   End If
   wFecDia = Format(Day(Date), "00")
   wFecMes = Trim(funnommes(Format(Month(Date), "00")))
   wFecAno = Right(Format(Year(Date), "0000"), 2)
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\DesctoDIECO.RPT"
   Crys1.Formulas(0) = "NOMBRE= '" + wNombre + "' "
   Crys1.Formulas(1) = "DNI= '" + wDni + "' "
   Crys1.Formulas(2) = "CODOFIN= '" + wCodofin + "' "
   Crys1.Formulas(3) = "DIREC= '" + wDirec + "' "
   Crys1.Formulas(4) = "DIST= '" + wDist + "' "
   Crys1.Formulas(5) = "CORREO= '" + wCorreo + "' "
   Crys1.Formulas(6) = "TELEFONO= '" + wTelefono + "' "
   Crys1.Formulas(7) = "CELULAR= '" + wCelular + "' "
   Crys1.Formulas(8) = "FECDIA= '" + wFecDia + "' "
   Crys1.Formulas(9) = "FECMES= '" + wFecMes + "' "
   Crys1.Formulas(10) = "FECANO= '" + wFecAno + "' "
   Crys1.Formulas(11) = "NOMASIG1= '" + wNomAsig1 + "' "
   Crys1.Formulas(12) = "NOMASIG2= '" + wNomAsig2 + "' "
   Crys1.Formulas(13) = "NOMASIG3= '" + wNomAsig3 + "' "
   Crys1.Formulas(14) = "NOMASIG4= '" + wNomAsig4 + "' "
   Crys1.Formulas(15) = "NOMASIG5= '" + wNomAsig5 + "' "
   Crys1.Formulas(16) = "DNIASIG1= '" + wDniAsig1 + "' "
   Crys1.Formulas(17) = "DNIASIG2= '" + wDniAsig2 + "' "
   Crys1.Formulas(18) = "DNIASIG3= '" + wDniAsig3 + "' "
   Crys1.Formulas(19) = "DNIASIG4= '" + wDniAsig4 + "' "
   Crys1.Formulas(20) = "DNIASIG5= '" + wDniAsig5 + "' "
   Crys1.Formulas(21) = "CIP= '" + wCip + "' "
   Crys1.Formulas(22) = "SITUAC= '" + wSituac + "' "
   Crys1.SelectionFormula = " {MAESOCIO.CODSOCIO}=" + Str(wSoc) + " "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdGrabar_Click()
   Dim wCarnetPIP As String, wCarnetPnp As Long, wTipCob As String, _
       wDirec As String, wRefer As String, wTelef As String, wTele2 As String, wCelul As String, _
       weMail As String, wMai2 As String, _
       wCodigo As Long, wIns As Integer, wSocio As Integer, wNumDoc As String, _
       wSituac As Integer

   wSocio = Val(txtCodSocio.Text)
   wCodigo = Val(txtCodigo.Text)
   wIns = Val(txtIns.Text)
   wCarnetPnp = Val(txtCarnetPNP.Text)
   wCarnetPIP = txtCarnetPIP.Text
   wTipCob = BuscaCodTipCob(cmbTipCob.List(cmbTipCob.ListIndex))
   wDirec = txtDirec.Text
   wRefer = txtRefer.Text
   wTelef = txtTelefono.Text
   wTele2 = txtTelefon2.Text
   wCelul = txtCelular.Text
   weMail = txteMail.Text
   wEMAI2 = txtEMail2.Text
   wNumDoc = txtNumDoc.Text
   wSituac = BuscaCodSitu(cmbSitu.List(cmbSitu.ListIndex))
      
   
   Db.BeginTrans
   Db.Execute ("UPDATE MAESOCIO " _
   & " SET CARNETPNP = " + Str(wCarnetPnp) + ", CARNETPIP = '" + wCarnetPIP + "', " _
   & "        TIPCOB = '" + wTipCob + "',           DIREC = '" + wDirec + "',  " _
   & "         REFER = '" + wRefer + "',         TELEFONO = '" + wTelef + "',  " _
   & "      TELEFON2 = '" + wTele2 + "',          CELULAR = '" + wCelul + "',  " _
   & "         EMAIL = '" + weMail + "',           EMAIL2 = '" + wEMAI2 + "', " _
   & "        NUMDOC = '" + wNumDoc + "',            SITU = " + Str(wSituac) + "  " _
   & " WHERE CODSOCIO = " + Str(wSocio) + " ")
   Db.CommitTrans

   cmdDIECO.Enabled = False
   cmdCAJMP.Enabled = False
   
   Select Case wTipCob
   Case "01"
        cmdDIECO.Enabled = True
   
        cmdDIECO.SetFocus
   Case "02"
        cmdCAJMP.Enabled = True
   
        cmdCAJMP.SetFocus
   Case "03"
   End Select
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cndOtro_Click()
   Limpiar
   
   cmdDIECO.Enabled = True
   cmdCAJMP.Enabled = True
   
   txtCodigo.SetFocus
End Sub

Private Sub Form_Activate()
   frmGesSolDscto.Left = (Screen.Width - Width) \ 2
   frmGesSolDscto.Top = 0
   
   Dim a As Integer
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
   
   a = Leerado8("SELECT * FROM MAESITU ORDER BY SITU ")
   If a > 0 Then
      ADO8.MoveFirst
      I = 0
      Do While Not ADO8.EOF
         cmbSitu.AddItem ADO8!nombre
         I = I + 1
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   
   Limpiar
   
   txtCodigo.SetFocus
End Sub

Private Sub txtCarnetPIP_GotFocus()
   txtCarnetPIP.SelStart = 0
   txtCarnetPIP.SelLength = 8
End Sub

Private Sub txtCarnetPIP_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtCarnetPNP.SetFocus
   Case 40
        cmbTipCob.SetFocus
   End Select
End Sub

Private Sub txtCarnetPIP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmbTipCob.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtCarnetPNP_GotFocus()
   txtCarnetPNP.SelStart = 0
   txtCarnetPNP.SelLength = 8
End Sub

Private Sub txtCarnetPNP_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 40
        txtCarnetPIP.SetFocus
   End Select
End Sub

Private Sub txtCarnetPNP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtCarnetPIP.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtCelular_GotFocus()
   txtCelular.SelStart = 0
   txtCelular.SelLength = 10
End Sub

Private Sub txtCelular_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtTelefon2.SetFocus
   Case 40
        txteMail.SetFocus
   End Select
End Sub

Private Sub txtCelular_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txteMail.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
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
      
      Llenar
   
      txtCarnetPNP.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtDirec_GotFocus()
   txtDirec.SelStart = 0
   txtDirec.SelLength = 50
End Sub

Private Sub txtDirec_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtNumDoc.SetFocus
   Case 40
        txtRefer.SetFocus
   End Select
End Sub

Private Sub txtDirec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtRefer.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txteMail_GotFocus()
   txteMail.SelStart = 0
   txteMail.SelLength = 50
End Sub

Private Sub txteMail_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtCelular.SetFocus
   Case 40
        txtEMail2.SetFocus
   End Select
End Sub

Private Sub txteMail_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtEMail2.SetFocus
   Else
      KeyAscii = Asc(LCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtEMail2_GotFocus()
   txtEMail2.SelStart = 0
   txtEMail2.SelLength = 50
End Sub

Private Sub txtEMail2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txteMail.SetFocus
   End Select
End Sub

Private Sub txtEMail2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdGrabar.SetFocus
   Else
      KeyAscii = Asc(LCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtNumDoc_GotFocus()
   txtNumDoc.SelStart = 0
   txtNumDoc.SelLength = Len(Trim(txtNumDoc.Text))
End Sub

Private Sub txtNumDoc_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        cmbTipCob.SetFocus
   Case 40
        cmbSitu.SetFocus
   End Select
End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmbSitu.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtRefer_GotFocus()
   txtRefer.SelStart = 0
   txtRefer.SelLength = 50
End Sub

Private Sub txtRefer_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDirec.SetFocus
   Case 40
        txtTelefono.SetFocus
   End Select
End Sub

Private Sub txtRefer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtTelefono.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtTelefon2_GotFocus()
   txtTelefon2.SelStart = 0
   txtTelefon2.SelLength = 20
End Sub

Private Sub txtTelefon2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtTelefono.SetFocus
   Case 40
        txtCelular.SetFocus
   End Select
End Sub

Private Sub txtTelefon2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtCelular.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtTelefono_GotFocus()
   txtTelefono.SelStart = 0
   txtTelefono.SelLength = 20
End Sub

Private Sub txtTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtRefer.SetFocus
   Case 40
        txtTelefon2.SetFocus
   End Select
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtTelefon2.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub


