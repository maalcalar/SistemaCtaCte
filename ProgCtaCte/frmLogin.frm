VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de Inventarios"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11295
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   120
      Picture         =   "frmLogin.frx":030A
      ScaleHeight     =   4275
      ScaleWidth      =   11115
      TabIndex        =   14
      Top             =   0
      Width           =   11175
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
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
      Left            =   7920
      TabIndex        =   10
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00000040&
      Caption         =   "&OK"
      Enabled         =   0   'False
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
      Left            =   6360
      TabIndex        =   9
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   4440
      Width           =   11295
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo cmbUsuario 
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NOMBRE"
         BoundColumn     =   "CODIGO"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin VB.Label Label2 
         Caption         =   "Clave "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Compañia y Mes de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   5640
      Width           =   11295
      Begin MSDataListLib.DataCombo cmbMes 
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Top             =   660
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   1020
         Width           =   615
      End
      Begin MSDataListLib.DataCombo cmbCias 
         DataField       =   "300"
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   300
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   ""
         BoundColumn     =   "CODCIA"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin VB.Label Label7 
         Caption         =   "Año "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Mes "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Nombre"
         DataField       =   "300"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OK As Boolean, ACCESO As Byte

Private Sub cmbAlm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdOK.Enabled = True
      cmdOK.SetFocus
   End If
End Sub

Private Sub cmbCias_Click(Area As Integer)
   If Area = 2 Then
      If cmbCias.Text <> "" Then
         cmbCias_KeyPress (13)
      End If
   End If
End Sub

Private Sub cmbCias_GotFocus()
    ADOMaster2.Filter = ""
End Sub

Private Sub cmbCias_KeyPress(KeyAscii As Integer)
   Dim a As Integer
   If cmbCias.Text = "" Then
      MsgBox "No Se Seleccionó Ninguna Compañia", vbCritical
      Exit Sub
   End If
    
   With ADOMaster2
      ADOMaster2.Filter = "CODIGOCIA = '" & Mid(cmbCias, 1, 2) & "' "
      wcodcia = .Fields("CODIGOCIA")
      wporigv = wporigv
   End With
    
   cmbMes.Enabled = True
   cmdOK.Enabled = True
   
   With ADOMaster
       .MoveFirst
       .Find "[PASSWORD]='" & txtPassword & "' "
       If .EOF Then
          MsgBox "Contraseña No Es Valida; Vuelva a Intentarlo", vbExclamation, "Inicio de Sesión"
          ACCESO = ACCESO + 1
          If ACCESO >= 3 Then
             MsgBox "Se Cancelara el Programa ....", vbExclamation, "Inicio de Sesión"
             Unload Me
             Exit Sub
          End If
          txtPassword = ""
          txtPassword.SetFocus
          Exit Sub
       End If
   End With
   With ADOMaster
'        MENUTAB = .Fields("INV_TABLAS")
'        MENUORD = .Fields("INV_ORDCOM")
'        MENUGUI = .Fields("INV_GUIAS")
'        MENUVTA = .Fields("INV_VENTAS")
'        MENUNOT = .Fields("INV_NOTAS")
'        MENUINV = .Fields("INV_INVFIS")
'        MENUCON = .Fields("INV_CONSULTAS")
'        MENUSER = .Fields("INV_SERVICIOS")
        SUPERVISOR = .Fields("SUPERVISOR")
        wnomusu = .Fields("ABREV")
        wcodusu = .Fields("CODIGO")
        wswprint = .Fields("SWPRINT")
   End With
   
   cmbMes.SetFocus
End Sub

Private Sub cmbMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      With ADOMaster
          .MoveFirst
          .Find "[PASSWORD]='" & txtPassword & "' "
          If .EOF Then
             MsgBox "Contraseña No Es Valida; Vuelva a Intentarlo", vbExclamation, "Inicio de Sesión"
             ACCESO = ACCESO + 1
             If ACCESO >= 3 Then
                MsgBox "Se Cancelara el Programa ....", vbExclamation, "Inicio de Sesión"
                Unload Me
                Exit Sub
             End If
             txtPassword = ""
             txtPassword.SetFocus
             Exit Sub
          End If
      End With
      With ADOMaster
           SUPERVISOR = .Fields("SUPERVISOR")
           wnomusu = .Fields("ABREV")
           wcodusu = .Fields("CODIGO")
           wswprint = .Fields("SWPRINT")
      End With
      txtAno.Enabled = True
      txtAno.SetFocus
   End If
End Sub

Private Sub cmbUsuario_Click(Area As Integer)
   If Area = 2 Then
      If cmbUsuario.Text <> "" Then
         cmbUsuario_KeyPress (13)
      End If
   End If
End Sub

Private Sub cmbUsuario_KeyPress(KeyAscii As Integer)
    If cmbUsuario.Text = "" Then
       MsgBox "No Se Seleccionó Ningún Usuario", vbCritical
       Exit Sub
    End If
    
    With ADOMaster
       .Filter = "CODIGO = '" & Mid(cmbUsuario.Text, 1, 3) & "' "
    End With
    txtPassword.SetFocus
End Sub

Private Sub cmdOK_Click()
   
   With ADOMaster
       .MoveFirst
       .Find "[PASSWORD]='" & txtPassword & "' "
       If .EOF Then
          MsgBox "Contraseña No Es Valida; Vuelva a Intentarlo", vbExclamation, "Inicio de Sesión"
          ACCESO = ACCESO + 1
          If ACCESO >= 3 Then
             MsgBox "Se Cancelara el Programa ....", vbExclamation, "Inicio de Sesión"
             Unload Me
             Exit Sub
          End If
          txtPassword = ""
          txtPassword.SetFocus
          Exit Sub
       End If
   End With
   With ADOMaster
        MENUMAE = ADOMaster!cxc_maestro
        MENUELE = ADOMaster!cxc_eleccion
        MENUGES = ADOMaster!cxc_gestion
        MENUAPO = ADOMaster!cxc_aporte
        MENUDIE = ADOMaster!cxc_dieco
        MENUCAJ = ADOMaster!cxc_cajamp
        MENUTES = ADOMaster!cxc_tesor
        MENUCON = ADOMaster!cxc_consulta
        MENUSER = ADOMaster!cxc_servicios
	MENUBCP = ADOMaster!cxc_bcp
	MENURPT = ADOMaster!cxc_repteso
	MENUCNT = ADOMaster!cxc_cont
        SUPERVISOR = .Fields("SUPERVISOR")
        wnomusu = .Fields("NOMBRE")
        wcodusu = .Fields("CODIGO")
        wswprint = .Fields("SWPRINT")
        If ADOMaster!codigo = wcodusu Then
           .Fields("MESALM") = Left(cmbMes.Text, 2)
           .Fields("ANOALM") = Left(txtAno, 4)
           .Fields("ALMALM") = walmcia
           .Update
        End If
   End With
   
   wcodcia = Mid(cmbCias.Text, 1, 2)
   wnomcia = cmbCias.Text
   wruccia = Mid(cmbCias.Text, 4, 11)
   wnomcia = Mid(cmbCias.Text, 16, 80)
   wmescia = Mid(cmbMes.Text, 1, 2)
   wdiacia = fundiames(wmescia)
      
   wanocia = txtAno
'   wmesnom = Mid(cmbMes.Text, 4, 9)
   Dim a As Integer
   With ADOMaster2
'      wporigv = ADOMaster2!porigv
      wdircia = "AV CASUARINAS 450 URB. VALLE HERMOSO - SANTIAGO DE SURCO"
      wtelcia = "3444100"
'      wnomlog = IIf(IsNull(ADOMaster2!nomlog), "", ADOMaster2!nomlog)
   End With
   
'   If wmescia = "01" Then
'      zMesTope = Format(Val(wanocia) - 1, "0000") + "12"
'   Else
'      zMesTope = wanocia + Format(Val(wmescia) - 1, "00")
'   End If
   
   Dim zDiaFin As Date, zDiaHoy As Date
   zDiaFin = fundiames(wmescia) + "/" + wmescia + "/" + wanocia
   zDiaHoy = Format(Date, "dd/mm/yyyy")
   
   If Format(zDiaHoy, "dd/mm/yyyy") < Format(zDiaFin, "dd/mm/yyyy") Then
      If wmescia > "01" Then
         zMesTope = wanocia + Format(Val(wmescia) - 1, "00")
      Else
         zMesTope = Format(Val(wanocia) - 1, "0000") + "12"
      End If
   Else
      zMesTope = wanocia + wmescia
   End If
   
   OK = True
   Me.Hide
End Sub

Private Sub cmdSalir_Click()
   OK = False
   Me.Hide
End Sub

Private Sub Form_Activate()
    Dim a As Integer
    
    a = LeeradoMaster3("SELECT *,MES+' '+NOMBRE AS NOMMES " _
                    & " FROM MAEMESES " _
                    & " ORDER BY MES")
    Set cmbMes.RowSource = ADOMaster3
    cmbMes.ListField = "NOMMES"
    
    a = LeeradoMaster("SELECT *,CODIGO+' '+NOMBRE AS USU " _
                    & " FROM USUARIOS " _
                    & " WHERE SWCTACTE = 1 " _
                    & " ORDER BY CODIGO")
    Set cmbUsuario.RowSource = ADOMaster
    cmbUsuario.ListField = "USU"
    
    ADOMaster.MoveFirst
    cmbUsuario.Text = ADOMaster.Fields("CODIGO") + " " + ADOMaster.Fields("NOMBRE")
    
    a = LeeradoMaster2("SELECT *,ANO,CODIGOCIA+' '+RUC+' '+NOMBRECIA AS CIA " _
                    & " FROM COMPANIAS " _
                    & " ORDER BY CODIGOCIA")
    Set cmbCias.RowSource = ADOMaster2
    cmbCias.ListField = "Cia"
    
    cmbUsuario.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ADOMaster = Nothing
    Set ADOMaster2 = Nothing
    Set ADOMaster3 = Nothing
End Sub

Private Sub txtAno_GotFocus()
    txtAno.SelStart = 0
    txtAno.SelLength = Len(txtAno)
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtAno = Format(txtAno.Text, "0000")
       If txtAno < "2000" Or txtAno > "2040" Then
          MsgBox "Año de Proceso Fuera De Rango", vbCritical
          txtAno = ""
          Exit Sub
       End If
       cmdOK.SetFocus
    End If
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If txtPassword.Text = "" Then
          MsgBox "Contraseña En Blanco"
          cmbCias.Text = ""
          cmbMes.Text = ""
          txtAno.Text = ""
          Exit Sub
       End If
       With ADOMaster
           .MoveFirst
           .Find "[PASSWORD]='" & txtPassword & "' "
           If .EOF Then
              MsgBox "Contraseña No Es Valida; Vuelva a Intentarlo", vbExclamation, "Inicio de Sesión"
              ACCESO = ACCESO + 1
              If ACCESO >= 3 Then
                 MsgBox "Se Cancelara el Programa ....", vbExclamation, "Inicio de Sesión"
                 Unload Me
                 Exit Sub
              End If
              txtPassword = ""
              Exit Sub
           End If
       End With
       With ADOMaster
            SUPERVISOR = .Fields("SUPERVISOR")
            wnomusu = .Fields("ABREV")
            wcodusu = .Fields("CODIGO")
            wswprint = .Fields("SWPRINT")
            
            wmescia = Format(Month(Date), "00")
            wdiacia = Format(Day(Date), "00")
            wanocia = Format(Year(Date), "0000")
            txtAno.Text = wanocia
            
            
'            wmescia = IIf(IsNull(ADOMaster!mesalm), "12", ADOMaster!mesalm)
'            wdiacia = fundiames(wmescia)
'            txtAno.Text = IIf(IsNull(ADOMaster!anoalm), "2016", ADOMaster!anoalm)
            
            If Len(Trim(funnommes(wmescia))) = 9 Then
               cmbMes.Text = wmescia + " " + Trim(funnommes(wmescia))
            Else
               cmbMes.Text = wmescia + " " + Trim(funnommes(wmescia)) + Space(9 - Len(Trim(funnommes(wmescia))))
            End If
            
       End With
       ADOMaster2.MoveFirst
       cmbCias.Text = ADOMaster2!codigocia + " " + ADOMaster2!ruc + " " + ADOMaster2!NombreCia
       
       wcodcia = ADOMaster2.Fields("CODIGOCIA")
       cmdOK.Enabled = True
'       cmdOK.SetFocus
       cmbCias.SetFocus
    End If
End Sub
