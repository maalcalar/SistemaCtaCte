VERSION 5.00
Begin VB.Form frmServActualiza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualiza Saldos de Cuentas Por Cobrar"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   10425
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "&Actualiza"
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
      Left            =   6720
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
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
      Left            =   8400
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmServActualiza.frx":0000
      Left            =   1440
      List            =   "frmServActualiza.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2400
      Width           =   7935
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   240
      Picture         =   "frmServActualiza.frx":0004
      ScaleHeight     =   1635
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Proceso"
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
      TabIndex        =   5
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblMensaje 
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2880
      Width           =   7935
   End
   Begin VB.Label Label25 
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
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Esta Opción Sirve Para Actualizar Los Saldos de Cuentas Por Cobrar."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "frmServActualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wcia As String

Private Sub cmdActualiza_Click()
   Dim aa As Long, wRegAct As Long, wRegTot As Long, _
       wSoc As Long, wCod As Long, wIns As Integer, _
       wMes As String, wcon As String, WE_S As String, _
       wMon As String, _
       wSdoOld As Currency, wSdoNew As Currency, wCargos As Currency, wAbonos As Currency, _
       zSdoOld As Currency, zSdoNew As Currency, zCargos As Currency, zAbonos As Currency, _
       wTipMov As String, wTipCob As String, wSerCob As String, _
       wNumCob As String, wLinCob As String, wFec As Date

   aa = Leerado8a("SELECT C.CODSOCIO, C.MES, C.CONCEPTO, C.CARGOS, " _
                & "       C.ABONOS, C.SDONEW, M.E_SOCIO, M.NOMBRE, " _
                & "       M.CODIGO, M.INS " _
                & " FROM CTASXCAB AS C LEFT JOIN MAESOCIO AS M " _
                & "   ON C.CODSOCIO = M.CODSOCIO " _
                & " ORDER BY C.CODSOCIO, C.MES, C.CONCEPTO ")
   If aa > 0 Then
      wRegAct = 1
      wRegTot = aa
      ADO8a.MoveFirst
      Do While Not ADO8a.EOF
         DoEvents
         lblMensaje.Caption = "Registro " + _
                              Trim(Format(wRegAct, "####,##0")) + " / " + _
                              Trim(Format(wRegTot, "####,##0"))
         lblMensaje.Refresh
   
         wSoc = ADO8a!codsocio
         wMes = ADO8a!mes
         wcon = ADO8a!concepto
         zSdoOld = 0: zSdoNew = 0: zCargos = 0: zAbonos = 0
   
         wSdoOld = 0: wSdoNew = 0: wCargos = 0: wAbonos = 0
         aa = Leerado7a("SELECT * FROM CTASXDET " _
            & " WHERE CODSOCIO = '" + Str(wSoc) + "' AND " _
            & "            MES = '" + wMes + "' AND " _
            & "       CONCEPTO = '" + wcon + "' " _
            & " ORDER BY FECHA, TIPMOV, TIPCOB, SERCOB, NUMCOB, LINCOB ")
         If aa > 0 Then
            ADO7a.MoveFirst
            Do While Not ADO7a.EOF
               wTipMov = ADO7a!tipmov
               wTipCob = ADO7a!tipcob
               wSerCob = ADO7a!sercob
               wNumCob = ADO7a!numcob
               wLinCob = ADO7a!lincob
               wFec = Format(ADO7a!fecha, "dd/mm/yyyy")
          
               wCargos = ADO7a!cargos
               wAbonos = ADO7a!abonos
               wSdoNew = wSdoOld + wCargos - wAbonos
               
               zCargos = zCargos + wCargos
               zAbonos = zAbonos + wAbonos
               zSdoNew = zSdoOld + zCargos - zAbonos
             
               Db.BeginTrans
               Db.Execute ("UPDATE CTASXDET " _
               & " SET SDOOLD = " + Str(wSdoOld) + ", " _
               & "     CARGOS = " + Str(wCargos) + ", " _
               & "     ABONOS = " + Str(wAbonos) + ", " _
               & "     SDONEW = " + Str(wSdoNew) + " " _
               & " WHERE CODSOCIO = '" + Str(wSoc) + "' AND " _
               & "            MES = '" + wMes + "' AND " _
               & "       CONCEPTO = '" + wcon + "' AND " _
               & "        FECHA = '" + Format(wFec, "dd/mm/yyyy") + "' AND " _
               & "       TIPMOV = '" + wTipMov + "' AND " _
               & "       TIPCOB = '" + wTipCob + "' AND " _
               & "       SERCOB = '" + wSerCob + "' AND " _
               & "       NUMCOB = '" + wNumCob + "' AND " _
               & "       LINCOB = '" + wLinCob + "' ")
               Db.CommitTrans
         
               wSdoOld = wSdoNew
               
               ADO7a.MoveNext
            Loop
         End If
             
         Db.BeginTrans
         Db.Execute ("UPDATE CTASXCAB " _
         & " SET CARGOS = " + Str(zCargos) + ", " _
         & "     ABONOS = " + Str(zAbonos) + ", " _
         & "     SDONEW = " + Str(zSdoNew) + " " _
         & " WHERE CODSOCIO = '" + Str(wSoc) + "' AND " _
         & "            MES = '" + wMes + "' AND " _
         & "       CONCEPTO = '" + wcon + "' ")
         Db.CommitTrans
   
         wRegAct = wRegAct + 1
         ADO8a.MoveNext
      Loop
   End If

   MsgBox "Proceso Termino OK", vbExclamation
   Unload Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Form_Initialize
End Sub

Private Sub Form_Initialize()
   frmServActualiza.Left = (Screen.Width - Width) \ 2
   frmServActualiza.Top = 0
   
   Dim a As Integer
   
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
   wcia = Left(cmbCia.Text, 2)
   cmbCia.Enabled = False

   cmdActualiza.SetFocus
End Sub


