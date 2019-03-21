VERSION 5.00
Begin VB.Form frmServAdel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignar Adelantos"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   12090
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmServAdel.frx":0000
      Left            =   1800
      List            =   "frmServAdel.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2520
      Width           =   8895
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
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdAsignar 
      Caption         =   "&Asignar"
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
      Left            =   7560
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   600
      Picture         =   "frmServAdel.frx":0004
      ScaleHeight     =   1635
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   360
      Width           =   2895
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
      Left            =   600
      TabIndex        =   7
      Top             =   2520
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
      Left            =   1800
      TabIndex        =   6
      Top             =   3000
      Width           =   8895
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
      Left            =   600
      TabIndex        =   5
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmServAdel.frx":160142
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
      Height          =   1695
      Left            =   3840
      TabIndex        =   1
      Top             =   360
      Width           =   6855
   End
End
Attribute VB_Name = "frmServAdel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAsignar_Click()
   Dim aa As Long, wRegAct As Long, wRegTot As Long, _
       wSoc As Integer, wCod As Integer, wIns As Integer, _
       wAde As Currency, wSdo As Currency, wAbo As Currency, wMes As String, wMon As String, _
       wOld As Currency, wNew As Currency, wMesOri As String, _
       wOldOri As Currency, wNewOri As Currency, WE_S As String, zFec As Date
  
   aa = Leerado8("SELECT * FROM CTASXCAB " _
                & " WHERE   SDONEW < 0 AND " _
                & "       CONCEPTO = '01' " _
                & " ORDER BY CODSOCIO, MES ")
   If aa = 0 Then
      MsgBox "No Existen Adelantos Por Asignar", vbExclamation
      Exit Sub
   End If
   
   
   If aa > 0 Then
      ADO8.MoveFirst
      wRegAct = 1
      wRegTot = aa
      Do While Not ADO8.EOF
         DoEvents
         lblMensaje.Caption = "Registro " + _
                              Trim(Format(wRegAct, "##,##0")) + " / " + _
                              Trim(Format(wRegTot, "##,##0"))
         lblMensaje.Refresh
         
         wSoc = ADO8!codsocio
         
         wMesOri = ADO8!mes
         wMon = ADO8!moneda
         wAde = -ADO8!sdonew
         wSdo = -ADO8!sdonew
         wOldOri = ADO8!sdonew
         wNewOri = 0
         WE_S = ADO8!e_socio
      
         If wSdo > 0 Then
            aa = Leerado7("SELECT * FROM CTASXCAB " _
                       & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                       & "       CONCEPTO = '01' AND " _
                       & "       SDONEW > 0 " _
                       & " ORDER BY MES ")
            If aa > 0 Then
               ADO7.MoveFirst
               Do While Not ADO7.EOF
                  wMes = ADO7!mes
                  If ADO7!sdonew >= wSdo Then
                     wAbo = wSdo
                  Else
                     wAbo = ADO7!sdonew
                  End If
                  wOld = ADO7!sdonew
                  wNew = wOld - wAbo
               
                  Db.BeginTrans
                  Db.Execute ("UPDATE CTASXCAB " _
                  & " SET ABONOS = ABONOS + " + Str(wAbo) + ", " _
                  & "     SDONEW = SDONEW - " + Str(wAbo) + " " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       CONCEPTO = '01' AND " _
                  & "            MES = '" + wMes + "' ")
                  Db.CommitTrans
               
                  Db.BeginTrans
                  Db.Execute ("INSERT INTO CTASXDET " _
                  & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
                  & "  TIPCAM, DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW, OBS ) " _
                  & " VALUES " _
                  & " (" + Str(wSoc) + ", '" + wMes + "', '01', '04', '', '', '', '2', " _
                  & "  '" + Format(Date, "dd/mm/yyyy") + "', " _
                  & "  0, " + Str(IIf(wMon = "S", 0, wAbo)) + ", " + Str(IIf(wMon = "S", wAbo, 0)) + ", " _
                  & "  " + Str(wOld) + ", 0, " + Str(wAbo) + ", " + Str(wNew) + ", '' )  ")
                  Db.CommitTrans
               
                  Call ActualizaSaldos(wSoc, wMes, "01")
               
                  wSdo = wSdo - wAbo
                  If wSdo = 0 Then
                     Exit Do
                  End If
                  ADO7.MoveNext
               Loop
            End If
               
            If wSdo > 0 Then
               Do While wSdo > 0
               
                  wMes = CreaAporte(wSoc, 1)
                  wAbo = 0: wMon = "": wOld = 0
                  aa = Leerado6a("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + WE_S + "' ")
                  If aa > 0 Then
                     wAbo = ADO6a!aporte
                     wMon = ADO6a!moneda
                  End If
                  wOld = wAbo
                  If wAbo > wSdo Then
                     wAbo = wSdo
                  End If
                  wNew = wOld - wAbo
                  zFec = Format("01/" + Mid(wMes, 6, 2) + "/" + Mid(wMes, 1, 4), "dd/mm/yyyy")
                  
                  Db.BeginTrans
                  Db.Execute ("UPDATE CTASXCAB " _
                  & " SET ABONOS = ABONOS + " + Str(wAbo) + ", " _
                  & "     SDONEW = SDONEW - " + Str(wAbo) + " " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       CONCEPTO = '01' AND " _
                  & "            MES = '" + wMes + "' ")
                  Db.CommitTrans
               
                  Db.BeginTrans
                  Db.Execute ("INSERT INTO CTASXDET " _
                  & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
                  & "  TIPCAM, DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW, OBS ) " _
                  & " VALUES " _
                  & " (" + Str(wSoc) + ", '" + wMes + "', '01', '04', '', '', '', '2', " _
                  & "  '" + Format(Date, "dd/mm/yyyy") + "', " _
                  & "  0, " + Str(IIf(wMon = "S", 0, wAbo)) + ", " + Str(IIf(wMon = "S", wAbo, 0)) + ", " _
                  & "  " + Str(wOld) + ", 0, " + Str(wAbo) + ", " + Str(wNew) + ", '' )  ")
                  Db.CommitTrans
               
                  Call ActualizaSaldos(wSoc, wMes, "01")
               
                  wSdo = wSdo - wAbo
               Loop
            End If
               
            Db.BeginTrans
            Db.Execute ("UPDATE CTASXCAB " _
            & " SET CARGOS = CARGOS + " + Str(wAde) + ", " _
            & "     SDONEW = SDONEW + " + Str(wAde) + " " _
            & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
            & "       CONCEPTO = '01' AND " _
            & "            MES = '" + wMesOri + "' ")
            Db.CommitTrans
            
            Db.BeginTrans
            Db.Execute ("INSERT INTO CTASXDET " _
            & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
            & "  TIPCAM, DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW, OBS ) " _
            & " VALUES " _
            & " (" + Str(wSoc) + ", '" + wMesOri + "', '01', '04', '', '', '', '2', " _
            & "  '" + Format(Date, "dd/mm/yyyy") + "', " _
            & "  0, " + Str(IIf(wMon = "S", 0, wAde)) + ", " + Str(IIf(wMon = "S", wAde, 0)) + ", " _
            & "  " + Str(wOldOri) + ", " + Str(wAde) + ", 0, " + Str(wNewOri) + ", '' )  ")
            Db.CommitTrans
         
            Call ActualizaSaldos(wSoc, wMesOri, "01")
         
         End If
               
         wRegAct = wRegAct + 1
         ADO8.MoveNext
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
   frmServAdel.Left = (Screen.Width - Width) \ 2
   frmServAdel.Top = 0
   
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
   cmbCia.Enabled = False

   cmdAsignar.SetFocus
End Sub

