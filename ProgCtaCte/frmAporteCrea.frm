VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAporteCrea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Aporte Mensuales"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   8220
   Begin VB.CommandButton cmdCrear 
      Caption         =   "&Crear"
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
      Left            =   5520
      TabIndex        =   2
      Top             =   3000
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
      Left            =   6840
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin MSMask.MaskEdBox txtDesde 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtHasta 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblHasta 
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
      Left            =   2760
      TabIndex        =   8
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblDesde 
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
      Left            =   2760
      TabIndex        =   7
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Mes Final"
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
      Left            =   720
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblMensaje 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   7095
   End
   Begin VB.Label Label5 
      Caption         =   "Mes Inicial"
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
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmAporteCrea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCrear_Click()


   Dim wswAporte As Boolean, wswDsctos As Boolean, _
       wFec As String, wMes As String, wmmm As String, wdes As String, whas As String, a As Long, II As Long, _
       wSoc As Integer, WE_S As String, wMon As String, wApo As Currency, _
       wRegAct As Long, wRegTot As Long
   
   wswAporte = False
   wswDsctos = False
   wdes = txtDesde.Text
   whas = txtHasta.Text
   
   a = Leerado7("SELECT * FROM CTASXCAB WHERE MES >= '" + wdes + "' AND MES <= '" + whas + "' AND ABONOS = 0 ")
   If a > 0 Then
      wswAporte = True
   End If
   
   If wswAporte = True Then
      If MsgBox("Desea Volver a Crear Las Cuotas?", vbYesNo) = vbYes Then
         
         DoEvents
         lblMensaje.Caption = "Borrando Aportes Anteriores"
         lblMensaje.Refresh
         
         a = Leerado8("SELECT * FROM CTASXCAB " _
                  & " WHERE MES >= '" + Left(wdes, 4) + "/" + Right(wdes, 2) + "' AND " _
                  & "       MES <= '" + Left(whas, 4) + "/" + Right(whas, 2) + "' AND " _
                  & "       CONCEPTO = '01' AND " _
                  & "       ABONOS = 0 ")
         If a > 0 Then
            ADO8.MoveFirst
            wRegAct = 1
            wRegTot = a
            Do While Not ADO8.EOF
                DoEvents
                lblMensaje.Caption = "Borrando Aportes Anteriores " + _
                                  Trim(Format(wRegAct, "##,##0")) + " / " + _
                                  Trim(Format(wRegTot, "##,##0"))
                lblMensaje.Refresh
               
               wSoc = ADO8!codsocio
               wMes = ADO8!mes
         
               Db.BeginTrans
               Db.Execute ("DELETE FROM CTASXCAB " _
               & " WHERE CODSOCIO =  " + Str(wSoc) + " AND " _
               & "            MES = '" + wMes + "' AND " _
               & "       CONCEPTO = '01' ")
               Db.CommitTrans
         
               Db.BeginTrans
               Db.Execute ("DELETE FROM CTASXDET " _
               & " WHERE CODSOCIO =  " + Str(wSoc) + " AND " _
               & "            MES = '" + wMes + "' AND " _
               & "       CONCEPTO = '01' ")
               Db.CommitTrans
         
               wRegAct = wRegAct + 1
               ADO8.MoveNext
            Loop
         End If
         
      Else
         Unload Me
         Exit Sub
      End If
   End If
   
   For II = Val(Left(wdes, 4) + Right(wdes, 2)) To Val(Left(whas, 4) + Right(whas, 2))
       wMes = Format(II, "0000/00")
       If Right(wMes, 2) > "12" Then
          wMes = Format(Val(Left(wMes, 4)) + 1, "0000") + "/" + "01"
          II = Val(Left(wMes, 4) + Right(wMes, 2))
       End If
   
       wFec = Format("01/" + Right(wMes, 2) + "/" + Left(wMes, 4), "dd/mm/yyyy")
   
       a = Leerado8("SELECT * FROM MAEMESES " _
                    & " WHERE ANO = '" + Left(wMes, 4) + "' AND " _
                    & "       MES = '" + Right(wMes, 2) + "' ")
       If a = 0 Then
          Db.BeginTrans
          Db.Execute ("INSERT INTO MAEMESES " _
          & " (ANO, MES, USUARIO, USUFECH, USUHORA) " _
          & " VALUES " _
          & " ('" + Left(wMes, 4) + "', '" + Right(wMes, 2) + "', " _
          & "  '" + wcodusu + "', '" + Format(wFec, "dd/mm/yyyy") + "', " _
          & "  '" + Format("00:00:00", "hh:mm:ss") + "' ) ")
          Db.CommitTrans
       Else
          Db.BeginTrans
          Db.Execute ("UPDATE MAEMESES " _
          & " SET USUARIO = '" + wcodusu + "', " _
          & "     USUFECH = '" + Format(wFec, "dd/mm/yyyy") + "', " _
          & "     USUHORA = '" + Format("00:00:00", "hh:mm:ss") + "' " _
          & " WHERE ANO = '" + Left(wMes, 4) + "' AND " _
          & "       MES = '" + Right(wMes, 2) + "' ")
          Db.CommitTrans
       End If
   
       a = Leerado8("SELECT M.CODSOCIO, M.E_SOCIO, E.MONEDA, E.APORTE " _
                    & " FROM MAESOCIO AS M INNER JOIN MAEE_SOCIO AS E " _
                    & "   ON M.E_SOCIO = E.E_SOCIO " _
                    & " WHERE E.APORTE > 0 ")
       If a > 0 Then
          ADO8.MoveFirst
          wRegAct = 1
          wRegTot = a
          Do While Not ADO8.EOF
             DoEvents
             lblMensaje.Caption = "Creando Aportes Mes " + _
                                  Left(Trim(funnommes(Mid(wMes, 6, 2))), 3) + "-" + _
                                  Mid(wMes, 1, 4) + " Registro " + _
                                  Trim(Format(wRegAct, "##,##0")) + " / " + _
                                  Trim(Format(wRegTot, "##,##0"))
             lblMensaje.Refresh
             
             wSoc = ADO8!codsocio
             WE_S = ADO8!e_socio
             wMon = ADO8!moneda
             wApo = ADO8!aporte
       
             a = Leerado6a("SELECT * FROM CTASXCAB " _
                        & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                        & "            MES = '" + wMes + "' AND " _
                        & "       CONCEPTO = '01' ")
             If a = 0 Then
                Db.BeginTrans
                Db.Execute ("INSERT INTO CTASXCAB " _
                & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, CARGOS, ABONOS, SDONEW ) " _
                & " VALUES " _
                & " (" + Str(wSoc) + ", '" + wMes + "', '01', '" + WE_S + "', '" + wMon + "', " _
                & "  " + Str(wApo) + ", 0, " + Str(wApo) + " ) ")
                Db.CommitTrans
             End If
                
             a = Leerado6a("SELECT * FROM CTASXDET " _
                        & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                        & "            MES = '" + wMes + "' AND " _
                        & "       CONCEPTO = '01' ")
             If a = 0 Then
                Db.BeginTrans
                Db.Execute ("INSERT INTO CTASXDET " _
                & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
                & "  DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW) " _
                & " VALUES " _
                & " (" + Str(wSoc) + ", '" + wMes + "', '01', '00', '', '', '', '1', " _
                & "  '" + Format(wFec, "dd/mm/yyyy") + "', 0, 0, 0, " + Str(wApo) + ", " _
                & "  0, " + Str(wApo) + " ) ")
                Db.CommitTrans
             End If
       
             wRegAct = wRegAct + 1
             ADO8.MoveNext
          Loop
       End If
   
   Next II
   
   lblMensaje.Caption = ""
   lblMensaje.Refresh
   
   MsgBox "Cuotas Aportaciones Creadas OK", vbExclamation
   Unload Me
   Exit Sub
   
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   frmAporteCrea.Left = (Screen.Width - Width) \ 2
   frmAporteCrea.Top = 0
   
   Dim a As Integer, wMesCab As String

   wMesCab = wanocia + wmescia
   a = Leerado7("SELECT MAX(MES) AS MES From CTASXCAB ")
   If a > 0 Then
      wMesCab = IIf(IsNull(ADO7!mes), wanocia + wmescia, ADO7!mes)
   End If
   If Right(wMesCab, 2) = "12" Then
      wMesCab = Format(Val(Left(wMesCab, 4)) + 1, "0000") + "01"
   Else
      wMesCab = Left(wMesCab, 4) + Format(Val(Right(wMesCab, 2)) + 1, "00")
   End If

   txtDesde.Text = Format(wMesCab, "0000/00")
   txtHasta.Text = Format(wMesCab, "0000/00")
   
   txtDesde.SetFocus
End Sub

Private Sub txtDesde_Change()
   Dim wMes As String, wAno As String
   If txtDesde.Text <> "____-__" Then
      wAno = Left(txtDesde.Text, 4)
      wMes = Right(txtDesde.Text, 2)
               
      If wMes <> "01" And wMes <> "02" And wMes <> "03" And wMes <> "04" And wMes <> "05" And wMes <> "06" And _
         wMes <> "07" And wMes <> "08" And wMes <> "09" And wMes <> "10" And wMes <> "11" And wMes <> "12" Then
         lblDesde.Caption = ""
      Else
         lblDesde.Caption = Trim(funnommes(wMes)) + " " + wAno
      End If
   Else
      lblDesde.Caption = ""
   End If
End Sub

Private Sub txtDesde_GotFocus()
   txtDesde.SelStart = 0
   txtDesde.SelLength = 7
End Sub

Private Sub txtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 40
        txtHasta.SetFocus
   End Select
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
   Dim wMes As String, wAno As String
   If KeyAscii = 13 Then
      If txtDesde.Text = "____-__" Then
         MsgBox "Mes Inicial En Blanco", vbExclamation
         txtDesde.Text = "____-__"
         txtDesde.SetFocus
         Exit Sub
      End If
      wAno = Left(txtDesde.Text, 4)
      wMes = Right(txtDesde.Text, 2)
               
      If wMes <> "01" And wMes <> "02" And wMes <> "03" And wMes <> "04" And wMes <> "05" And wMes <> "06" And _
         wMes <> "07" And wMes <> "08" And wMes <> "09" And wMes <> "10" And wMes <> "11" And wMes <> "12" Then
         MsgBox "Mes Digitado Es Invalido", vbExclamation
         txtDesde.Text = "____-__"
         txtDesde.SetFocus
         Exit Sub
      End If
      If wAno < "2004" Or wAno > "2030" Then
         MsgBox "Año Digitado Es Invalido", vbExclamation
         txtDesde.Text = "____-__"
         txtDesde.SetFocus
         Exit Sub
      End If
      txtHasta.SetFocus
   End If
End Sub

Private Sub txtHasta_Change()
   Dim wMes As String, wAno As String
   If txtHasta.Text <> "____-__" Then
      wAno = Left(txtHasta.Text, 4)
      wMes = Right(txtHasta.Text, 2)
               
      If wMes <> "01" And wMes <> "02" And wMes <> "03" And wMes <> "04" And _
         wMes <> "05" And wMes <> "06" And wMes <> "07" And wMes <> "08" And _
         wMes <> "09" And wMes <> "10" And wMes <> "11" And wMes <> "12" Then
         lblHasta.Caption = ""
      Else
         lblHasta.Caption = Trim(funnommes(wMes)) + " " + wAno
      End If
   Else
      lblHasta.Caption = ""
   End If
End Sub

Private Sub txtHasta_GotFocus()
   txtHasta.SelStart = 0
   txtHasta.SelLength = 7
End Sub

Private Sub txtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDesde.SetFocus
   End Select
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
   Dim wMes As String, wAno As String
   If KeyAscii = 13 Then
      If txtHasta.Text = "____-__" Then
         MsgBox "Mes Final En Blanco", vbExclamation
         txtHasta.Text = "____-__"
         txtHasta.SetFocus
         Exit Sub
      End If
      wAno = Left(txtHasta.Text, 4)
      wMes = Right(txtHasta.Text, 2)
               
      If wMes <> "01" And wMes <> "02" And wMes <> "03" And wMes <> "04" And _
         wMes <> "05" And wMes <> "06" And wMes <> "07" And wMes <> "08" And _
         wMes <> "09" And wMes <> "10" And wMes <> "11" And wMes <> "12" Then
         MsgBox "Mes Digitado Es Invalido", vbExclamation
         txtHasta.Text = "____-__"
         txtHasta.SetFocus
         Exit Sub
      End If
      If wAno < "2004" Or wAno > "2030" Then
         MsgBox "Año Digitado Es Invalido", vbExclamation
         txtHasta.Text = "____-__"
         txtHasta.SetFocus
         Exit Sub
      End If
      cmdCrear.SetFocus
   End If
End Sub
