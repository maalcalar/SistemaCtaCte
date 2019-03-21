VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCtaCreaCuota 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Provisiones Mensuales"
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
      Mask            =   "####-##"
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
      Mask            =   "####-##"
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
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1920
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
Attribute VB_Name = "frmCtaCreaCuota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCrear_Click()

   Dim wFec As String, wMes As String, wdes As String, whas As String, a As Long, II As Long
   wdes = Left(txtDesde.Text, 4) + Right(txtDesde.Text, 2)
   whas = Left(txtHasta.Text, 4) + Right(txtHasta.Text, 2)
   
   a = Leerado7("SELECT * FROM CTASXCAB WHERE MES >= '" + wdes + "' AND MES <= '" + whas + "' ")
   If a > 0 Then
      If MsgBox("Desea Volver a Crear Las Cuotas?", vbYesNo) = vbYes Then
         Db.BeginTrans
         Db.Execute ("DELETE FROM CTASXCAB " _
         & " WHERE MES >= '" + wdes + "' AND " _
         & "       MES <= '" + whas + "' ")
         Db.CommitTrans
      
         Db.BeginTrans
         Db.Execute ("DELETE FROM CTASXDET " _
         & " WHERE MES >= '" + wdes + "' AND " _
         & "       MES <= '" + whas + "' AND " _
         & "       LINEA = '01' ")
         Db.CommitTrans
      
         Db.BeginTrans
         Db.Execute ("DELETE FROM MESES " _
         & "  WHERE MES >= '" + wdes + "' AND " _
         & "        MES <= '" + whas + "' ")
         Db.CommitTrans
      Else
         Unload Me
         Exit Sub
      End If
   End If
   
   For II = Val(wdes) To Val(whas)
       wMes = Format(II, "000000")
       If Right(wMes, 2) > "12" Then
          wMes = Format(Val(Left(wMes, 4)) + 1, "0000") + "01"
          II = Val(wMes)
       End If
       wFec = Format("01/" + Right(wMes, 2) + "/" + Left(wMes, 4))
   
       DoEvents
       lblMensaje.Caption = "Creando Mes " + Left(wMes, 4) + "-" + Right(wMes, 2)
       lblMensaje.Refresh
       
       Db.BeginTrans
       Db.Execute ("INSERT INTO CTASXCAB " _
       & " (LOCAL1, MES) " _
       & " SELECT " _
       & "  LOCAL1, '" + wMes + "' " _
       & " FROM LOCALES ")
       Db.CommitTrans
   
       Db.BeginTrans
       Db.Execute ("INSERT INTO CTASXDET " _
       & " (LOCAL1, MES, LINEA, TIPCOB, NUMCOB, LINCOB, FECHA, FORPAG, DOLARE, SOLESS) " _
       & " SELECT " _
       & "  LOCAL1, '" + wMes + "', '01', '', '', '', '" + wFec + "', '', 0, 0  " _
       & " FROM LOCALES ")
       Db.CommitTrans
   
       Db.BeginTrans
       Db.Execute ("UPDATE CTASXCAB " _
       & " SET MANCARGOS = LOCALES.CUOMAN, " _
       & "     MANABONOS = 0, " _
       & "     MANSDONEW = LOCALES.CUOMAN " _
       & " FROM CTASXCAB INNER JOIN LOCALES " _
       & "   ON CTASXCAB.LOCAL1 = LOCALES.LOCAL1 " _
       & " WHERE LOCALES.CUOMAN <> 0 AND " _
       & "         CTASXCAB.MES = '" + wMes + "' ")
       Db.CommitTrans
   
'       Db.BeginTrans
'       Db.Execute ("UPDATE CTASXCAB INNER JOIN LOCALES " _
'       & "              ON CTASXCAB.LOCAL1 = LOCALES.LOCAL1 " _
'       & " SET CTASXCAB.LUZCARGOS = LOCALES.CUOLUZ, " _
'       & "     CTASXCAB.LUZABONOS = 0, " _
'       & "     CTASXCAB.LUZSDONEW = LOCALES.CUOLUZ " _
'       & " WHERE LOCALES.CUOLUZ <> 0 AND " _
'       & "         CTASXCAB.MES = '" + wmes + "' ")
'       Db.CommitTrans
   
'       Db.BeginTrans
'       Db.Execute ("UPDATE CTASXCAB INNER JOIN LOCALES " _
'       & "              ON CTASXCAB.LOCAL1 = LOCALES.LOCAL1 " _
'       & " SET CTASXCAB.PUBCARGOS = LOCALES.CUOPUB, " _
'       & "     CTASXCAB.PUBABONOS = 0, " _
'       & "     CTASXCAB.PUBSDONEW = LOCALES.CUOPUB " _
'       & " WHERE LOCALES.CUOPUB <> 0 AND " _
'       & "         CTASXCAB.MES = '" + wmes + "' ")
'       Db.CommitTrans
   
       Db.BeginTrans
       Db.Execute ("UPDATE CTASXDET " _
       & " SET MANSDOOLD = 0, " _
       & "     MANCARGOS = LOCALES.CUOMAN, " _
       & "     MANABONOS = 0, " _
       & "     MANSDONEW = LOCALES.CUOMAN " _
       & " FROM CTASXDET INNER JOIN LOCALES " _
       & "   ON CTASXDET.LOCAL1 = LOCALES.LOCAL1 " _
       & " WHERE LOCALES.CUOMAN <> 0 AND " _
       & "         CTASXDET.MES = '" + wMes + "' AND " _
       & "       CTASXDET.LINEA = '01' ")
       Db.CommitTrans
   
'       Db.BeginTrans
'       Db.Execute ("UPDATE CTASXDET INNER JOIN LOCALES " _
'       & "              ON CTASXDET.LOCAL1 = LOCALES.LOCAL1 " _
'       & " SET CTASXDET.LUZSDOOLD = 0, " _
'       & "     CTASXDET.LUZCARGOS = LOCALES.CUOLUZ, " _
'       & "     CTASXDET.LUZABONOS = 0, " _
'       & "     CTASXDET.LUZSDONEW = LOCALES.CUOLUZ " _
'       & " WHERE LOCALES.CUOLUZ <> 0 AND " _
'       & "         CTASXDET.MES = '" + wmes + "' AND " _
'       & "       CTASXDET.LINEA = '01' ")
'       Db.CommitTrans
   
   
'       Db.BeginTrans
'       Db.Execute ("UPDATE CTASXDET INNER JOIN LOCALES " _
'       & "              ON CTASXDET.LOCAL1 = LOCALES.LOCAL1 " _
'       & " SET CTASXDET.PUBSDOOLD = 0, " _
'       & "     CTASXDET.PUBCARGOS = LOCALES.CUOPUB, " _
'       & "     CTASXDET.PUBABONOS = 0, " _
'       & "     CTASXDET.PUBSDONEW = LOCALES.CUOPUB " _
'       & " WHERE LOCALES.CUOPUB <> 0 AND " _
'       & "         CTASXDET.MES = '" + wmes + "' AND " _
'       & "       CTASXDET.LINEA = '01' ")
'       Db.CommitTrans
   
       Db.BeginTrans
       Db.Execute ("INSERT INTO MESES " _
       & " (MES) " _
       & " VALUES " _
       & " ('" + wMes + "') ")
       Db.CommitTrans
   
   Next
   
   MsgBox "Cuota Creadas OK", vbExclamation
   Unload Me
   Exit Sub
   
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   frmCtaCreaCuota.Left = (Screen.Width - Width) \ 2
   frmCtaCreaCuota.Top = 0
   
   Dim a As Integer, wMes As String

   wMes = "200408"
   a = Leerado7("SELECT MAX(MES) AS MES From CTASXCAB ")
   If a > 0 Then
      wMes = IIf(IsNull(ADO7!mes), "200408", ADO7!mes)
   End If
   If Right(wMes, 2) = "12" Then
      wMes = Format(Val(Left(wMes, 4)) + 1, "0000") + "01"
   Else
      wMes = Left(wMes, 4) + Format(Val(Right(wMes, 2)) + 1, "00")
   End If

   txtDesde.Text = Format(wMes, "0000-00")
   txtHasta.Text = Format(wMes, "0000-00")
   
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
