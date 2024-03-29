VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDiecoRecibe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibir Descuento DIECO"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   13725
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtro"
      Height          =   855
      Left            =   0
      TabIndex        =   35
      Top             =   7320
      Width           =   6495
      Begin VB.TextBox txtFiltrar 
         Height          =   285
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   38
         Top             =   500
         Width           =   4095
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Filtrar x Nombre"
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   500
         Width           =   1575
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Mostrar Todos"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   200
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Modificar RECIBE Un Socio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdRevision 
      Caption         =   "Revisi�n"
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
      Left            =   12240
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdExtorna 
      Caption         =   "Extornar Descuento"
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
      Left            =   8760
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7560
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
      Left            =   12240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Actualiza Descuento"
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5055
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8916
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin VB.CommandButton cmdRecibir 
      Caption         =   "Recibir Archivo"
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox cmbMeses 
      Height          =   315
      ItemData        =   "frmDiecoRecibe.frx":0000
      Left            =   960
      List            =   "frmDiecoRecibe.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txtAnoCab 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin MSMask.MaskEdBox txtFecDsc 
      Height          =   285
      Left            =   3600
      TabIndex        =   31
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label14 
      Caption         =   "Fecha Proceso"
      Height          =   210
      Left            =   3600
      TabIndex        =   32
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label lblCanApo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8160
      TabIndex        =   30
      Top             =   660
      Width           =   1095
   End
   Begin VB.Label lblEnviado 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9360
      TabIndex        =   29
      Top             =   660
      Width           =   1095
   End
   Begin VB.Label lblCanAsi 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8160
      TabIndex        =   28
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Cant.Titulares"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   8160
      TabIndex        =   27
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Cant.Asignados"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   8160
      TabIndex        =   26
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Total Envio S/."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   9360
      TabIndex        =   25
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblNoDscto 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   10560
      TabIndex        =   24
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "No Cobrado"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   10560
      TabIndex        =   23
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblRecibido 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9360
      TabIndex        =   22
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Cobrado"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9360
      TabIndex        =   21
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RECIBIR DESCUENTO DE DIECO"
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
      Height          =   375
      Left            =   3600
      TabIndex        =   20
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Asig 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblAsig1 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   6480
      Width           =   4215
   End
   Begin VB.Label lblAsig2 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   6720
      Width           =   4215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Asig 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label lblAsig3 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   6960
      Width           =   4215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Asig 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label lblAsig4 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6000
      TabIndex        =   12
      Top             =   6480
      Width           =   4215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Asig 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   5280
      TabIndex        =   11
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblAsig5 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6000
      TabIndex        =   10
      Top             =   6720
      Width           =   4215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Asig 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   6720
      Width           =   735
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
      Left            =   6840
      TabIndex        =   8
      Top             =   7200
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "Mes"
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
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label25 
      Caption         =   "A�o"
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
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "frmDiecoRecibe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Limpiar()
   lblCanApo.Caption = ""
   lblCanAsi.Caption = ""
   lblEnviado.Caption = ""
   lblRecibido.Caption = ""
   lblNoDscto.Caption = ""
End Sub

Private Sub cmbMeses_Click()
   cmbMeses_KeyPress (13)
End Sub

Private Sub cmbMeses_KeyPress(KeyAscii As Integer)
   Dim aa As Integer, bb As Integer, wAno As String, wMes As String
   wAno = txtAnoCab.Text
   wMes = Left(cmbMeses.Text, 2)
   Limpiar
   txtFecDsc.Text = Format("25/" + wMes + "/" + wAno, "dd/mm/yyyy")
   
   If KeyAscii = 13 Then
      LlenaCab
      LlenaCab1
      TotalCab
   End If
End Sub

Private Sub cmdDetalle_Click()
   zDetaCambio = True
   zDetaCodSoc = ADO2!codsocio
   zDetaTipDsc = "01"
   zDetaAnoDsc = txtAnoCab.Text
   zDetaMesDsc = Left(cmbMeses.Text, 2)
   zDetaSw = False

   frmDIECOModif.Show vbModal

   If zDetaSw = True Then
      
      ADO2.Requery
      LlenaCab
      LlenaCab1
'      LabelCab
      TotalCab
      ADO2.Find "CODSOCIO=" + Str(zDetaCodSoc) + ""
   End If
End Sub

Private Sub cmdExtorna_Click()
   Dim wAno As String, wMes As String, wFec As Date

   wAno = txtAnoCab.Text
   wMes = Left(cmbMeses.Text, 2)
   wFec = Format(fundiames(wMes) + "/" + wMes + "/" + wAno, "dd/mm/yyyy")
   
   lblMensaje.Caption = "Borrando Proceso Anterior..."
   lblMensaje.Refresh
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM CTASXDET " _
   & " WHERE      MES = '" + wAno + "/" + wMes + "' AND " _
   & "       CONCEPTO = '01' AND " _
   & "         TIPMOV = '2' AND " _
   & "         TIPCOB = '01' AND " _
   & "         SERCOB = '001' AND " _
   & "         NUMCOB = '" + Right(wAno, 2) + wMes + "00001' AND " _
   & "         LINCOB = '0001' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE DIECOCAB " _
   & " SET DSCSOCIO = 0, DIFSOCIO = 0, " _
   & "     DSCASIG1 = 0, DIFASIG1 = 0, " _
   & "     DSCASIG2 = 0, DIFASIG2 = 0, " _
   & "     DSCASIG3 = 0, DIFASIG3 = 0, " _
   & "     DSCASIG4 = 0, DIFASIG4 = 0, " _
   & "     DSCASIG5 = 0, DIFASIG5 = 0, " _
   & "     DSCDIECO = 0, DSCDIFER = 0 " _
   & " WHERE MES = '" + wAno + wMes + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET DSCSOCIO = 0, DIFSOCIO = 0, " _
   & "     DSCASIG1 = 0, DIFASIG1 = 0, " _
   & "     DSCASIG2 = 0, DIFASIG2 = 0, " _
   & "     DSCASIG3 = 0, DIFASIG3 = 0, " _
   & "     DSCASIG4 = 0, DIFASIG4 = 0, " _
   & "     DSCASIG5 = 0, DIFASIG5 = 0, " _
   & "     DSCDIECO = 0, DSCDIFER = 0 " _
   & " WHERE MES = '" + wAno + wMes + "' AND " _
   & "       USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_APOR_PLA " _
   & " SET IMPO" + wMes + " = 0 " _
   & " WHERE  CUOANO = '" + wAno + "' AND " _
   & "       TIPAPOR = '1' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_APOR_PLA " _
   & " SET TOTIMPO = IMPO01 + IMPO02 + IMPO03 + IMPO04 + IMPO05 + IMPO06 + " _
   & "               IMPO07 + IMPO08 + IMPO09 + IMPO10 + IMPO11 + IMPO12  " _
   & " WHERE  CUOANO = '" + wAno + "' AND " _
   & "       TIPAPOR = '1' ")
   Db.CommitTrans
   
   lblMensaje.Caption = ""
   lblMensaje.Refresh

'   Limpiar
'   LlenaCab
'   LlenaCab1
'   TotalCab
   
   MsgBox "Descuento DIECO - MES " + wAno + "-" + wMes + " Extornado OK", vbExclamation
   Unload Me
End Sub

Private Sub cmdGrabar_Click()
   Dim zz As Integer, wRegAct As Integer, wRegTot As Integer, wAno As String, wMes As String, _
      wSoc As Integer, wCod As Long, wIns As Integer, wNom As String, wSit As Integer, wEsp As Integer, _
      wDscDieco As Currency, wTotEnvio As Currency, wDscDifer As Currency, wSdoxDist As Currency, _
      wDscSocio As Currency, wDscAsig1 As Currency, wDscAsig2 As Currency, _
      wDscAsig3 As Currency, wDscAsig4 As Currency, wDscAsig5 As Currency, _
      wDifSocio As Currency, wDifAsig1 As Currency, wDifAsig2 As Currency, _
      wDifAsig3 As Currency, wDifAsig4 As Currency, wDifAsig5 As Currency, _
      wNetSocio As Currency, wNetAsig1 As Currency, wNetAsig2 As Currency, _
      wNetAsig3 As Currency, wNetAsig4 As Currency, wNetAsig5 As Currency, _
      wCodAsig1 As Long, wCodAsig2 As Long, wCodAsig3 As Long, wCodAsig4 As Long, wCodAsig5 As Long, _
      wInsAsig1 As Integer, wInsAsig2 As Integer, wInsAsig3 As Integer, wInsAsig4 As Integer, wInsAsig5 As Integer, _
      wSocAsig1 As Integer, wSocAsig2 As Integer, wSocAsig3 As Integer, wSocAsig4 As Integer, wSocAsig5 As Integer, _
      wTotAsig1 As Currency, wTotAsig2 As Currency, wTotAsig3 As Currency, wTotAsig4 As Currency, wTotAsig5 As Currency, _
      wNomAsig1 As String, wNomAsig2 As String, wNomAsig3 As String, wNomAsig4 As String, wNomAsig5 As String, _
      wqqq As Variant, wFec As Date, wApo As Currency
   
   wAno = txtAnoCab.Text
   wMes = Left(cmbMeses.Text, 2)
   wFec = Format(fundiames(wMes) + "/" + wMes + "/" + wAno, "dd/mm/yyyy")
   
   lblMensaje.Caption = "Borrando Proceso Anterior..."
   lblMensaje.Refresh
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM CTASXDET " _
   & " WHERE      MES = '" + wAno + "/" + wMes + "' AND " _
   & "       CONCEPTO = '01' AND " _
   & "         TIPMOV = '2' AND " _
   & "         TIPCOB = '01' AND " _
   & "         SERCOB = '001' AND " _
   & "         NUMCOB = '" + Right(wAno, 2) + wMes + "00001' AND " _
   & "         LINCOB = '0001' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE DIECOCAB " _
   & " SET DSCSOCIO = 0, DIFSOCIO = 0, " _
   & "     DSCASIG1 = 0, DIFASIG1 = 0, " _
   & "     DSCASIG2 = 0, DIFASIG2 = 0, " _
   & "     DSCASIG3 = 0, DIFASIG3 = 0, " _
   & "     DSCASIG4 = 0, DIFASIG4 = 0, " _
   & "     DSCASIG5 = 0, DIFASIG5 = 0 " _
   & " WHERE MES = '" + wAno + wMes + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET DSCSOCIO = 0, DIFSOCIO = 0, " _
   & "     DSCASIG1 = 0, DIFASIG1 = 0, " _
   & "     DSCASIG2 = 0, DIFASIG2 = 0, " _
   & "     DSCASIG3 = 0, DIFASIG3 = 0, " _
   & "     DSCASIG4 = 0, DIFASIG4 = 0, " _
   & "     DSCASIG5 = 0, DIFASIG5 = 0 " _
   & " WHERE MES = '" + wAno + wMes + "' AND " _
   & "       USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_APOR_PLA " _
   & " SET IMPO" + wMes + " = 0 " _
   & " WHERE  CUOANO = '" + wAno + "' AND " _
   & "       TIPAPOR = '1' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_APOR_PLA " _
   & " SET TOTIMPO = IMPO01 + IMPO02 + IMPO03 + IMPO04 + IMPO05 + IMPO06 + " _
   & "               IMPO07 + IMPO08 + IMPO09 + IMPO10 + IMPO11 + IMPO12  " _
   & " WHERE  CUOANO = '" + wAno + "' AND " _
   & "       TIPAPOR = '1' ")
   Db.CommitTrans
   
   wApo = 0
   zz = Leerado8("SELECT * FROM MAEE_SOCIO " _
                & " WHERE E_SOCIO = 'TIT' ")
   If zz > 0 Then
      wApo = ADO8!aporte
   End If
   Set ADO8 = Nothing
   
   lblMensaje.Caption = "Actualizando Saldos"
   lblMensaje.Refresh
   
   zz = Leerado8("SELECT * FROM TMP_DIECOCAB " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "       MES = '" + wAno + wMes + "' " _
                & " ORDER BY NOMBRE ")
   If zz > 0 Then
      wRegAct = 1
      wRegTot = zz
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         DoEvents
         lblMensaje.Caption = "Actualizando Saldos - Registro " + _
                              Trim(Format(wRegAct, "####0")) + " / " + _
                              Trim(Format(wRegTot, "####0"))
         lblMensaje.Refresh
         
         wSoc = ADO8!codsocio
         wCod = ADO8!codigo
         wIns = ADO8!ins
         wNom = Trim(ADO8!nombre)
         wSit = ADO8!situ
         wEsp = ADO8!situesp
         wTotEnvio = ADO8!totenvio
         wDscDieco = ADO8!dscdieco
         wDscDifer = ADO8!dscdifer
         wSdoxDist = ADO8!dscdieco
         wNetSocio = ADO8!netsocio
         wDscSocio = 0: wDscAsig1 = 0: wDscAsig2 = 0: wDscAsig3 = 0: wDscAsig4 = 0: wDscAsig5 = 0
         wDifSocio = 0: wDifAsig1 = 0: wDifAsig2 = 0: wDifAsig3 = 0: wDifAsig4 = 0: wDifAsig5 = 0
         wSocAsig1 = ADO8!codasig1: wNetAsig1 = ADO8!netasig1
         wSocAsig2 = ADO8!codasig2: wNetAsig2 = ADO8!netasig2
         wSocAsig3 = ADO8!codasig3: wNetAsig3 = ADO8!netasig3
         wSocAsig4 = ADO8!codasig4: wNetAsig4 = ADO8!netasig4
         wSocAsig5 = ADO8!codasig5: wNetAsig5 = ADO8!netasig5
         wCodAsig1 = 0: wCodAsig2 = 0: wCodAsig3 = 0: wCodAsig4 = 0: wCodAsig5 = 0
         wInsAsig1 = 0: wInsAsig2 = 0: wInsAsig3 = 0: wInsAsig4 = 0: wInsAsig5 = 0
         wNomAsig1 = "": wNomAsig2 = "": wNomAsig3 = "": wNomAsig4 = "": wNomAsig5 = ""
     
         Db.BeginTrans
         Db.Execute ("UPDATE MAESOCIO " _
         & " SET SITU = " + Str(wSit) + ", SITUESP = " + Str(wEsp) + " " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " ")
         Db.CommitTrans
     
         ' Primero Todos Los Aportes del Mes y Luego el Resto
         If wNetSocio > 0 And wSdoxDist > 0 Then
            If wSdoxDist >= wApo Then
               If wNetSocio >= wApo Then
                  wDscSocio = wDscSocio + wApo
                  wSdoxDist = wSdoxDist - wApo
               Else
                  wDscSocio = wDscSocio + wNetSocio
                  wSdoxDist = wSdoxDist - wNetSocio
               End If
            Else
               wDscSocio = wDscSocio + wSdoxDist
               wSdoxDist = wSdoxDist - wSdoxDist
            End If
         End If
         If wNetAsig1 > 0 And wSdoxDist > 0 Then
            If wSdoxDist >= wApo Then
               If wNetAsig1 >= wApo Then
                  wDscAsig1 = wDscAsig1 + wApo
                  wSdoxDist = wSdoxDist - wApo
               Else
                  wDscAsig1 = wDscAsig1 + wNetAsig1
                  wSdoxDist = wSdoxDist - wNetAsig1
               End If
            Else
               If wSdoxDist > wNetAsig1 Then
                  wSdoxDist = wSdoxDist + (wNetAsig1 - wDscAsig1)
                  wDscAsig1 = wDscAsig1 + (wNetAsig1 - wDscAsig1)
               Else
                  wDscAsig1 = wDscAsig1 + wSdoxDist
                  wSdoxDist = wSdoxDist - wSdoxDist
               End If
            End If
         End If
         
         If wNetAsig2 > 0 And wSdoxDist > 0 Then
            If wSdoxDist >= wApo Then
               If wNetAsig2 >= wApo Then
                  wDscAsig2 = wDscAsig2 + wApo
                  wSdoxDist = wSdoxDist - wApo
               Else
                  wDscAsig2 = wDscAsig2 + wNetAsig2
                  wSdoxDist = wSdoxDist - wNetAsig2
               End If
            Else
               If wSdoxDist > wNetAsig2 Then
                  wSdoxDist = wSdoxDist + (wNetAsig2 - wDscAsig2)
                  wDscAsig2 = wDscAsig2 + (wNetAsig2 - wDscAsig2)
               Else
                  wDscAsig2 = wDscAsig2 + wSdoxDist
                  wSdoxDist = wSdoxDist - wSdoxDist
               End If
            End If
         End If
         
         If wNetAsig3 > 0 And wSdoxDist > 0 Then
            If wSdoxDist >= wApo Then
               If wNetAsig3 >= wApo Then
                  wDscAsig3 = wDscAsig3 + wApo
                  wSdoxDist = wSdoxDist - wApo
               Else
                  wDscAsig3 = wDscAsig3 + wNetAsig3
                  wSdoxDist = wSdoxDist - wNetAsig3
               End If
            Else
               If wSdoxDist > wNetAsig3 Then
                  wSdoxDist = wSdoxDist + (wNetAsig3 - wDscAsig3)
                  wDscAsig3 = wDscAsig3 + (wNetAsig3 - wDscAsig3)
               Else
                  wDscAsig3 = wDscAsig3 + wSdoxDist
                  wSdoxDist = wSdoxDist - wSdoxDist
               End If
            End If
         End If
         
         If wNetAsig4 > 0 And wSdoxDist > 0 Then
            If wSdoxDist >= wApo Then
               If wNetAsig4 >= wApo Then
                  wDscAsig4 = wDscAsig4 + wApo
                  wSdoxDist = wSdoxDist - wApo
               Else
                  wDscAsig4 = wDscAsig4 + wNetAsig4
                  wSdoxDist = wSdoxDist - wNetAsig4
               End If
            Else
               If wSdoxDist > wNetAsig4 Then
                  wSdoxDist = wSdoxDist + (wNetAsig4 - wDscAsig4)
                  wDscAsig4 = wDscAsig4 + (wNetAsig4 - wDscAsig4)
               Else
                  wDscAsig4 = wDscAsig4 + wSdoxDist
                  wSdoxDist = wSdoxDist - wSdoxDist
               End If
            End If
         End If
         
         If wNetAsig5 > 0 And wSdoxDist > 0 Then
            If wSdoxDist >= wApo Then
               If wNetAsig5 >= wApo Then
                  wDscAsig5 = wDscAsig5 + wApo
                  wSdoxDist = wSdoxDist - wApo
               Else
                  wDscAsig5 = wDscAsig5 + wNetAsig5
                  wSdoxDist = wSdoxDist - wNetAsig5
               End If
            Else
               If wSdoxDist > wNetAsig5 Then
                  wSdoxDist = wSdoxDist + (wNetAsig5 - wDscAsig5)
                  wDscAsig5 = wDscAsig5 + (wNetAsig5 - wDscAsig5)
               Else
                  wDscAsig5 = wDscAsig5 + wSdoxDist
                  wSdoxDist = wSdoxDist - wSdoxDist
               End If
            End If
         End If
         
         If (wNetSocio - wDscSocio) > 0 And wSdoxDist > 0 Then
            If wSdoxDist >= (wNetSocio - wDscSocio) Then
               wSdoxDist = wSdoxDist - (wNetSocio - wDscSocio)
               wDscSocio = wNetSocio
            Else
               If wSdoxDist > (wNetSocio - wDscSocio) Then
                  wSdoxDist = wSdoxDist + (wNetSocio - wDscSocio)
                  wDscSocio = wDscSocio + (wNetSocio - wDscSocio)
               Else
                  wDscSocio = wDscSocio + wSdoxDist
                  wSdoxDist = 0
               End If
            End If
         End If
         
         If (wNetAsig1 - wDscAsig1) > 0 And wSdoxDist > 0 Then
            If wSdoxDist > (wNetAsig1 - wDscAsig1) Then
               wSdoxDist = wSdoxDist - (wNetAsig1 - wDscAsig1)
               wDscAsig1 = wNetAsig1
            Else
               If wSdoxDist > (wNetAsig1 - wDscAsig1) Then
                  wSdoxDist = wSdoxDist + (wNetAsig1 - wDscAsig1)
                  wDscAsig1 = wDscAsig1 + (wNetAsig1 - wDscAsig1)
               Else
                  wDscAsig1 = wDscAsig1 + wSdoxDist
                  wSdoxDist = 0
               End If
            End If
         End If
         
         If (wNetAsig2 - wDscAsig2) > 0 And wSdoxDist > 0 Then
            If wSdoxDist > (wNetAsig2 - wDscAsig2) Then
               wSdoxDist = wSdoxDist - (wNetAsig2 - wDscAsig2)
               wDscAsig2 = wNetAsig2
            Else
               If wSdoxDist > (wNetAsig2 - wDscAsig2) Then
                  wSdoxDist = wSdoxDist + (wNetAsig2 - wDscAsig2)
                  wDscAsig2 = wDscAsig2 + (wNetAsig2 - wDscAsig2)
               Else
                  wDscAsig2 = wDscAsig2 + wSdoxDist
                  wSdoxDist = 0
               End If
            End If
         End If
         
         If (wNetAsig3 - wDscAsig3) > 0 And wSdoxDist > 0 Then
            If wSdoxDist > (wNetAsig3 - wDscAsig3) Then
               wSdoxDist = wSdoxDist - (wNetAsig3 - wDscAsig3)
               wDscAsig3 = wNetAsig3
            Else
               If wSdoxDist > (wNetAsig3 - wDscAsig3) Then
                  wSdoxDist = wSdoxDist + (wNetAsig3 - wDscAsig3)
                  wDscAsig3 = wDscAsig3 + (wNetAsig3 - wDscAsig3)
               Else
                  wDscAsig3 = wDscAsig3 + wSdoxDist
                  wSdoxDist = 0
               End If
            End If
         End If
         
         If (wNetAsig4 - wDscAsig4) > 0 And wSdoxDist > 0 Then
            If wSdoxDist > (wNetAsig4 - wDscAsig4) Then
               wSdoxDist = wSdoxDist - (wNetAsig4 - wDscAsig4)
               wDscAsig4 = wNetAsig4
            Else
               If wSdoxDist > (wNetAsig4 - wDscAsig4) Then
                  wSdoxDist = wSdoxDist + (wNetAsig4 - wDscAsig4)
                  wDscAsig4 = wDscAsig4 + (wNetAsig4 - wDscAsig4)
               Else
                  wDscAsig4 = wDscAsig4 + wSdoxDist
                  wSdoxDist = 0
               End If
            End If
         End If
         
         If (wNetAsig5 - wDscAsig5) > 0 And wSdoxDist > 0 Then
            If wSdoxDist > (wNetAsig5 - wDscAsig5) Then
               wSdoxDist = wSdoxDist - (wNetAsig5 - wDscAsig5)
               wDscAsig5 = wNetAsig5
            Else
               If wSdoxDist > (wNetAsig5 - wDscAsig5) Then
                  wSdoxDist = wSdoxDist + (wNetAsig5 - wDscAsig5)
                  wDscAsig5 = wDscAsig5 + (wNetAsig5 - wDscAsig5)
               Else
                  wDscAsig5 = wDscAsig5 + wSdoxDist
                  wSdoxDist = 0
               End If
            End If
         End If
         
         wDifSocio = wNetSocio - wDscSocio
         wDifAsig1 = wNetAsig1 - wDscAsig1
         wDifAsig2 = wNetAsig2 - wDscAsig2
         wDifAsig3 = wNetAsig3 - wDscAsig3
         wDifAsig4 = wNetAsig4 - wDscAsig4
         wDifAsig5 = wNetAsig5 - wDscAsig5
                  
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_DIECOCAB " _
         & " SET DSCSOCIO = " + Str(wDscSocio) + ", DIFSOCIO = " + Str(wDifSocio) + ", " _
         & "     DSCASIG1 = " + Str(wDscAsig1) + ", DIFASIG1 = " + Str(wDifAsig1) + ", " _
         & "     DSCASIG2 = " + Str(wDscAsig2) + ", DIFASIG2 = " + Str(wDifAsig2) + ", " _
         & "     DSCASIG3 = " + Str(wDscAsig3) + ", DIFASIG3 = " + Str(wDifAsig3) + ", " _
         & "     DSCASIG4 = " + Str(wDscAsig4) + ", DIFASIG4 = " + Str(wDifAsig4) + ", " _
         & "     DSCASIG5 = " + Str(wDscAsig5) + ", DIFASIG5 = " + Str(wDifAsig5) + ", " _
         & "     DSCDIFER = " + Str(wSdoxDist) + " " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "            MES = '" + wAno + wMes + "' AND " _
         & "            USU = '" + wcodusu + "' ")
         Db.CommitTrans
   
         Db.BeginTrans
         Db.Execute ("UPDATE DIECOCAB " _
         & " SET DSCSOCIO = " + Str(wDscSocio) + ", DIFSOCIO = " + Str(wDifSocio) + ", " _
         & "     DSCASIG1 = " + Str(wDscAsig1) + ", DIFASIG1 = " + Str(wDifAsig1) + ", " _
         & "     DSCASIG2 = " + Str(wDscAsig2) + ", DIFASIG2 = " + Str(wDifAsig2) + ", " _
         & "     DSCASIG3 = " + Str(wDscAsig3) + ", DIFASIG3 = " + Str(wDifAsig3) + ", " _
         & "     DSCASIG4 = " + Str(wDscAsig4) + ", DIFASIG4 = " + Str(wDifAsig4) + ", " _
         & "     DSCASIG5 = " + Str(wDscAsig5) + ", DIFASIG5 = " + Str(wDifAsig5) + ", " _
         & "     DSCDIFER = " + Str(wSdoxDist) + " " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "            MES = '" + wAno + wMes + "' ")
         Db.CommitTrans
   
         If wDscSocio > 0 Then
            Call DistribuyeDieco(wAno, wMes, wFec, wSoc, wDscSocio)
         End If
         If wDscAsig1 > 0 Then
            Call DistribuyeDieco(wAno, wMes, wFec, wSocAsig1, wDscAsig1)
         End If
         If wDscAsig2 > 0 Then
            Call DistribuyeDieco(wAno, wMes, wFec, wSocAsig2, wDscAsig2)
         End If
         If wDscAsig3 > 0 Then
            Call DistribuyeDieco(wAno, wMes, wFec, wSocAsig3, wDscAsig3)
         End If
         If wDscAsig4 > 0 Then
            Call DistribuyeDieco(wAno, wMes, wFec, wSocAsig4, wDscAsig4)
         End If
         If wDscAsig5 > 0 Then
            Call DistribuyeDieco(wAno, wMes, wFec, wSocAsig5, wDscAsig5)
         End If
         
         wRegAct = wRegAct + 1
         ADO8.MoveNext
      Loop
   End If

   Db.BeginTrans
   Db.Execute ("UPDATE DIECOCAB " _
   & " SET DSCDIFER = TOTENVIO - DSCDIECO " _
   & " WHERE MES = '" + wAno + wMes + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET DSCDIFER = TOTENVIO - DSCDIECO " _
   & " WHERE MES = '" + wAno + wMes + "' AND " _
   & "       USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   lblMensaje.Caption = ""
   lblMensaje.Refresh

   Limpiar
   LlenaCab
   LlenaCab1
   TotalCab
   
'   cmdRevision.Visible = True
'   DataGrid2.Visible = True
   
   MsgBox "Descuento DIECO Actualizado OK", vbExclamation
End Sub

Private Sub cmdRecibir_Click()
   Dim wAno As String, wMes As String, wRuta As String, _
       wCod As Long, wIns As Integer, wSit As Integer, wEsp As Integer, wCip As Long, _
       wImp As Currency, wSin As Currency, wCom As Currency, Cadena As String, _
       zz As Integer, zRegAct As Integer, zRegTot As Integer, wNom As String, _
       wTotEnvio As Currency, wDscDieco As Currency, wDscDifer As Currency, _
       wNetSocio As Currency, wNetAsig1 As Currency, wNetAsig2 As Currency, wNetAsig3 As Currency, wNetAsig4 As Currency, wNetAsig5 As Currency, _
       wDscSocio As Currency, wDscAsig1 As Currency, wDscAsig2 As Currency, wDscAsig3 As Currency, wDscAsig4 As Currency, wDscAsig5 As Currency, _
       wDifSocio As Currency, wDifAsig1 As Currency, wDifAsig2 As Currency, wDifAsig3 As Currency, wDifAsig4 As Currency, wDifAsig5 As Currency, _
       wFecDsc As Date

   wAno = txtAnoCab.Text
   wMes = Left(cmbMeses.Text, 2)
   
   wRuta = xraizDIECO + wAno + "\" + wAno + "-" + wMes + "\15020001.TXT"
   
   If Len(Dir$(wRuta)) = 0 Then
      MsgBox "Archivo " + vbNewLine + _
             wRuta + vbNewLine + _
             "No Existe En Ruta Indicada", vbExclamation
      Exit Sub
   End If
   
   If Not IsDate(txtFecDsc.Text) Then
      MsgBox "Fecha Digitada Es Invalida", vbExclamation
      Exit Sub
   End If
   wFecDsc = Format(txtFecDsc.Text, "dd/mm/yyyy")

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET DSCDIECO = 0, DSCDIFER = 0 " _
   & " WHERE MES = '" + wAno + wMes + "' AND " _
   & "       USU = '" + wcodusu + "' ")
   Db.CommitTrans

   zRegAct = 1
   zRegTot = ADO2.RecordCount
   Open wRuta For Input As #1 ' Abre el archivo para recibir los datos.
    ' Repite el bucle hasta el final del archivo.
   Do While Not EOF(1)
      DoEvents
      lblMensaje.Caption = "Registro " + _
                           Trim(Format(zRegAct, "####0")) + " / " + _
                           Trim(Format(zRegTot, "####0"))
      lblMensaje.Refresh
      
      Line Input #1, Cadena
      
      wCod = Val(Trim(Mid(Cadena, 1, 8)))
      wIns = Val(Mid(Cadena, 9, 1))
      wSit = Val(Mid(Cadena, 18, 1))
      wImp = Format(Val(Mid(Cadena, 19, 10) + "." + Mid(Cadena, 29, 2)), "#####0.00")
      wSin = Format(Val(Mid(Cadena, 31, 10) + "." + Mid(Cadena, 41, 2)), "#####0.00")
      wCip = Val(Mid(Cadena, 43, 8))
      wEsp = Val(Mid(Cadena, 51, 1))
      wCom = Format(Val(Mid(Cadena, 52, 4) + "." + Mid(Cadena, 56, 2)), "#####0.00")
      wNom = ""
      
      zz = Leerado8("SELECT * FROM MAESOCIO " _
                    & " WHERE CODIGO = " + Str(wCod) + " AND " _
                    & "          INS = " + Str(wIns) + " ")
      If zz = 0 Then
         zz = Leerado8("SELECT * FROM MAEPNP " _
                    & " WHERE CODIGO = " + Str(wCod) + " AND " _
                    & "          INS = " + Str(wIns) + " ")
         If zz = 0 Then
            MsgBox "Descuento DIECO Sin Socio " + vbNewLine + Str(Trim(wCod)) + "-" + Str(wIns) + " Enviado", vbExclamation
         End If
         wSoc1 = ADO8!codsocio1
         zz = Leerado8("SELECT * FROM MAESOCIO " _
                    & " WHERE CODSOCIO = " + Str(wSoc1) + " ")
         If zz = 0 Then
            MsgBox "Socio Derivado de PNP No Existe", vbExclamation
            Exit Sub
         End If
         wCod = ADO8!codigo
         wIns = ADO8!ins
         wNom = Trim(ADO8!nombre)
      Else
         wNom = Trim(ADO8!nombre)
      End If
      
      zz = Leerado8("SELECT * FROM TMP_DIECOCAB " _
                    & " WHERE CODIGO = " + Str(wCod) + " AND " _
                    & "          INS = " + Str(wIns) + " AND " _
                    & "          USU = '" + wcodusu + "' ")
      If zz = 0 Then
         zz = Leerado8("SELECT * FROM TMP_DIECOCAB " _
                    & " WHERE CODPNP = " + Str(wCod) + " AND " _
                    & "       INSPNP = " + Str(wIns) + " AND MES = '" + wAno + wMes + "' AND " _
                    & "          USU = '" + wcodusu + "' ")
         If zz = 0 Then
            MsgBox "Descuento DIECO Sin Socio " + vbNewLine + Str(Trim(wCod)) + "-" + Str(wIns) + " Enviado", vbExclamation
         Else
            wTotEnvio = ADO8!totenvio
            wDscDieco = wImp
            wDscDifer = wTotEnvio - wDscDieco
            
            Db.BeginTrans
            Db.Execute ("UPDATE TMP_DIECOCAB " _
            & " SET DSCDIECO = " + Str(wImp) + ", DSCDIFER = " + Str(wSin) + ", " _
            & "         SITU = " + Str(wSit) + ",  SITUESP = " + Str(wEsp) + ", " _
            & "       FECDSC = '" + Format(wFecDsc, "dd/mm/yyyy") + "' " _
            & " WHERE    USU = '" + wcodusu + "' AND " _
            & "       CODPNP = " + Str(wCod) + " AND " _
            & "       INSPNP = " + Str(wIns) + " AND " _
            & "          MES = '" + wAno + wMes + "' ")
            Db.CommitTrans
      
            Db.BeginTrans
            Db.Execute ("UPDATE DIECOCAB " _
            & " SET DSCDIECO = " + Str(wImp) + ", DSCDIFER = " + Str(wSin) + ", " _
            & "         SITU = " + Str(wSit) + ",  SITUESP = " + Str(wEsp) + ",  " _
            & "       FECDSC = '" + Format(wFecDsc, "dd/mm/yyyy") + "' " _
            & " WHERE CODPNP = " + Str(wCod) + " AND " _
            & "       INSPNP = " + Str(wIns) + " AND " _
            & "          MES = '" + wAno + wMes + "' ")
            Db.CommitTrans
         End If
      Else
         wTotEnvio = ADO8!totenvio
         wDscDieco = wImp
         wDscDifer = wTotEnvio - wDscDieco
      
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_DIECOCAB " _
         & " SET DSCDIECO = " + Str(wImp) + ", DSCDIFER = " + Str(wSin) + ", " _
         & "         SITU = " + Str(wSit) + ",  SITUESP = " + Str(wEsp) + ", " _
         & "       FECDSC = '" + Format(wFecDsc, "dd/mm/yyyy") + "' " _
         & " WHERE    USU = '" + wcodusu + "' AND " _
         & "       CODIGO = " + Str(wCod) + " AND " _
         & "          INS = " + Str(wIns) + " AND " _
         & "          MES = '" + wAno + wMes + "' ")
         Db.CommitTrans
      
         Db.BeginTrans
         Db.Execute ("UPDATE DIECOCAB " _
         & " SET DSCDIECO = " + Str(wImp) + ", DSCDIFER = " + Str(wSin) + ", " _
         & "         SITU = " + Str(wSit) + ",  SITUESP = " + Str(wEsp) + ", " _
         & "       FECDSC = '" + Format(wFecDsc, "dd/mm/yyyy") + "' " _
         & " WHERE CODIGO = " + Str(wCod) + " AND " _
         & "          INS = " + Str(wIns) + " AND " _
         & "          MES = '" + wAno + wMes + "' ")
         Db.CommitTrans
      End If
      
      zRegAct = zRegAct + 1
   Loop
   
   Close #1
   
   Db.BeginTrans
   Db.Execute ("UPDATE DIECOCAB " _
   & " SET DSCDIFER = TOTENVIO - DSCDIECO " _
   & " WHERE MES = '" + wAno + wMes + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET DSCDIFER = TOTENVIO - DSCDIECO " _
   & " WHERE MES = '" + wAno + wMes + "' AND " _
   & "       USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   lblMensaje.Caption = ""
   lblMensaje.Refresh

'   DataGrid1.Refresh
   Limpiar
   LlenaCab
   LlenaCab1
   TotalCab
   cmdGrabar.Enabled = True
   cmdGrabar.SetFocus
End Sub

Private Sub cmdRevision_Click()
   Dim aa As Integer, _
       wAno As String, wMes As String, _
       wSoc As Integer, wCod As Long, wIns As Integer, wNom As String, wImp As Currency, _
       wTotEnvio As Currency, wDscDieco As Currency
   wTip = "01"
   wAno = txtAnoCab.Text
   wMes = Left(cmbMeses.Text, 2)
   
   aa = Leerado8("SELECT * FROM TMP_DIECO WHERE USU = '" + wcodusu + "' ")
   If aa = 0 Then
      MsgBox "DIECO Mensual " + Left(Trim(funnommes(wMes)), 3) + " " + wAno + " Sin Registros"
      Exit Sub
   End If
   
   aa = Leerado8("SELECT SUM(TOTENVIO) AS TOTENVIO, SUM(DSCDIECO) AS DSCDIECO " _
                & " FROM TMP_DIECO " _
                & " WHERE USU = '" + wcodusu + "' ")
   If aa > 0 Then
      wTotEnvio = IIf(IsNull(ADO8!totenvio), 0, ADO8!totenvio)
      wDscDieco = IIf(IsNull(ADO8!dscdieco), 0, ADO8!dscdieco)
   End If
   Set ADO8 = Nothing
   
   If wTotEnvio = 0 Then
      MsgBox "Total Enviado En Cero"
      Exit Sub
   End If
   If wDscDieco = 0 Then
      MsgBox "Total Recibido En Cero"
      Exit Sub
   End If
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_CONTROL WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   aa = Leerado8("SELECT * FROM DIECOCAB WHERE MES = '" + wAno + wMes + "' ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wSoc = ADO8!codsocio
         wCod = 0: wIns = 0: wNom = ""
         aa = Leerado7("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
         If aa > 0 Then
            wCod = ADO7!codigo
            wIns = ADO7!ins
            wNom = ADO7!nombre
         End If
         Set ADO7 = Nothing
         
         If ADO8!dscdieco <> ADO8!dscsocio + ADO8!dscasig1 + ADO8!dscasig2 + ADO8!dscasig3 + ADO8!dscasig4 + ADO8!dscasig5 Then
            MsgBox "Descuadre - Socio " + Str(ADO8!codsocio) + " " + Trim(wNom)
            Exit Sub
         End If
   
         If ADO8!dscsocio > 0 And ADO8!codsocio = 0 Then
            MsgBox "Error Socio - Socio " + Str(ADO8!codsocio)
            Exit Sub
         End If
         If ADO8!dscasig1 > 0 And ADO8!codasig1 = 0 Then
            MsgBox "Error Asignado 1 - Socio " + Str(ADO8!codsocio)
            Exit Sub
         End If
         If ADO8!dscasig2 > 0 And ADO8!codasig2 = 0 Then
            MsgBox "Error Asignado 2 - Socio " + Str(ADO8!codsocio)
            Exit Sub
         End If
         If ADO8!dscasig3 > 0 And ADO8!codasig3 = 0 Then
            MsgBox "Error Asignado 3 - Socio " + Str(ADO8!codsocio)
            Exit Sub
         End If
         If ADO8!dscasig4 > 0 And ADO8!codasig4 = 0 Then
            MsgBox "Error Asignado 4 - Socio " + Str(ADO8!codsocio)
            Exit Sub
         End If
         If ADO8!dscasig5 > 0 And ADO8!codasig5 = 0 Then
            MsgBox "Error Asignado 5 - Socio " + Str(ADO8!codsocio)
            Exit Sub
         End If
   
         If ADO8!dscsocio > 0 Then
            wSoc = ADO8!codsocio
            wImp = ADO8!dscsocio
            wCod = 0: wIns = 0: wNom = ""
            aa = Leerado7("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
            If aa > 0 Then
               wCod = ADO7!codigo
               wIns = ADO7!ins
               wNom = ADO7!nombre
            End If
            Set ADO7 = Nothing
                 
            aa = Leerado7("SELECT * FROM TMP_CONTROL WHERE USU = '" + wcodusu + "' AND CODSOCIO = " + Str(wSoc) + " ")
            If aa = 0 Then
               Db.BeginTrans
               Db.Execute ("INSERT INTO TMP_CONTROL " _
               & " (ANO, MES, TIPCOB, CODSOCIO, CODIGO, INS, NOMBRE, NUM, USU ) " _
               & " VALUES " _
               & " ('" + wAno + "', '" + wMes + "', '" + wTip + "', " + Str(wSoc) + ", " _
               & "  " + Str(wCod) + ", " + Str(wIns) + ", " _
               & "  '" + wNom + "', 0, '" + wcodusu + "' ) ")
               Db.CommitTrans
            End If
            Db.BeginTrans
            Db.Execute ("UPDATE TMP_CONTROL " _
            & " SET NUM = NUM + 1, " _
            & "     IMPORTE = IMPORTE + " + Str(wImp) + " " _
            & " WHERE USU = '" + wcodusu + "' AND " _
            & "       CODSOCIO = " + Str(wSoc) + " ")
            Db.CommitTrans
         End If
         
         If ADO8!dscasig1 > 0 Then
            wSoc = ADO8!codasig1
            wImp = ADO8!dscasig1
            wCod = 0: wIns = 0: wNom = ""
            aa = Leerado7("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
            If aa > 0 Then
               wCod = ADO7!codigo
               wIns = ADO7!ins
               wNom = ADO7!nombre
            End If
            Set ADO7 = Nothing
                 
            aa = Leerado7("SELECT * FROM TMP_CONTROL WHERE USU = '" + wcodusu + "' AND CODSOCIO = " + Str(wSoc) + " ")
            If aa = 0 Then
               Db.BeginTrans
               Db.Execute ("INSERT INTO TMP_CONTROL " _
               & " (ANO, MES, TIPCOB, CODSOCIO, CODIGO, INS, NOMBRE, NUM, USU ) " _
               & " VALUES " _
               & " ('" + wAno + "', '" + wMes + "', '" + wTip + "', " + Str(wSoc) + ", " _
               & "  " + Str(wCod) + ", " + Str(wIns) + ", " _
               & "  '" + wNom + "', 0, '" + wcodusu + "' ) ")
               Db.CommitTrans
            End If
            Db.BeginTrans
            Db.Execute ("UPDATE TMP_CONTROL " _
            & " SET NUM = NUM + 1, " _
            & "     IMPORTE = IMPORTE + " + Str(wImp) + " " _
            & " WHERE USU = '" + wcodusu + "' AND " _
            & "       CODSOCIO = " + Str(wSoc) + " ")
            Db.CommitTrans
         End If
   
         If ADO8!dscasig2 > 0 Then
            wSoc = ADO8!codasig2
            wImp = ADO8!dscasig2
            wCod = 0: wIns = 0: wNom = ""
            aa = Leerado7("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
            If aa > 0 Then
               wCod = ADO7!codigo
               wIns = ADO7!ins
               wNom = ADO7!nombre
            End If
            Set ADO7 = Nothing
                 
            aa = Leerado7("SELECT * FROM TMP_CONTROL WHERE USU = '" + wcodusu + "' AND CODSOCIO = " + Str(wSoc) + " ")
            If aa = 0 Then
               Db.BeginTrans
               Db.Execute ("INSERT INTO TMP_CONTROL " _
               & " (ANO, MES, TIPCOB, CODSOCIO, CODIGO, INS, NOMBRE, NUM, USU ) " _
               & " VALUES " _
               & " ('" + wAno + "', '" + wMes + "', '" + wTip + "', " + Str(wSoc) + ", " _
               & "  " + Str(wCod) + ", " + Str(wIns) + ", " _
               & "  '" + wNom + "', 0, '" + wcodusu + "' ) ")
               Db.CommitTrans
            End If
            Db.BeginTrans
            Db.Execute ("UPDATE TMP_CONTROL " _
            & " SET NUM = NUM + 1, " _
            & "     IMPORTE = IMPORTE + " + Str(wImp) + " " _
            & " WHERE USU = '" + wcodusu + "' AND " _
            & "       CODSOCIO = " + Str(wSoc) + " ")
            Db.CommitTrans
         End If
   
         If ADO8!dscasig3 > 0 Then
            wSoc = ADO8!codasig3
            wImp = ADO8!dscasig3
            wCod = 0: wIns = 0: wNom = ""
            aa = Leerado7("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
            If aa > 0 Then
               wCod = ADO7!codigo
               wIns = ADO7!ins
               wNom = ADO7!nombre
            End If
            Set ADO7 = Nothing
                 
            aa = Leerado7("SELECT * FROM TMP_CONTROL WHERE USU = '" + wcodusu + "' AND CODSOCIO = " + Str(wSoc) + " ")
            If aa = 0 Then
               Db.BeginTrans
               Db.Execute ("INSERT INTO TMP_CONTROL " _
               & " (ANO, MES, TIPCOB, CODSOCIO, CODIGO, INS, NOMBRE, NUM, USU ) " _
               & " VALUES " _
               & " ('" + wAno + "', '" + wMes + "', '" + wTip + "', " + Str(wSoc) + ", " _
               & "  " + Str(wCod) + ", " + Str(wIns) + ", " _
               & "  '" + wNom + "', 0, '" + wcodusu + "' ) ")
               Db.CommitTrans
            End If
            Db.BeginTrans
            Db.Execute ("UPDATE TMP_CONTROL " _
            & " SET NUM = NUM + 1, " _
            & "     IMPORTE = IMPORTE + " + Str(wImp) + " " _
            & " WHERE USU = '" + wcodusu + "' AND " _
            & "       CODSOCIO = " + Str(wSoc) + " ")
            Db.CommitTrans
         End If
   
         If ADO8!dscasig4 > 0 Then
            wSoc = ADO8!codasig4
            wImp = ADO8!dscasig4
            wCod = 0: wIns = 0: wNom = ""
            aa = Leerado7("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
            If aa > 0 Then
               wCod = ADO7!codigo
               wIns = ADO7!ins
               wNom = ADO7!nombre
            End If
            Set ADO7 = Nothing
                 
            aa = Leerado7("SELECT * FROM TMP_CONTROL WHERE USU = '" + wcodusu + "' AND CODSOCIO = " + Str(wSoc) + " ")
            If aa = 0 Then
               Db.BeginTrans
               Db.Execute ("INSERT INTO TMP_CONTROL " _
               & " (ANO, MES, TIPCOB, CODSOCIO, CODIGO, INS, NOMBRE, NUM, USU ) " _
               & " VALUES " _
               & " ('" + wAno + "', '" + wMes + "', '" + wTip + "', " + Str(wSoc) + ", " _
               & "  " + Str(wCod) + ", " + Str(wIns) + ", " _
               & "  '" + wNom + "', 0, '" + wcodusu + "' ) ")
               Db.CommitTrans
            End If
            Db.BeginTrans
            Db.Execute ("UPDATE TMP_CONTROL " _
            & " SET NUM = NUM + 1, " _
            & "     IMPORTE = IMPORTE + " + Str(wImp) + " " _
            & " WHERE USU = '" + wcodusu + "' AND " _
            & "       CODSOCIO = " + Str(wSoc) + " ")
            Db.CommitTrans
         End If
   
         If ADO8!dscasig5 > 0 Then
            wSoc = ADO8!codasig5
            wImp = ADO8!dscasig5
            wCod = 0: wIns = 0: wNom = ""
            aa = Leerado7("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
            If aa > 0 Then
               wCod = ADO7!codigo
               wIns = ADO7!ins
               wNom = ADO7!nombre
            End If
            Set ADO7 = Nothing
                 
            aa = Leerado7("SELECT * FROM TMP_CONTROL WHERE USU = '" + wcodusu + "' AND CODSOCIO = " + Str(wSoc) + " ")
            If aa = 0 Then
               Db.BeginTrans
               Db.Execute ("INSERT INTO TMP_CONTROL " _
               & " (ANO, MES, TIPCOB, CODSOCIO, CODIGO, INS, NOMBRE, NUM, USU ) " _
               & " VALUES " _
               & " ('" + wAno + "', '" + wMes + "', '" + wTip + "', " + Str(wSoc) + ", " _
               & "  " + Str(wCod) + ", " + Str(wIns) + ", " _
               & "  '" + wNom + "', 0, '" + wcodusu + "' ) ")
               Db.CommitTrans
            End If
            Db.BeginTrans
            Db.Execute ("UPDATE TMP_CONTROL " _
            & " SET NUM = NUM + 1, " _
            & "     IMPORTE = IMPORTE + " + Str(wImp) + " " _
            & " WHERE USU = '" + wcodusu + "' AND " _
            & "       CODSOCIO = " + Str(wSoc) + " ")
            Db.CommitTrans
         End If
   
         ADO8.MoveNext
      Loop
   End If
   aa = Leerado6a("SELECT CODIGO, INS, NOMBRE " _
                & " FROM TMP_CONTROL " _
                & " WHERE USU = '" + wcodusu + "' AND NUM > 1 " _
                & " ORDER BY NOMBRE ")
   If aa > 0 Then
      DataGrid2.Visible = True
      DataGrid2.Caption = "REGISTROS ERRADOS"
      
      Set DataGrid2.DataSource = ADO6a
    
      DataGrid2.SetFocus
   Else
      MsgBox "No Existen Asignados Con Errores"
      
      DataGrid2.Visible = False
   End If
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub DataGrid1_DblClick()
   zDetaCambio = False
   zDetaCodSoc = ADO2!codsocio
   zDetaTipDsc = "01"
   zDetaAnoDsc = txtAnoCab.Text
   zDetaMesDsc = Left(cmbMeses.Text, 2)

   frmDIECODetalle.Show vbModal
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   Select Case ColIndex
   Case 0
        ADO2.Sort = "CODSOCIO"
   Case 1
        ADO2.Sort = "CODIGO"
   Case 3
        ADO2.Sort = "NOMBRE"
   End Select
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If Not ADO2.EOF Then
      lblAsig1.Caption = ADO2!nomasig1
      lblAsig2.Caption = ADO2!nomasig2
      lblAsig3.Caption = ADO2!nomasig3
      lblAsig4.Caption = ADO2!nomasig4
      lblAsig5.Caption = ADO2!nomasig5
   End If
End Sub

Private Sub Form_Activate()
   Form_Initialize
End Sub

Private Sub Form_Initialize()
   frmDiecoRecibe.Left = (Screen.Width - Width) \ 2
   frmDiecoRecibe.Top = 0
   
   txtAnoCab.Text = wanocia
   
   Dim a As Integer
   cmbMeses.Clear
   a = Leerado("select * from MAEMESES " _
            & " WHERE ANO = '" + wanocia + "' AND " _
            & "       MES >= '01' AND " _
            & "       MES <= '12' " _
            & " ORDER BY MES ")
   ADO1.MoveFirst
   Do While Not ADO1.EOF
      cmbMeses.AddItem ADO1!mes + " " + Trim(funnommes(ADO1!mes))
       ADO1.MoveNext
   Loop
   txtAnoCab.SetFocus
End Sub

Private Sub LlenaCab()
   Dim wAno As String, wMes As String, zz As Integer
   wAno = txtAnoCab.Text
   wMes = Left(cmbMeses.Text, 2)
      
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_DIECOCAB WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECOCAB " _
   & " (MES, CODSOCIO, CODIGO, INS, E_SOCIO, NOMBRE, FECENV, FECDSC, " _
   & "  TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "  TOTENVIO, DSCDIECO, DSCDIFER, SITU    , SITUESP , TIPCOB  , CODPNP, INSPNP, " _
   & "  CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
   & "  DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
   & "  NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, DSCASIG1, DSCASIG2, DSCASIG3, DSCASIG4, DSCASIG5, " _
   & "  DIFASIG1, DIFASIG2, DIFASIG3, DIFASIG4, DIFASIG5, " _
   & "  NOMASIG1, NOMASIG2, NOMASIG3, NOMASIG4, NOMASIG5, USU ) " _
   & " SELECT " _
   & "  D.MES, D.CODSOCIO, M.CODIGO, M.INS, M.E_SOCIO, M.NOMBRE, D.FECENV, D.FECDSC, " _
   & "  TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "  TOTENVIO, DSCDIECO, DSCDIFER, D.SITU  , D.SITUESP, D.TIPCOB, D.CODPNP, D.INSPNP, " _
   & "  CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
   & "  DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
   & "  NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, DSCASIG1, DSCASIG2, DSCASIG3, DSCASIG4, DSCASIG5, " _
   & "  DIFASIG1, DIFASIG2, DIFASIG3, DIFASIG4, DIFASIG5, " _
   & "  ''      , ''      , ''      , ''      , ''      , '" + wcodusu + "'  " _
   & " FROM DIECOCAB AS D INNER JOIN MAESOCIO AS M ON D.CODSOCIO = M.CODSOCIO " _
   & " WHERE D.MES = '" + wAno + wMes + "'  ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG1 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG1 = M.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' AND " _
   & "       T.CODASIG1 <> 0 ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG2 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG2 = M.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' AND " _
   & "       T.CODASIG2 <> 0 ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG3 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG3 = M.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' AND " _
   & "       T.CODASIG3 <> 0 ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG4 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG4 = M.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' AND " _
   & "       T.CODASIG4 <> 0 ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG5 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG5 = M.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' AND " _
   & "       T.CODASIG5 <> 0 ")
   Db.CommitTrans
   
   zz = Leerado2("SELECT CODSOCIO, CODIGO  , INS     , NOMBRE  , " _
                & "      TOTENVIO, DSCDIECO, DSCDIFER " _
                & "      TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
                & "      CODASIG1, TOTASIG1, DEUASIG1, ADEASIG1, NETASIG1, DSCASIG1, DIFASIG1, " _
                & "      CODASIG2, TOTASIG2, DEUASIG2, ADEASIG2, NETASIG2, DSCASIG2, DIFASIG2, " _
                & "      CODASIG3, TOTASIG3, DEUASIG3, ADEASIG3, NETASIG3, DSCASIG3, DIFASIG3, " _
                & "      CODASIG4, TOTASIG4, DEUASIG4, ADEASIG4, NETASIG4, DSCASIG4, DIFASIG4, " _
                & "      CODASIG5, TOTASIG5, DEUASIG5, ADEASIG5, NETASIG5, DSCASIG5, DIFASIG5, " _
                & "      NOMASIG1, NOMASIG2, NOMASIG3, NOMASIG4, NOMASIG5, MES, " _
                & "      SITU    , SITUESP , TIPCOB  , CODPNP  , INSPNP  , FECENV  , FECDSC  , E_SOCIO " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE MES = '" + wAno + wMes + "' AND USU = '" + wcodusu + "' " _
                & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO2
   If zz = 0 Then
      MsgBox "No Existe Envio a DIECO del Mes " + Trim(funnommes(wMes)) + "-" + wAno
   End If
End Sub

Private Sub LlenaCab1()
   DataGrid1.Columns(0).Width = 750   ' CODSOCIO
   DataGrid1.Columns(0).Alignment = dbgRight
   DataGrid1.Columns(0).Caption = "SOCIO"
    
   DataGrid1.Columns(1).Width = 800   ' CODIGO
   DataGrid1.Columns(1).Alignment = dbgRight
   DataGrid1.Columns(1).Caption = "CODIGO"
    
   DataGrid1.Columns(2).Width = 400   ' INS
   DataGrid1.Columns(2).Alignment = dbgCenter
   DataGrid1.Columns(2).Caption = "INS"
    
   DataGrid1.Columns(3).Width = 6000  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE ASOCIADO"
    
   DataGrid1.Columns(4).Width = 800    ' TOTENVIO
   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(4).Caption = "T.ENVIO"
    
   DataGrid1.Columns(5).Width = 800    ' TOTDIECO
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(5).Caption = "DESCTO"
    
   DataGrid1.Columns(6).Width = 800    ' TOTDIFER
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(6).Caption = "NO DSCTO"
    
   DataGrid1.Columns(7).Visible = False
   DataGrid1.Columns(8).Visible = False
   DataGrid1.Columns(9).Visible = False
   DataGrid1.Columns(10).Visible = False
   DataGrid1.Columns(11).Visible = False
   DataGrid1.Columns(12).Visible = False
   DataGrid1.Columns(13).Visible = False
   DataGrid1.Columns(14).Visible = False
   DataGrid1.Columns(15).Visible = False
   DataGrid1.Columns(16).Visible = False
   DataGrid1.Columns(17).Visible = False
   DataGrid1.Columns(18).Visible = False
   DataGrid1.Columns(19).Visible = False
   DataGrid1.Columns(20).Visible = False
   DataGrid1.Columns(21).Visible = False
   DataGrid1.Columns(22).Visible = False
   DataGrid1.Columns(23).Visible = False
   DataGrid1.Columns(24).Visible = False
   DataGrid1.Columns(25).Visible = False
   DataGrid1.Columns(26).Visible = False
   DataGrid1.Columns(27).Visible = False
   DataGrid1.Columns(28).Visible = False
   DataGrid1.Columns(29).Visible = False
   DataGrid1.Columns(30).Visible = False
   DataGrid1.Columns(31).Visible = False
   DataGrid1.Columns(32).Visible = False
   DataGrid1.Columns(33).Visible = False
   DataGrid1.Columns(34).Visible = False
   DataGrid1.Columns(35).Visible = False
   DataGrid1.Columns(36).Visible = False
   DataGrid1.Columns(37).Visible = False
   DataGrid1.Columns(38).Visible = False
   DataGrid1.Columns(39).Visible = False
   DataGrid1.Columns(40).Visible = False
   DataGrid1.Columns(41).Visible = False
   DataGrid1.Columns(42).Visible = False
   DataGrid1.Columns(43).Visible = False
   DataGrid1.Columns(44).Visible = False
   DataGrid1.Columns(45).Visible = False
   DataGrid1.Columns(46).Visible = False
   DataGrid1.Columns(47).Visible = False
   DataGrid1.Columns(48).Visible = False
   DataGrid1.Columns(49).Visible = False
   DataGrid1.Columns(50).Visible = False
   DataGrid1.Columns(51).Visible = False
   DataGrid1.Columns(52).Visible = False
   DataGrid1.Columns(53).Visible = False
   DataGrid1.Columns(54).Visible = False
   DataGrid1.Columns(55).Visible = False
   DataGrid1.Columns(56).Visible = False
   DataGrid1.Columns(57).Visible = False
   DataGrid1.Columns(58).Visible = False
   DataGrid1.Columns(59).Visible = False
   DataGrid1.Columns(60).Visible = False
   
   DataGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
   DataGrid1.SetFocus
End Sub

Private Sub TotalCab()
   Dim zz As Integer, _
       zAno As String, zMes As String, _
       zTotEnv As Currency, zTotDsc As Currency, zTotNoD As Currency, _
       zCanApo As Integer, zCanAsi As Integer
   
   zAno = txtAnoCab.Text
   zMes = Left(cmbMeses.Text, 2)
   
   zz = Leerado8("SELECT SUM(TOTENVIO) AS TOTENVIO, " _
                & "      SUM(DSCDIECO) AS DSCDIECO, " _
                & "      SUM(DSCDIFER) AS DSCDIFER, " _
                & "      COUNT(*) AS CAN " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "       MES = '" + zAno + zMes + "' ")
   If zz > 0 Then
      zTotEnv = IIf(IsNull(ADO8!totenvio), 0, ADO8!totenvio)
      zTotDsc = IIf(IsNull(ADO8!dscdieco), 0, ADO8!dscdieco)
      zTotNoD = IIf(IsNull(ADO8!dscdifer), 0, ADO8!dscdifer)
      zCanApo = IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT COUNT(*) AS CAN " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE (USU = '" + wcodusu + "') AND " _
                & "       (MES = '" + zAno + zMes + "') AND " _
                & "       (TOTASIG1 <> 0) ")
   If zz > 0 Then
      zCanAsi = IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT COUNT(*) AS CAN " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE (USU = '" + wcodusu + "') AND " _
                & "       (MES = '" + zAno + zMes + "') AND " _
                & "       (TOTASIG2 <> 0) ")
   If zz > 0 Then
      zCanAsi = zCanAsi + IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT COUNT(*) AS CAN " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE (USU = '" + wcodusu + "') AND " _
                & "       (MES = '" + zAno + zMes + "') AND " _
                & "       (TOTASIG3 <> 0) ")
   If zz > 0 Then
      zCanAsi = zCanAsi + IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT COUNT(*) AS CAN " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE (USU = '" + wcodusu + "') AND " _
                & "       (MES = '" + zAno + zMes + "') AND " _
                & "       (TOTASIG4 <> 0) ")
   If zz > 0 Then
      zCanAsi = zCanAsi + IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT COUNT(*) AS CAN " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE (USU = '" + wcodusu + "') AND " _
                & "       (MES = '" + zAno + zMes + "') AND " _
                & "       (TOTASIG5 <> 0) ")
   If zz > 0 Then
      zCanAsi = zCanAsi + IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   lblEnviado.Caption = Format(zTotEnv, "###,##0.00;;\ ")
   lblRecibido.Caption = Format(zTotDsc, "###,##0.00;;\ ")
   lblNoDscto.Caption = Format(zTotNoD, "###,##0.00;;\ ")
   lblCanApo.Caption = Format(zCanApo, "##,##0")
   lblCanAsi.Caption = Format(zCanAsi, "##,##0")
End Sub

Private Sub optFiltro_Click()
   If optTodos.Value = True Then
      txtFiltrar.Text = ""
      txtFiltrar.Enabled = False
      DataGrid1.SetFocus
   Else
      txtFiltrar.Enabled = True
      optFiltro.Value = True
      txtFiltrar.SetFocus
   End If
End Sub

Private Sub optFiltro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      optTodos_Click
   End If
End Sub

Private Sub optTodos_Click()
   If optTodos.Value = True Then
      txtFiltrar.Text = ""
      txtFiltrar.Enabled = False
      LlenaCab
      LlenaCab1
      TotalCab
   Else
      txtFiltrar.Enabled = True
      optFiltro.Value = True
   End If
End Sub

Private Sub optTodos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If optTodos.Value = True Then
         txtFiltrar.Text = ""
         txtFiltrar.Enabled = False
         ADO2.Filter = ""
         Set DataGrid1.DataSource = ADO2
         DataGrid1.SetFocus
      Else
         txtFiltrar.Enabled = True
         optFiltro.Value = True
         txtFiltrar.SetFocus
      End If
   End If
End Sub

Private Sub txtAnoCab_GotFocus()
   txtAnoCab.SelStart = 0
   txtAnoCab.SelLength = 4
End Sub

Private Sub txtAnoCab_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txtAnoCab.Text)) = 0 Then
         MsgBox "A�o En Blanco", vbExclamation
         txtAnoCab.Text = wanocia
         Exit Sub
      End If
      If txtAnoCab.Text < "2018" Or txtAnoCab.Text > "2030" Then
         MsgBox "A�o Digitado Fuera de Rango", vbExclamation
         txtAnoCab.Text = wanocia
         Exit Sub
      End If
      cmbMeses.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFecDsc_GotFocus()
   txtFecDsc.SelStart = 0
   txtFecDsc.SelLength = 10
End Sub

Private Sub txtFecDsc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsDate(txtFecDsc.Text) Then
         MsgBox "Fecha Digitada Es Invalida", vbExclamation
         txtFecDsc.Text = "__/__/____"
         Exit Sub
      End If
      cmdRecibir.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFiltrar_Change()
   Dim aa As Integer, wAno As String, wMes As String
   wAno = txtAnoCab.Text
   wMes = Left(cmbMeses.Text, 2)
   
   optFiltro.Value = True
   
   aa = Leerado2("SELECT CODSOCIO, CODIGO  , INS     , NOMBRE  , " _
                & "      TOTENVIO, DSCDIECO, DSCDIFER " _
                & "      TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
                & "      CODASIG1, TOTASIG1, DEUASIG1, ADEASIG1, NETASIG1, DSCASIG1, DIFASIG1, " _
                & "      CODASIG2, TOTASIG2, DEUASIG2, ADEASIG2, NETASIG2, DSCASIG2, DIFASIG2, " _
                & "      CODASIG3, TOTASIG3, DEUASIG3, ADEASIG3, NETASIG3, DSCASIG3, DIFASIG3, " _
                & "      CODASIG4, TOTASIG4, DEUASIG4, ADEASIG4, NETASIG4, DSCASIG4, DIFASIG4, " _
                & "      CODASIG5, TOTASIG5, DEUASIG5, ADEASIG5, NETASIG5, DSCASIG5, DIFASIG5, " _
                & "      NOMASIG1, NOMASIG2, NOMASIG3, NOMASIG4, NOMASIG5, MES, " _
                & "      SITU    , SITUESP , TIPCOB  , CODPNP  , INSPNP  , FECENV  , FECDSC  , E_SOCIO " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE MES = '" + wAno + wMes + "' AND " _
                & "       USU = '" + wcodusu + "' AND " _
                & "       NOMBRE LIKE '%" + Trim(txtFiltrar.Text) + "%' " _
                & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO2
   
   LlenaCab1
   TotalCab
   txtFiltrar.SetFocus
End Sub

Private Sub txtFiltrar_GotFocus()
   txtFiltrar.SelStart = 0
   txtFiltrar.SelLength = Len(Trim(txtFiltrar.Text))
End Sub
