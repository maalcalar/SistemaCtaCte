VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMPConxSoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Caja Militar Policial Por Socio"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   13350
   Begin VB.TextBox txtGrado 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   6960
      MaxLength       =   3
      TabIndex        =   21
      Top             =   1240
      Width           =   495
   End
   Begin VB.TextBox txtNumdoc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8280
      MaxLength       =   8
      TabIndex        =   20
      Top             =   660
      Width           =   975
   End
   Begin VB.TextBox txtIns 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7920
      MaxLength       =   1
      TabIndex        =   19
      Top             =   660
      Width           =   375
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   6960
      MaxLength       =   8
      TabIndex        =   18
      Top             =   660
      Width           =   975
   End
   Begin VB.TextBox txtTipCob 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9360
      MaxLength       =   3
      TabIndex        =   17
      Top             =   1240
      Width           =   495
   End
   Begin VB.TextBox txtE_socio 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9240
      MaxLength       =   3
      TabIndex        =   16
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox txtCodSocio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   11
      Top             =   1240
      Width           =   735
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
      Left            =   12000
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   840
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
      Left            =   9240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6720
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
      Left            =   10560
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6720
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
      Left            =   11880
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   7646
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
   Begin Crystal.CrystalReport Crys1 
      Left            =   12720
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Estado de Cuenta"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSMask.MaskEdBox txtDesde 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtHasta 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   920
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####-##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Enviado"
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
      Left            =   9240
      TabIndex        =   36
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label lblNetSocio 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9240
      TabIndex        =   35
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   " Cobrado"
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
      Left            =   10440
      TabIndex        =   34
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label lblDscSocio 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10440
      TabIndex        =   33
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "No Cobrado"
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
      Left            =   11640
      TabIndex        =   32
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label lblDifSocio 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11640
      TabIndex        =   31
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label lblGrado 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7440
      TabIndex        =   30
      Top             =   1245
      Width           =   1935
   End
   Begin VB.Label Label15 
      Caption         =   "Grado"
      Height          =   195
      Left            =   7200
      TabIndex        =   29
      Top             =   1065
      Width           =   1335
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "D.N.I."
      Height          =   195
      Left            =   8280
      TabIndex        =   28
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Ins"
      Height          =   195
      Left            =   7920
      TabIndex        =   27
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Codofin"
      Height          =   195
      Left            =   6960
      TabIndex        =   26
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Tipo de Cobro"
      Height          =   195
      Left            =   9360
      TabIndex        =   25
      Top             =   1065
      Width           =   1335
   End
   Begin VB.Label lblTipCob 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9840
      TabIndex        =   24
      Top             =   1245
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Estado del Socio"
      Height          =   195
      Left            =   9240
      TabIndex        =   23
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblE_socio 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9720
      TabIndex        =   22
      Top             =   660
      Width           =   2175
   End
   Begin VB.Label lblHasta 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   920
      Width           =   2295
   End
   Begin VB.Label lblDesde 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblCodSocio 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   1240
      Width           =   4695
   End
   Begin VB.Label Label4 
      Caption         =   "Cod.Socio"
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
      Left            =   120
      TabIndex        =   12
      Top             =   1240
      Width           =   975
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
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   920
      Width           =   1095
   End
   Begin VB.Label Label3 
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
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
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
      Left            =   360
      TabIndex        =   6
      Top             =   6480
      Width           =   7575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Consulta Caja Militar Policial Por Socio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmMPConxSoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Limpiar()
   txtDesde.Text = "____-__"
   txtHasta.Text = "____-__"
   txtCodSocio.Text = ""
   txtCodigo.Text = ""
   txtIns.Text = ""
   txtNumdoc.Text = ""
   txtGrado.Text = ""
   txtE_socio.Text = ""
   txtTipCob.Text = ""
End Sub

Private Sub cmdBuscar_Click()
   Dim wDesAno As String, wDesMes As String, _
       wHasAno As String, wHasMes As String, _
       wSoc As Integer, zz As Integer

   wDesAno = Left(txtDesde.Text, 4)
   wDesMes = Right(txtDesde.Text, 2)
   wHasAno = Left(txtHasta.Text, 4)
   wHasMes = Right(txtHasta.Text, 2)
   wSoc = Val(txtCodSocio.Text)

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_CAJMPSOC WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_CAJMPSOC " _
   & " (MES, CODENVIO, LIN, CODSOCIO, CODIGO, INS, CARNETPNP, NUMDOC, CODBENI, NOMBRE, " _
   & "  FECENV, FECDSC, TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "  TOTENVIO, DSCCAJMP, DSCDIFER, USU ) " _
   & " SELECT " _
   & "  C.MES, C.CODSOCIO, '0', C.CODSOCIO, C.CODIGO, C.INS, C.CARNETPNP, C.NUMDOC, " _
   & "  C.CODBENI, M.NOMBRE, C.FECENV, C.FECDSC, C.TOTAPORT, C.TOTDEUDA, " _
   & "  C.TOTADELA, C.NETSOCIO, C.DSCSOCIO, C.DIFSOCIO, C.TOTENVIO, C.DSCCAJMP, " _
   & "  C.DSCDIFER, '" + wcodusu + "' " _
   & " FROM CAJMPCAB AS C INNER JOIN MAESOCIO AS M " _
   & "   ON C.CODSOCIO = M.CODSOCIO " _
   & " WHERE C.MES >= '" + wDesAno + wDesMes + "' AND " _
   & "       C.MES <= '" + wHasAno + wHasMes + "' AND " _
   & "       C.CODSOCIO = " + Str(wSoc) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_CAJMPSOC " _
   & " (MES, CODENVIO, LIN, CODSOCIO, CODIGO, INS, CARNETPNP, NUMDOC, CODBENI, NOMBRE, " _
   & "  FECENV, FECDSC, TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "  TOTENVIO, DSCCAJMP, DSCDIFER, USU ) " _
   & " SELECT " _
   & "  C.MES, C.CODSOCIO, '1', C.CODASIG1, M.CODIGO, M.INS, M.CARNETPNP, M.NUMDOC, " _
   & "  M.CODBENI, M.NOMBRE, C.FECENV, C.FECDSC, C.TOTASIG1, C.DEUASIG1, " _
   & "  C.ADEASIG1, C.NETASIG1, C.DSCASIG1, C.DIFASIG1, 0, 0, 0, '" + wcodusu + "' " _
   & " FROM CAJMPCAB AS C INNER JOIN MAESOCIO AS M " _
   & "   ON C.CODASIG1 = M.CODSOCIO " _
   & " WHERE C.MES >= '" + wDesAno + wDesMes + "' AND " _
   & "       C.MES <= '" + wHasAno + wHasMes + "' AND " _
   & "       C.CODASIG1 = " + Str(wSoc) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_CAJMPSOC " _
   & " (MES, CODENVIO, LIN, CODSOCIO, CODIGO, INS, CARNETPNP, NUMDOC, CODBENI, NOMBRE, " _
   & "  FECENV, FECDSC, TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "  TOTENVIO, DSCCAJMP, DSCDIFER, USU ) " _
   & " SELECT " _
   & "  C.MES, C.CODSOCIO, '2', C.CODASIG2, M.CODIGO, M.INS, M.CARNETPNP, M.NUMDOC, " _
   & "  M.CODBENI, M.NOMBRE, C.FECENV, C.FECDSC, C.TOTASIG2, C.DEUASIG2, " _
   & "  C.ADEASIG2, C.NETASIG2, C.DSCASIG2, C.DIFASIG2, 0, 0, 0, '" + wcodusu + "' " _
   & " FROM CAJMPCAB AS C INNER JOIN MAESOCIO AS M " _
   & "   ON C.CODASIG2 = M.CODSOCIO " _
   & " WHERE C.MES >= '" + wDesAno + wDesMes + "' AND " _
   & "       C.MES <= '" + wHasAno + wHasMes + "' AND " _
   & "       C.CODASIG2 = " + Str(wSoc) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_CAJMPSOC " _
   & " (MES, CODENVIO, LIN, CODSOCIO, CODIGO, INS, CARNETPNP, NUMDOC, CODBENI, NOMBRE, " _
   & "  FECENV, FECDSC, TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "  TOTENVIO, DSCCAJMP, DSCDIFER, USU ) " _
   & " SELECT " _
   & "  C.MES, C.CODSOCIO, '3', C.CODASIG3, M.CODIGO, M.INS, M.CARNETPNP, M.NUMDOC, " _
   & "  M.CODBENI, M.NOMBRE, C.FECENV, C.FECDSC, C.TOTASIG3, C.DEUASIG3, " _
   & "  C.ADEASIG3, C.NETASIG3, C.DSCASIG3, C.DIFASIG3, 0, 0, 0, '" + wcodusu + "' " _
   & " FROM CAJMPCAB AS C INNER JOIN MAESOCIO AS M " _
   & "   ON C.CODASIG3 = M.CODSOCIO " _
   & " WHERE C.MES >= '" + wDesAno + wDesMes + "' AND " _
   & "       C.MES <= '" + wHasAno + wHasMes + "' AND " _
   & "       C.CODASIG3 = " + Str(wSoc) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_CAJMPSOC " _
   & " (MES, CODENVIO, LIN, CODSOCIO, CODIGO, INS, CARNETPNP, NUMDOC, CODBENI, NOMBRE, " _
   & "  FECENV, FECDSC, TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "  TOTENVIO, DSCCAJMP, DSCDIFER, USU ) " _
   & " SELECT " _
   & "  C.MES, C.CODSOCIO, '4', C.CODASIG4, M.CODIGO, M.INS, M.CARNETPNP, M.NUMDOC, " _
   & "  M.CODBENI, M.NOMBRE, C.FECENV, C.FECDSC, C.TOTASIG4, C.DEUASIG4, " _
   & "  C.ADEASIG4, C.NETASIG4, C.DSCASIG4, C.DIFASIG4, 0, 0, 0, '" + wcodusu + "' " _
   & " FROM CAJMPCAB AS C INNER JOIN MAESOCIO AS M " _
   & "   ON C.CODASIG4 = M.CODSOCIO " _
   & " WHERE C.MES >= '" + wDesAno + wDesMes + "' AND " _
   & "       C.MES <= '" + wHasAno + wHasMes + "' AND " _
   & "       C.CODASIG4 = " + Str(wSoc) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_CAJMPSOC " _
   & " (MES, CODENVIO, LIN, CODSOCIO, CODIGO, INS, CARNETPNP, NUMDOC, CODBENI, NOMBRE, " _
   & "  FECENV, FECDSC, TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "  TOTENVIO, DSCCAJMP, DSCDIFER, USU ) " _
   & " SELECT " _
   & "  C.MES, C.CODSOCIO, '5', C.CODASIG5, M.CODIGO, M.INS, M.CARNETPNP, M.NUMDOC, " _
   & "  M.CODBENI, M.NOMBRE, C.FECENV, C.FECDSC, C.TOTASIG5, C.DEUASIG5, " _
   & "  C.ADEASIG5, C.NETASIG5, C.DSCASIG5, C.DIFASIG5, 0, 0, 0, '" + wcodusu + "' " _
   & " FROM CAJMPCAB AS C INNER JOIN MAESOCIO AS M " _
   & "   ON C.CODASIG5 = M.CODSOCIO " _
   & " WHERE C.MES >= '" + wDesAno + wDesMes + "' AND " _
   & "       C.MES <= '" + wHasAno + wHasMes + "' AND " _
   & "       C.CODASIG5 = " + Str(wSoc) + " ")
   Db.CommitTrans

   zz = Leerado2("SELECT MES     , CODSOCIO, NOMBRE, " _
                & "      TOTAPORT, TOTDEUDA, TOTADELA , NETSOCIO, DSCSOCIO, DIFSOCIO, " _
                & "      CODENVIO, LIN     , " _
                & "      CODIGO  , INS     , CARNETPNP, NUMDOC  , CODBENI , FECENV  , " _
                & "      FECDSC  , TOTENVIO, DSCCAJMP , DSCDIFER, USU " _
                & " FROM TMP_CAJMPSOC " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY MES, NOMBRE ")
   Set DataGrid1.DataSource = ADO2

   DataGrid1.Columns(0).Width = 750   ' MES
   DataGrid1.Columns(0).Alignment = dbgRight
   DataGrid1.Columns(0).Caption = "MES"
    
   DataGrid1.Columns(1).Width = 750   ' CODSOCIO
   DataGrid1.Columns(1).Alignment = dbgRight
   DataGrid1.Columns(1).Caption = "SOCIO"
    
   DataGrid1.Columns(2).Width = 4500  ' NOMBRE
   DataGrid1.Columns(2).Alignment = dbgLeft
   DataGrid1.Columns(2).Caption = "NOMBRE ASOCIADO"
    
   DataGrid1.Columns(3).Width = 800    ' TOTAPORT
   DataGrid1.Columns(3).Alignment = dbgRight
   DataGrid1.Columns(3).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(3).Caption = "TOT.APORT"
    
   DataGrid1.Columns(4).Width = 800    ' TOTDEUDA
   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(4).Caption = "DEUDAS"
    
   DataGrid1.Columns(5).Width = 800    ' TOTADELA
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(5).Caption = "ADELANTOS"
    
   DataGrid1.Columns(6).Width = 750    ' NETSOCIO
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(6).Caption = "NETO ENVIO"
    
   DataGrid1.Columns(7).Width = 750    ' DSCSOCIO
   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(7).Caption = "COBRADO"
    
   DataGrid1.Columns(8).Width = 750    ' DIFSOCIO
   DataGrid1.Columns(8).Alignment = dbgRight
   DataGrid1.Columns(8).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(8).Caption = "NO COBRADO"
    
   DataGrid1.Columns(9).Width = 750   ' CODENVIO
   DataGrid1.Columns(9).Alignment = dbgRight
   DataGrid1.Columns(9).Caption = "ENVIO"
    
   DataGrid1.Columns(10).Width = 350   ' LIN
   DataGrid1.Columns(10).Alignment = dbgRight
   DataGrid1.Columns(10).Caption = "LIN"
    
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

   TotalCab
   DataGrid1.SetFocus
End Sub

Private Sub TotalCab()
   Dim zz As Integer, _
       zNetSocio As Currency, _
       zDscSocio As Currency, _
       zDifSocio As Currency
   
   zNetSocio = 0: zDscSocio = 0: zDifSocio = 0
   zz = Leerado8("SELECT SUM(NETSOCIO) AS NETSOCIO, " _
                & "      SUM(DSCSOCIO) AS DSCSOCIO, " _
                & "      SUM(DIFSOCIO) AS DIFSOCIO " _
                & " FROM TMP_CAJMPSOC " _
                & " WHERE USU = '" + wcodusu + "' ")
   If zz > 0 Then
      zNetSocio = IIf(IsNull(ADO8!netsocio), 0, ADO8!netsocio)
      zDscSocio = IIf(IsNull(ADO8!dscsocio), 0, ADO8!dscsocio)
      zDifSocio = IIf(IsNull(ADO8!difsocio), 0, ADO8!difsocio)
   End If
   Set ADO8 = Nothing
   
   lblNetSocio.Caption = Format(zNetSocio, "####,##0.00;;\ ")
   lblDscSocio.Caption = Format(zDscSocio, "####,##0.00;;\ ")
   lblDifSocio.Caption = Format(zDifSocio, "####,##0.00;;\ ")
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(13) As String, _
       wRegAct As Integer, wRegTot As Integer
   Dim wAsi As Integer, wCod As Long, wIns As Integer, _
       wNom As String, wCip As Long, wDni As String, wMes As String, _
       wTotAport As Currency, wTotDeuda As Currency, wTotAdela As Currency, _
       wNetSocio As Currency, wDscSocio As Currency, wDifSocio As Currency
   
   Heading(0) = "NUM"
   Heading(1) = "COD.ENVIO"
   Heading(2) = "COD.SOCIO"
   Heading(3) = "CODOFIN"
   Heading(4) = "CARNET PNP"
   Heading(5) = "DNI"
   Heading(6) = "H"
   Heading(7) = "NOMBRE"
   Heading(8) = "APORT.MES"
   Heading(9) = "DEUDAS"
   Heading(10) = "ADELANTO"
   Heading(11) = "TOT.ENVIO"
   Heading(12) = "TOT.DSCTO"
   Heading(13) = "NO COBRADO"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 14)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 14)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "DESCUENTO CAJA MILITAR POLICIAL POR SOCIO"
        For I = 1 To 14 Step 1
            .Cells(3, I) = Heading(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 6
        objExcel.Columns("B").ColumnWidth = 7
        objExcel.Columns("C").ColumnWidth = 7
        objExcel.Columns("D").ColumnWidth = 10
        objExcel.Columns("E").ColumnWidth = 10
        objExcel.Columns("F").ColumnWidth = 10
        objExcel.Columns("G").ColumnWidth = 5
        objExcel.Columns("H").ColumnWidth = 50
        objExcel.Columns("I").ColumnWidth = 11
        objExcel.Columns("J").ColumnWidth = 11
        objExcel.Columns("K").ColumnWidth = 11
        objExcel.Columns("L").ColumnWidth = 11
        objExcel.Columns("M").ColumnWidth = 11
        objExcel.Columns("N").ColumnWidth = 11
   End With
   
   
   aa = Leerado3("SELECT * FROM TMP_CAJMPSOC " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY MES ")
   If aa > 0 Then
      wRegAct = 1
      wRegTot = aa
      V = 4
      H = 1
      wreg = 1
      wTotEnvio = 0: wDscCajMP = 0: wDscDifer = 0
      Do While Not ADO3.EOF
         DoEvents
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + _
                              Format(wRegAct, "####0") + " / " + _
                              Format(wRegTot, "####0")
         lblMensaje.Refresh
         
         objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 13)).NumberFormat = "#####,##0.00"
         
         objExcel.Cells(V, H + 0) = Format(wRegAct, "####0")
         objExcel.Cells(V, H + 1) = ADO3!codenvio
         objExcel.Cells(V, H + 2) = ADO3!codsocio
         objExcel.Cells(V, H + 3) = Trim(Format(ADO3!codigo, "#######0")) + "-" + Format(ADO3!ins, "9")
         objExcel.Cells(V, H + 4) = ADO3!carnetpnp
         objExcel.Cells(V, H + 5) = ADO3!numdoc
         objExcel.Cells(V, H + 6) = IIf(ADO3!lin = "0", "", "H")
         objExcel.Cells(V, H + 7) = ADO3!nombre
         objExcel.Cells(V, H + 8) = ADO3!totaport
         objExcel.Cells(V, H + 9) = ADO3!totdeuda
         objExcel.Cells(V, H + 10) = ADO3!totadela
         objExcel.Cells(V, H + 11) = ADO3!netsocio
         objExcel.Cells(V, H + 12) = ADO3!dscsocio
         objExcel.Cells(V, H + 13) = ADO3!difsocio
         
         wRegAct = wRegAct + 1
         wTotAport = wTotAport + ADO3!totaport
         wTotDeuda = wTotDeuda + ADO3!totdeuda
         wTotAdela = wTotAdela + ADO3!totadela
         wNetSocio = wNetSocio + ADO3!netsocio
         wDscSocio = wDscSocio + ADO3!dscsocio
         wDifSocio = wDifSocio + ADO3!difsocio
         V = V + 1
         
         wreg = wreg + 1
         ADO3.MoveNext
      Loop
      V = V + 1
      
      objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 13)).NumberFormat = "#####,##0.00"
      objExcel.Range(objExcel.Cells(V, H + 7), objExcel.Cells(V, H + 13)).Font.Color = RGB(255, 0, 0)
      objExcel.Range(objExcel.Cells(V, H + 7), objExcel.Cells(V, H + 13)).Font.Bold = True
      objExcel.Range(objExcel.Cells(V, H + 7), objExcel.Cells(V, H + 13)).Borders.Color = RGB(255, 0, 0)
            
      objExcel.Cells(V, H + 7) = "TOTALES FINALES"
      objExcel.Cells(V, H + 8) = wTotAport
      objExcel.Cells(V, H + 9) = wTotDeuda
      objExcel.Cells(V, H + 10) = wTotAdela
      objExcel.Cells(V, H + 11) = wNetSocio
      objExcel.Cells(V, H + 12) = wDscSocio
      objExcel.Cells(V, H + 13) = wDifSocio
      
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

Private Sub cmdImprimir_Click()
   Dim wAno As String, wMes As String

   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\CajaMPxSoc.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
'   Crys1.Formulas(2) = "NOMMES= 'MES " + Trim(funnommes(wMes)) + " DEL " + wAno + "' "
   Crys1.SelectionFormula = " {TMP_CAJMPSOC.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Form_Initialize
End Sub

Private Sub Form_Initialize()
   frmMPConxSoc.Left = (Screen.Width - Width) \ 2
   frmMPConxSoc.Top = 0
   
   Limpiar
   
   txtDesde.SetFocus
End Sub

Private Sub CalxSoc()
End Sub

Private Sub txtCodSocio_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
   If aa > 0 Then
      lblCodSocio.Caption = ADO6a!nombre
   Else
      lblCodSocio.Caption = ""
      Limpiar
   End If
   Set ADO6a = Nothing
End Sub

Private Sub txtCodSocio_GotFocus()
   txtCodSocio.SelStart = 0
   If Len(Trim(txtCodSocio.Text)) > 0 Then
      txtCodSocio.SelLength = Len(Trim(txtCodSocio.Text))
   Else
      txtCodSocio.SelLength = 8
   End If
End Sub

Private Sub txtCodSocio_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtCodSocio.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtCodSocio_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtCodSocio.Text)) = 0 Then
         MsgBox "Codigo Socio En Blanco", vbExclamation
         txtCodSocio.Text = ""
         Exit Sub
      End If
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
      If aa = 0 Then
         MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
         txtCodSocio.Text = ""
         Exit Sub
      End If
      txtCodigo.Text = ADO8!codigo
      txtIns.Text = ADO8!ins
      txtNumdoc.Text = ADO8!numdoc
      txtE_socio.Text = ADO8!e_socio
      txtGrado.Text = ADO8!grado
      txtTipCob.Text = ADO8!tipcob
   
      cmdBuscar.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDesde_Change()
   Dim wMes As String, wAno As String
   wAno = Left(txtDesde.Text, 4)
   wMes = Right(txtDesde.Text, 2)
   
   lblDesde.Caption = Trim(funnommes(wMes)) + " Del " + wAno
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
      If wAno < "2016" Or wAno > "2030" Then
         MsgBox "Año del Mes Inicial fuera de Rango", vbExclamation
         txtDesde.Text = "____-__"
         txtDesde.SetFocus
         Exit Sub
      End If
      If wMes <> "01" And wMes <> "02" And wMes <> "03" And wMes <> "04" And _
         wMes <> "05" And wMes <> "06" And wMes <> "07" And wMes <> "08" And _
         wMes <> "09" And wMes <> "10" And wMes <> "11" And wMes <> "12" Then
         MsgBox "Mes del Mes Inicial fuera de Rango", vbExclamation
         txtDesde.Text = "____-__"
         txtDesde.SetFocus
         Exit Sub
      End If
      txtHasta.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtE_socio_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + txtE_socio.Text + "' ")
   If aa > 0 Then
      lblE_socio.Caption = ADO6a!nombre
   Else
      lblE_socio.Caption = ""
   End If
   Set ADO6a = Nothing
End Sub

Private Sub txtGrado_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAEGRADO WHERE GRADO = " + Str(Val(txtGrado.Text)) + " ")
   If aa > 0 Then
      lblGrado.Caption = ADO6a!nombre
   Else
      lblGrado.Caption = ""
   End If
   Set ADO6a = Nothing
End Sub

Private Sub txtTipCob_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAETIPCOB WHERE TIPCOB = '" + txtTipCob.Text + "' ")
   If aa > 0 Then
      lblTipCob.Caption = ADO6a!nombre
   Else
      lblTipCob.Caption = ""
   End If
   Set ADO6a = Nothing
End Sub
Private Sub txtHasta_Change()
   Dim wMes As String, wAno As String
   wAno = Left(txtHasta.Text, 4)
   wMes = Right(txtHasta.Text, 2)
   
   lblHasta.Caption = Trim(funnommes(wMes)) + " Del " + wAno
End Sub

Private Sub txtHasta_GotFocus()
   txtHasta.SelStart = 0
   txtHasta.SelLength = 7
End Sub

Private Sub txtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDesde.SetFocus
   Case 40
        txtCodSocio.SetFocus
   End Select
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
   Dim wMes As String, wAno As String, _
       wDesMes As String, wDesAno As String, _
       wHasMes As String, wHasAno As String
   If KeyAscii = 13 Then
      If txtHasta.Text = "____-__" Then
         MsgBox "Mes Final En Blanco", vbExclamation
         txtHasta.Text = "____-__"
         txtHasta.SetFocus
         Exit Sub
      End If
      wAno = Left(txtHasta.Text, 4)
      wMes = Right(txtHasta.Text, 2)
      If wAno < "2016" Or wAno > "2030" Then
         MsgBox "Año del Mes Final fuera de Rango", vbExclamation
         txtHasta.Text = "____-__"
         txtHasta.SetFocus
         Exit Sub
      End If
      If wMes <> "01" And wMes <> "02" And wMes <> "03" And wMes <> "04" And _
         wMes <> "05" And wMes <> "06" And wMes <> "07" And wMes <> "08" And _
         wMes <> "09" And wMes <> "10" And wMes <> "11" And wMes <> "12" Then
         MsgBox "Mes del Mes Inicial fuera de Rango", vbExclamation
         txtHasta.Text = "____-__"
         txtHasta.SetFocus
         Exit Sub
      End If
      wDesAno = Left(txtDesde.Text, 4)
      wDesMes = Right(txtDesde.Text, 2)
      wHasAno = Left(txtHasta.Text, 4)
      wHasMes = Right(txtHasta.Text, 2)
      If wDesAno + wDesMes > wHasAno + wHasMes Then
         MsgBox "Rango de Meses Digitado Es Invalido", vbExclamation
         txtHasta.Text = "____-__"
         txtHasta.SetFocus
         Exit Sub
      End If
      txtCodSocio.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub
