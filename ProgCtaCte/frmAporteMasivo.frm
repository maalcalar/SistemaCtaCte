VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAporteMasivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estado De Aportes Masivo"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   15825
   Begin VB.TextBox txtCodSocio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   9
      TabIndex        =   18
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdExpDet 
      Caption         =   "&Exportar Detalle"
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
      Left            =   10680
      TabIndex        =   17
      Top             =   8520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdImpDet 
      Caption         =   "&Imprimir Provisión Mensual"
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
      Left            =   10680
      TabIndex        =   16
      Top             =   7800
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   9840
      TabIndex        =   12
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox cmbTipCob 
      Height          =   315
      ItemData        =   "frmAporteMasivo.frx":0000
      Left            =   1800
      List            =   "frmAporteMasivo.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   860
      Width           =   6015
   End
   Begin VB.CommandButton cmdAporte 
      Caption         =   "Imprimir Estado de Pagos"
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
      Left            =   9360
      TabIndex        =   8
      Top             =   7800
      Width           =   1095
   End
   Begin VB.ComboBox cmbE_Socio 
      Height          =   315
      ItemData        =   "frmAporteMasivo.frx":0004
      Left            =   1800
      List            =   "frmAporteMasivo.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   500
      Width           =   6015
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
      Left            =   13320
      TabIndex        =   3
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir Resumen"
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
      Left            =   12000
      TabIndex        =   2
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar Resumen"
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
      TabIndex        =   1
      Top             =   8520
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   10610
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
      Left            =   13320
      Top             =   0
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
   Begin Crystal.CrystalReport Crys2 
      Left            =   12480
      Top             =   0
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
   Begin MSMask.MaskEdBox txtTope 
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Top             =   200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblCodSocio 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      TabIndex        =   20
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Index           =   0
      Left            =   300
      TabIndex        =   19
      Top             =   1200
      Width           =   1410
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
      TabIndex        =   15
      Top             =   200
      Width           =   1410
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
      Left            =   2640
      TabIndex        =   14
      Top             =   200
      Width           =   1935
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tipo Cobro"
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
      Index           =   18
      Left            =   240
      TabIndex        =   11
      Top             =   860
      Width           =   1410
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
      Left            =   240
      TabIndex        =   9
      Top             =   7920
      Width           =   8055
   End
   Begin VB.Label Label23 
      Caption         =   "Cantidad Socios"
      Height          =   255
      Left            =   8040
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblTotal 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   8040
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Estado de Socio"
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
      Index           =   16
      Left            =   240
      TabIndex        =   5
      Top             =   500
      Width           =   1410
   End
End
Attribute VB_Name = "frmAporteMasivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub cmbE_Socio_Click()
'   cmbE_Socio_KeyPress (13)
'End Sub

Private Sub cmbE_Socio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmbTipCob.SetFocus
   End If
End Sub

Private Sub cmbTipCob_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtCodSocio.SetFocus
   End If
End Sub

Private Sub cmdAporte_Click()
   Dim aa As Long, wRegAct As Long, wRegTot As Long, wSoc As Integer

   Dim wCod As Long, wIns As Long, wNom As String, wLin As Integer, _
       wRec As String, wMon As String, wImp As Currency, wFec As Date, _
       wObs As String, wNde As Currency, wnCr As Currency, wDeu As Currency, _
       wCer As Currency, wAde As Currency, wFecTope As Date, _
       wMesTope As String, wAnoTope As String, wDiaTope As String, _
       wVip As String, wCartaDieco As String, wFracSw As Boolean, wRen As Currency, _
       wSdoOld As Currency, wSdoGra As Currency, _
       wFracCargos As Currency, wFracAbonos As Currency, wFracSdoNew As Currency
   
   lblMensaje.Caption = "Preparando Archivo....."
   lblMensaje.Refresh

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_ESTADO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   wMesTope = Right(txtTope.Text, 2)
   wAnoTope = Left(txtTope.Text, 4)
   wDiaTope = fundiames(wMesTope)
   wFecTope = Format(wDiaTope + "/" + wMesTope + "/" + wAnoTope, "dd/mm/yyyy")
   wVip = ""

   aa = Leerado8("SELECT * FROM TMP_MASIVO WHERE USU = '" + wcodusu + "' ")
   If aa > 0 Then
      ADO8.MoveFirst
      wRegAct = 1
      wRegTot = aa
      Do While Not ADO8.EOF
         wSoc = ADO8!codsocio
         wCod = ADO8!codigo
         wIns = ADO8!ins
         wFracSw = False
         wRen = 0
         wFracCargos = 0: wFracAbonos = 0: wFracSdoNew = 0
   
         lblMensaje.Caption = "Socio " + Str(ADO8!codsocio) + " " + ADO8!nombre
         lblMensaje.Refresh
         
'         Db.BeginTrans
'         Db.Execute ("DELETE FROM TMP_FRACDET WHERE USU = '" + wcodusu + "'")
'         Db.CommitTrans
         
         Db.BeginTrans
         Db.Execute ("DELETE FROM TMP_ESTADO WHERE USU = '" + wcodusu + "' AND CODSOCIO = " + Str(wSoc) + " ")
         Db.CommitTrans

         aa = Leerado6a("SELECT SUM(CARGOS - ABONOS) AS DIFER " _
                    & " FROM CTASXDET " _
                    & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                    & "       CONCEPTO = '02' ")
         If aa > 0 Then
            wRen = IIf(IsNull(ADO6a!difer), 0, ADO6a!difer)
         End If
         Set ADO6a = Nothing
         
         wSdoOld = 0: wSdoGra = 0
         aa = Leerado7a("SELECT " _
                    & "  " + Str(wSoc) + ", '6', D.LINEA, " _
                    & "  D.NUMERO, D.LINEA, D.VCMTO, D.CARGOS, D.ABONOS, " _
                    & "  D.SDONEW, C.SDOPEN, D.FECCOB, '" + wcodusu + "' " _
                    & " FROM FRACDET AS D INNER JOIN FRACCAB AS C " _
                    & "   ON D.NUMERO = C.NUMERO " _
                    & " WHERE C.CODSOCIO = " + Str(wSoc) + " " _
                    & " ORDER BY D.LINEA")
         If aa > 0 Then
            ADO7a.MoveFirst
            
            aa = Leerado6a("SELECT * FROM FRACCAB WHERE NUMERO = '" + ADO7a!numero + "'  ")
            If aa > 0 Then
               wSdoOld = ADO6a!sdopen
            End If
                    
            Do While Not ADO7a.EOF
               wSdoGra = wSdoOld - ADO7a!cargos
         
               Db.BeginTrans
               Db.Execute ("INSERT INTO TMP_ESTADO " _
               & " (CODSOCIO, TIPOREG, LINEA, " _
               & "  FRACNUMERO, FRACLINEA, FRACVCMTO, FRACCARGOS, " _
               & "  FRACABONOS, FRACSDONEW, FRACSDOOLD, FRACSDOGRA, FRACFECCOB, USU) " _
               & " VALUES " _
               & " (" + Str(wSoc) + ", '6', '" + ADO7a!linea + "', " _
               & "  '" + ADO7a!numero + "', '" + ADO7a!linea + "', " _
               & "  '" + Format(ADO7a!vcmto, "dd/mm/yyyy") + "', " _
               & "  " + Str(ADO7a!cargos) + ", " + Str(ADO7a!abonos) + ", " _
               & "  " + Str(ADO7a!sdonew) + ", " + Str(wSdoOld) + ", " _
               & "  " + Str(wSdoGra) + ", '" + Format(ADO7a!feccob, "dd/mm/yyyy") + "', '" + wcodusu + "' ) ")
               Db.CommitTrans
         
               wSdoOld = wSdoGra
               wFracCargos = wFracCargos + ADO7a!cargos
               wFracAbonos = wFracAbonos + ADO7a!abonos
               wFracSdoNew = wFracCargos - wFracAbonos
         
               ADO7a.MoveNext
            Loop
         End If
         
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_MASIVO " _
         & " SET RENOVA = " + Str(wRen) + "," _
         & "     FRACCARGOS = " + Str(wFracCargos) + ", " _
         & "     FRACABONOS = " + Str(wFracAbonos) + ", " _
         & "     FRACSDONEW = " + Str(wFracSdoNew) + " " _
         & " WHERE      USU = '" + wcodusu + "' AND " _
         & "       CODSOCIO = " + Str(wSoc) + " ")
         Db.CommitTrans
         
         aa = Leerado7("SELECT Z.* " _
                & " FROM ZZZ_MRECIBOS AS Z INNER JOIN ZZZ_CONCEPTO AS M " _
                & "   ON Z.CONCEPTO = M.CCONCE " _
                & " WHERE Z.CODIGO = " + Str(wCod) + " AND " _
                & "          Z.INS = " + Str(wIns) + " AND " _
                & "      (Z.MARCA2 <> 'A' OR Z.MARCA2 IS NULL) AND " _
                & "      (M.MARCA = 'S') " _
                & " ORDER BY Z.FECHA_PAGO, Z.SERIE, Z.NRO_COMP ")
        
'                & "      (Z.FECHA_PAGO <= '" + Format(wFecTope, "dd/mm/yyyy") + "')  " _

        If aa > 0 Then
            ADO7.MoveFirst
            wLin = 1
            Do While Not ADO7.EOF
               wRec = ADO7!serie + "-" + Format(ADO7!nro_comp, "000000")
               wMon = IIf(ADO7!moneda = "S/." Or ADO7!moneda = "S", "S", "D")
               wImp = ADO7!monto
               wFec = Format(ADO7!fecha_pago, "dd/mm/yyyy")
               wObs = Trim(IIf(IsNull(ADO7!obs), "", ADO7!obs))
   
                  Db.BeginTrans
                  Db.Execute ("INSERT INTO TMP_ESTADO " _
                  & " (CODSOCIO, LINEA, CODIGO, INS, TIPOREG, RECIBO, MONEDA, IMPORTE, FECHA, CONCEPTO, USU) " _
                  & " VALUES " _
                  & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', " + Str(wCod) + ", " + Str(wIns) + ", " _
                  & "  '2', '" + wRec + "', '" + wMon + "', " _
                  & "  " + Str(wImp) + ", '" + Format(wFec, "dd/mm/yyyy") + "', " _
                  & "  '" + GlosaLibre(wObs) + "', '" + wcodusu + "' ) ")
                  Db.CommitTrans
   
               wLin = wLin + 1
               ADO7.MoveNext
            Loop
         End If
         Set ADO7 = Nothing
   
         aa = Leerado7("SELECT * FROM ZZZ_BCORECAU " _
                & " WHERE CODIGO = " + Str(wCod) + " AND " _
                & "          INS = " + Str(wIns) + " AND " _
                & "        FECHA <= '" + Format(wFecTope, "dd/mm/yyyy") + "' " _
                & " ORDER BY FECHA, RECIBO ")
         If aa > 0 Then
            ADO7.MoveFirst
            wLin = 1
            Do While Not ADO7.EOF
               wRec = Format(ADO7!recibo, "000000")
               wMon = IIf(ADO7!moneda = "S/.", "S", "D")
               wImp = ADO7!aporte
               wFec = Format(ADO7!fecha, "dd/mm/yyyy")
               wnCr = ADO7!ncredito
               wNde = ADO7!ndebito
               wDeu = ADO7!deuda_pt2
               wCer = ADO7!dins_cer
               wAde = ADO7!adelanto
   
               Db.BeginTrans
               Db.Execute ("INSERT INTO TMP_ESTADO " _
               & " (CODSOCIO, LINEA, CODIGO, INS, TIPOREG, BCORECIBO, BCOMONEDA, BCONCREDITO, BCONDEBITO, " _
               & "  BCOAPORTE, BCOFECHA, USU) " _
               & " VALUES " _
               & " (" + Str(wSoc) + ", " + Format(wLin, "0000") + ", " + Str(wCod) + ", " + Str(wIns) + ", " _
               & "  '3', '" + wRec + "', '" + wMon + "', " _
               & "  " + Str(wnCr) + ", " + Str(wNde) + ", " + Str(wImp) + ", " _
               & "  '" + Format(wFec, "dd/mm/yyyy") + "', '" + wcodusu + "' ) ")
               Db.CommitTrans
               
               wLin = wLin + 1
               ADO7.MoveNext
            Loop
         End If
         Set ADO7 = Nothing
         
         aa = Leerado7("SELECT * FROM ZZZ_DEVOL " _
                & " WHERE CODIGO = " + Str(wCod) + " AND " _
                & "          INS = " + Str(wIns) + " AND " _
                & "        FECHA <= '" + Format(wFecTope, "dd/mm/yyyy") + "' " _
                & " ORDER BY FECHA, SERIE, NRO_COMP ")
        If aa > 0 Then
            ADO7.MoveFirst
            wLin = 1
            Do While Not ADO7.EOF
               wRec = ADO7!serie + "-" + Format(ADO7!nro_comp, "000000")
               wMon = "S"
               wImp = ADO7!importe
               wFec = Format(ADO7!fecha, "dd/mm/yyyy")
               wObs = Trim(IIf(IsNull(ADO7!glosa), "", ADO7!glosa))
   
                  Db.BeginTrans
                  Db.Execute ("INSERT INTO TMP_ESTADO " _
                  & " (CODSOCIO, LINEA, CODIGO, INS, TIPOREG, RECIBO, DEVMONEDA, DEVIMPORTE, FECHA, CONCEPTO, USU) " _
                  & " VALUES " _
                  & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', " + Str(wCod) + ", " + Str(wIns) + ", " _
                  & "  '4', '" + wRec + "', '" + wMon + "', " _
                  & "  " + Str(wImp) + ", '" + Format(wFec, "dd/mm/yyyy") + "', " _
                  & "  '" + GlosaLibre(wObs) + "', '" + wcodusu + "' ) ")
                  Db.CommitTrans
   
               wLin = wLin + 1
               ADO7.MoveNext
            Loop
         End If
         Set ADO7 = Nothing
         
         aa = Leerado7("SELECT * FROM ZZZ_APOR_PLA " _
                & " WHERE CODIGO = " + Str(wCod) + " AND " _
                & "          INS = " + Str(wIns) + " AND " _
                & "       CUOANO <= '" + wAnoTope + "'  " _
                & " ORDER BY CUOANO ")
         If aa > 0 Then
            ADO7.MoveFirst
            wLin = 1
            Do While Not ADO7.EOF
               
               If ADO7!cuoano < wAnoTope Then
               
                  Db.BeginTrans
                  Db.Execute ("INSERT INTO TMP_ESTADO " _
                  & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                  & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                  & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                  & "  TOTAL, DEUDA, USU ) " _
                  & " VALUES " _
                  & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                  & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                  & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                  & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", " + Str(ADO7!impo06) + ", " _
                  & "  " + Str(ADO7!impo07) + ", " + Str(ADO7!impo08) + ", " + Str(ADO7!impo09) + ", " _
                  & "  " + Str(ADO7!impo10) + ", " + Str(ADO7!impo11) + ", " + Str(ADO7!impo12) + ", " _
                  & "  " + Str(ADO7!totimpo) + ", " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                  Db.CommitTrans
               Else
                  If ADO7!cuoano = wAnoTope Then
                     Select Case wMesTope
                     Case "01"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  " + Str(ADO7!impo01) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "02"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", 0, " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "03"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "04"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03 + ADO7!impo04) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "05"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", 0, " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03 + ADO7!impo04 + ADO7!impo05) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "06"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", " + Str(ADO7!impo06) + ", " _
                          & "  0, 0, 0, " _
                          & "  0, 0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03 + ADO7!impo04 + ADO7!impo05 + ADO7!impo06) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "07"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", " + Str(ADO7!impo06) + ", " _
                          & "  " + Str(ADO7!impo07) + ", 0, 0, " _
                          & "  0,0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03 + ADO7!impo04 + ADO7!impo05 + ADO7!impo06 + ADO7!impo07) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "08"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", " + Str(ADO7!impo06) + ", " _
                          & "  " + Str(ADO7!impo07) + ", " + Str(ADO7!impo08) + ", 0, " _
                          & "  0,0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03 + ADO7!impo04 + ADO7!impo05 + ADO7!impo06 + ADO7!impo07 + ADO7!impo08) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "09"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", " + Str(ADO7!impo06) + ", " _
                          & "  " + Str(ADO7!impo07) + ", " + Str(ADO7!impo08) + ", " + Str(ADO7!impo09) + ", " _
                          & "  0,0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03 + ADO7!impo04 + ADO7!impo05 + ADO7!impo06 + ADO7!impo07 + ADO7!impo08 + ADO7!impo09) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "10"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", " + Str(ADO7!impo06) + ", " _
                          & "  " + Str(ADO7!impo07) + ", " + Str(ADO7!impo08) + ", " + Str(ADO7!impo09) + ", " _
                          & "  " + Str(ADO7!impo10) + ",0, 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03 + ADO7!impo04 + ADO7!impo05 + ADO7!impo06 + ADO7!impo07 + ADO7!impo08 + ADO7!impo09 + ADO7!impo10) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "11"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", " + Str(ADO7!impo06) + ", " _
                          & "  " + Str(ADO7!impo07) + ", " + Str(ADO7!impo08) + ", " + Str(ADO7!impo09) + ", " _
                          & "  " + Str(ADO7!impo10) + ", " + Str(ADO7!impo11) + ", 0, " _
                          & "  " + Str(ADO7!impo01 + ADO7!impo02 + ADO7!impo03 + ADO7!impo04 + ADO7!impo05 + ADO7!impo06 + ADO7!impo07 + ADO7!impo08 + ADO7!impo09 + ADO7!impo10 + ADO7!impo11) + ", " _
                          & "  " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     Case "12"
                          Db.BeginTrans
                          Db.Execute ("INSERT INTO TMP_ESTADO " _
                          & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
                          & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
                          & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
                          & "  TOTAL, DEUDA, USU ) " _
                          & " VALUES " _
                          & " (" + Str(wSoc) + ", '" + Format(wLin, "0000") + "', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
                          & "  '" + ADO7!cuoano + "',  '" + ADO7!tipapor + "', '', " _
                          & "  " + Str(ADO7!impo01) + ", " + Str(ADO7!impo02) + ", " + Str(ADO7!impo03) + ", " _
                          & "  " + Str(ADO7!impo04) + ", " + Str(ADO7!impo05) + ", " + Str(ADO7!impo06) + ", " _
                          & "  " + Str(ADO7!impo07) + ", " + Str(ADO7!impo08) + ", " + Str(ADO7!impo09) + ", " _
                          & "  " + Str(ADO7!impo10) + ", " + Str(ADO7!impo11) + ", " + Str(ADO7!impo12) + ", " _
                          & "  " + Str(ADO7!totimpo) + ", " + Str(IIf(IsNull(ADO7!deuda_pt2), 0, ADO7!deuda_pt2)) + ", '" + wcodusu + "') ")
                          Db.CommitTrans
                     End Select
                  
                  End If
               End If
               
               wLin = wLin + 1
               ADO7.MoveNext
            Loop
         End If
         Set ADO7 = Nothing
   
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_ESTADO " _
         & " SET TOTAL = IMP01 + IMP02 + IMP03 + IMP04 + IMP05 + IMP06 + " _
         & "             IMP07 + IMP08 + IMP09 + IMP10 + IMP11 + IMP12 " _
         & " WHERE USU = '" + wcodusu + "' ")
         Db.CommitTrans
   
         ADO8.MoveNext
      Loop
   End If

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ESTADO " _
   & " SET NOMCOB = 'DIECO 1' " _
   & " WHERE (TIPCOB = '1') AND " _
   & "       (USU = '" + wcodusu + "') ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ESTADO " _
   & " SET NOMCOB = 'DIECO 2' " _
   & " WHERE (TIPCOB = '2') AND " _
   & "       (USU = '" + wcodusu + "') ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_ESTADO " _
   & " SET NOMCOB = 'CAJA MP' " _
   & " WHERE (TIPCOB = '4') AND " _
   & "       (USU = '" + wcodusu + "') ")
   Db.CommitTrans
   
   aa = Leerado8("SELECT * FROM TMP_MASIVO WHERE USU = '" + wcodusu + "' ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wSoc = ADO8!codsocio
         wCod = ADO8!codigo
         wIns = ADO8!ins

         aa = Leerado7("SELECT * FROM TMP_ESTADO WHERE USU = '" + wcodusu + "' AND CODSOCIO = " + Str(wSoc) + " ")
         If aa = 0 Then
            Db.BeginTrans
            Db.Execute ("INSERT INTO TMP_ESTADO " _
            & " (CODSOCIO, LINEA, TIPOREG, CODIGO, INS, ANO, TIPCOB, NOMCOB, " _
            & "  IMP01, IMP02, IMP03, IMP04, IMP05, IMP06, " _
            & "  IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, " _
            & "  TOTAL, DEUDA, USU ) " _
            & " VALUES " _
            & " (" + Str(wSoc) + ", '0001', '1', " + Str(wCod) + ", " + Str(wIns) + ", " _
            & "  '',  '', '', " _
            & "  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '" + wcodusu + "') ")
            Db.CommitTrans
         End If

         ADO8.MoveNext
      Loop
   End If

   wVip = "": wCartaDieco = ""
   aa = Leerado8a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
   If aa > 0 Then
      wVip = IIf(ADO8a!vip = True, "SOCIO VIP", "")
      wCartaDieco = IIf(ADO8a!cartadieco = True, "ASOCIADO SIN CARTA AUTORIZACION DIECO", "")
   End If

   lblMensaje.Caption = ""
   lblMensaje.Refresh

   Crys2.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys2.ReportFileName = xraiz + "ReportCtaCte\EstadoCtaMasivo.RPT"
   Crys2.SelectionFormula = " {TMP_ESTADO.USU}='" + wcodusu + "' "
   Crys2.WindowState = crptMaximized
   Crys2.Action = 1

End Sub

Private Sub cmdBuscar_Click()
   lblTotal.Caption = ""
   Set DataGrid1.DataSource = Nothing
   
   LlenaCab
   TotalCab
   DataGrid1.SetFocus
End Sub

Private Sub cmdExpDet_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(17) As String, wRegAct As Integer, wRegTot As Integer
   Dim wFecIng As Date, wE_S As String, wNom As String, _
   wImp01 As Currency, wImp02 As Currency, wImp03 As Currency, wImp04 As Currency, _
   wImp05 As Currency, wImp06 As Currency, wImp07 As Currency, wImp08 As Currency, _
   wImp09 As Currency, wImp10 As Currency, wImp11 As Currency, wImp12 As Currency, _
   zImp01 As Currency, zImp02 As Currency, zImp03 As Currency, zImp04 As Currency, _
   zImp05 As Currency, zImp06 As Currency, zImp07 As Currency, zImp08 As Currency, _
   zImp09 As Currency, zImp10 As Currency, zImp11 As Currency, zImp12 As Currency, _
   wTotal As Currency, zTotal As Currency
   Heading(0) = "COD.SOCIO"
   Heading(1) = "CODIGO"
   Heading(2) = "INS"
   Heading(3) = "E_SOC"
   Heading(4) = "NOMBRE"
   Heading(5) = "ENE"
   Heading(6) = "FEB"
   Heading(7) = "MAR"
   Heading(8) = "ABR"
   Heading(9) = "MAY"
   Heading(10) = "JUN"
   Heading(11) = "JUL"
   Heading(12) = "AGO"
   Heading(13) = "SET"
   Heading(14) = "OCT"
   Heading(15) = "NOV"
   Heading(16) = "DIC"
   Heading(17) = "TOT"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 18)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 18)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "DETALLE DE APORTES DE ASOCIADOS - MES TOPE " + txtTope.Text
        For I = 1 To 18 Step 1
            .Cells(3, I) = Heading(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 8
        objExcel.Columns("B").ColumnWidth = 10
        objExcel.Columns("C").ColumnWidth = 4
        objExcel.Columns("D").ColumnWidth = 6
        objExcel.Columns("E").ColumnWidth = 55
        objExcel.Columns("F").ColumnWidth = 10
        objExcel.Columns("G").ColumnWidth = 10
        objExcel.Columns("H").ColumnWidth = 10
        objExcel.Columns("I").ColumnWidth = 10
        objExcel.Columns("J").ColumnWidth = 10
        objExcel.Columns("K").ColumnWidth = 10
        objExcel.Columns("L").ColumnWidth = 10
        objExcel.Columns("M").ColumnWidth = 10
        objExcel.Columns("N").ColumnWidth = 10
        objExcel.Columns("O").ColumnWidth = 10
        objExcel.Columns("P").ColumnWidth = 10
        objExcel.Columns("Q").ColumnWidth = 10
        objExcel.Columns("R").ColumnWidth = 10
   End With
   
   aa = Leerado3("SELECT C.CODSOCIO, C.CODIGO, C.INS, C.E_SOCIO, C.NOMBRE, D.ANO, D.IMP01, D.IMP02, D.IMP03, D.IMP04, D.IMP05, IMP06, IMP07, IMP08, IMP09, IMP10, IMP11, IMP12, TOTAL " _
                & " FROM TMP_MASIVODET AS D INNER JOIN TMP_MASIVO AS C " _
                & "   ON D.CODSOCIO = C.CODSOCIO AND " _
                & "      D.USU = C.USU " _
                & " WHERE D.USU = '" + wcodusu + "' " _
                & " ORDER BY C.E_SOCIO, C.NOMBRE, D.ANO ")
   If aa > 0 Then
      wRegAct = 1
      wRegTot = aa
      V = 4: H = 1
      zImp01 = 0: zImp02 = 0: zImp03 = 0: zImp04 = 0: zImp05 = 0: zImp06 = 0
      zImp07 = 0: zImp08 = 0: zImp09 = 0: zImp10 = 0: zImp11 = 0: zImp12 = 0: zTotal = 0
      Do While Not ADO3.EOF
         
         wNom = ADO3!nombre
         wImp01 = 0: wImp02 = 0: wImp03 = 0: wImp04 = 0: wImp05 = 0: wImp06 = 0
         wImp07 = 0: wImp08 = 0: wImp09 = 0: wImp10 = 0: wImp11 = 0: wImp12 = 0: wTotal = 0
         Do While ADO3!nombre = wNom
            DoEvents
            lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wRegAct, "####0") + " / " + Format(wRegTot, "####0")
            lblMensaje.Refresh
         
            objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 17)).NumberFormat = "####,##0.00"
         
            objExcel.Cells(V, H + 0) = ADO3!codsocio
            objExcel.Cells(V, H + 1) = ADO3!codigo
            objExcel.Cells(V, H + 2) = ADO3!ins
            objExcel.Cells(V, H + 3) = ADO3!e_socio
            objExcel.Cells(V, H + 4) = ADO3!nombre
            objExcel.Cells(V, H + 5) = ADO3!imp01
            objExcel.Cells(V, H + 6) = ADO3!imp02
            objExcel.Cells(V, H + 7) = ADO3!imp03
            objExcel.Cells(V, H + 8) = ADO3!imp04
            objExcel.Cells(V, H + 9) = ADO3!imp05
            objExcel.Cells(V, H + 10) = ADO3!imp06
            objExcel.Cells(V, H + 11) = ADO3!imp07
            objExcel.Cells(V, H + 12) = ADO3!imp08
            objExcel.Cells(V, H + 13) = ADO3!imp09
            objExcel.Cells(V, H + 14) = ADO3!imp10
            objExcel.Cells(V, H + 15) = ADO3!imp11
            objExcel.Cells(V, H + 16) = ADO3!imp12
            objExcel.Cells(V, H + 17) = ADO3!Total
         
            wImp01 = wImp01 + ADO3!imp01
            wImp02 = wImp02 + ADO3!imp02
            wImp03 = wImp03 + ADO3!imp03
            wImp04 = wImp04 + ADO3!imp04
            wImp05 = wImp05 + ADO3!imp05
            wImp06 = wImp06 + ADO3!imp06
            wImp07 = wImp07 + ADO3!imp07
            wImp08 = wImp08 + ADO3!imp08
            wImp09 = wImp09 + ADO3!imp09
            wImp10 = wImp10 + ADO3!imp10
            wImp11 = wImp11 + ADO3!imp11
            wImp12 = wImp12 + ADO3!imp12
            wTotal = wTotal + ADO3!Total
            
            zImp01 = zImp01 + ADO3!imp01
            zImp02 = zImp02 + ADO3!imp02
            zImp03 = zImp03 + ADO3!imp03
            zImp04 = zImp04 + ADO3!imp04
            zImp05 = zImp05 + ADO3!imp05
            zImp06 = zImp06 + ADO3!imp06
            zImp07 = zImp07 + ADO3!imp07
            zImp08 = zImp08 + ADO3!imp08
            zImp09 = zImp09 + ADO3!imp09
            zImp10 = zImp10 + ADO3!imp10
            zImp11 = zImp11 + ADO3!imp11
            zImp12 = zImp12 + ADO3!imp12
            zTotal = zTotal + ADO3!Total
            
            wRegAct = wRegAct + 1
            V = V + 1
            ADO3.MoveNext
            If ADO3.EOF Then
               Exit Do
            End If
         Loop
         
         V = V + 1
         objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 17)).NumberFormat = "####,##0.00"
         objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 17)).Font.Bold = True
         objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 17)).Font.Color = RGB(255, 0, 0)
         objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 17)).Borders.LineStyle = xlContinuous
         objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 17)).Borders.Color = RGB(255, 0, 0)
         
         objExcel.Cells(V, H + 4) = "TOTALES " + wNom
         objExcel.Cells(V, H + 5) = wImp01
         objExcel.Cells(V, H + 6) = wImp02
         objExcel.Cells(V, H + 7) = wImp03
         objExcel.Cells(V, H + 8) = wImp04
         objExcel.Cells(V, H + 9) = wImp05
         objExcel.Cells(V, H + 10) = wImp06
         objExcel.Cells(V, H + 11) = wImp07
         objExcel.Cells(V, H + 12) = wImp08
         objExcel.Cells(V, H + 13) = wImp09
         objExcel.Cells(V, H + 14) = wImp10
         objExcel.Cells(V, H + 15) = wImp11
         objExcel.Cells(V, H + 16) = wImp12
         objExcel.Cells(V, H + 17) = wTotal
         V = V + 2
         
      Loop
      
      If wTotal <> zTotal Then
         V = V + 1
         objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 17)).NumberFormat = "####,##0.00"
         objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 17)).Font.Bold = True
         objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 17)).Font.Color = RGB(255, 0, 0)
         objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 17)).Borders.LineStyle = xlContinuous
         objExcel.Range(objExcel.Cells(V, H + 4), objExcel.Cells(V, H + 17)).Borders.Color = RGB(255, 0, 0)
         
         objExcel.Cells(V, H + 4) = "TOTALES FINALES "
         objExcel.Cells(V, H + 5) = zImp01
         objExcel.Cells(V, H + 6) = zImp02
         objExcel.Cells(V, H + 7) = zImp03
         objExcel.Cells(V, H + 8) = zImp04
         objExcel.Cells(V, H + 9) = zImp05
         objExcel.Cells(V, H + 10) = zImp06
         objExcel.Cells(V, H + 11) = zImp07
         objExcel.Cells(V, H + 12) = zImp08
         objExcel.Cells(V, H + 13) = zImp09
         objExcel.Cells(V, H + 14) = zImp10
         objExcel.Cells(V, H + 15) = zImp11
         objExcel.Cells(V, H + 16) = zImp12
         objExcel.Cells(V, H + 17) = zTotal
         V = V + 1
      End If
      
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

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(11) As String, wRegAct As Integer, wRegTot As Integer
   Dim wFecIng As Date, wE_S As String, _
   wTotApo As Currency, wTotCob As Currency, _
   zTotApo As Currency, zTotCob As Currency
   Heading(0) = "NUM."
   Heading(1) = "COD.SOCIO"
   Heading(2) = "CODIGO"
   Heading(3) = "INS"
   Heading(4) = "E_SOC"
   Heading(5) = "NOMBRE"
   Heading(6) = "FEC.ING"
   Heading(7) = "DESDE"
   Heading(8) = "HASTA"
   Heading(9) = "TOT.APORTES"
   Heading(10) = "TOT.COBRADO"
   Heading(11) = "DIFERENCIA"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 12)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 12)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "RESUMEN DE ASOCIADOS - MES TOPE " + txtTope.Text
        For I = 1 To 12 Step 1
            .Cells(3, I) = Heading(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 6
        objExcel.Columns("B").ColumnWidth = 8
        objExcel.Columns("C").ColumnWidth = 10
        objExcel.Columns("D").ColumnWidth = 4
        objExcel.Columns("E").ColumnWidth = 6
        objExcel.Columns("F").ColumnWidth = 55
        objExcel.Columns("G").ColumnWidth = 11
        objExcel.Columns("H").ColumnWidth = 10
        objExcel.Columns("I").ColumnWidth = 10
        objExcel.Columns("J").ColumnWidth = 11
        objExcel.Columns("K").ColumnWidth = 11
        objExcel.Columns("L").ColumnWidth = 11
   End With
   
   aa = Leerado3("SELECT * FROM TMP_MASIVO WHERE USU = '" + wcodusu + "' ORDER BY E_SOCIO, NOMBRE ")
   If aa > 0 Then
      wRegAct = 1
      wRegTot = aa
      V = 4: H = 1: zTotApo = 0: zTotCob = 0: wTotApo = 0: wTotCob = 0
      Do While Not ADO3.EOF
         
         wE_S = ADO3!e_socio
         wTotApo = 0: wTotCob = 0
         Do While ADO3!e_socio = wE_S
            DoEvents
            lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wRegAct, "####0") + " / " + Format(wRegTot, "####0")
            lblMensaje.Refresh
            wFecIng = Format(ADO3!fecing, "dd/mm/yyyy")
         
            objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 11)).NumberFormat = "####,##0.00"
         
            objExcel.Cells(V, H + 0) = wRegAct
            objExcel.Cells(V, H + 1) = ADO3!codsocio
            objExcel.Cells(V, H + 2) = ADO3!codigo
            objExcel.Cells(V, H + 3) = ADO3!ins
            objExcel.Cells(V, H + 4) = ADO3!e_socio
            objExcel.Cells(V, H + 5) = ADO3!nombre
            objExcel.Cells(V, H + 6) = wFecIng
            objExcel.Cells(V, H + 7) = ADO3!desde
            objExcel.Cells(V, H + 8) = ADO3!hasta
            objExcel.Cells(V, H + 9) = ADO3!totapo
            objExcel.Cells(V, H + 10) = ADO3!totcob
            objExcel.Cells(V, H + 11) = ADO3!difer
         
            wTotApo = wTotApo + ADO3!totapo
            wTotCob = wTotCob + ADO3!totcob
            
            zTotApo = zTotApo + ADO3!totapo
            zTotCob = zTotCob + ADO3!totcob
            wRegAct = wRegAct + 1
            V = V + 1
            ADO3.MoveNext
            If ADO3.EOF Then
               Exit Do
            End If
         Loop
         
         V = V + 1
         objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 10)).NumberFormat = "####,##0.00"
         objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 10)).Font.Bold = True
         objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 10)).Font.Color = RGB(255, 0, 0)
         objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 10)).Borders.LineStyle = xlContinuous
         objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 10)).Borders.Color = RGB(255, 0, 0)
         
         objExcel.Cells(V, H + 8) = "TOTALES " + wE_S
         objExcel.Cells(V, H + 9) = wTotApo
         objExcel.Cells(V, H + 10) = wTotCob
         V = V + 1
         
      Loop
      
      If wTotApo <> zTotApo Or wTotCob <> zTotCob Then
         V = V + 1
         objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 10)).NumberFormat = "####,##0.00"
         objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 10)).Font.Bold = True
         objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 10)).Font.Color = RGB(255, 0, 0)
         objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 10)).Borders.LineStyle = xlContinuous
         objExcel.Range(objExcel.Cells(V, H + 8), objExcel.Cells(V, H + 10)).Borders.Color = RGB(255, 0, 0)
         
         objExcel.Cells(V, H + 8) = "TOTALES FINALES "
         objExcel.Cells(V, H + 9) = zTotApo
         objExcel.Cells(V, H + 10) = zTotCob
         V = V + 1
      End If
      
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

Private Sub cmdImpDet_Click()
   Dim wFec As String
   wFec = Format(Date, "dd/mm/yyyy")
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\MasivoDetalle.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= 'FECHA " + wFec + "' "
   Crys1.SelectionFormula = " {TMP_MASIVODET.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdImprimir_Click()
   Dim wFec As String
   wFec = Format(Date, "dd/mm/yyyy")
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\MasivoResumen.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= 'FECHA " + wFec + "' "
   Crys1.SelectionFormula = " {TMP_MASIVO.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   Select Case ColIndex
   Case 0
        ADO2.Sort = "CODSOCIO"
   Case 1
        ADO2.Sort = "CODIGO"
   Case 2
        ADO2.Sort = "INS"
   Case 3
        ADO2.Sort = "NOMBRE"
   Case 10
        ADO2.Sort = "DIFER"
   End Select
End Sub

Private Sub Form_Activate()
   frmAporteMasivo.Left = (Screen.Width - Width) \ 2
   frmAporteMasivo.Top = 0
   
   Dim a As Integer, I As Integer, wAno As String, wMes As String
   
'   wAno = Format(Year(Date), "0000")
'   wMes = Format(Month(Date), "00")
'   If wMes > "02" Then
'      wMes = Format(Val(wMes) - 2, "00")
'   Else
'      wMes = Format(Val(wMes) - 2 + 12, "00")
'      wAno = Format(Val(wAno) - 1, "00")
'   End If
'   txtTope.Text = wAno + "/" + wMes
   
   
   txtTope.Text = Left(zMesTope, 4) + "/" + Right(zMesTope, 2)
   txtTope.Enabled = False
   
   
   cmbE_Socio.Clear
   cmbE_Socio.AddItem "000 TODOS LOS SOCIOS"
   cmbE_Socio.AddItem "111 TODOS LOS SOCIOS ACTIVOS"
   cmbE_Socio.AddItem "222 TODOS LOS SOCIOS NO ACTIVOS"
   cmbE_Socio.AddItem "333 SOCIOS CON FRACCIONAMIENTO "
   
   a = Leerado6a("SELECT * FROM MAEE_SOCIO ORDER BY E_SOCIO ")
   If a > 0 Then
      ADO6a.MoveFirst
      Do While Not ADO6a.EOF
         cmbE_Socio.AddItem Trim(ADO6a!nombre)
         
         ADO6a.MoveNext
      Loop
   End If
   Set ADO6a = Nothing
   
   cmbTipCob.Clear
   cmbTipCob.AddItem "00 TODOS LOS TIPOS COBROS"
   a = Leerado8("SELECT * FROM MAETIPCOB ORDER BY TIPCOB ")
   If a > 0 Then
      ADO8.MoveFirst
      I = 1
      Do While Not ADO8.EOF
         cmbTipCob.AddItem ADO8!nombre
         I = I + 1
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   
   cmbE_Socio.ListIndex = 0
   cmbTipCob.ListIndex = 0
   
   cmdAporte.Enabled = False
   cmdImpDet.Enabled = False
   cmdImprimir.Enabled = False
   cmdExportar.Enabled = False
   cmdExpDet.Enabled = False
   
'   txtTope.SetFocus
   cmbE_Socio.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_MASIVO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim aa As Long, wRegAct As Long, wRegTot As Long, w As String, wE_S As String, wTipCob As String, _
       wSoc As Integer, wApo As Currency, wCob As Currency, wDif As Currency, _
       wMesUno As String, wMesDos As String, zSoc As Integer
   
   zSoc = Val(txtCodSocio.Text)
   w = ""
   If Left(cmbE_Socio.Text, 3) = "000" Or Left(cmbE_Socio.Text, 3) = "111" Or Left(cmbE_Socio.Text, 3) = "222" Or Left(cmbE_Socio.Text, 3) = "333" Then
      wE_S = ""
      w = ""
      Select Case Left(cmbE_Socio.Text, 3)
      Case "000"
      Case "111"
           w = " WHERE E_SOCIO = 'ADH' OR E_SOCIO = 'CI1' OR E_SOCIO = 'CIV' OR " _
           & "         E_SOCIO = 'HER' OR E_SOCIO = 'HIJ' OR E_SOCIO = 'NIE' OR " _
           & "         E_SOCIO = 'TIT' OR E_SOCIO = 'TRA' OR E_SOCIO = 'VIU' "
      Case "222"
           w = " WHERE E_SOCIO = 'ESP' OR E_SOCIO = 'EXC' OR E_SOCIO = 'EXP' OR " _
           & "         E_SOCIO = 'FAL' OR E_SOCIO = 'REN' OR E_SOCIO = 'SEP' "
      Case "333"
           w = " WHERE CODSOCIO IN (SELECT DISTINCT C.CODSOCIO FROM FRACDET AS D INNER JOIN FRACCAB AS C ON D.NUMERO = C.NUMERO WHERE D.SDONEW > 0) "
      End Select
   Else
      wE_S = BuscaCodEsocioMasivo(cmbE_Socio.List(cmbE_Socio.ListIndex))
      w = " WHERE E_SOCIO = '" + wE_S + "' "
   End If
   
   If Left(cmbTipCob.Text, 2) = "00" Then
      wTipCob = "00"
   Else
      wTipCob = BuscaCodTipCobMasivo(cmbTipCob.List(cmbTipCob.ListIndex))
      If w = "" Then
         w = " WHERE TIPCOB = '" + wTipCob + "' "
      Else
         w = w + " AND TIPCOB = '" + wTipCob + "' "
      End If
   End If
   
   If zSoc <> 0 Then
      If w = "" Then
         w = " WHERE CODSOCIO = " + Str(zSoc) + " "
      Else
         w = w + " AND CODSOCIO = " + Str(zSoc) + " "
      End If
   End If
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_MASIVO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_MASIVO " _
   & " (CODSOCIO, CODIGO, INS, NUMDOC, NOMBRE, E_SOCIO, DEUDA_PT2, ADELANTO, " _
   & "  FECING, FECREIN, FECBAJ, FECRENO, FECRENU, FECEXCLU, FECEXPUL, " _
   & "  TOTAPO, TOTCOB, DESDE, HASTA, USU) " _
   & " SELECT " _
   & "  CODSOCIO, CODIGO, INS, NUMDOC, NOMBRE, E_SOCIO, DEUDA_PT2, ADELANTO, " _
   & "  FECING, FECREIN, FECBAJ, FECRENO, FECRENU, FECEXCLU, FECEXPUL, " _
   & "  0, 0, '', '', '" + wcodusu + "' " _
   & " FROM MAESOCIO " _
   & " " + w + " ")
   Db.CommitTrans

   aa = Leerado2("SELECT * " _
                & " FROM TMP_MASIVO " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY E_SOCIO, NOMBRE ")
   If aa > 0 Then
      ADO2.MoveFirst
      wRegAct = 1
      wRegTot = aa
      Do While Not ADO2.EOF
         DoEvents
         lblMensaje.Caption = Trim(Format(wRegAct, "####0")) + " / " + _
                              Trim(Format(wRegTot, "####0")) + _
                              " Socio " + Trim(ADO2!nombre)
         lblMensaje.Refresh
         
         wSoc = ADO2!codsocio
         wApo = CalTotAportes(wSoc, txtTope.Text, 1)
         wMesUno = CalTotAportes(wSoc, txtTope.Text, 3)
         wMesDos = CalTotAportes(wSoc, txtTope.Text, 4)
         wCob = CalTotCobros(wSoc, txtTope.Text)
         wDif = wApo - wCob
   
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_MASIVO " _
         & " SET TOTAPO = " + Str(wApo) + ", " _
         & "     TOTCOB = " + Str(wCob) + ", " _
         & "      DIFER = " + Str(wDif) + ", " _
         & "      DESDE = '" + wMesUno + "', " _
         & "      HASTA = '" + wMesDos + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "            USU = '" + wcodusu + "' ")
         Db.CommitTrans
   
         wRegAct = wRegAct + 1
         ADO2.MoveNext
      Loop
   End If

   DoEvents
   lblMensaje.Caption = ""
   lblMensaje.Refresh

   cmdAporte.Enabled = True
   cmdImpDet.Enabled = True
   cmdImprimir.Enabled = True
   cmdExportar.Enabled = True
   cmdExpDet.Enabled = True
   
   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, FECING, FECRENU, " _
                & "      FECEXCLU, FECEXPUL, FECRENO, FECBAJ, FECREIN, USU " _
                & " FROM TMP_MASIVO " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY E_SOCIO, NOMBRE ")
   Set DataGrid1.DataSource = ADO2

   DataGrid1.Columns(0).Width = 750   ' CODSOCIO
   DataGrid1.Columns(0).Alignment = dbgRight
   DataGrid1.Columns(0).Caption = "SOCIO"
       
   DataGrid1.Columns(1).Width = 1000   ' CODIGO
   DataGrid1.Columns(1).Alignment = dbgRight
   DataGrid1.Columns(1).Caption = "CODIGO"
       
   DataGrid1.Columns(2).Width = 350   ' INS
   DataGrid1.Columns(2).Alignment = dbgCenter
   DataGrid1.Columns(2).Caption = "INS"
       
   DataGrid1.Columns(3).Width = 4800  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE"
    
   DataGrid1.Columns(4).Width = 450   ' E_SOCIO
   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Caption = "E_S"
    
   DataGrid1.Columns(5).Width = 1040   ' FECING
   DataGrid1.Columns(5).Alignment = dbgCenter
   DataGrid1.Columns(5).Caption = "FEC.ING"
   DataGrid1.Columns(5).NumberFormat = "dd/mm/yyyy"

   DataGrid1.Columns(6).Width = 1040   ' FECRENU
   DataGrid1.Columns(6).Alignment = dbgCenter
   DataGrid1.Columns(6).Caption = "FEC.RENUNC"
   DataGrid1.Columns(6).NumberFormat = "dd/mm/yyyy"

   DataGrid1.Columns(7).Width = 1040   ' FECEXCLU
   DataGrid1.Columns(7).Alignment = dbgCenter
   DataGrid1.Columns(7).Caption = "FEC.EXCLUS"
   DataGrid1.Columns(7).NumberFormat = "dd/mm/yyyy"

   DataGrid1.Columns(8).Width = 1040   ' FECEXPUL
   DataGrid1.Columns(8).Alignment = dbgCenter
   DataGrid1.Columns(8).Caption = "FEC.EXPULS"
   DataGrid1.Columns(8).NumberFormat = "dd/mm/yyyy"

   DataGrid1.Columns(9).Width = 1040   ' FECRENO
   DataGrid1.Columns(9).Alignment = dbgCenter
   DataGrid1.Columns(9).Caption = "FEC.RENOVA"
   DataGrid1.Columns(9).NumberFormat = "dd/mm/yyyy"

   DataGrid1.Columns(10).Width = 1040   ' FECBAJ
   DataGrid1.Columns(10).Alignment = dbgCenter
   DataGrid1.Columns(10).Caption = "FEC.BAJA"
   DataGrid1.Columns(10).NumberFormat = "dd/mm/yyyy"

   DataGrid1.Columns(11).Width = 1040   ' FECREIN
   DataGrid1.Columns(11).Alignment = dbgCenter
   DataGrid1.Columns(11).Caption = "FEC.REING"
   DataGrid1.Columns(11).NumberFormat = "dd/mm/yyyy"

   DataGrid1.Columns(12).Visible = False
End Sub

Private Function CalTotAportes(zSoc As Integer, zMesCorte As String, sw As Integer) As Variant
   On Error GoTo err

   Dim zFecIng As Date, zAnoIng As String, zMesIng As String, zDiaIng As Integer, zE_s, _
       zAnoAct As String, zMesAct As String, aa As Integer, II As Long, OO As Integer, _
       wmmm As String, waaa As String, wMax As Integer, wMin As Integer, _
       zTotApo As Currency, zApo As Currency, zCanApo As Integer, _
       zMesUno As String, zMesDos As String, _
       zFecRen As Date, zAnoRen, zMesRen As String, _
       zFecRei As Date, zAnoRei, zMesRei As String
   zAnoAct = Left(zMesCorte, 4)
   zMesAct = Right(zMesCorte, 2)
   zTotApo = 0: zCanApo = 0

   If sw = 1 Then
      Db.BeginTrans
      Db.Execute ("DELETE FROM TMP_MASIVODET WHERE CODSOCIO = " + Str(zSoc) + " AND USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If

   aa = Leerado7a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(zSoc) + " ")
   If aa > 0 Then
      zE_s = ADO7a!e_socio
      zFecIng = Format(ADO7a!fecing, "dd/mm/yyyy")
      If IsDate(ADO7a!fecrenu) Then
         zFecRen = Format(ADO7a!fecrenu, "dd/mm/yyyy")
      End If
      If IsDate(ADO7a!fecrein) Then
         zFecRei = Format(ADO7a!fecrein, "dd/mm/yyyy")
      End If
   End If
   Set ADO7a = Nothing

   zAnoIng = Format(Year(zFecIng), "0000")
   zMesIng = Format(Month(zFecIng), "00")
   zDiaIng = Day(zFecIng)

   If IsDate(zFecRei) Then
      zMesRei = Format(Month(zFecRei), "00")
      zAnoRei = Format(Year(zFecRei), "0000")
   End If

   If IsDate(zFecRen) Then
      zMesRen = Format(Month(zFecRen), "00")
      zAnoRen = Format(Year(zFecRen), "0000")
   
      zAnoAct = zAnoRen
      zMesAct = zMesRen
   End If

   If zAnoIng < "1997" Then
      zDiaIng = 1
      zMesIng = "01"
      zAnoIng = "1997"
   End If
   
   zMesUno = "": zMesDos = ""
   For II = Val(zAnoIng) To Val(zAnoAct)
       waaa = Format(II, "0000")
   
       If zAnoIng <> zAnoAct Then
          Select Case waaa
          Case zAnoIng
               If zDiaIng <= 16 Then
                  wMin = Val(zMesIng)
               Else
                  wMin = Val(zMesIng) + 1
               End If
               wMax = 12
          Case zAnoAct
               wMin = 1
               wMax = Val(zMesAct)
          Case Else
               wMin = 1
               wMax = 12
          End Select
       Else
          wMin = Val(zMesIng)
          wMax = Val(zMesAct)
       End If
     
       For OO = wMin To wMax
           wmmm = Format(OO, "00")
       
           If zCanApo = 0 Then
              zMesUno = waaa + "/" + wmmm
           End If
           zMesDos = waaa + "/" + wmmm
              
           zCanApo = zCanApo + 1
           
           If sw = 1 Then
              zApo = 0
              aa = Leerado6a("SELECT * FROM MAEAPORTE " _
                           & " WHERE ANO = '" + waaa + "' AND MES = '" + wmmm + "' ")
              If aa > 0 Then
                 Select Case zE_s
                 Case "CIV"
                      zApo = ADO6a!aporte_tit
                 Case "CI1"
                      zApo = ADO6a!aporte_tit
                 Case "TRA"
                      zApo = ADO6a!aporte_tit
                 Case Else
                      zApo = ADO6a!aporte_tit
                 End Select
                 zTotApo = zTotApo + zApo
           
                 aa = Leerado5a("SELECT * FROM TMP_MASIVODET " _
                            & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
                            & "            ANO = '" + waaa + "' AND " _
                            & "            USU = '" + wcodusu + "' ")
                 If aa = 0 Then
                    Db.BeginTrans
                    Db.Execute ("INSERT INTO TMP_MASIVODET " _
                    & " (CODSOCIO, ANO, USU) " _
                    & " VALUES " _
                    & " (" + Str(zSoc) + ", '" + waaa + "', " _
                    & "  '" + wcodusu + "' ) ")
                    Db.CommitTrans
                 End If
              
                 Db.BeginTrans
                 Db.Execute ("UPDATE TMP_MASIVODET " _
                 & " SET IMP" + wmmm + " = " + Str(zApo) + ", " _
                 & "     TOTAL = TOTAL + " + Str(zApo) + " " _
                 & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
                 & "            ANO = '" + waaa + "' AND " _
                 & "            USU = '" + wcodusu + "' ")
                 Db.CommitTrans
              
              End If
           End If
       Next
   Next
   Select Case sw
   Case 1 ' Total de Aportes
        CalTotAportes = zTotApo
   Case 2 ' Cantidad de Aportes
        CalTotAportes = zCanApo
   Case 3 ' Mes Inicial
        CalTotAportes = zMesUno
   Case 4 ' Mes Final
        CalTotAportes = zMesDos
   End Select
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Private Function CalTotCobros(zSoc As Integer, zMesCorte As String) As Currency
   On Error GoTo err

   Dim aa As Integer, II As Integer, waaa As String, zCob As Currency, _
       zAnoAct As String, zMesAct As String, zFecAct As Date, _
       zAnoIng As String, zMesIng As String, zFecIng As Date, zCod As Long, zIns As Integer

   zAnoAct = Left(zMesCorte, 4)
   zMesAct = Right(zMesCorte, 2)
   zFecAct = Format(fundiames(zMesAct) + "/" + zMesAct + "/" + zAnoAct, "dd/mm/yyyy")

   aa = Leerado7a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(zSoc) + " ")
   If aa > 0 Then
      zCod = ADO7a!codigo
      zIns = ADO7a!ins
      zFecIng = Format(ADO7a!fecing, "dd/mm/yyyy")
      zAnoIng = Format(Year(zFecIng), "0000")
      zMesIng = Format(Month(zMesCorte), "00")
   End If
   Set ADO7a = Nothing

   zCob = 0
   For II = Val(zAnoIng) To Val(zAnoAct)
       waaa = Format(II, "0000")
   
       If II = zAnoAct Then
          aa = Leerado7a("SELECT * " _
                    & " FROM ZZZ_APOR_PLA " _
                    & " WHERE CODIGO = " + Str(zCod) + " AND " _
                    & "          INS = " + Str(zIns) + " AND " _
                    & "       CUOANO = '" + waaa + "'  ")
          If aa > 0 Then
             Select Case zMesAct
             Case "01"
                  zCob = zCob + ADO7a!impo01
             Case "02"
                  zCob = zCob + ADO7a!impo01 + ADO7a!impo02
             Case "03"
                  zCob = zCob + ADO7a!impo01 + ADO7a!impo02 + ADO7a!impo03
             Case "04"
                  zCob = zCob + ADO7a!impo01 + ADO7a!impo02 + ADO7a!impo03 + ADO7a!impo04
             Case "05"
                  zCob = zCob + ADO7a!impo01 + ADO7a!impo02 + ADO7a!impo03 + ADO7a!impo04 + ADO7a!impo05
             Case "06"
                  zCob = zCob + ADO7a!impo01 + ADO7a!impo02 + ADO7a!impo03 + ADO7a!impo04 + ADO7a!impo05 + ADO7a!impo06
             Case "07"
                  zCob = zCob + ADO7a!impo01 + ADO7a!impo02 + ADO7a!impo03 + ADO7a!impo04 + ADO7a!impo05 + ADO7a!impo06 + _
                                ADO7a!impo07
             Case "08"
                  zCob = zCob + ADO7a!impo01 + ADO7a!impo02 + ADO7a!impo03 + ADO7a!impo04 + ADO7a!impo05 + ADO7a!impo06 + _
                                ADO7a!impo07 + ADO7a!impo08
             Case "09"
                  zCob = zCob + ADO7a!impo01 + ADO7a!impo02 + ADO7a!impo03 + ADO7a!impo04 + ADO7a!impo05 + ADO7a!impo06 + _
                                ADO7a!impo07 + ADO7a!impo08 + ADO7a!impo09
             Case "10"
                  zCob = zCob + ADO7a!impo01 + ADO7a!impo02 + ADO7a!impo03 + ADO7a!impo04 + ADO7a!impo05 + ADO7a!impo06 + _
                                ADO7a!impo07 + ADO7a!impo08 + ADO7a!impo09 + ADO7a!impo10
             Case "11"
                  zCob = zCob + ADO7a!impo01 + ADO7a!impo02 + ADO7a!impo03 + ADO7a!impo04 + ADO7a!impo05 + ADO7a!impo06 + _
                                ADO7a!impo07 + ADO7a!impo08 + ADO7a!impo09 + ADO7a!impo10 + ADO7a!impo11
             Case "12"
                  zCob = zCob + ADO7a!impo01 + ADO7a!impo02 + ADO7a!impo03 + ADO7a!impo04 + ADO7a!impo05 + ADO7a!impo06 + _
                                ADO7a!impo07 + ADO7a!impo08 + ADO7a!impo09 + ADO7a!impo10 + ADO7a!impo11 + ADO7a!impo12
             End Select
          End If
          Set ADO7a = Nothing
       Else
          aa = Leerado7a("SELECT SUM(TOTIMPO) AS TOT " _
                    & " FROM ZZZ_APOR_PLA " _
                    & " WHERE CODIGO = " + Str(zCod) + " AND " _
                    & "          INS = " + Str(zIns) + "AND " _
                    & "       CUOANO = '" + waaa + "'  ")
          If aa > 0 Then
             zCob = zCob + IIf(IsNull(ADO7a!tot), 0, ADO7a!tot)
          End If
          Set ADO7a = Nothing
       End If
   Next II

   aa = Leerado7a("SELECT SUM(Z.MONTO) AS MONTO " _
                & " FROM ZZZ_MRECIBOS AS Z INNER JOIN ZZZ_CONCEPTO AS M " _
                & "   ON Z.CONCEPTO = M.CCONCE " _
                & " WHERE Z.CODIGO = " + Str(zCod) + " AND " _
                & "          Z.INS = " + Str(zIns) + " AND " _
                & "        Z.FECHA_PAGO <= '" + Format(zFecAct, "dd/mm/yyyy") + "' AND " _
                & "       (Z.MARCA2 <> 'A' OR Z.MARCA2 IS NULL) AND " _
                & "       (M.MARCA = 'S') ")
   If aa > 0 Then
      zCob = zCob + IIf(IsNull(ADO7a!monto), 0, ADO7a!monto)
   End If
   Set ADO7a = Nothing

   aa = Leerado7a("SELECT SUM(APORTE) AS MONTO " _
                & " FROM ZZZ_BCORECAU " _
                & " WHERE CODIGO = " + Str(zCod) + " AND " _
                & "          INS = " + Str(zIns) + " AND " _
                & "        FECHA <= '" + Format(zFecAct, "dd/mm/yyyy") + "' ")
   If aa > 0 Then
      zCob = zCob + IIf(IsNull(ADO7a!monto), 0, ADO7a!monto)
   End If
   Set ADO7a = Nothing

   aa = Leerado7a("SELECT SUM(IMPORTE) AS IMPORTE " _
                & " FROM ZZZ_DEVOL " _
                & " WHERE CODIGO = " + Str(zCod) + " AND " _
                & "          INS = " + Str(zIns) + " AND " _
                & "      FECHA <= '" + Format(zFecAct, "dd/mm/yyyy") + "' ")
   If aa > 0 Then
      zCob = zCob - IIf(IsNull(ADO7a!importe), 0, ADO7a!importe)
   End If
   Set ADO7a = Nothing

   CalTotCobros = zCob
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Private Sub TotalCab()
   Dim zz As Integer

   zz = Leerado7a("SELECT COUNT(*) AS NUM FROM TMP_MASIVO WHERE USU = '" + wcodusu + "' ")
   If zz > 0 Then
      zz = ADO7a!num
   End If

   lblTotal.Caption = Format(zz, "###,##0;;\ ")
End Sub

Private Function BuscaEsocioMasivo(zE_Socio As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer, zNum As Integer
   zRes = 0
   zNum = 2
   zz = Leerado5a("SELECT * FROM MAEE_SOCIO ORDER BY E_SOCIO ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If ADO5a!e_socio = zE_Socio Then
            zRes = zNum
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   BuscaEsocioMasivo = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Private Function BuscaCodEsocioMasivo(zNom As String) As String
   On Error GoTo err
   
   Dim zz As Integer, zRes As String, zLen As Integer, zNum As Integer
   zRes = 2: zLen = Len(Trim(zNom))
   zz = Leerado5a("SELECT * FROM MAEE_SOCIO ORDER BY E_SOCIO ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If Mid(ADO5a!nombre, 1, zLen) = Trim(zNom) Then
            zRes = ADO5a!e_socio
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   
   BuscaCodEsocioMasivo = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaTipCobMasivo(zTipCob As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer, zNum As Integer
   zRes = 1
   zNum = 0
   zz = Leerado5a("SELECT * FROM MAETIPCOB ORDER BY TIPCOB ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If ADO5a!tipcob = zTipCob Then
            zRes = zNum
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   BuscaTipCobMasivo = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaCodTipCobMasivo(zNom As String) As String
   On Error GoTo err
   
   Dim zz As Integer, zRes As String, zLen As Integer, zNum As Integer
   zRes = 1: zLen = Len(Trim(zNom))
   zz = Leerado5a("SELECT * FROM MAETIPCOB ORDER BY TIPCOB ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If Mid(ADO5a!nombre, 1, zLen) = Trim(zNom) Then
            zRes = ADO5a!tipcob
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   
   BuscaCodTipCobMasivo = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Private Sub txtCodSocio_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
   If aa > 0 Then
      lblCodSocio.Caption = ADO6a!nombre
   Else
      lblCodSocio.Caption = ""
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
   Dim KeyNew As Integer
   KeyNew = Asc(UCase(Chr(KeyCode)))
   
   Select Case KeyCode
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtCodSocio.Text = xseleccion
        End If
   Case 65 To 90
        xlista = "1"
        xseleccion = Chr(KeyNew)
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtCodSocio.Text = xseleccion
        End If
          
   End Select
End Sub

Private Sub txtCodSocio_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If Val(txtCodSocio.Text) > 0 Then
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
         If aa = 0 Then
            MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
            txtCodSocio.Text = ""
            Exit Sub
         End If
      End If
      cmdBuscar.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
'         KeyAscii = 0
'      End If
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

Private Sub txtTope_GotFocus()
   txtTope.SelStart = 0
   txtTope.SelLength = 10
End Sub

Private Sub txtTope_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 40
        cmbE_Socio.SetFocus
   End Select
End Sub

Private Sub txtTope_KeyPress(KeyAscii As Integer)
   Dim wMes As String, wAno As String
   If KeyAscii = 13 Then
      If txtTope.Text = "____/__" Then
         MsgBox "Mes Tope En Blanco", vbExclamation
         txtTope.Text = "____/__"
         Exit Sub
      End If
      wAno = Left(txtTope.Text, 4)
      wMes = Right(txtTope.Text, 2)
      If wMes <> "01" And wMes <> "02" And wMes <> "03" And wMes <> "04" And _
         wMes <> "05" And wMes <> "06" And wMes <> "07" And wMes <> "08" And _
         wMes <> "09" And wMes <> "10" And wMes <> "11" And wMes <> "12" Then
         MsgBox "Mes Digitado Es Errado", vbExclamation
         txtTope.Text = "____/__"
         Exit Sub
      End If
      If wAno < "2017" And wAno > "2030" Then
         MsgBox "Año Digitado Es Errado", vbExclamation
         txtTope.Text = "____/__"
         Exit Sub
      End If
      cmbE_Socio.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii), 1) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub
