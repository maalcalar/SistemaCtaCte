VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBCPRetorno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retorno Diario BCP"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11205
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5655
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   9975
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
   Begin VB.CommandButton cmdLeer 
      Caption         =   "Leer EXCEL"
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
      Left            =   4680
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
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
      Left            =   9240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1095
   End
   Begin MSMask.MaskEdBox txtDesde 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtHasta 
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
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
      Left            =   600
      TabIndex        =   8
      Top             =   7320
      Width           =   6135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Hasta"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Desde"
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
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RETORNO COBROS POR BCP"
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
      Left            =   2400
      TabIndex        =   0
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmBCPRetorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdLeer_Click()
   On Error GoTo err

   Dim wFec As Date, wCod As Long, wIns As Integer, wSoc As Integer, wImp As Currency, _
       wAno As String, wMes As String, wDoc As String, wGlo As String, wE_S As String, _
       wApo As Currency, wSdo As Currency, wMon As String, wMesCob As String, wLin As Integer, _
       wDeu As Currency, wAde As Currency, wMin As Long, wMax As Long, wNumOpe As String, _
       wAnoOld As String, wMesOld As String, wTipOld As String, wSerOld As String, wNumOld As String, _
       wTotFrac As Currency, wCuoFrac As Currency, wCanFrac As Integer, _
       wNumFrac As String, wLinFrac As String, wMaxFrac As String, wMinFrac As String, _
       II As Integer

   wNom = InputBox("El Archivo Deberá estar Ubicado En La Carpeta EXCEL" + _
                    vbNewLine + vbNewLine + vbNewLine + _
                   "Nombre del Archivo Excel" + vbNewLine, _
                   "Importar Archivo BCP", "BCP_dd_mm")
   If Len(Trim(wNom)) = 0 Then
      MsgBox "Nombre En Blanco", vbExclamation
      Exit Sub
   End If
   
   If Len(Dir(xraiz & "EXCEL\" & wNom & ".XLS")) = 0 Then
      MsgBox "Archivo de Importar BCP No Existe", vbExclamation
      Exit Sub
   End If
   
   Set objExcel = New Excel.Application
   objExcel.Visible = False
   objExcel.Workbooks.Open xraiz & "EXCEL\" & wNom & ".XLS"
   objExcel.Worksheets(1).Activate 'coloca la primera hoja del libro
   
   wMax = 0: wMin = 0
   For aa = 6 To 10000
       If IsDate(objExcel.Cells(aa, 1)) Then
          If wMax = 0 Then
             wMax = aa
          Else
             wMin = aa
          End If
       Else
          Exit For
       End If
   Next

   For aa = wMin To wMax Step -1
       If objExcel.Cells(aa, 1) = "" Then
          Exit For
       End If
       wRegAct = aa - 5
       
       DoEvents
       lblMensaje.Caption = "Importando " + wNom + " - Reg " + Format(wRegAct, "####0")
       lblMensaje.Refresh
       
       If IsDate(objExcel.Cells(aa, 1)) And Mid(objExcel.Cells(aa, 3), 1, 8) = "EFECTIVO" Then
          wFec = Format(objExcel.Cells(aa, 1), "dd/mm/yyyy")
          wAno = Format(Year(wFec), "0000")
          wMes = Format(Month(wFec), "00")
          wCod = Mid(objExcel.Cells(aa, 3), 9, 14)
          wImp = Val(objExcel.Cells(aa, 4))
          wNumOpe = Trim(objExcel.Cells(aa, 7))
          
          wSdo = wImp
'          wDoc = Format(Val(wDoc) + 1, "0000000000")
          wGlo = "APORTE POR BCP - FECHA " + Format(wFec, "dd/mm/yyyy")
          wFec = Format(txtHasta.Text, "dd/mm/yyyy")
   
          zz = Leerado6a("SELECT * FROM COBRODET " _
                        & " WHERE (CONPAGO = '155' OR CONPAGO = '128') AND " _
                        & "       (NUMOPE = '" + wNumOpe + "') ")
          If zz > 0 Then
             ADO6a.MoveFirst
             Do While Not ADO6a.EOF
             
                wAnoOld = ADO6a!ano
                wMesOld = ADO6a!mes
                wTipOld = ADO6a!tipcob
                wSerOld = ADO6a!sercob
                wNumOld = ADO6a!numcob
             
                Db.BeginTrans
                Db.Execute ("DELETE FROM COBROCAB " _
                & " WHERE    ANO = '" + wAnoOld + "' AND " _
                & "          MES = '" + wMesOld + "' AND " _
                & "       TIPCOB = '" + wTipOld + "' AND " _
                & "       SERCOB = '" + wSerOld + "' AND " _
                & "       NUMCOB = '" + wNumOld + "' ")
                Db.CommitTrans
          
                Db.BeginTrans
                Db.Execute ("DELETE FROM COBRODET " _
                & " WHERE    ANO = '" + wAnoOld + "' AND " _
                & "          MES = '" + wMesOld + "' AND " _
                & "       TIPCOB = '" + wTipOld + "' AND " _
                & "       SERCOB = '" + wSerOld + "' AND " _
                & "       NUMCOB = '" + wNumOld + "' ")
                Db.CommitTrans
          
                Db.BeginTrans
                Db.Execute ("DELETE FROM ZZZ_MRECIBOS " _
                & " WHERE    SERIE = '" + wSerOld + "' AND " _
                & "       NRO_COMP = " + Str(Val(wNumOld)) + " AND " _
                & "       YEAR(FECHA_PAGO) = " + Str(Val(wAnoOld)) + " ")
                Db.CommitTrans
          
                wDoc = wNumOld
             
                ADO6a.MoveNext
             Loop
          End If
          wDoc = BusUltRecibo()
   
          zz = Leerado6a("SELECT * FROM MAESOCIO WHERE CODIGO = " + Str(wCod) + " ")
          If zz > 0 Then
             wIns = ADO6a!ins
             wSoc = ADO6a!codsocio
             wE_S = ADO6a!e_socio
          End If
          Set ADO6a = Nothing
      
          zz = Leerado6a("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + wE_S + "' ")
          If zz > 0 Then
             wMon = ADO6a!moneda
             wApo = ADO6a!aporte
          End If
          Set ADO6a = Nothing
          
          wLin = 1
          wMesCob = wAno + "/" + wMes
          wGlo = wGlo + " " + wMesCob
             
'---------------------------------------
' Verifica Fraccionamientos Pendientes

          wTotFrac = 0: wCuoFrac = 0
          zz = Leerado6a("SELECT COUNT(*) AS CAN, MIN(D.LINEA) AS MINIMO, MAX(D.LINEA) AS MAXIMO, SUM(SDONEW) AS SDONEW " _
                     & " FROM FRACDET AS D INNER JOIN FRACCAB AS C " _
                     & "   ON D.NUMERO = C.NUMERO " _
                     & " WHERE C.CODSOCIO = " + Str(wSoc) + " AND " _
                     & "       D.SDONEW > 0 AND " _
                     & "       D.VCMTO <= '" + Format(wFec, "dd/mm/yyyy") + "' ")
          If zz > 0 Then
             wMinFrac = IIf(IsNull(ADO6a!minimo), "", ADO6a!minimo)
             wMaxFrac = IIf(IsNull(ADO6a!maximo), "", ADO6a!maximo)
             wCanFrac = IIf(IsNull(ADO6a!can), 0, ADO6a!can)
             wTotFrac = IIf(IsNull(ADO6a!sdonew), 0, ADO6a!sdonew)
          End If

          wSdo = wImp
          
          If wCanFrac > 0 Then
          zz = Leerado6a("SELECT D.NUMERO, D.LINEA, D.VCMTO, D.SDONEW " _
                     & " FROM FRACDET AS D INNER JOIN FRACCAB AS C " _
                     & "   ON D.NUMERO = C.NUMERO " _
                     & " WHERE C.CODSOCIO = " + Str(wSoc) + " AND " _
                     & "       D.SDONEW > 0 AND " _
                     & "       D.VCMTO <= '" + Format(wFec, "dd/mm/yyyy") + "' " _
                     & " ORDER BY D.NUMERO, D.LINEA ")
          If zz > 0 Then
             Do While Not ADO6a.EOF
                wNumFrac = ADO6a!numero
                wLinFrac = ADO6a!linea
                wCuoFrac = IIf(IsNull(ADO6a!sdonew), 0, ADO6a!sdonew)
             
                If wSdo >= ADO6a!sdonew Then
                   Db.BeginTrans
                   Db.Execute ("INSERT INTO COBRODET " _
                   & " (ANO, MES, TIPCOB, SERCOB, NUMCOB, LINCOB, MESCOB, CONPAGO, DOLARE, SOLESS, MONDOC, " _
                   & "  SDOOLD, CARGOS, ABONOS, SDONEW, IMPORTE, CONCEPTO, PARIENTE, LINPARIE, NOMBRE, NUMOPE, NUMFRA, LINFRA ) " _
                   & " VALUES " _
                   & " ('" + wAno + "', '" + wMes + "', '2', '004', '" + wDoc + "', " _
                   & "  '" + Format(wLin, "00") + "', '" + Format(ADO6a!vcmto, "yyyy/mm") + "', '128', 0, " + Str(wCuoFrac) + ", " _
                   & "  '" + wMon + "', " + Str(wCuoFrac) + ", 0, " + Str(wCuoFrac) + ", 0, " + Str(wCuoFrac) + ", " _
                   & "  '03', '', '', '', '" + wNumOpe + "', '" + wNumFrac + "', '" + wLinFrac + "' ) ")
                   Db.CommitTrans
                
                   Db.BeginTrans
                   Db.Execute ("INSERT INTO TMP_COBRODET " _
                   & " (ANO, MES, TIPCOB, SERCOB, NUMCOB, LINCOB, MESCOB, CONPAGO, DOLARE, SOLESS, MONDOC, " _
                   & "  SDOOLD, CARGOS, ABONOS, SDONEW, IMPORTE, CONCEPTO, PARIENTE, LINPARIE, NOMBRE, " _
                   & "  NUMOPE, NUMFRA, LINFRA, USU ) " _
                   & " VALUES " _
                   & " ('" + wAno + "', '" + wMes + "', '2', '004', '" + wDoc + "', " _
                   & "  '" + Format(wLin, "00") + "', '" + Format(ADO6a!vcmto, "yyyy/mm") + "', '128', 0, " + Str(wCuoFrac) + ", " _
                   & "  '" + wMon + "', " + Str(wCuoFrac) + ", 0, " + Str(wCuoFrac) + ", 0, " + Str(wCuoFrac) + ", " _
                   & "  '03', '', '', '', '" + wNumOpe + "', '" + wNumFrac + "', '" + wLinFrac + "', '" + wcodusu + "' ) ")
                   Db.CommitTrans

                   Db.BeginTrans
                   Db.Execute ("UPDATE FRACDET " _
                   & " SET ABONOS = " + Str(wCuoFrac) + ", " _
                   & "     SDONEW = CARGOS - " + Str(wCuoFrac) + ", " _
                   & "     SERCOB = '004', NUMCOB = '" + wDoc + "', " _
                   & "     FECCOB = '" + Format(wFec, "dd/mm/yyyy") + "' " _
                   & " WHERE NUMERO = '" + wNumFrac + "' AND " _
                   & "        LINEA = '" + Format(Val(wLinFrac), "@@") + "' ")
                   Db.CommitTrans
                   
                   aa = Leerado7a("SELECT * FROM CTASXCAB " _
                                & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                                & "            MES = '" + Format(ADO6a!vcmto, "yyyy/mm") + "' AND " _
                                & "       CONCEPTO = '03' ")
                   If aa = 0 Then
                      wqqq = CreaAporteMes(wSoc, wMes, "03", 1)
                   End If
               
                   Db.BeginTrans
                   Db.Execute ("INSERT INTO CTASXDET " _
                   & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
                   & "  TIPCAM, DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW ) " _
                   & " VALUES " _
                   & " (" + Str(wSoc) + ", '" + Format(ADO6a!vcmto, "yyyy/mm") + "', '03', " _
                   & "  '03', '004', '" + wDoc + "', '" + Format(wLin, "") + "', " _
                   & "  '2', '" + Format(wFec, "dd/mm/yyyy") + "', 0, " _
                   & "  0, " + Str(wCuoFrac) + ", " _
                   & "  0, 0, " + Str(wCuoFrac) + ", 0 ) ")
                   Db.CommitTrans
         
                   Call ActualizaSaldos(wSoc, Format(ADO6a!vcmto, "yyyy/mm"), "03")
            
            
                   zz = Leerado8("SELECT * FROM ZZZ_MRECIBOS " _
                                & " WHERE    SERIE = '004' AND " _
                                & "       CONCEPTO = 128 AND " _
                                & "       NRO_COMP = " + Str(Val(wDoc)) + " AND " _
                                & "       YEAR(FECHA_PAGO) = " + Str(Val(wanocia)) + " AND " _
                                & "       OBS = 'BCP - PAGO FRACC." + wNumFra + " MES " + Format(ADO6a!vcmto, "yyyy/mm") + "' ")
                   If zz = 0 Then
                      Db.BeginTrans
                      Db.Execute ("INSERT INTO ZZZ_MRECIBOS " _
                      & " (CODIGO, INS, CONCEPTO, SERIE, AUXILIAR, NRO_COMP, MONTO, MONEDA, T_CAMBIO, " _
                      & "  FECHA_PAGO, FECHA_CADU, OBS, D_IMPOR, DEUDA_PT2, DINS_CER, ADELANTO, " _
                      & "  MARCA1, MARCA2, MARCA3, MARCA4, OBS1 ) " _
                      & " VALUES " _
                      & " (" + Str(wCod) + ", " + Str(wIns) + ", 128, '004', 0, " _
                      & "  " + Str(Val(wDoc)) + ", " + Str(wCuoFrac) + ", " _
                      & "  'S/.', 0, " _
                      & "  '" + Format(wFec, "dd/mm/yyyy") + "', null, 'BCP OP " + wNumOpe + " - PAGO FRACC." + wNumFra + " MES " + Format(ADO6a!vcmto, "yyyy/mm") + "', " _
                      & "  '', 0, 0, 0, " _
                      & "  '" + Format(Date, "dd/mm/yyyy") + "', 'N', '" + wcodusu + "', " _
                      & "  '" + Format(Time, "hh:mm:ss") + "', '' ) ")
                      Db.CommitTrans
                   Else
                      Db.BeginTrans
                      Db.Execute ("UPDATE ZZZ_MRECIBOS " _
                      & " SET CODIGO = " + Str(wCod) + ", INS = " + Str(wIns) + ", CONCEPTO = 128, " _
                      & "     AUXILIAR = 0, MONTO = " + Str(wCuoFrac) + ", " _
                      & "     MONEDA = 'S/.', " _
                      & "     T_CAMBIO = 0, FECHA_PAGO = '" + Format(wFec, "dd/mm/yyyy") + "', " _
                      & "     FECHA_CADU = null, OBS = 'BCP OP " + wNumOpe + " - PAGO FRACC." + wNumFra + " MES " + Format(ADO6a!vcmto, "yyyy/mm") + "', D_IMPOR = '', " _
                      & "     DEUDA_PT2 = 0, DINS_CER = 0, ADELANTO = 0, " _
                      & "     MARCA1 = '" + Format(Date, "dd/mm/yyyy") + "', MARCA2 = 'N', " _
                      & "     MARCA3 = '" + wcodusu + "', MARCA4 = '" + Format(Time, "hh:mm:ss") + "', " _
                      & "     OBS1 = '' " _
                      & " WHERE    SERIE = '004' AND " _
                      & "       CONCEPTO = 128 AND " _
                      & "       NRO_COMP = " + Str(Val(wDoc)) + " AND " _
                      & "       YEAR(FECHA_PAGO) = " + Str(Val(wanocia)) + " AND " _
                      & "       OBS = 'BCP - PAGO FRACC." + wNumFra + " MES " + Format(ADO6a!vcmto, "yyyy/mm") + "' ")
                      Db.CommitTrans
                   End If
                   
                   wSdo = wSdo - wCuoFrac
                   wLin = wLin + 1
                End If
                
                ADO6a.MoveNext
             Loop
          End If

          End If
'---------------------------------------
             
          Db.BeginTrans
          Db.Execute ("INSERT INTO COBRODET " _
          & " (ANO, MES, TIPCOB, SERCOB, NUMCOB, LINCOB, MESCOB, CONPAGO, DOLARE, SOLESS, MONDOC, " _
          & "  SDOOLD, CARGOS, ABONOS, SDONEW, IMPORTE, CONCEPTO, PARIENTE, LINPARIE, NOMBRE, " _
          & "  NUMOPE, NUMFRA, LINFRA ) " _
          & " VALUES " _
          & " ('" + wAno + "', '" + wMes + "', '2', '004', '" + wDoc + "', " _
          & "  '" + Format(wLin, "00") + "', '" + wMesCob + "', '155', 0, " + Str(wSdo) + ", " _
          & "  '" + wMon + "', " + Str(wSdo) + ", 0, " + Str(wSdo) + ", 0, " + Str(wSdo) + ", " _
          & "  '01', '', '', '', '" + wNumOpe + "', '', '' ) ")
          Db.CommitTrans

          Db.BeginTrans
          Db.Execute ("INSERT INTO TMP_COBRODET " _
          & " (ANO, MES, TIPCOB, SERCOB, NUMCOB, LINCOB, MESCOB, CONPAGO, DOLARE, SOLESS, MONDOC, " _
          & "  SDOOLD, CARGOS, ABONOS, SDONEW, IMPORTE, CONCEPTO, PARIENTE, LINPARIE, NOMBRE, " _
          & "  NUMOPE, NUMFRA, LINFRA, USU ) " _
          & " VALUES " _
          & " ('" + wAno + "', '" + wMes + "', '2', '004', '" + wDoc + "', " _
          & "  '" + Format(wLin, "00") + "', '" + wMesCob + "', '155', 0, " + Str(wSdo) + ", " _
          & "  '" + wMon + "', " + Str(wSdo) + ", 0, " + Str(wSdo) + ", 0, " + Str(wSdo) + ", " _
          & "  '01', '', '', '', '" + wNumOpe + "', '', '', '" + wcodusu + "' ) ")
          Db.CommitTrans

          zz = Leerado6a("SELECT * FROM CTASXDET " _
                      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                      & "            MES = '" + wMesCob + "' AND " _
                      & "       CONCEPTO = '01' AND " _
                      & "         TIPCOB = '03' AND " _
                      & "         SERCOB = '004' AND " _
                      & "         NUMCOB = '" + wDoc + "' ")
          If zz = 0 Then
             Db.BeginTrans
             Db.Execute ("INSERT INTO CTASXDET " _
             & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
             & "  TIPCAM, DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW ) " _
             & " VALUES " _
             & " (" + Str(wSoc) + ", '" + wMesCob + "', '01', " _
             & "  '03', '004', '" + wDoc + "', '" + Format(wLin, "00") + "', " _
             & "  '2', '" + Format(wFec, "dd/mm/yyyy") + "', 0, " _
             & "  0, " + Str(wSdo) + ", " _
             & "  0, 0, " + Str(wSdo) + ", 0 ) ")
             Db.CommitTrans
          Else
             Db.BeginTrans
             Db.Execute ("UPDATE CTASXDET " _
             & " SET TIPMOV = '2', FECHA = '" + Format(wFec, "dd/mm/yyyy") + "', " _
             & "     TIPCAM = 0, DOLARE = 0, SOLESS = " + Str(wSdo) + ", " _
             & "     SDOOLD = 0, CARGOS = 0, ABONOS = " + Str(wSdo) + ", SDONEW = 0 " _
             & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
             & "            MES = '" + wMesCob + "' AND " _
             & "       CONCEPTO = '01' AND " _
             & "         TIPCOB = '03' AND " _
             & "         SERCOB = '004' AND " _
             & "         NUMCOB = '" + wDoc + "' AND " _
             & "         LINCOB = '" + Format(wLin, "00") + "' ")
             Db.CommitTrans
          End If
         
          Db.BeginTrans
          Db.Execute ("INSERT INTO COBROCAB " _
          & " (ANO, MES, TIPCOB, SERCOB, NUMCOB, FECHA, MONEDA, IMPORTE, GLOSA, " _
          & "  CODSOCIO, TIPCAM, DOLARE, SOLESS, FORPAG ) " _
          & " VALUES " _
          & " ('" + wAno + "', '" + wMes + "', '2', '004', '" + wDoc + "', " _
          & "  '" + Format(wFec, "dd/mm/yyyy") + "', 'S', " + Str(wImp) + ", " _
          & "  '" + wGlo + "', " + Str(wSoc) + ", 0, 0, " + Str(wImp) + ", '08'  ) ")
          Db.CommitTrans

          zz = Leerado8("SELECT * FROM ZZZ_MRECIBOS " _
                       & " WHERE    SERIE = '004' AND " _
                       & "       CONCEPTO = 155 AND " _
                       & "       NRO_COMP = " + Str(Val(wDoc)) + " AND " _
                       & "       YEAR(FECHA_PAGO) = " + Str(Val(wanocia)) + " ")
          If zz = 0 Then
             Db.BeginTrans
             Db.Execute ("INSERT INTO ZZZ_MRECIBOS " _
             & " (CODIGO, INS, CONCEPTO, SERIE, AUXILIAR, NRO_COMP, MONTO, MONEDA, T_CAMBIO, " _
             & "  FECHA_PAGO, FECHA_CADU, OBS, D_IMPOR, DEUDA_PT2, DINS_CER, ADELANTO, " _
             & "  MARCA1, MARCA2, MARCA3, MARCA4, OBS1 ) " _
             & " VALUES " _
             & " (" + Str(wCod) + ", " + Str(wIns) + ", 155, '004', 0, " _
             & "  " + Str(Val(wDoc)) + ", " + Str(wSdo) + ", " _
             & "  'S/.', 0, " _
             & "  '" + Format(wFec, "dd/mm/yyyy") + "', null, 'BCP OP " + wNumOpe + " APORTES " + wMesCob + "', " _
             & "  '', " + Str(wDeu) + ", 0, " + Str(wAde) + ", " _
             & "  '" + Format(Date, "dd/mm/yyyy") + "', 'N', '" + wcodusu + "', " _
             & "  '" + Format(Time, "hh:mm:ss") + "', '' ) ")
             Db.CommitTrans
          Else
             Db.BeginTrans
             Db.Execute ("UPDATE ZZZ_MRECIBOS " _
             & " SET CODIGO = " + Str(wCod) + ", INS = " + Str(wIns) + ", CONCEPTO = 155, " _
             & "     AUXILIAR = 0, MONTO = " + Str(wSdo) + ", " _
             & "     MONEDA = 'S/.', " _
             & "     T_CAMBIO = 0, FECHA_PAGO = '" + Format(wFec, "dd/mm/yyyy") + "', " _
             & "     FECHA_CADU = null, OBS = 'BCP OP " + wNumOpe + " APORTES " + wMesCob + "', D_IMPOR = '', " _
             & "     DEUDA_PT2 = 0, DINS_CER = 0, ADELANTO = 0, " _
             & "     MARCA1 = '" + Format(Date, "dd/mm/yyyy") + "', MARCA2 = 'N', " _
             & "     MARCA3 = '" + wcodusu + "', MARCA4 = '" + Format(Time, "hh:mm:ss") + "', " _
             & "     OBS1 = '' " _
             & " WHERE    SERIE = '004' AND " _
             & "       CONCEPTO = 155 AND " _
             & "       NRO_COMP = " + Str(Val(wDoc)) + " AND " _
             & "       YEAR(FECHA_PAGO) = " + Str(Val(wanocia)) + " ")
             Db.CommitTrans
          End If
       
       End If
   
   Next
   objExcel.Workbooks.Close
   Set objExcel = Nothing
    
   DoEvents
   lblMensaje.Caption = ""
   lblMensaje.Refresh
    
   MsgBox "Archivo " + wNom + " Se Ha Importado OK", vbExclamation
   Unload Me
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Function BusUltRecibo() As String
   On Error GoTo err

   Dim zz As Integer, zDoc As String
   zDoc = "000000000"
   
   zz = Leerado6a("SELECT MAX(CAST(NUMCOB AS INT)) AS NUMCOB " _
                & " From COBROCAB " _
                & " WHERE TIPCOB = '2' AND " _
                & "       SERCOB = '004' AND ANO = '" + wanocia + "'")
   If zz > 0 Then
      zDoc = IIf(IsNull(ADO6a!numcob), "0000000000", ADO6a!numcob)
   End If
   Set ADO6a = Nothing

   zDoc = Format(Val(zDoc) + 1, "0000000000")

   BusUltRecibo = zDoc
   Exit Function
err:
   MsgBox err.Description + " " + err.Description
   Resume Next
End Function

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Form_Initialize
End Sub

Private Sub Form_Initialize()
   frmBCPRetorno.Left = (Screen.Width - Width) \ 2
   frmBCPRetorno.Top = 0
   
   txtDesde.Text = "__/__/____"
   txtHasta.Text = "__/__/____"
   
   txtDesde.SetFocus
End Sub

Private Sub txtDesde_GotFocus()
   txtDesde.SelStart = 0
   txtDesde.SelLength = 10
End Sub

Private Sub txtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 40
        txtHasta.SetFocus
   End Select
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtDesde.Text = "__/__/____" Then
         MsgBox "Fecha Inicial En Blanco", vbExclamation
         txtDesde.Text = "__/__/____"
         Exit Sub
      End If
      txtHasta.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtHasta_GotFocus()
   txtHasta.SelStart = 0
   txtHasta.SelLength = 10
End Sub

Private Sub txtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDesde.SetFocus
   End Select
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtHasta.Text = "__/__/____" Then
         MsgBox "Fecha Final En Blanco", vbExclamation
         txtHasta.Text = "__/__/____"
         Exit Sub
      End If
      cmdLeer.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub
