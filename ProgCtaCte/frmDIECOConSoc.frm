VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDIECOConSoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DIECO - Consulta de Socios que se Envia"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   13365
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
      Left            =   11760
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7440
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
      Height          =   615
      Left            =   9120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7440
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
      Left            =   3840
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   12975
      _ExtentX        =   22886
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
      TabIndex        =   8
      Top             =   7320
      Width           =   7815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Asignados"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10800
      TabIndex        =   5
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Titulares"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10800
      TabIndex        =   4
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label lblCanAsi 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11880
      TabIndex        =   3
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label lblCanApo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11880
      TabIndex        =   2
      Top             =   6720
      Width           =   855
   End
End
Attribute VB_Name = "frmDIECOConSoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LlenaCab()
   Dim zz As Integer, wRegAct As Integer, wRegTot As Integer, _
       wAno As String, wMes As String, wNom As String, _
       wFecEnv As Date, wFecDsc As Date, _
       wSoc As Integer, wCod As Long, wIns As Integer, wApo As Currency, wMon As String, _
       wCodAsig1 As Integer, wCodAsig2 As Integer, wCodAsig3 As Integer, wCodAsig4 As Integer, wCodAsig5 As Integer, _
       wNomAsig1 As String, wNomAsig2 As String, wNomAsig3 As String, wNomAsig4 As String, wNomAsig5 As String, _
       wTotAsig1 As Currency, wTotAsig2 As Currency, wTotAsig3 As Currency, wTotAsig4 As Currency, wTotAsig5 As Currency, _
       wDeuAsig1 As Currency, wDeuAsig2 As Currency, wDeuAsig3 As Currency, wDeuAsig4 As Currency, wDeuAsig5 As Currency, _
       wAdeAsig1 As Currency, wAdeAsig2 As Currency, wAdeAsig3 As Currency, wAdeAsig4 As Currency, wAdeAsig5 As Currency, _
       wNetAsig1 As Currency, wNetAsig2 As Currency, wNetAsig3 As Currency, wNetAsig4 As Currency, wNetAsig5 As Currency, _
       wTotDeuda As Currency, wTotAdela As Currency, wTotEnvio As Currency, _
       wNetSocio As Currency, wLin As Integer, WE_S As String, wCodPnp As Long, wInsPnp As Integer, _
       wMesOld As String
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_DIECOSOC WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   zz = Leerado8("SELECT M.CODSOCIO, M.CODIGO, M.INS, M.NOMBRE, E.APORTE, E.MONEDA, M.ADELANTO, " _
             & "         M.DEUDA_PT2, M.TIPCOB, M.E_SOCIO " _
             & " FROM MAESOCIO AS M INNER JOIN MAEE_SOCIO AS E " _
             & "   ON M.E_SOCIO = E.E_SOCIO " _
             & " WHERE (M.TIPCOB = '01' OR M.TIPCOB = '04') AND " _
             & "       (M.E_SOCIO <> 'ESP') AND " _
             & "       (M.E_SOCIO <> 'EXC') AND " _
             & "       (M.E_SOCIO <> 'EXP') AND " _
             & "       (M.E_SOCIO <> 'FAL') AND " _
             & "       (M.E_SOCIO <> 'HON') AND " _
             & "       (M.E_SOCIO <> 'REN') AND " _
             & "       (M.E_SOCIO <> 'SEP') AND " _
             & "       (M.INS <> 7) AND " _
             & "       (M.SITU <> 6) ")
   If zz > 0 Then
      ADO8.MoveFirst
      wRegAct = 1
      wRegTot = zz
      Do While Not ADO8.EOF
         DoEvents
         lblMensaje.Caption = "Registro " + _
                              Trim(Format(wRegAct, "####0")) + " / " + _
                              Trim(Format(wRegTot, "####0"))
         lblMensaje.Refresh
         
         wSoc = ADO8!codsocio
         wCod = ADO8!codigo
         wIns = ADO8!ins
         wNom = Trim(ADO8!nombre)
         wApo = ADO8!aporte
         wMon = ADO8!moneda
         WE_S = ADO8!e_socio
         
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECOSOC " _
   & " (MES, CODENVIO, LIN, CODSOCIO, CODIGO, INS, NOMBRE, " _
   & "  FECENV, FECDSC, TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "  TOTENVIO, DSCDIECO, DSCDIFER, USU ) " _
   & " SELECT " _
   & "  C.MES, C.CODSOCIO, '1', C.CODASIG1, M.CODIGO, M.INS, " _
   & "  M.NOMBRE, C.FECENV, C.FECDSC, C.TOTASIG1, C.DEUASIG1, " _
   & "  C.ADEASIG1, C.NETASIG1, C.DSCASIG1, C.DIFASIG1, 0, 0, 0, '" + wcodusu + "' " _
   & " FROM DIECOCAB AS C INNER JOIN MAESOCIO AS M " _
   & "   ON C.CODASIG1 = M.CODSOCIO " _
   & " WHERE C.MES >= '" + wDesAno + wDesMes + "' AND " _
   & "       C.MES <= '" + wHasAno + wHasMes + "' AND " _
   & "       C.CODASIG1 = " + Str(wSoc) + " ")
   Db.CommitTrans
   
         
         
         
         
         wTotEnvio = 0: wTotAdela = 0: wTotDeuda = 0
         wCodAsig1 = 0: wNomAsig1 = "": wTotAsig1 = 0: wDeuAsig1 = 0: wAdeAsig1 = 0: wNetAsig1 = 0
         wCodAsig2 = 0: wNomAsig2 = "": wTotAsig2 = 0: wDeuAsig2 = 0: wAdeAsig2 = 0: wNetAsig2 = 0
         wCodAsig3 = 0: wNomAsig3 = "": wTotAsig3 = 0: wDeuAsig3 = 0: wAdeAsig3 = 0: wNetAsig3 = 0
         wCodAsig4 = 0: wNomAsig4 = "": wTotAsig4 = 0: wDeuAsig4 = 0: wAdeAsig4 = 0: wNetAsig4 = 0
         wCodAsig5 = 0: wNomAsig5 = "": wTotAsig5 = 0: wDeuAsig5 = 0: wAdeAsig5 = 0: wNetAsig5 = 0
         
         wTotDeuda = SaldoFoto(wSoc, wMesOld)
         If wTotDeuda < 0 Then
            wTotAdela = -wTotDeuda
            wTotDeuda = 0
         End If
'         wTotAdela = ADO8!adelanto
'         wTotDeuda = ADO8!deuda_pt2
         
         ' Si el Socio Tiene Adelantos Mayor a Aporte No Se Envia
         ' En Caso El socio tiene adelantos se descuentan Aporte - Adelanto
         '
         If wTotAdela >= wApo Then
            wNetSocio = 0
         Else
            wNetSocio = wApo - wTotAdela
         End If
         
         ' Si la deuda es mayor a 6 Cuotas NO Se Envia
         If wTotDeuda > 0 And wTotDeuda < Round(6 * wApo, 2) Then
            If wTotDeuda > wApo Then
               wNetSocio = wNetSocio + wApo
            Else
               wNetSocio = wNetSocio + wTotDeuda
            End If
         End If
         wCodPnp = 0
         wInsPnp = 0
         
         ' Solo Para Socios TipCob 04 PNP (Parientes PNP que paga la cuota)
         If ADO8!tipcob = "04" Then
            zz = Leerado7("SELECT * FROM MAEPNP " _
                        & " WHERE CODSOCIO1 = " + Str(wSoc) + " OR " _
                        & "       CODSOCIO2 = " + Str(wSoc) + " OR " _
                        & "       CODSOCIO3 = " + Str(wSoc) + " OR " _
                        & "       CODSOCIO4 = " + Str(wSoc) + " OR " _
                        & "       CODSOCIO5 = " + Str(wSoc) + " ")
            If zz = 0 Then
               MsgBox "Codofin " + Trim(Str(wCod)) + "-" + Trim(Str(wIns)) + " " + wNom + vbNewLine + _
                      "No Tiene PNP ASOCIADO", vbExclamation
               Exit Sub
            End If
            wCodAsig1 = wSoc
            wNomAsig1 = wNom
            wTotAsig1 = wApo
         
            wCodPnp = ADO7!codsocio
            wInsPnp = ADO7!ins
         
            wCod = ADO7!codigo
            wIns = ADO7!ins
            wNom = ADO7!nombre
            wApo = 0
            wMon = ""
            wTotDeuda = 0
            wTotAdela = 0
            wNetSocio = 0
         Else
            
            ' Hijos Asignados
            zz = Leerado7("SELECT D.LIN, D.CODHIJO, M.NOMBRE, M.ADELANTO, M.DEUDA_PT2 " _
                 & " FROM MAEASIGNADO AS D INNER JOIN MAESOCIO AS M " _
                 & "   ON D.CODHIJO = M.CODSOCIO " _
                 & " WHERE D.CODSOCIO = " + Str(wSoc) + " AND " _
                 & "         D.ESTADO = 'H' AND FECTOP IS NULL " _
                 & " ORDER BY D.LIN ")
            If zz > 0 Then
               ADO7.MoveFirst
               wLin = 1
               Do While Not ADO7.EOF
                  Select Case wLin
                  Case 1
                       wCodAsig1 = ADO7!codhijo
                       wNomAsig1 = ADO7!nombre
                       wTotAsig1 = wApo
                       
'                       wDeuAsig1 = ADO7!deuda_pt2
'                       wAdeAsig1 = ADO7!adelanto
                       wAdeAsig1 = 0
                       wDeuAsig1 = SaldoFoto(wCodAsig1, wMesOld)
                       If wDeuAsig1 < 0 Then
                          wAdeAsig1 = -wDeuAsig1
                          wDeuAsig1 = 0
                       End If
                       
                       If wAdeAsig1 >= wApo Then
                          wNetAsig1 = 0
                       Else
                          wNetAsig1 = wApo - wAdeAsig1
                       End If
                       If wDeuAsig1 > 0 And wDeuAsig1 < Round(6 * wApo, 2) Then
                          If wDeuAsig1 > wApo Then
                             wNetAsig1 = wNetAsig1 + wApo
                          Else
                             wNetAsig1 = wNetAsig1 + wDeuAsig1
                          End If
                       End If
                  Case 2
                       wCodAsig2 = ADO7!codhijo
                       wNomAsig2 = ADO7!nombre
                       
                       wTotAsig2 = wApo
'                       wDeuAsig2 = ADO7!deuda_pt2
'                       wAdeAsig2 = ADO7!adelanto
                       wAdeAsig2 = 0
                       wDeuAsig2 = SaldoFoto(wCodAsig2, wMesOld)
                       If wDeuAsig2 < 0 Then
                          wAdeAsig2 = -wDeuAsig2
                          wDeuAsig2 = 0
                       End If
                       
                       If wAdeAsig2 >= wApo Then
                          wNetAsig2 = 0
                       Else
                          wNetAsig2 = wApo - wAdeAsig2
                       End If
                       If wDeuAsig2 > 0 And wDeuAsig2 < Round(6 * wApo, 2) Then
                          If wDeuAsig2 > wApo Then
                             wNetAsig2 = wNetAsig2 + wApo
                          Else
                             wNetAsig2 = wNetAsig2 + wDeuAsig2
                          End If
                       End If
                  Case 3
                       wCodAsig3 = ADO7!codhijo
                       wNomAsig3 = ADO7!nombre
                       
                       wTotAsig3 = wApo
'                       wDeuAsig3 = ADO7!deuda_pt2
'                       wAdeAsig3 = ADO7!adelanto
                       wAdeAsig3 = 0
                       wDeuAsig3 = SaldoFoto(wCodAsig3, wMesOld)
                       If wDeuAsig3 < 0 Then
                          wAdeAsig3 = -wDeuAsig3
                          wDeuAsig3 = 0
                       End If
                       
                       If wAdeAsig3 >= wApo Then
                          wNetAsig3 = 0
                       Else
                          wNetAsig3 = wApo - wAdeAsig3
                       End If
                       If wDeuAsig3 > 0 And wDeuAsig3 < Round(6 * wApo, 2) Then
                          If wDeuAsig3 > wApo Then
                             wNetAsig3 = wNetAsig3 + wApo
                          Else
                             wNetAsig3 = wNetAsig3 + wDeuAsig3
                          End If
                       End If
                  Case 4
                       wCodAsig4 = ADO7!codhijo
                       wNomAsig4 = ADO7!nombre
                       
                       wTotAsig4 = wApo
'                       wDeuAsig4 = ADO7!deuda_pt2
'                       wAdeAsig4 = ADO7!adelanto
                       wAdeAsig4 = 0
                       wDeuAsig4 = SaldoFoto(wCodAsig4, wMesOld)
                       If wDeuAsig4 < 0 Then
                          wAdeAsig4 = -wDeuAsig4
                          wDeuAsig4 = 0
                       End If
                       
                       If wAdeAsig4 >= wApo Then
                          wNetAsig4 = 0
                       Else
                          wNetAsig4 = wApo - wAdeAsig4
                       End If
                       If wDeuAsig4 > 0 And wDeuAsig4 < Round(6 * wApo, 2) Then
                          If wDeuAsig4 > wApo Then
                             wNetAsig4 = wNetAsig4 + wApo
                          Else
                             wNetAsig4 = wNetAsig4 + wDeuAsig4
                          End If
                       End If
                  Case 5
                       wCodAsig5 = ADO7!codhijo
                       wNomAsig5 = ADO7!nombre
                       
                       wTotAsig5 = wApo
'                       wDeuAsig5 = ADO7!deuda_pt2
'                       wAdeAsig5 = ADO7!adelanto
                       wAdeAsig5 = 0
                       wDeuAsig5 = SaldoFoto(wCodAsig5, wMesOld)
                       If wDeuAsig5 < 0 Then
                          wAdeAsig5 = -wDeuAsig5
                          wDeuAsig5 = 0
                       End If
                       
                       If wAdeAsig5 >= wApo Then
                          wNetAsig5 = 0
                       Else
                          wNetAsig5 = wApo - wAdeAsig5
                       End If
                       If wDeuAsig5 > 0 And wDeuAsig5 < Round(6 * wApo, 2) Then
                          If wDeuAsig5 > wApo Then
                             wNetAsig5 = wNetAsig5 + wApo
                          Else
                             wNetAsig5 = wNetAsig5 + wDeuAsig5
                          End If
                       End If
                  End Select
              
                  wLin = wLin + 1
                  ADO7.MoveNext
               Loop
            End If
         End If
         
         wTotEnvio = wNetSocio + _
                     wNetAsig1 + wNetAsig2 + wNetAsig3 + wNetAsig4 + wNetAsig5
   
         If wTotEnvio > 0 Then
            Db.BeginTrans
            Db.Execute ("INSERT INTO TMP_DIECOCAB " _
            & " (MES, CODSOCIO, CODIGO, INS, NOMBRE, FECENV, FECDSC, " _
            & "   TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, TOTENVIO, " _
            & "   NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, " _
            & "   ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
            & "   DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
            & "   TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
            & "   CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
            & "   NOMASIG1, NOMASIG2, NOMASIG3, NOMASIG4, NOMASIG5, " _
            & "   TIPCOB  , CODPNP  , INSPNP  , E_SOCIO, USU ) " _
            & " VALUES " _
            & " ('" + wAno + wMes + "', " + Str(wSoc) + ", " + Str(wCod) + ", " + Str(wIns) + ", " _
            & "  '" + Trim(wNom) + "', '" + Format(Date, "dd/mm/yyyy") + "', " _
            & "  '" + Format(Date, "dd/mm/yyyy") + "', " + Str(wApo) + ", " + Str(wTotDeuda) + ", " _
            & "  " + Str(wTotAdela) + ", " + Str(wNetSocio) + ", " + Str(wTotEnvio) + ", " _
            & "  " + Str(wNetAsig1) + ", " + Str(wNetAsig2) + ", " + Str(wNetAsig3) + ", " + Str(wNetAsig4) + ", " + Str(wNetAsig5) + ",  " _
            & "  " + Str(wAdeAsig1) + ", " + Str(wAdeAsig2) + ", " + Str(wAdeAsig3) + ", " + Str(wAdeAsig4) + ", " + Str(wAdeAsig5) + ",  " _
            & "  " + Str(wDeuAsig1) + ", " + Str(wDeuAsig2) + ", " + Str(wDeuAsig3) + ", " + Str(wDeuAsig4) + ", " + Str(wDeuAsig5) + ",  " _
            & "  " + Str(wTotAsig1) + ", " + Str(wTotAsig2) + ", " + Str(wTotAsig3) + ", " + Str(wTotAsig4) + ", " + Str(wTotAsig5) + ",  " _
            & "  " + Str(wCodAsig1) + ", " + Str(wCodAsig2) + ", " + Str(wCodAsig3) + ", " + Str(wCodAsig4) + ", " + Str(wCodAsig5) + ",  " _
            & "  '" + Trim(wNomAsig1) + "', '" + Trim(wNomAsig2) + "', '" + Trim(wNomAsig3) + "', " _
            & "  '" + Trim(wNomAsig4) + "', '" + Trim(wNomAsig5) + "', '" + ADO8!tipcob + "', " _
            & "  " + Str(wCodPnp) + ", " + Str(wInsPnp) + ", '" + WE_S + "', '" + wcodusu + "' ) ")
            Db.CommitTrans
         End If
   
         wRegAct = wRegAct + 1
         ADO8.MoveNext
      Loop
   End If

   lblMensaje.Caption = ""
   lblMensaje.Refresh

   zz = Leerado2("SELECT CODSOCIO, CODIGO  , INS     , NOMBRE  , NETSOCIO, " _
                & "      NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, " _
                & "      TOTENVIO, " _
                & "      TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
                & "      ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
                & "      DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
                & "      USU     , MES     , TOTAPORT, TOTDEUDA, TOTADELA, " _
                & "      DSCDIECO, DSCSOCIO, DSCDIFER, TIPCOB  , CODPNP  , INSPNP,  " _
                & "      DSCASIG1, DSCASIG2, DSCASIG3, DSCASIG4, DSCASIG5, " _
                & "      CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
                & "      NOMASIG1, NOMASIG2, NOMASIG3, NOMASIG4, NOMASIG5 " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE MES = '" + wAno + wMes + "' AND USU = '" + wcodusu + "' " _
                & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO2

'   LlenaCab
   LlenaCab1
   LabelCab
   TotalCab
End Sub
