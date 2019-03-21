VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBCPEnvio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envio Mensual BCP"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   12375
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
      Left            =   10800
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCreaTXT 
      Caption         =   "&Crear TXT"
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
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7200
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
      Left            =   1440
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
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
      Left            =   9360
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCalculo 
      Caption         =   "Calcular Envio"
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
      Left            =   4800
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar Cálculo"
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
      Left            =   6120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox cmbMeses 
      Height          =   315
      ItemData        =   "frmBCPEnvio.frx":0000
      Left            =   840
      List            =   "frmBCPEnvio.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   660
      Width           =   2535
   End
   Begin VB.TextBox txtAnoCab 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5775
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   10186
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
   Begin MSMask.MaskEdBox txtFecha 
      Height          =   285
      Left            =   3480
      TabIndex        =   17
      Top             =   660
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label14 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   480
      Width           =   975
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
      Left            =   2760
      TabIndex        =   16
      Top             =   7320
      Width           =   5415
   End
   Begin VB.Label lblCanApo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8160
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9840
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Titulares"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8160
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "S/. Enviado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9600
      TabIndex        =   7
      Top             =   360
      Width           =   1455
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
      Left            =   240
      TabIndex        =   4
      Top             =   660
      Width           =   495
   End
   Begin VB.Label Label25 
      Caption         =   "Año"
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
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ENVIA DESCUENTOS A BCP"
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
      Left            =   2640
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmBCPEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Inicio(sw As Boolean)
   cmdCalculo.Enabled = sw
   cmdEliminar.Enabled = sw
   cmdExportar.Enabled = sw
   cmdCreaTXT.Enabled = sw
   cmdGrabar.Enabled = sw
End Sub

Private Sub cmbMeses_Click()
   cmbMeses_KeyPress (13)
End Sub

Private Sub cmbMeses_KeyPress(KeyAscii As Integer)
   Dim zz As Integer, wAno As String, wMes As String
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
   
   If KeyAscii = 13 Then
      Set DataGrid1.DataSource = Nothing
      
      Inicio True
      
      zz = Leerado2("SELECT * FROM BCPCAB " _
                & "  WHERE MES = '" + wAno + wMes + "' ")
      If zz > 0 Then
      
         txtFecha.Text = Format(ADO2!fecenv, "dd/mm/yyyy")
      
         lblMensaje.Caption = "Trae Calculo BCP - Mes " + Left(Trim(funnommes(wMes)), 3) + " " + wAno
         lblMensaje.Refresh
      
         LlenaCab
         LlenaCab1
         TotalCab
         ADO2.MoveFirst
      
         lblMensaje.Caption = ""
         lblMensaje.Refresh
      
      Else
         txtFecha.Text = Format("01/" + wMes + "/" + wAno, "dd/mm/yyyy")
         
         txtFecha.SetFocus
      End If
   End If
End Sub

Private Sub cmdCalculo_Click()
   Dim zz As Integer, wRegAct As Integer, wRegTot As Integer, _
       wAno As String, wMes As String, wNom As String, _
       wFecEnv As Date, wFecDsc As Date, _
       wSoc As Integer, wCod As Long, wIns As Integer, wApo As Currency, wMon As String, _
       wTotAdela As Currency, wTotEnvio As Currency, _
       wNetSocio As Currency, wTotDeuda As Currency, wFrac As Currency
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
   wFecEnv = Format(txtFecha.Text, "dd/mm/yyyy")
   
   zz = Leerado8("SELECT * FROM BCPCAB " _
             & "  WHERE MES = '" + wAno + wMes + "' ")
   If zz > 0 Then
      If MsgBox("Ya Existe Proceso BCP del Mes" + vbNewLine + _
                "Desea Volver a Crearlo???", vbYesNo + vbQuestion, "Crear Archivo Descuento BCP") = vbNo Then
         Exit Sub
      End If
   End If
   Set ADO8 = Nothing
   
   Set DataGrid1.DataSource = Nothing
   
   lblMensaje.Caption = "Calculando Descuentos BCP - Mes " + Trim(funnommes(wMes)) + " " + wAno
   lblMensaje.Refresh
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_BCPCAB WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM BCPCAB WHERE MES = '" + wAno + wMes + "' ")
   Db.CommitTrans
   
   zz = Leerado8("SELECT M.CODSOCIO, M.CODIGO, M.INS, M.NOMBRE, M.TIPCOB, M.E_SOCIO, E.APORTE, E.MONEDA, M.ADELANTO, M.DEUDA_PT2 " _
             & " FROM MAESOCIO AS M INNER JOIN MAEE_SOCIO AS E " _
             & "   ON M.E_SOCIO = E.E_SOCIO " _
             & " WHERE E.MONEDA = 'S' AND " _
             & "       E.APORTE > 0 " _
             & " ORDER BY M.CODSOCIO ")
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
         wNom = ADO8!nombre
         wApo = ADO8!aporte
         wMon = ADO8!moneda
         wTotDeuda = 0
         wTotAdela = 0

         wTotDeuda = SaldoFoto(wSoc, zMesTope)
         If wTotDeuda < 0 Then
            wTotAdela = -wTotDeuda
            wTotDeuda = 0
         End If
            
         wFrac = 0
         zz = Leerado6a("SELECT SUM(SDONEW) AS SDONEW " _
                     & " FROM FRACDET AS D INNER JOIN FRACCAB AS C " _
                     & "   ON D.NUMERO = C.NUMERO " _
                     & " WHERE C.CODSOCIO = " + Str(wSoc) + " AND " _
                     & "       D.SDONEW > 0 AND " _
                     & "       D.VCMTO <= '" + Format(wFecEnv, "dd/mm/yyyy") + "' ")
         If zz > 0 Then
            wFrac = IIf(IsNull(ADO6a!sdonew), 0, ADO6a!sdonew)
         End If
         Set ADO6a = Nothing
            
         wTotEnvio = wApo + wTotDeuda - wTotAdela + wFrac
            
         If wTotAdela = 0 Then
            Db.BeginTrans
            Db.Execute ("INSERT INTO TMP_BCPCAB " _
            & " (MES, CODSOCIO, CODIGO, INS, NOMBRE, FECENV, FECDSC, " _
            & "   TOTDEUDA, TOTADELA, TOTAPORT, TOTENVIO, TOTFRACC, USU ) " _
            & " VALUES " _
            & " ('" + wAno + wMes + "', " + Str(wSoc) + ", " + Str(wCod) + ", " + Str(wIns) + ", " _
            & "  '" + Trim(wNom) + "', '" + Format(wFecEnv, "dd/mm/yyyy") + "', " _
            & "  '" + Format(wFecEnv, "dd/mm/yyyy") + "', " _
            & "  " + Str(wTotDeuda) + ", " + Str(wTotAdela) + ", " _
            & "  " + Str(wApo) + ", " + Str(wTotEnvio) + ", " + Str(wFrac) + ", " _
            & "  '" + wcodusu + "' ) ")
            Db.CommitTrans
         End If
   
         wRegAct = wRegAct + 1
         ADO8.MoveNext
      Loop
   End If

   lblMensaje.Caption = ""
   lblMensaje.Refresh

   zz = Leerado2("SELECT CODSOCIO, CODIGO  , INS     , NOMBRE  , TOTDEUDA, " _
                & "      TOTADELA, TOTAPORT, TOTENVIO, TOTFRACC, USU     , MES " _
                & " FROM TMP_BCPCAB " _
                & " WHERE MES = '" + wAno + wMes + "' AND USU = '" + wcodusu + "' " _
                & " ORDER BY CODIGO, INS ")
   Set DataGrid1.DataSource = ADO2

'   LlenaCab
   LlenaCab1
   TotalCab
End Sub

Private Sub cmdCreaTXT_Click()
   Dim zz As Long, wRegAct As Long, wRegTot As Long, _
       wAno As String, wMes As String, wCod As Integer, wIns As Integer, _
       wDir As String, wDir2 As String, wFile As String, _
       wCanSoc As Integer, wTotEnv As Currency, wFecEnv As Date, _
       wNom As String
         
   wAno = txtAnoCab.Text
   wMes = Left(cmbMeses.Text, 2)
   wFecEnv = Format(txtFecha.Text, "dd/mm/yyyy")
   
   wDir = xraizBCP + wAno
   wDir2 = xraizBCP + wAno + "\" + wAno + "-" + wMes
   wFile = wDir2 + "\CDPG.TXT"
   
   If Len(Dir(wDir, vbDirectory)) = 0 Then
      MkDir wDir
   End If
   
   If Len(Dir(wDir2, vbDirectory)) = 0 Then
      MkDir wDir2
   End If
   
   If Len(Dir$(wFile)) > 0 Then
      Kill wFile
   End If
   
   Open wFile For Output As #1
   
   zz = Leerado8("SELECT COUNT(*) AS CAN, SUM(TOTENVIO) AS TOTENVIO " _
                & " FROM TMP_BCPCAB " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "       MES = '" + wAno + wMes + "' ")
   If zz > 0 Then
      wCanSoc = ADO8!can
      wTotEnv = ADO8!totenvio
   End If
   Set ADO8 = Nothing
      
   Print #1, "CC" + "194" + "0" + "1177014" + "C" + _
             "ASOCIACION DE OFICIALES PIP             " + _
             Format(wFecEnv, "yyyymmdd") + _
             Format(wCanSoc, "000000000") + _
             Format(Int(wTotEnv), "0000000000000") + _
             Format((wTotEnv - Int(wTotEnv)) * 100, "00") + _
             Space(114)
   
   zz = Leerado8("SELECT * FROM TMP_BCPCAB WHERE USU = '" + wcodusu + "' AND MES = '" + wAno + wMes + "' ORDER BY CODIGO, INS ")
   If zz > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wNom = LaEne(ADO8!nombre)
   
         Print #1, "DD" + "194" + "0" + "1177014" + _
                   Format(ADO8!codigo, "00000000000000") + _
                   LlenaDat(wNom, 40) + _
                   LlenaDat(wNom, 30) + _
                   Format(wFecEnv, "yyyymmdd") + _
                   wAno + wMes + fundiames(wMes) + _
                   Format(Int(ADO8!totenvio), "0000000000000") + _
                   Format((ADO8!totenvio - Int(ADO8!totenvio)) * 100, "00") + _
                   "000000000000000" + _
                   Format(Int(ADO8!totaport), "0000000") + _
                   Format((ADO8!totaport - Int(ADO8!totaport)) * 100, "00") + _
                   Space(48)
         
         ADO8.MoveNext
      Loop
   End If
   
   Close #1

   MsgBox "Proceso Termino OK", vbExclamation

   MsgBox "El Archivo Creado se encontrará en " + _
          App.Path + "\BCP\" + wAno + "-" + wMes
End Sub

Private Function LaEne(zNom As String) As String
   On Error GoTo err
   
   Dim II As Integer, zNew As String
   zNew = ""
   For II = 1 To Len(Trim(zNom))
       If Mid(zNom, II, 1) = "Ñ" Then
          zNew = zNew + "N"
       Else
          zNew = zNew + Mid(zNom, II, 1)
       End If
   Next II
   
   LaEne = zNew
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Private Sub cmdEliminar_Click()
   Dim wAno As String, wMes As String
   
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
   
   If MsgBox("Esta Seguro de Querer Eliminar Calculo " + vbNewLine + _
             "del Mes " + Trim(funnommes(wMes)) + " " + wAno + _
             "???", vbYesNo + vbQuestion, "Eliminar Archivo BCP") = vbNo Then
      Exit Sub
   End If
   Set DataGrid1.DataSource = Nothing
   
   lblTotal.Caption = ""
   lblCanApo.Caption = ""

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_BCPCAB WHERE USU = '" + wcodusu + "' AND MES = '" + wAno + wMes + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM BCPCAB WHERE MES = '" + wAno + wMes + "' ")
   Db.CommitTrans

   MsgBox "Calculo de Mes " + Trim(funnommes(wMes)) + "-" + wAno + " Eliminado OK", vbExclamation
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(6) As String, wRegAct As Integer, wRegTot As Integer
   Dim wNom As String, wMes As String, _
       wTotSoc As Currency, wTotAsi As Currency, wTotEnvio As Currency
   wMes = Left(cmbMeses.Text, 2)
   
   Heading(0) = "NRO"
   Heading(1) = "CODOFIN"
   Heading(2) = "NOMBRE"
   Heading(3) = "TOT.DEUDA"
   Heading(4) = "ADELANTOS"
   Heading(5) = "CUOTA MES"
   Heading(6) = "TOT.ENVIO"
   aa = Leerado3("SELECT * FROM TMP_BCPCAB WHERE USU = '" + wcodusu + "' ORDER BY NOMBRE ")
   If aa > 0 Then
      wRegAct = 1
      wRegTot = aa
      Set objExcel = New Excel.Application
      objExcel.SheetsInNewWorkbook = 1
      objExcel.Workbooks.Add
      With objExcel
           
           .Range(objExcel.Cells(1, 1), .Cells(1, 7)).Merge
           .Range(objExcel.Cells(1, 1), .Cells(1, 7)).HorizontalAlignment = xlCenter
           
           .Range(objExcel.Cells(2, 1), .Cells(2, 7)).Merge
           .Range(objExcel.Cells(2, 1), .Cells(2, 7)).HorizontalAlignment = xlCenter
           
           .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
           .Range(.Cells(3, 1), .Cells(3, 7)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 7)).Font.Bold = True
           .Cells(1, 1) = "REPORTE DE ENVIOS AL BCP " + funnommes(wMes) + "-" + wanocia
           .Cells(2, 1) = "POR CONCEPTO DE CUOTAS DE APORTACION MENSUAL"
           
           For I = 1 To 7 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           
           .Range("A1:G1").Merge
           .Range("A2:G2").Merge
           
'           .Range(objExcel.Cells(3, 1), .Cells(3, 8)).Range
'           .Range(objExcel.Cells(3, 1), .Cells(3, 8)).VerticalAlignment = xlCenter
           
           
           objExcel.Columns("A").ColumnWidth = 8
           objExcel.Columns("B").ColumnWidth = 11
           objExcel.Columns("C").ColumnWidth = 50
           objExcel.Columns("D").ColumnWidth = 12
           objExcel.Columns("E").ColumnWidth = 12
           objExcel.Columns("F").ColumnWidth = 12
           objExcel.Columns("G").ColumnWidth = 12
      End With
      V = 4
      H = 1
      wRegAct = 1
      wTotEnvio = 0
      Do While Not ADO3.EOF
         DoEvents
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + _
                              Format(wRegAct, "####0") + " / " + _
                              Format(wRegTot, "####0")
         lblMensaje.Refresh
         
         objExcel.Range(objExcel.Cells(V, H + 3), objExcel.Cells(V, H + 6)).NumberFormat = "#####,##0.00"
         
         objExcel.Cells(V, H + 0) = wRegAct
         objExcel.Cells(V, H + 1) = Format(ADO3!codigo, "#######0") + Format(ADO3!ins, "0")
         objExcel.Cells(V, H + 2) = ADO3!nombre
         objExcel.Cells(V, H + 3) = ADO3!totdeuda
         objExcel.Cells(V, H + 4) = ADO3!totadela
         objExcel.Cells(V, H + 5) = ADO3!totaport
         objExcel.Cells(V, H + 6) = ADO3!totenvio
         
         wTotEnvio = wTotEnvio + ADO3!totenvio
         wRegAct = wRegAct + 1
         V = V + 1
         ADO3.MoveNext
      Loop
      
      V = V + 1
      objExcel.Range(objExcel.Cells(V, H + 6), objExcel.Cells(V, H + 6)).NumberFormat = "#####,##0.00"
      
      objExcel.Cells(V, H + 2) = "TOTAL ENVIO"
      objExcel.Cells(V, H + 6) = wTotEnvio
      
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

Private Sub cmdGrabar_Click()

   Dim zAno As String, zMes As String, zz As Long
   zAno = wanocia
   zMes = Left(cmbMeses.Text, 2)
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM BCPCAB WHERE MES = '" + zAno + zMes + "' ")
   Db.CommitTrans
 
   Db.BeginTrans
   Db.Execute ("INSERT INTO BCPCAB " _
   & " (MES, CODSOCIO, CODIGO, INS, NOMBRE, FECENV, FECDSC, " _
   & "   TOTDEUDA, TOTADELA, TOTAPORT, TOTENVIO, TOTFRACC ) " _
   & " SELECT " _
   & "  MES, CODSOCIO, CODIGO, INS, NOMBRE, FECENV, FECDSC, " _
   & "   TOTDEUDA, TOTADELA, TOTAPORT, TOTENVIO, TOTFRACC " _
   & " FROM TMP_BCPCAB " _
   & " WHERE MES = '" + zAno + zMes + "' AND " _
   & "       USU = '" + wcodusu + "' ")
   Db.CommitTrans

   MsgBox "Proceso BCP Grabado OK", vbExclamation
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Form_Initialize
End Sub

Private Sub Form_Initialize()
   frmBCPEnvio.Left = (Screen.Width - Width) \ 2
   frmBCPEnvio.Top = 0
   
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
   
   Inicio False
   
   cmbMeses.SetFocus
End Sub

Private Sub LlenaCab()
   Dim wAno As String, wMes As String, zz As Integer
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
      
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_BCPCAB WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_BCPCAB " _
   & " (MES     , CODSOCIO, CODIGO  , INS     , NOMBRE  , FECENV  , FECDSC  , " _
   & "  TOTDEUDA, TOTAPORT, TOTENVIO, TOTFRACC, USU ) " _
   & " SELECT " _
   & "   D.MES, D.CODSOCIO, M.CODIGO, M.INS, M.NOMBRE, D.FECENV, D.FECDSC, " _
   & "   TOTDEUDA, TOTAPORT, TOTENVIO, TOTFRACC, '" + wcodusu + "'  " _
   & " FROM BCPCAB AS D INNER JOIN MAESOCIO AS M " _
   & "   ON D.CODSOCIO = M.CODSOCIO " _
   & " WHERE D.MES = '" + wAno + wMes + "'  ")
   Db.CommitTrans

   zz = Leerado2("SELECT CODSOCIO, CODIGO  , INS     , NOMBRE  , TOTDEUDA, " _
                & "      TOTADELA, TOTAPORT, TOTENVIO, TOTFRACC, USU     , MES " _
                & " FROM TMP_BCPCAB " _
                & " WHERE MES = '" + wAno + wMes + "' AND USU = '" + wcodusu + "' " _
                & " ORDER BY CODIGO, INS ")
   Set DataGrid1.DataSource = ADO2
End Sub

Private Sub LlenaCab1()
   DataGrid1.Columns(0).Width = 750   ' CODSOCIO
   DataGrid1.Columns(0).Alignment = dbgRight
   DataGrid1.Columns(0).Caption = "SOCIO"
    
   DataGrid1.Columns(1).Width = 900   ' CODIGO
   DataGrid1.Columns(1).Alignment = dbgRight
   DataGrid1.Columns(1).Caption = "CODIGO"
    
   DataGrid1.Columns(2).Width = 400   ' INS
   DataGrid1.Columns(2).Alignment = dbgCenter
   DataGrid1.Columns(2).Caption = "INS"
    
   DataGrid1.Columns(3).Width = 4400  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE ASOCIADO"
    
   DataGrid1.Columns(4).Width = 850     ' DEUDA
   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(4).Caption = "T.DEUDA"
    
   DataGrid1.Columns(5).Width = 850     ' ADELANTO
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(5).Caption = "T.DEUDA"
    
   DataGrid1.Columns(6).Width = 850     ' APORTE
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(6).Caption = "APORTE MES"

   DataGrid1.Columns(7).Width = 850     ' TOTFRACC
   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(7).Caption = "FRACC"

   DataGrid1.Columns(8).Width = 850     ' TOT ENVIO
   DataGrid1.Columns(8).Alignment = dbgRight
   DataGrid1.Columns(8).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(8).Caption = "TOT.ENVIO"

   DataGrid1.Columns(9).Visible = False
   DataGrid1.Columns(10).Visible = False
   
   DataGrid1.SetFocus
End Sub

Private Sub TotalCab()
   Dim zz As Integer, _
       zAno As String, zMes As String, _
       zTot As Currency, zCanApo As Integer, zCanAsi As Integer
   
   zAno = wanocia
   zMes = Left(cmbMeses.Text, 2)
   
   zz = Leerado8("SELECT SUM(TOTENVIO) AS TOTENVIO, COUNT(*) AS CAN " _
                & " FROM TMP_BCPCAB " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "       MES = '" + zAno + zMes + "' ")
   If zz > 0 Then
      zTot = IIf(IsNull(ADO8!totenvio), 0, ADO8!totenvio)
      zCanApo = IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   lblTotal.Caption = Format(zTot, "###,##0.00;;\ ")
   lblCanApo.Caption = Format(zCanApo, "##,##0")
End Sub

Private Sub txtFecha_GotFocus()
   txtFecha.SelStart = 0
   txtFecha.SelLength = 10
End Sub

Private Sub txtFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 38
         If txtNumCob.Enabled = True Then
            txtNumCob.SetFocus
         End If
    Case 40
         txtCodSocio.SetFocus
    End Select
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFecha.Text = "__/__/____" Then
         MsgBox "Fecha En Blanco", vbExclamation
         txtFecha.SetFocus
         Exit Sub
      End If
      If Not IsDate(Trim(txtFecha)) Then
         MsgBox "Campo Digitado No Es Fecha Valida", vbExclamation
         txtFecha.Text = "__/__/____"
         txtFecha.SetFocus
         Exit Sub
      End If
      txtFecha.Text = Format(txtFecha.Text, "dd/mm/yyyy")
      If Format(Month(txtFecha.Text), "00") <> Mid(cmbMeses.Text, 1, 2) Then
         MsgBox "Mes Digitado No Corresponde", vbInformation
         txtFecha.Text = "__/__/____"
         txtFecha.SetFocus
         Exit Sub
      End If
      If Format(Year(txtFecha.Text), "0000") <> wanocia Then
         MsgBox "Año Digitado No Corresponde", vbInformation
         txtFecha.Text = "__/__/____"
         txtFecha.SetFocus
         Exit Sub
      End If
      cmdCalculo.SetFocus
   Else
      If InStr(1, "0123456789." + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub


