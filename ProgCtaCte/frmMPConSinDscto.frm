VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmMPConSinDscto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Socios Caja Militar Policial Sin Descuento"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   13350
   Begin VB.CommandButton Command1 
      Caption         =   "&Exportar con Cobros Tesoreria"
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
      Left            =   9240
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton cmdImpTes 
      Caption         =   "&Imprimir con Cobros Tesoreria"
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
      Left            =   10560
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   7920
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   720
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7200
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   7200
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1095
   End
   Begin VB.ComboBox cmbMeses 
      Height          =   315
      ItemData        =   "frmMPConSinDscto.frx":0000
      Left            =   720
      List            =   "frmMPConSinDscto.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtAnoCab 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5535
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   9763
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
      Left            =   12360
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
      Left            =   600
      TabIndex        =   16
      Top             =   7320
      Width           =   7575
   End
   Begin VB.Label lblDscDifer 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10560
      TabIndex        =   11
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Total No Cobrado"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8880
      TabIndex        =   10
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblDscCajMP 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10560
      TabIndex        =   9
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Cobrado"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8880
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblTotEnvio 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10560
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Enviado"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8880
      TabIndex        =   6
      Top             =   480
      Width           =   1575
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
      Left            =   120
      TabIndex        =   4
      Top             =   840
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
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Consulta Socios CMP Sin Descuentos"
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
      TabIndex        =   2
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmMPConSinDscto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbMeses_Click()
   cmbMeses_KeyPress (13)
End Sub

Private Sub cmbMeses_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdBuscar.SetFocus
   End If
End Sub

Private Sub cmdBuscar_Click()
   Dim wAno As String, wMes As String, zz As Integer
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
      
   Call CalxSoc
   
   zz = Leerado2("SELECT CODSOCIO, NOMBRE, " _
                & "      TOTAPORT, TOTDEUDA, TOTADELA , NETSOCIO, DSCSOCIO, DIFSOCIO, " _
                & "      CODENVIO, LIN     , " _
                & "      CODIGO  , INS     , CARNETPNP, NUMDOC  , CODBENI , FECENV  , " _
                & "      FECDSC  , TOTENVIO, DSCCAJMP , DSCDIFER, USU " _
                & " FROM TMP_CAJMPSOC " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY MES, NOMBRE ")
   Set DataGrid1.DataSource = ADO2

   DataGrid1.Columns(0).Width = 750   ' CODSOCIO
   DataGrid1.Columns(0).Alignment = dbgRight
   DataGrid1.Columns(0).Caption = "SOCIO"
    
   DataGrid1.Columns(1).Width = 4900  ' NOMBRE
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Caption = "NOMBRE ASOCIADO"
    
   DataGrid1.Columns(2).Width = 800    ' TOTAPORT
   DataGrid1.Columns(2).Alignment = dbgRight
   DataGrid1.Columns(2).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(2).Caption = "TOT.APORT"
    
   DataGrid1.Columns(3).Width = 800    ' TOTDEUDA
   DataGrid1.Columns(3).Alignment = dbgRight
   DataGrid1.Columns(3).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(3).Caption = "DEUDAS"
    
   DataGrid1.Columns(4).Width = 800    ' TOTADELA
   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(4).Caption = "ADELANTOS"
    
   DataGrid1.Columns(5).Width = 750    ' NETSOCIO
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(5).Caption = "NETO ENVIO"
    
   DataGrid1.Columns(6).Width = 750    ' DSCSOCIO
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(6).Caption = "COBRADO"
    
   DataGrid1.Columns(7).Width = 750    ' DIFSOCIO
   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(7).Caption = "NO COBRADO"
    
   DataGrid1.Columns(8).Width = 750   ' CODENVIO
   DataGrid1.Columns(8).Alignment = dbgRight
   DataGrid1.Columns(8).Caption = "ENVIO"
    
   DataGrid1.Columns(9).Width = 350   ' LIN
   DataGrid1.Columns(9).Alignment = dbgRight
   DataGrid1.Columns(9).Caption = "LIN"
    
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

   cmdImprimir.Enabled = True
   cmdExportar.Enabled = True
   Command1.Enabled = True
   cmdImpTes.Enabled = True
   
   TotalCab
   DataGrid1.SetFocus
End Sub

Private Sub TotalCab()
   Dim zz As Integer, wMes As String, wAno As String, _
       zNetSocio As Currency, _
       zDscSocio As Currency, _
       zDifSocio As Currency
   
   wAno = txtAnoCab.Text
   wMes = Left(cmbMeses.Text, 2)
   zNetSocio = 0: zDscSocio = 0: zDifSocio = 0
   
   zz = Leerado8("SELECT SUM(TOTENVIO) AS TOTENVIO, " _
                & "       SUM(DSCCAJMP) AS DSCCAJMP, " _
                & "       SUM(DSCDIFER) AS DSCDIFER " _
                & " FROM CAJMPCAB " _
                & " WHERE MES = '" + wAno + wMes + "' ")
   If zz > 0 Then
      zNetSocio = IIf(IsNull(ADO8!totenvio), 0, ADO8!totenvio)
      zDscSocio = IIf(IsNull(ADO8!dsccajmp), 0, ADO8!dsccajmp)
      zDifSocio = IIf(IsNull(ADO8!dscdifer), 0, ADO8!dscdifer)
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT  SUM(DIFSOCIO) AS DIFSOCIO " _
                & " FROM TMP_CAJMPSOC " _
                & " WHERE USU = '" + wcodusu + "' ")
   If zz > 0 Then
      zDifSocio = IIf(IsNull(ADO8!difsocio), 0, ADO8!difsocio)
   End If
   Set ADO8 = Nothing
   
   lblTotEnvio.Caption = Format(zNetSocio, "####,##0.00;;\ ")
   lblDscCajMP.Caption = Format(zDscSocio, "####,##0.00;;\ ")
   lblDscDifer.Caption = Format(zDifSocio, "####,##0.00;;\ ")
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(20) As String, _
       wRegAct As Integer, wRegTot As Integer, wAno As String, wMes As String
   Dim wAsi As Integer, wCod As Long, wIns As Integer, _
       wNom As String, wCip As Long, wDni As String, _
       wTotAport As Currency, wTotDeuda As Currency, wTotAdela As Currency, _
       wNetSocio As Currency, wDscSocio As Currency, wDifSocio As Currency, wSoc As Integer, _
       wTel As String, wMai As String, wRef As String, wUbi As String, wDir As String, _
       wUlt As String, wDon As String
   
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
   
   Heading(0) = "NUM"
   Heading(1) = "COD.ENVIO"
   Heading(2) = "COD.SOCIO"
   Heading(3) = "CODOFIN"
   Heading(4) = "CARNET PNP"
   Heading(5) = "DNI"
   Heading(6) = "H"
   Heading(7) = "NOMBRE"
   Heading(8) = "ULT.APORTE"
   Heading(9) = "DONDE"
   Heading(10) = "APORT.MES"
   Heading(11) = "DEUDAS"
   Heading(12) = "ADELANTO"
   Heading(13) = "TOT.ENVIO"
   Heading(14) = "TOT.DSCTO"
   Heading(15) = "NO COBRADO"
   Heading(16) = "TELEFONOS"
   Heading(17) = "DIRECCION"
   Heading(18) = "UBIGEO"
   Heading(19) = "REFERENCIA"
   Heading(20) = "MAIL"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 21)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 21)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "DESCUENTO CAJA MILITAR POLICIAL - MES " + Trim(funnommes(wMes)) + " " + wanocia
        For I = 1 To 21 Step 1
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
        objExcel.Columns("O").ColumnWidth = 11
        objExcel.Columns("P").ColumnWidth = 11
        objExcel.Columns("Q").ColumnWidth = 25
        objExcel.Columns("R").ColumnWidth = 40
        objExcel.Columns("S").ColumnWidth = 18
        objExcel.Columns("T").ColumnWidth = 30
        objExcel.Columns("U").ColumnWidth = 30
   End With
   
   aa = Leerado3("SELECT * FROM TMP_CAJMPSOC " _
                & " WHERE MES = '" + wanocia + wMes + "' AND USU = '" + wcodusu + "' " _
                & " ORDER BY NOMBRE, LIN ")
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
         
         wSoc = ADO3!codsocio
         wTel = "": wDir = "": wRef = "": wMai = "": wUbi = ""
         aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
         If aa > 0 Then
            wTel = Trim(Trim(ADO8!telefono) + " " + Trim(ADO8!telefon2) + " " + Trim(ADO8!celular))
            wDir = Trim(ADO8!direc)
            wRef = Trim(ADO8!refer)
            wMai = Trim(ADO8!email) + " " + Trim(ADO8!email2)
            wUbi = ADO8!ubigeo
            wCod = ADO8!codigo
            wIns = ADO8!ins
         End If
         Set ADO8 = Nothing
         
         If Len(Trim(wUbi)) > 0 Then
            aa = Leerado8("SELECT * FROM MAEUBIGEO WHERE CODIGO = '" + wUbi + "' ")
            If aa > 0 Then
               wUbi = Trim(ADO8!nombre)
            End If
            Set ADO8 = Nothing
         End If
         
         wUlt = "": wDon = ""
         aa = Leerado8("SELECT * FROM ZZZ_APOR_PLA " _
                    & " WHERE CODIGO = " + Str(wCod) + " AND " _
                    & "          INS = " + Str(wIns) + " " _
                    & " ORDER BY TIPAPOR, CUOANO ")
         If aa > 0 Then
            ADO8.MoveFirst
            Do While Not ADO8.EOF
               If Not IsNull(ADO8!impo01) And ADO8!impo01 > 0 Then
                  wUlt = ADO8!cuoano + "-01"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo02) And ADO8!impo02 > 0 Then
                  wUlt = ADO8!cuoano + "-02"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo03) And ADO8!impo03 > 0 Then
                  wUlt = ADO8!cuoano + "-03"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo04) And ADO8!impo04 > 0 Then
                  wUlt = ADO8!cuoano + "-04"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo05) And ADO8!impo05 > 0 Then
                  wUlt = ADO8!cuoano + "-05"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo06) And ADO8!impo06 > 0 Then
                  wUlt = ADO8!cuoano + "-06"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo07) And ADO8!impo07 > 0 Then
                  wUlt = ADO8!cuoano + "-07"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo08) And ADO8!impo08 > 0 Then
                  wUlt = ADO8!cuoano + "-08"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo09) And ADO8!impo09 > 0 Then
                  wUlt = ADO8!cuoano + "-09"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo10) And ADO8!impo10 > 0 Then
                  wUlt = ADO8!cuoano + "-10"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo11) And ADO8!impo11 > 0 Then
                  wUlt = ADO8!cuoano + "-11"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo12) And ADO8!impo12 > 0 Then
                  wUlt = ADO8!cuoano + "-12"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
         
               ADO8.MoveNext
            Loop
         End If
         
         objExcel.Range(objExcel.Cells(V, H + 10), objExcel.Cells(V, H + 15)).NumberFormat = "#####,##0.00"
         
         objExcel.Cells(V, H + 0) = Format(wRegAct, "####0")
         objExcel.Cells(V, H + 1) = ADO3!codenvio
         objExcel.Cells(V, H + 2) = ADO3!codsocio
         objExcel.Cells(V, H + 3) = Trim(Format(ADO3!codigo, "#######0")) + "-" + Format(ADO3!ins, "9")
         objExcel.Cells(V, H + 4) = ADO3!carnetpnp
         objExcel.Cells(V, H + 5) = ADO3!numdoc
         objExcel.Cells(V, H + 6) = IIf(ADO3!lin = "0", "", "H")
         objExcel.Cells(V, H + 7) = ADO3!nombre
         objExcel.Cells(V, H + 8) = wUlt
         objExcel.Cells(V, H + 9) = wDon
         objExcel.Cells(V, H + 10) = ADO3!totaport
         objExcel.Cells(V, H + 11) = ADO3!totdeuda
         objExcel.Cells(V, H + 12) = ADO3!totadela
         objExcel.Cells(V, H + 13) = ADO3!netsocio
         objExcel.Cells(V, H + 14) = ADO3!dscsocio
         objExcel.Cells(V, H + 15) = ADO3!difsocio
         objExcel.Cells(V, H + 16) = wTel
         objExcel.Cells(V, H + 17) = wDir
         objExcel.Cells(V, H + 18) = wUbi
         objExcel.Cells(V, H + 19) = wRef
         objExcel.Cells(V, H + 20) = wMai
         
         wRegAct = wRegAct + 1
         wNetSocio = wNetSocio + ADO3!netsocio
         wDscSocio = wDscSocio + ADO3!dscsocio
         wDifSocio = wDifSocio + ADO3!difsocio
         V = V + 1
         
         wreg = wreg + 1
         ADO3.MoveNext
      Loop
      V = V + 1
      
      objExcel.Range(objExcel.Cells(V, H + 13), objExcel.Cells(V, H + 15)).NumberFormat = "#####,##0.00"
      objExcel.Range(objExcel.Cells(V, H + 7), objExcel.Cells(V, H + 15)).Font.Color = RGB(255, 0, 0)
      objExcel.Range(objExcel.Cells(V, H + 7), objExcel.Cells(V, H + 15)).Font.Bold = True
      objExcel.Range(objExcel.Cells(V, H + 7), objExcel.Cells(V, H + 15)).Borders.Color = RGB(255, 0, 0)
            
      objExcel.Cells(V, H + 7) = "TOTALES FINALES"
      objExcel.Cells(V, H + 13) = wNetSocio
      objExcel.Cells(V, H + 14) = wDscSocio
      objExcel.Cells(V, H + 15) = wDifSocio
      
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
   wMes = Left(cmbMeses.Text, 2)
   wAno = txtAnoCab.Text

   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\CajaMPxSinDscto.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= 'MES " + Trim(funnommes(wMes)) + " DEL " + wAno + "' "
   Crys1.SelectionFormula = " {TMP_CAJMPSOC.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1
End Sub

Private Sub cmdImpTes_Click()
   Dim wAno As String, wMes As String
   wMes = Left(cmbMeses.Text, 2)
   wAno = txtAnoCab.Text

   Call CalxPag
   
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\CajaMPxSinDsctoTeso.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= 'MES " + Trim(funnommes(wMes)) + " DEL " + wAno + "' "
   Crys1.SelectionFormula = " {TMP_CAJMPSOC.USU}='" + wcodusu + "' "
   Crys1.WindowState = crptMaximized
   Crys1.Action = 1

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_CAJMPSOC WHERE USU = '" + wcodusu + "' AND LIN > '5' ")
   Db.CommitTrans

End Sub

Private Sub CalxPag()
   Dim aa As Integer, wRegAct As Integer, wRegTot As Integer, _
       wCod As Long, wIns As Integer, _
       wAno As String, wMes As String, _
       wFec As Date, wSer As String, wDoc As String, wcon As String, wNom As String, _
       wGlo As String, wMon As String, wImp As Currency, wLinTes As Integer
 
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
 
   DoEvents
   lblMensaje.Caption = ""
   lblMensaje.Refresh
    
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_CAJMPSOC WHERE USU = '" + wcodusu + "' AND LIN > '5' ")
   Db.CommitTrans
 
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_CAJMPSOC " _
   & " SET FECPAG = NULL, SERPAG = '', NUMPAG = '', CONPAG = '', NOMPAG = '', " _
   & "     MONPAG = '', IMPPAG = 0, GLOPAG = '' " _
   & " WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
 
   aa = Leerado8("SELECT * FROM TMP_CAJMPSOC WHERE USU = '" + wcodusu + "' AND LIN <= '5' ")
   If aa > 0 Then
      wRegAct = 1
      wRegTot = aa
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         DoEvents
         lblMensaje.Caption = "Preparando Registro - " + _
                              Trim(Format(wRegAct, "#####0")) + " / " + _
                              Trim(Format(wRegTot, "#####0"))
         lblMensaje.Refresh
         wCod = ADO8!codigo
         wIns = ADO8!ins
         wLinTes = 6
   
         aa = Leerado7("SELECT * FROM ZZZ_MRECIBOS " _
                        & " WHERE CODIGO = " + Str(wCod) + " AND " _
                        & "          INS = " + Str(wIns) + " AND " _
                        & "        YEAR(FECHA_PAGO) = " + Str(Val(wAno)) + " AND " _
                        & "       MONTH(FECHA_PAGO) = " + Str(Val(wMes)) + " AND " _
                        & "       MONTO > 0 ")
         If aa > 0 Then
            wFec = ADO7!fecha_pago
            wcon = ADO7!concepto
            wSer = Format(ADO7!serie, "0000")
            wDoc = Format(ADO7!nro_comp, "000000000")
            wMon = IIf(ADO7!moneda = "S/.", "S", "D")
            wImp = ADO7!monto
            wGlo = ADO7!obs
            wNom = ""
                     
            aa = Leerado6("SELECT * FROM ZZZ_CONCEPTO WHERE CONCEPTO = '" + wcon + "' ")
            If aa > 0 Then
               wNom = ADO6!desconce
            End If
            Set ADO6 = Nothing
         
            If Len(Trim(ADO8!serpag)) = 0 Or IsNull(ADO8!serpag) Then
               Db.BeginTrans
               Db.Execute ("UPDATE TMP_CAJMPSOC " _
               & " SET FECPAG = '" + Format(wFec, "dd/mm/yyyy") + "', " _
               & "     SERPAG = '" + wSer + "', NUMPAG = '" + wDoc + "', " _
               & "     CONPAG = '" + wcon + "', NOMPAG = '" + wNom + "', " _
               & "     MONPAG = '" + wMon + "', GLOPAG = '" + wGlo + "', " _
               & "     IMPPAG = " + Str(wImp) + " " _
               & " WHERE    USU = '" + wcodusu + "' AND " _
               & "       CODIGO = " + Str(wCod) + " AND " _
               & "          INS = " + Str(wIns) + " ")
               Db.CommitTrans
            Else
               Db.BeginTrans
               Db.Execute ("INSERT INTO TMP_CAJMPSOC " _
               & " (USU, MES, CODENVIO, LIN, CODSOCIO, CODIGO, INS, CARNETPNP, NUMDOC, " _
               & "  CODBENI, NOMBRE, TOTAPORT, TOTDEUDA, TOTADELA, " _
               & "  NETSOCIO, DSCSOCIO, DIFSOCIO, TOTENVIO, DSCCAJMP, DSCDIFER, FECPAG, " _
               & "  SERPAG, NUMPAG, CONPAG, NOMPAG, MONPAG, IMPPAG, GLOPAG ) " _
               & " VALUES " _
               & " ('" + wcodusu + "', '" + ADO8!mes + "', " + Str(ADO8!codenvio) + ", " _
               & "  '" + Format(wLinTes, "0") + "', " + Str(ADO8!codsocio) + ", " _
               & "  " + Str(ADO8!codigo) + ", " + Str(ADO8!ins) + ", '" + ADO8!carnetpnp + "', " _
               & "  '" + ADO8!numdoc + "', '" + wcodbeni + "', '" + ADO8!nombre + "', " _
               & "  0, 0, 0, 0, 0, 0, 0, 0, 0, '" + Format(wFec, "dd/mm/yyyy") + "', " _
               & "  '" + wSer + "', '" + wDoc + "', '" + wcon + "', '" + wNom + "', " _
               & "  '" + wMon + "', " + Str(wImp) + ", '" + wGlo + "'  ) ")
               Db.CommitTrans
               wLinTes = wLinTes + 1
            End If
         End If
    
         wRegAct = wRegAct + 1
         ADO8.MoveNext
      Loop
   End If

   DoEvents
   lblMensaje.Caption = ""
   lblMensaje.Refresh
 
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(21) As String, _
       wRegAct As Integer, wRegTot As Integer, wAno As String, wMes As String
   Dim wAsi As Integer, wCod As Long, wIns As Integer, _
       wNom As String, wCip As Long, wDni As String, _
       wTotAport As Currency, wTotDeuda As Currency, wTotAdela As Currency, _
       wNetSocio As Currency, wDscSocio As Currency, wDifSocio As Currency, _
       wSolPag As Currency, wDolPag, wSoc As Integer, wFec As Date, _
       wCarnetPnp As String, wNumDoc As String, _
       wUlt As String, wDon As String
   
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
   
   Call CalxPag
   
   Heading(0) = "NUM"
   Heading(1) = "COD.ENVIO"
   Heading(2) = "COD.SOCIO"
   Heading(3) = "CODOFIN"
   Heading(4) = "CARNET PNP"
   Heading(5) = "DNI"
   Heading(6) = "H"
   Heading(7) = "NOMBRE"
   Heading(8) = "ULT.APORTE"
   Heading(9) = "DONDE"
   Heading(10) = "APORT.MES"
   Heading(11) = "DEUDAS"
   Heading(12) = "ADELANTO"
   Heading(13) = "TOT.ENVIO"
   Heading(14) = "TOT.DSCTO"
   Heading(15) = "NO COBRADO"
   Heading(16) = "FECHA.PAGO"
   Heading(17) = "DCMTO"
   Heading(18) = "CONCEPTO"
   Heading(19) = "MON"
   Heading(20) = "IMPORTE"
   Heading(21) = "CONCEPTO"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 22)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 22)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "DESCUENTO CAJA MILITAR POLICIAL - MES " + Trim(funnommes(wMes)) + " " + wanocia
        For I = 1 To 22 Step 1
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
        objExcel.Columns("O").ColumnWidth = 11
        objExcel.Columns("P").ColumnWidth = 11
        objExcel.Columns("Q").ColumnWidth = 11
        objExcel.Columns("R").ColumnWidth = 14
        objExcel.Columns("S").ColumnWidth = 22
        objExcel.Columns("T").ColumnWidth = 5
        objExcel.Columns("U").ColumnWidth = 11
        objExcel.Columns("V").ColumnWidth = 50
   End With
   
   aa = Leerado3("SELECT * FROM TMP_CAJMPSOC " _
                & " WHERE MES = '" + wanocia + wMes + "' AND USU = '" + wcodusu + "' " _
                & " ORDER BY NOMBRE, LIN ")
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
         
         wSoc = ADO3!codsocio
         wCarnetPnp = ""
         wNumDoc = ""
         aa = Leerado7("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
         If aa > 0 Then
            wCarnetPnp = IIf(IsNull(ADO7!carnetpnp2), "", ADO7!carnetpnp)
            wNumDoc = IIf(IsNull(ADO7!numdoc), "", ADO7!numdoc)
            wCod = ADO7!codigo
            wIns = ADO7!ins
         End If
         Set ADO7 = Nothing
         
         wUlt = "": wDon = ""
         aa = Leerado8("SELECT * FROM ZZZ_APOR_PLA " _
                    & " WHERE CODIGO = " + Str(wCod) + " AND " _
                    & "          INS = " + Str(wIns) + " " _
                    & " ORDER BY TIPAPOR, CUOANO ")
         If aa > 0 Then
            ADO8.MoveFirst
            Do While Not ADO8.EOF
               If Not IsNull(ADO8!impo01) And ADO8!impo01 > 0 Then
                  wUlt = ADO8!cuoano + "-01"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo02) And ADO8!impo02 > 0 Then
                  wUlt = ADO8!cuoano + "-02"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo03) And ADO8!impo03 > 0 Then
                  wUlt = ADO8!cuoano + "-03"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo04) And ADO8!impo04 > 0 Then
                  wUlt = ADO8!cuoano + "-04"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo05) And ADO8!impo05 > 0 Then
                  wUlt = ADO8!cuoano + "-05"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo06) And ADO8!impo06 > 0 Then
                  wUlt = ADO8!cuoano + "-06"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo07) And ADO8!impo07 > 0 Then
                  wUlt = ADO8!cuoano + "-07"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo08) And ADO8!impo08 > 0 Then
                  wUlt = ADO8!cuoano + "-08"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo09) And ADO8!impo09 > 0 Then
                  wUlt = ADO8!cuoano + "-09"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo10) And ADO8!impo10 > 0 Then
                  wUlt = ADO8!cuoano + "-10"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo11) And ADO8!impo11 > 0 Then
                  wUlt = ADO8!cuoano + "-11"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
               If Not IsNull(ADO8!impo12) And ADO8!impo12 > 0 Then
                  wUlt = ADO8!cuoano + "-12"
                  wDon = IIf(ADO8!tipapor = "1", "DIECO", "CAJAMP")
               End If
         
               ADO8.MoveNext
            Loop
         End If
         
         objExcel.Range(objExcel.Cells(V, H + 10), objExcel.Cells(V, H + 15)).NumberFormat = "#####,##0.00"
         objExcel.Range(objExcel.Cells(V, H + 20), objExcel.Cells(V, H + 20)).NumberFormat = "#####,##0.00"
         
         objExcel.Cells(V, H + 0) = Format(wRegAct, "####0")
         objExcel.Cells(V, H + 1) = ADO3!codenvio
         objExcel.Cells(V, H + 2) = ADO3!codsocio
         objExcel.Cells(V, H + 3) = Trim(Format(ADO3!codigo, "#######0")) + "-" + Format(ADO3!ins, "9")
         objExcel.Cells(V, H + 4) = wCarnetPnp
         objExcel.Cells(V, H + 5) = wNumDoc
         objExcel.Cells(V, H + 6) = IIf(ADO3!lin = "0", "", "H")
         objExcel.Cells(V, H + 7) = ADO3!nombre
         objExcel.Cells(V, H + 8) = wUlt
         objExcel.Cells(V, H + 9) = wDon
         objExcel.Cells(V, H + 10) = ADO3!totaport
         objExcel.Cells(V, H + 11) = ADO3!totdeuda
         objExcel.Cells(V, H + 12) = ADO3!totadela
         objExcel.Cells(V, H + 13) = ADO3!netsocio
         objExcel.Cells(V, H + 14) = ADO3!dscsocio
         objExcel.Cells(V, H + 15) = ADO3!difsocio
         If IsDate(ADO3!fecpag) Then
            wFec = Format(ADO3!fecpag)
            objExcel.Cells(V, H + 16) = wFec
            objExcel.Cells(V, H + 17) = ADO3!serpag + " " + ADO3!numpag
            objExcel.Cells(V, H + 18) = ADO3!nompag
            objExcel.Cells(V, H + 19) = IIf(ADO3!monpag = "S", "S/.", "US$")
            objExcel.Cells(V, H + 20) = ADO3!imppag
            objExcel.Cells(V, H + 21) = ADO3!glopag
         End If
         wRegAct = wRegAct + 1
         
         If ADO3!monpag = "S" Then
            wSolPag = wSolPag + ADO3!imppag
         Else
            wDolPag = wDolPag + ADO3!imppag
         End If
         wNetSocio = wNetSocio + ADO3!netsocio
         wDscSocio = wDscSocio + ADO3!dscsocio
         wDifSocio = wDifSocio + ADO3!difsocio
         V = V + 1
         
         wreg = wreg + 1
         ADO3.MoveNext
      Loop
      V = V + 1
      
      objExcel.Range(objExcel.Cells(V, H + 10), objExcel.Cells(V, H + 15)).NumberFormat = "#####,##0.00"
      objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 15)).Font.Color = RGB(255, 0, 0)
      objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 15)).Font.Bold = True
      objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 15)).Borders.Color = RGB(255, 0, 0)
            
      objExcel.Cells(V, H + 7) = "TOTALES FINALES"
      objExcel.Cells(V, H + 13) = wNetSocio
      objExcel.Cells(V, H + 14) = wDscSocio
      objExcel.Cells(V, H + 15) = wDifSocio
      objExcel.Cells(V, H + 19) = "S/."
      objExcel.Cells(V, H + 20) = wSolPag
      V = V + 1
      
      If wDolPag > 0 Then
         objExcel.Range(objExcel.Cells(V, H + 20), objExcel.Cells(V, H + 20)).NumberFormat = "#####,##0.00"
         
         objExcel.Cells(V, H + 19) = "US$"
         objExcel.Cells(V, H + 20) = wDolPag
         V = V + 1
      End If
      
   End If
   
   Set ADO3 = Nothing
   objExcel.Visible = True
   Set objExcel = Nothing
   
   lblMensaje.Caption = ""
   lblMensaje.Refresh
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub Form_Activate()
   Form_Initialize
End Sub

Private Sub Form_Initialize()
   frmMPConSinDscto.Left = (Screen.Width - Width) \ 2
   frmMPConSinDscto.Top = 0
   
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
   
   cmdImprimir.Enabled = False
   cmdExportar.Enabled = False
   Command1.Enabled = False
   cmdImpTes.Enabled = False
   
   cmbMeses.SetFocus
End Sub

Private Sub CalxSoc()
   Dim wAno As String, wMes As String

   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)

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
   & " WHERE C.MES = '" + wAno + wMes + "' AND " _
   & "       C.DIFSOCIO <> 0 ")
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
   & " WHERE C.MES = '" + wAno + wMes + "' AND " _
   & "       C.CODASIG1 <> 0 AND " _
   & "       C.DIFASIG1 <> 0  ")
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
   & " WHERE C.MES = '" + wAno + wMes + "' AND " _
   & "       C.CODASIG2 <> 0 AND " _
   & "       C.DIFASIG2 <> 0   ")
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
   & " WHERE C.MES = '" + wAno + wMes + "' AND " _
   & "       C.CODASIG3 <> 0 AND " _
   & "       C.DIFASIG3 <> 0   ")
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
   & " WHERE C.MES = '" + wAno + wMes + "' AND " _
   & "       C.CODASIG4 <> 0 AND " _
   & "       C.DIFASIG4 <> 0   ")
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
   & " WHERE C.MES = '" + wAno + wMes + "' AND " _
   & "       C.CODASIG5 <> 0 AND " _
   & "       C.DIFASIG5 <> 0   ")
   Db.CommitTrans
End Sub
