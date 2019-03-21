VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmConSocioVIP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relación de Socios VIP"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   12585
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
      Left            =   11040
      TabIndex        =   6
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar "
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
      Left            =   8400
      TabIndex        =   5
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir "
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
      Left            =   9720
      TabIndex        =   4
      Top             =   6600
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
      Left            =   10200
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmConSocioVIP.frx":0000
      Left            =   1440
      List            =   "frmConSocioVIP.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   165
      Width           =   7335
   End
   Begin Crystal.CrystalReport Crys1 
      Left            =   9240
      Top             =   120
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   12255
      _ExtentX        =   21616
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
      Caption         =   "RELACION DE SOCIOS VIP"
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
      TabIndex        =   7
      Top             =   6720
      Width           =   7575
   End
   Begin VB.Label Label38 
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
      Left            =   240
      TabIndex        =   1
      Top             =   165
      Width           =   855
   End
End
Attribute VB_Name = "frmConSocioVIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
   LlenaCab
   cmdImprimir.Enabled = True
   cmdExportar.Enabled = True

   DataGrid1.SetFocus
End Sub

Private Sub cmdImprimir_Click()
   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\SociosVIP.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.SelectionFormula = " {TMP_SOCVIP.USU}='" + wcodusu + "' "
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
   Case 3
        ADO2.Sort = "NOMBRE"
   Case 5
        ADO2.Sort = "NOMGRA"
   End Select
End Sub

Private Sub Form_Activate()
   Dim a As Integer, wAno As String, wMes As String
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
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_SOCVIP WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   cmdImprimir.Enabled = False
   cmdExportar.Enabled = False
   
   cmdBuscar.SetFocus
End Sub

Private Sub Form_Load()
   frmConSocioVIP.Left = (Screen.Width - Width) \ 2
   frmConSocioVIP.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_SOCVIP WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim aa As Integer
 
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_SOCVIP WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_SOCVIP " _
   & " (CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, NOME_S, GRADO, NOMGRA, FECING, USU) " _
   & " SELECT " _
   & "  M.CODSOCIO, M.CODIGO, M.INS, M.NOMBRE, M.E_SOCIO, E.NOMBRE, M.GRADO, " _
   & "  G.NOMBRE, M.FECING, '" + wcodusu + "' " _
   & " FROM MAESOCIO AS M INNER JOIN MAEE_SOCIO AS E ON M.E_SOCIO = E.E_SOCIO " _
   & "                    INNER JOIN MAEGRADO   AS G ON M.GRADO = G.GRADO " _
   & " WHERE VIP = 1 ")
   Db.CommitTrans

   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, NOME_S, NOMGRA, FECING, " _
                & "      GRADO, E_SOCIO, USU " _
                & " FROM TMP_SOCVIP " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO2

   DataGrid1.Columns(0).Width = 700   ' CODSOCIO
   DataGrid1.Columns(0).Alignment = dbgLeft
   DataGrid1.Columns(0).Caption = "COD.SOCIO"
    
   DataGrid1.Columns(1).Width = 1000  ' CODIGO
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Caption = "CODIGO"
    
   DataGrid1.Columns(2).Width = 400   ' INS
   DataGrid1.Columns(2).Alignment = dbgLeft
   DataGrid1.Columns(2).Caption = "INS"
    
   DataGrid1.Columns(3).Width = 4500  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE"
    
   DataGrid1.Columns(4).Width = 1300  ' E_SOCIO
   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Caption = "ESTADO"
    
   DataGrid1.Columns(5).Width = 1800 ' NOMGRA
   DataGrid1.Columns(5).Alignment = dbgLeft
   DataGrid1.Columns(5).Caption = "GRADO"
    
   DataGrid1.Columns(6).Width = 1100 ' FECING
   DataGrid1.Columns(6).Alignment = dbgCenter
   DataGrid1.Columns(6).Caption = "FEC.ING"
    
   DataGrid1.Columns(7).Visible = False
   DataGrid1.Columns(8).Visible = False
   DataGrid1.Columns(9).Visible = False
End Sub
