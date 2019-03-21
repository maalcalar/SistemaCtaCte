VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmDIECOConxMes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta DIECO Por Mes"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   13350
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
      TabIndex        =   29
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
      Height          =   615
      Left            =   9240
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7080
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
      Height          =   615
      Left            =   10560
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7080
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
      Height          =   615
      Left            =   11880
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1095
   End
   Begin VB.ComboBox cmbMeses 
      Height          =   315
      ItemData        =   "frmDiecoConxMes.frx":0000
      Left            =   720
      List            =   "frmDiecoConxMes.frx":0002
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
      Height          =   5055
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   8916
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
      TabIndex        =   30
      Top             =   7320
      Width           =   7575
   End
   Begin VB.Label lblDscDifer 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10560
      TabIndex        =   25
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Total No Cobrado"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8880
      TabIndex        =   24
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblDscDIECO 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10560
      TabIndex        =   23
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Cobrado"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8880
      TabIndex        =   22
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblCanApo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6720
      TabIndex        =   21
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblTotEnvio 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10560
      TabIndex        =   20
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblCanAsi 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6720
      TabIndex        =   19
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Titulares"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   18
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Asignados"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   17
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Enviado"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8880
      TabIndex        =   16
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
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
      Left            =   240
      TabIndex        =   15
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label lblAsig1 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   6360
      Width           =   5175
   End
   Begin VB.Label lblAsig2 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   6600
      Width           =   5175
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
      Left            =   240
      TabIndex        =   12
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label lblAsig3 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   6840
      Width           =   5175
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
      Left            =   240
      TabIndex        =   10
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label lblAsig4 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7560
      TabIndex        =   9
      Top             =   6360
      Width           =   5175
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
      Left            =   6840
      TabIndex        =   8
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label lblAsig5 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      Top             =   6600
      Width           =   5175
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
      Left            =   6840
      TabIndex        =   6
      Top             =   6600
      Width           =   735
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
      Caption         =   "Consulta DIECO Por Mes"
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
Attribute VB_Name = "frmDIECOConxMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub cmbMeses_Click()
'   cmbMeses_KeyPress (13)
'End Sub

Private Sub cmbMeses_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      
      Set DataGrid1.DataSource = Nothing
      
      cmdBuscar.SetFocus
   End If
End Sub

Private Sub cmdBuscar_Click()
   Dim wAno As String, wMes As String, zz As Integer
   wAno = txtAnoCab.Text
   wMes = Left(cmbMeses.Text, 2)
      
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_DIECOCAB WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECOCAB " _
   & " (MES, CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, FECENV, FECDSC, " _
   & "   TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "   DIFASIG1, DIFASIG2, DIFASIG3, DIFASIG4, DIFASIG5, " _
   & "   DSCASIG1, DSCASIG2, DSCASIG3, DSCASIG4, DSCASIG5, " _
   & "   NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, " _
   & "   ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
   & "   DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
   & "   TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
   & "   CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
   & "   NOMASIG1, NOMASIG2, NOMASIG3, NOMASIG4, NOMASIG5, " _
   & "   TOTENVIO, DSCDIECO, DSCDIFER, TIPCOB  , USU ) " _
   & " SELECT " _
   & "  MES, CODSOCIO, CODIGO, INS, '', E_SOCIO, FECENV, FECDSC, " _
   & "   TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "   DIFASIG1, DIFASIG2, DIFASIG3, DIFASIG4, DIFASIG5, " _
   & "   DSCASIG1, DSCASIG2, DSCASIG3, DSCASIG4, DSCASIG5, " _
   & "   NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, " _
   & "   ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
   & "   DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
   & "   TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
   & "   CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
   & "   '', '', '', '', '', TOTENVIO, DSCDIECO, DSCDIFER, TIPCOB  , '" + wcodusu + "'  " _
   & " FROM DIECOCAB " _
   & " WHERE MES = '" + wAno + wMes + "'  ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMBRE = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODSOCIO = M.CODSOCIO " _
   & " WHERE T.MES = '" + wAno + wMes + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG1 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG1 = M.CODSOCIO " _
   & " WHERE T.MES = '" + wAno + wMes + "' AND " _
   & "       T.CODASIG1 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG2 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG2 = M.CODSOCIO " _
   & " WHERE T.MES = '" + wAno + wMes + "' AND " _
   & "       T.CODASIG2 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG3 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG3 = M.CODSOCIO " _
   & " WHERE T.MES = '" + wAno + wMes + "' AND " _
   & "       T.CODASIG3 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG4 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG4 = M.CODSOCIO " _
   & " WHERE T.MES = '" + wAno + wMes + "' AND " _
   & "       T.CODASIG4 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG5 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG5 = M.CODSOCIO " _
   & " WHERE T.MES = '" + wAno + wMes + "' AND " _
   & "       T.CODASIG5 <> 0 ")
   Db.CommitTrans

   zz = Leerado2("SELECT CODSOCIO, NOMBRE  , TOTENVIO, DSCDIECO, DSCDIFER, " _
                & "      DSCSOCIO, DSCASIG1, DSCASIG2, DSCASIG3, DSCASIG4, DSCASIG5, " _
                & "      NETSOCIO, NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, " _
                & "      CODASIG1, NOMASIG1, TOTASIG1, DEUASIG1, ADEASIG1, DIFASIG1, " _
                & "      CODASIG2, NOMASIG2, TOTASIG2, DEUASIG2, ADEASIG2, DIFASIG2, " _
                & "      CODASIG3, NOMASIG3, TOTASIG3, DEUASIG3, ADEASIG3, DIFASIG3, " _
                & "      CODASIG4, NOMASIG4, TOTASIG4, DEUASIG4, ADEASIG4, DIFASIG4, " _
                & "      CODASIG5, NOMASIG5, TOTASIG5, DEUASIG5, ADEASIG5, DIFASIG5, " _
                & "      TOTAPORT, TOTDEUDA, TOTADELA, DIFSOCIO, " _
                & "      TIPCOB  , USU     , MES     , CODIGO  , INS     , " _
                & "      FECENV  , FECDSC  " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE MES = '" + wAno + wMes + "' AND USU = '" + wcodusu + "' " _
                & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO2

   DataGrid1.Columns(0).Width = 750   ' CODSOCIO
   DataGrid1.Columns(0).Alignment = dbgRight
   DataGrid1.Columns(0).Caption = "SOCIO"
    
   DataGrid1.Columns(1).Width = 4500  ' NOMBRE
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Caption = "NOMBRE ASOCIADO"
    
   DataGrid1.Columns(2).Width = 800    ' TOTENVIO
   DataGrid1.Columns(2).Alignment = dbgRight
   DataGrid1.Columns(2).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(2).Caption = "TOT.ENVIO"
    
   DataGrid1.Columns(3).Width = 800    ' DSCCAJMP
   DataGrid1.Columns(3).Alignment = dbgRight
   DataGrid1.Columns(3).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(3).Caption = "TOT.DSCTO"
    
   DataGrid1.Columns(4).Width = 800    ' DSCDIFER
   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(4).Caption = "NO COBRADO"
    
   DataGrid1.Columns(5).Width = 750    ' DSCSOCIO
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(5).Caption = "DSC.SOC"
    
   DataGrid1.Columns(6).Width = 750    ' DSCASIG1
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(6).Caption = "DSC.AS1"
    
   DataGrid1.Columns(7).Width = 750    ' DSCASIG2
   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(7).Caption = "DSC.AS2"
    
   DataGrid1.Columns(8).Width = 750    ' DSCASIG3
   DataGrid1.Columns(8).Alignment = dbgRight
   DataGrid1.Columns(8).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(8).Caption = "DSC.AS3"
    
   DataGrid1.Columns(9).Width = 750    ' DSCASIG4
   DataGrid1.Columns(9).Alignment = dbgRight
   DataGrid1.Columns(9).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(9).Caption = "DSC.AS4"
    
   DataGrid1.Columns(10).Width = 750    ' DSCASIG5
   DataGrid1.Columns(10).Alignment = dbgRight
   DataGrid1.Columns(10).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(10).Caption = "DSC.AS5"
    
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

   TotalCab
   DataGrid1.SetFocus
End Sub

Private Sub TotalCab()
   Dim zz As Integer, _
       zAno As String, zMes As String, _
       zTotEnvio As Currency, _
       zDscDieco As Currency, _
       zDscDifer As Currency, _
       zCanApo As Integer, zCanAsi As Integer
   
   zAno = txtAnoCab.Text
   zMes = Left(cmbMeses.Text, 2)
   
   zTotEnvio = 0: zDscDieco = 0: zDscDifer = 0
   zCanApo = 0: zCanAsi = 0
   zz = Leerado8("SELECT SUM(TOTENVIO) AS TOTENVIO, " _
                & "      SUM(DSCDIECO) AS DSCDIECO, " _
                & "      SUM(DSCDIFER) AS DSCDIFER, " _
                & "      COUNT(*) AS CAN " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "       MES = '" + zAno + zMes + "' ")
   If zz > 0 Then
      zTotEnvio = IIf(IsNull(ADO8!totenvio), 0, ADO8!totenvio)
      zDscDieco = IIf(IsNull(ADO8!dscdieco), 0, ADO8!dscdieco)
      zDscDifer = IIf(IsNull(ADO8!dscdifer), 0, ADO8!dscdifer)
      zCanApo = IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT COUNT(*) AS CAN " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE (USU = '" + wcodusu + "') AND " _
                & "       (MES = '" + zAno + zMes + "') AND " _
                & "       (CODASIG1 <> 0) ")
   If zz > 0 Then
      zCanAsi = zCanAsi + IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT COUNT(*) AS CAN " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE (USU = '" + wcodusu + "') AND " _
                & "       (MES = '" + zAno + zMes + "') AND " _
                & "       (CODASIG2 <> 0) ")
   If zz > 0 Then
      zCanAsi = zCanAsi + IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT COUNT(*) AS CAN " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE (USU = '" + wcodusu + "') AND " _
                & "       (MES = '" + zAno + zMes + "') AND " _
                & "       (CODASIG3 <> 0) ")
   If zz > 0 Then
      zCanAsi = zCanAsi + IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT COUNT(*) AS CAN " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE (USU = '" + wcodusu + "') AND " _
                & "       (MES = '" + zAno + zMes + "') AND " _
                & "       (CODASIG4 <> 0) ")
   If zz > 0 Then
      zCanAsi = zCanAsi + IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT COUNT(*) AS CAN " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE (USU = '" + wcodusu + "') AND " _
                & "       (MES = '" + zAno + zMes + "') AND " _
                & "       (CODASIG5 <> 0) ")
   If zz > 0 Then
      zCanAsi = zCanAsi + IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   lblTotEnvio.Caption = Format(zTotEnvio, "####,##0.00;;\ ")
   lblDscDIECO.Caption = Format(zDscDieco, "####,##0.00;;\ ")
   lblDscDifer.Caption = Format(zDscDifer, "####,##0.00;;\ ")
   lblCanApo.Caption = Format(zCanApo, "##,##0")
   lblCanAsi.Caption = Format(zCanAsi, "##,##0")
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(12) As String, _
       wRegAct As Integer, wRegTot As Integer
   Dim wAsi As Integer, wCod As Long, wIns As Integer, _
       wNom As String, wCip As Long, wDni As String, wMes As String, _
       wTotAport As Currency, wTotDeuda As Currency, wTotAdela As Currency, _
       wNetSocio As Currency, wDscSocio As Currency, wDifSocio As Currency
   
   wMes = Left(cmbMeses.Text, 2)
      
   Call CalxSoc
   
   Heading(0) = "NUM"
   Heading(1) = "COD.ENVIO"
   Heading(2) = "COD.SOCIO"
   Heading(3) = "CODOFIN"
   Heading(4) = "H"
   Heading(5) = "NOMBRE"
   Heading(6) = "E_SOCIO"
   Heading(7) = "APORT.MES"
   Heading(8) = "DEUDAS"
   Heading(9) = "ADELANTO"
   Heading(10) = "TOT.ENVIO"
   Heading(11) = "TOT.DSCTO"
   Heading(12) = "NO COBRADO"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 13)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 13)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "DESCUENTO DIECO - MES " + Trim(funnommes(wMes)) + " " + wanocia
        For I = 1 To 13 Step 1
            .Cells(3, I) = Heading(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 6
        objExcel.Columns("B").ColumnWidth = 7
        objExcel.Columns("C").ColumnWidth = 7
        objExcel.Columns("D").ColumnWidth = 10
        objExcel.Columns("E").ColumnWidth = 5
        objExcel.Columns("F").ColumnWidth = 50
        objExcel.Columns("G").ColumnWidth = 7
        
        objExcel.Columns("H").ColumnWidth = 11
        objExcel.Columns("I").ColumnWidth = 11
        objExcel.Columns("J").ColumnWidth = 11
        objExcel.Columns("K").ColumnWidth = 11
        objExcel.Columns("L").ColumnWidth = 11
        objExcel.Columns("M").ColumnWidth = 11
   End With
   
   
   aa = Leerado3("SELECT * FROM TMP_DIECOSOC " _
                & " WHERE MES = '" + wanocia + wMes + "' AND " _
                & "       USU = '" + wcodusu + "' " _
                & " ORDER BY NOMENVIO, LIN ")
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
         
         objExcel.Range(objExcel.Cells(V, H + 7), objExcel.Cells(V, H + 12)).NumberFormat = "#####,##0.00"
         
         objExcel.Cells(V, H + 0) = Format(wRegAct, "####0")
         objExcel.Cells(V, H + 1) = ADO3!codenvio
         objExcel.Cells(V, H + 2) = ADO3!codsocio
         objExcel.Cells(V, H + 3) = Trim(Format(ADO3!codigo, "#######0")) + "-" + Format(ADO3!ins, "9")
         objExcel.Cells(V, H + 4) = IIf(ADO3!lin = "0", "", "H")
         objExcel.Cells(V, H + 5) = ADO3!nombre
         objExcel.Cells(V, H + 6) = ADO3!e_socio
         objExcel.Cells(V, H + 7) = ADO3!totaport
         objExcel.Cells(V, H + 8) = ADO3!totdeuda
         objExcel.Cells(V, H + 9) = ADO3!totadela
         objExcel.Cells(V, H + 10) = ADO3!netsocio
         objExcel.Cells(V, H + 11) = ADO3!dscsocio
         objExcel.Cells(V, H + 12) = ADO3!difsocio
         
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
      
      objExcel.Range(objExcel.Cells(V, H + 7), objExcel.Cells(V, H + 12)).NumberFormat = "#####,##0.00"
      objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 12)).Font.Color = RGB(255, 0, 0)
      objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 12)).Font.Bold = True
      objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 12)).Borders.Color = RGB(255, 0, 0)
            
      objExcel.Cells(V, H + 5) = "TOTALES FINALES"
      objExcel.Cells(V, H + 7) = wTotAport
      objExcel.Cells(V, H + 8) = wTotDeuda
      objExcel.Cells(V, H + 9) = wTotAdela
      objExcel.Cells(V, H + 10) = wNetSocio
      objExcel.Cells(V, H + 11) = wDscSocio
      objExcel.Cells(V, H + 12) = wDifSocio
      
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

   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)

   Call CalxSoc

   Crys1.Connect = "dsn=" + xodbc + "; uid=" + xUser + "; pwd=" + xPwd + ";dsq=" + xodbc + ""
   Crys1.ReportFileName = xraiz + "ReportCtaCte\DiecoxMes.RPT"
   Crys1.Formulas(0) = "NOMBRECIA= '" + wnomcia + "' "
   Crys1.Formulas(1) = "RUCCIA= 'RUC " + wruccia + "' "
   Crys1.Formulas(2) = "NOMMES= 'MES " + Trim(funnommes(wMes)) + " DEL " + wAno + "' "
   Crys1.SelectionFormula = " {TMP_DIECOSOC.USU}='" + wcodusu + "' "
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
        ADO2.Sort = "NOMBRE"
   Case 2
        ADO2.Sort = "TOTENVIO DESC"
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
   frmDIECOConxMes.Left = (Screen.Width - Width) \ 2
   frmDIECOConxMes.Top = 0
   
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
      cmbMeses.AddItem ADO1!MES + " " + Trim(funnommes(ADO1!MES))
       ADO1.MoveNext
   Loop
   
   txtAnoCab.SetFocus
End Sub

Private Sub CalxSoc()
   Dim wAno As String, wMes As String

   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_DIECOSOC WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECOSOC " _
   & " (MES, CODENVIO, LIN, CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, " _
   & "  FECENV, FECDSC, TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "  TOTENVIO, DSCDIECO, DSCDIFER, USU ) " _
   & " SELECT " _
   & "  C.MES, C.CODSOCIO, '0', C.CODSOCIO, C.CODIGO, C.INS, " _
   & "  M.NOMBRE, M.E_SOCIO, C.FECENV, C.FECDSC, C.TOTAPORT, C.TOTDEUDA, " _
   & "  C.TOTADELA, C.NETSOCIO, C.DSCSOCIO, C.DIFSOCIO, C.TOTENVIO, C.DSCDIECO, " _
   & "  C.DSCDIFER, '" + wcodusu + "' " _
   & " FROM DIECOCAB AS C LEFT JOIN MAESOCIO AS M " _
   & "   ON C.CODSOCIO = M.CODSOCIO " _
   & " WHERE C.MES = '" + wAno + wMes + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECOSOC " _
   & " (MES, CODENVIO, LIN, CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, " _
   & "  FECENV, FECDSC, TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "  TOTENVIO, DSCDIECO, DSCDIFER, USU ) " _
   & " SELECT " _
   & "  C.MES, C.CODSOCIO, '1', C.CODASIG1, M.CODIGO, M.INS, " _
   & "  M.NOMBRE, M.E_SOCIO, C.FECENV, C.FECDSC, C.TOTASIG1, C.DEUASIG1, " _
   & "  C.ADEASIG1, C.NETASIG1, C.DSCASIG1, C.DIFASIG1, 0, 0, 0, '" + wcodusu + "' " _
   & " FROM DIECOCAB AS C INNER JOIN MAESOCIO AS M " _
   & "   ON C.CODASIG1 = M.CODSOCIO " _
   & " WHERE C.MES = '" + wAno + wMes + "' AND " _
   & "       C.CODASIG1 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECOSOC " _
   & " (MES, CODENVIO, LIN, CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, " _
   & "  FECENV, FECDSC, TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "  TOTENVIO, DSCDIECO, DSCDIFER, USU ) " _
   & " SELECT " _
   & "  C.MES, C.CODSOCIO, '2', C.CODASIG2, M.CODIGO, M.INS, " _
   & "  M.NOMBRE, M.E_SOCIO, C.FECENV, C.FECDSC, C.TOTASIG2, C.DEUASIG2, " _
   & "  C.ADEASIG2, C.NETASIG2, C.DSCASIG2, C.DIFASIG2, 0, 0, 0, '" + wcodusu + "' " _
   & " FROM DIECOCAB AS C INNER JOIN MAESOCIO AS M " _
   & "   ON C.CODASIG2 = M.CODSOCIO " _
   & " WHERE C.MES = '" + wAno + wMes + "' AND " _
   & "       C.CODASIG2 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECOSOC " _
   & " (MES, CODENVIO, LIN, CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, " _
   & "  FECENV, FECDSC, TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "  TOTENVIO, DSCDIECO, DSCDIFER, USU ) " _
   & " SELECT " _
   & "  C.MES, C.CODSOCIO, '3', C.CODASIG3, M.CODIGO, M.INS, " _
   & "  M.NOMBRE, M.E_SOCIO, C.FECENV, C.FECDSC, C.TOTASIG3, C.DEUASIG3, " _
   & "  C.ADEASIG3, C.NETASIG3, C.DSCASIG3, C.DIFASIG3, 0, 0, 0, '" + wcodusu + "' " _
   & " FROM DIECOCAB AS C INNER JOIN MAESOCIO AS M " _
   & "   ON C.CODASIG3 = M.CODSOCIO " _
   & " WHERE C.MES = '" + wAno + wMes + "' AND " _
   & "       C.CODASIG3 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECOSOC " _
   & " (MES, CODENVIO, LIN, CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, " _
   & "  FECENV, FECDSC, TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "  TOTENVIO, DSCDIECO, DSCDIFER, USU ) " _
   & " SELECT " _
   & "  C.MES, C.CODSOCIO, '4', C.CODASIG4, M.CODIGO, M.INS, " _
   & "  M.NOMBRE, M.E_SOCIO, C.FECENV, C.FECDSC, C.TOTASIG4, C.DEUASIG4, " _
   & "  C.ADEASIG4, C.NETASIG4, C.DSCASIG4, C.DIFASIG4, 0, 0, 0, '" + wcodusu + "' " _
   & " FROM DIECOCAB AS C INNER JOIN MAESOCIO AS M " _
   & "   ON C.CODASIG4 = M.CODSOCIO " _
   & " WHERE C.MES = '" + wAno + wMes + "' AND " _
   & "       C.CODASIG4 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECOSOC " _
   & " (MES, CODENVIO, LIN, CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, " _
   & "  FECENV, FECDSC, TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "  TOTENVIO, DSCDIECO, DSCDIFER, USU ) " _
   & " SELECT " _
   & "  C.MES, C.CODSOCIO, '5', C.CODASIG5, M.CODIGO, M.INS, " _
   & "  M.NOMBRE, M.E_SOCIO, C.FECENV, C.FECDSC, C.TOTASIG5, C.DEUASIG5, " _
   & "  C.ADEASIG5, C.NETASIG5, C.DSCASIG5, C.DIFASIG5, 0, 0, 0, '" + wcodusu + "' " _
   & " FROM DIECOCAB AS C INNER JOIN MAESOCIO AS M " _
   & "   ON C.CODASIG5 = M.CODSOCIO " _
   & " WHERE C.MES = '" + wAno + wMes + "' AND " _
   & "       C.CODASIG5 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOSOC " _
   & " SET NOMENVIO = M.NOMBRE " _
   & " FROM TMP_DIECOSOC AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODENVIO = M.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans

End Sub

Private Sub txtAnoCab_GotFocus()
   txtAnoCab.SelStart = 0
   txtAnoCab.SelLength = 4
End Sub

Private Sub txtAnoCab_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txtAnoCab.Text)) = 0 Then
         MsgBox "Año En Blanco", vbExclamation
         txtAnoCab.Text = wanocia
         Exit Sub
      End If
      If txtAnoCab.Text < "2015" Or txtAnoCab.Text > "2020" Then
         MsgBox "Año Digitado Fuera de Rango", vbExclamation
         txtAnoCab.Text = wanocia
         Exit Sub
      End If
      
      Set DataGrid1.DataSource = Nothing
      
      cmbMeses.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub
