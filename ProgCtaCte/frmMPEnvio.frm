VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMPEnvio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envio Caja Militar Policial"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   13695
   Begin VB.CommandButton cmdRecupera 
      Caption         =   "&Recupera"
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
      Left            =   13800
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtro"
      Height          =   1095
      Left            =   120
      TabIndex        =   35
      Top             =   6960
      Width           =   6495
      Begin VB.TextBox txtFiltrar 
         Height          =   285
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   38
         Top             =   600
         Width           =   4095
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Filtrar x Nombre"
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Mostrar Todos"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Modificar Envio"
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
      Left            =   12360
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCrear 
      Caption         =   "Crear Nuevo Mes"
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar C�lculo"
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
      Left            =   6480
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   600
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
      Left            =   10440
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7320
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
      Left            =   8160
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7320
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
      Left            =   5160
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   600
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
      Left            =   6840
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7320
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
      Left            =   11760
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5055
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   11895
      _ExtentX        =   20981
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
   Begin VB.TextBox txtAnoCab 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.ComboBox cmbMeses 
      Height          =   315
      ItemData        =   "frmMPEnvio.frx":0000
      Left            =   960
      List            =   "frmMPEnvio.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblCanApo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10920
      TabIndex        =   33
      Top             =   315
      Width           =   1095
   End
   Begin VB.Label lblEnviado 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   12240
      TabIndex        =   32
      Top             =   315
      Width           =   1095
   End
   Begin VB.Label lblCanAsi 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10920
      TabIndex        =   31
      Top             =   825
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Cant.Titulares"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10920
      TabIndex        =   30
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Cant.Asignados"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10920
      TabIndex        =   29
      Top             =   615
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Total Envio S/."
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   12240
      TabIndex        =   28
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblNoDscto 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   12240
      TabIndex        =   27
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "No Cobrado"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   12240
      TabIndex        =   26
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblRecibido 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   12240
      TabIndex        =   25
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Cobrado"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   12240
      TabIndex        =   24
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ENVIO A CAJA MILITAR POLICIAL"
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
      Left            =   3480
      TabIndex        =   22
      Top             =   0
      Width           =   5775
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
      Left            =   6000
      TabIndex        =   21
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblAsig5 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6720
      TabIndex        =   20
      Top             =   6480
      Width           =   4695
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
      Left            =   6000
      TabIndex        =   19
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label lblAsig4 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6720
      TabIndex        =   18
      Top             =   6240
      Width           =   4695
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
      TabIndex        =   17
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label lblAsig3 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      TabIndex        =   16
      Top             =   6720
      Width           =   4815
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
      TabIndex        =   15
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblAsig2 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   6480
      Width           =   4815
   End
   Begin VB.Label lblAsig1 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   6240
      Width           =   4815
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
      TabIndex        =   12
      Top             =   6240
      Width           =   735
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
      Left            =   6960
      TabIndex        =   10
      Top             =   6840
      Width           =   5655
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
      TabIndex        =   3
      Top             =   360
      Width           =   495
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
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "frmMPEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wMax As Integer, wGrabar As Boolean, wEstado As Boolean

Private Sub cmbMeses_Click()
   cmbMeses_KeyPress (13)
End Sub

Private Sub cmbMeses_KeyPress(KeyAscii As Integer)
   Dim zz As Integer, wAno As String, wMes As String
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
   
   If KeyAscii = 13 Then
      Set DataGrid1.DataSource = Nothing
      
      wEstado = False
      
      zz = Leerado8("SELECT * FROM CONTROL_ENVIO " _
                    & " WHERE MES = '" + wanocia + wMes + "' AND TIPO = '02' ")
      If zz > 0 Then
         wEstado = IIf(ADO8!estado = "C", False, True)
      End If
      wEstado = True
      
      Set ADO8 = Nothing
      zz = Leerado2("SELECT * FROM CAJMPCAB " _
                & "  WHERE MES = '" + wAno + wMes + "' ")
      If zz > 0 Then
      
         lblMensaje.Caption = "Trae Calculo CAJA Militar Policial - Mes " + Left(Trim(funnommes(wMes)), 3) + " " + wAno
         lblMensaje.Refresh
      
         LlenaCab
         LlenaCab1
         TotalCab
         ADO2.MoveFirst
         LabelCab
      
         lblMensaje.Caption = ""
         lblMensaje.Refresh
      
      Else
         If cmdCalculo.Enabled = False Then
            cmdCalculo.Enabled = True
         End If
         cmdCalculo.SetFocus
      End If
   End If
End Sub

Private Sub cmdCalculo_Click()
   Dim zz As Integer, wRegAct As Integer, wRegTot As Integer, _
       wAno As String, wMes As String, wNom As String, _
       wFecEnv As Date, wFecDsc As Date, _
       wSoc As Integer, wCod As Long, wIns As Integer, wCar As Long, wDoc As String, wBen As String, wApo As Currency, wMon As String, _
       wCodAsig1 As Integer, wCodAsig2 As Integer, wCodAsig3 As Integer, wCodAsig4 As Integer, wCodAsig5 As Integer, _
       wNomAsig1 As String, wNomAsig2 As String, wNomAsig3 As String, wNomAsig4 As String, wNomAsig5 As String, _
       wTotAsig1 As Currency, wTotAsig2 As Currency, wTotAsig3 As Currency, wTotAsig4 As Currency, wTotAsig5 As Currency, _
       wDeuAsig1 As Currency, wDeuAsig2 As Currency, wDeuAsig3 As Currency, wDeuAsig4 As Currency, wDeuAsig5 As Currency, _
       wAdeAsig1 As Currency, wAdeAsig2 As Currency, wAdeAsig3 As Currency, wAdeAsig4 As Currency, wAdeAsig5 As Currency, _
       wNetAsig1 As Currency, wNetAsig2 As Currency, wNetAsig3 As Currency, wNetAsig4 As Currency, wNetAsig5 As Currency, _
       wTotAport As Currency, wTotEnvio As Currency, wNetSocio As Currency, _
       wTotDeuda As Currency, wTotAdela As Currency, wLin As Integer, wMesOld As String, wCarF As String
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
   
   If wMes > "01" Then
      wMesOld = wAno + Format(Val(wMes) - 1, "00")
   Else
      wMesOld = Format(Val(wAno) - 1, "00") + "12"
   End If
   
   wFecEnv = "01/" + wMes + "/" + wAno
   
   zz = Leerado8("SELECT * FROM CAJMPCAB " _
             & "  WHERE MES = '" + wAno + wMes + "' ")
   If zz > 0 Then
      If MsgBox("Ya Existe Proceso CAJA MILITAR POLICIAL del Mes" + vbNewLine + _
                "Desea Volver a Crearlo???", vbYesNo + vbQuestion, "Crear Archivo Descuento CAJA MILITAR POLICIAL") = vbNo Then
         Exit Sub
      End If
   End If
   Set ADO8 = Nothing
   
   Set DataGrid1.DataSource = Nothing
   
   lblMensaje.Caption = "Calculando Descuentos CAJA MILITAR POLICIAL - Mes " + Trim(funnommes(wMes)) + " " + wAno
   lblMensaje.Refresh
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_CAJMPCAB WHERE MES = '" + wAno + wMes + "' AND USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM CAJMPCAB WHERE MES = '" + wAno + wMes + "' ")
   Db.CommitTrans
   
   zz = Leerado8("SELECT M.CODSOCIO, M.CODIGO, M.INS, M.CARNETPNP, M.NUMDOC, M.NOMBRE, M.TIPCOB, " _
             & "         E.APORTE, E.MONEDA, M.ADELANTO, M.DEUDA_PT2, M.E_SOCIO, M.CARNETPNP2 " _
             & " FROM MAESOCIO AS M INNER JOIN MAEE_SOCIO AS E " _
             & "   ON M.E_SOCIO = E.E_SOCIO " _
             & " WHERE (M.TIPCOB = '02') AND " _
             & "       (M.E_SOCIO <> 'EXC') AND " _
             & "       (M.E_SOCIO <> 'REN') AND " _
             & "       (M.E_SOCIO <> 'FAL') AND " _
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
         wCar = ADO8!carnetpnp
         wDoc = ADO8!numdoc
         wNom = ADO8!nombre
         wApo = ADO8!aporte
         wMon = ADO8!moneda
         wTotEnvio = 0: wTotAport = 0: wTotDeuda = 0: wTotAdela = 0: wNetSocio = 0
         wCodAsig1 = 0: wNomAsig1 = "": wTotAsig1 = 0: wDeuAsig1 = 0: wAdeAsig1 = 0: wNetAsig1 = 0
         wCodAsig2 = 0: wNomAsig2 = "": wTotAsig2 = 0: wDeuAsig2 = 0: wAdeAsig2 = 0: wNetAsig2 = 0
         wCodAsig3 = 0: wNomAsig3 = "": wTotAsig3 = 0: wDeuAsig3 = 0: wAdeAsig3 = 0: wNetAsig3 = 0
         wCodAsig4 = 0: wNomAsig4 = "": wTotAsig4 = 0: wDeuAsig4 = 0: wAdeAsig4 = 0: wNetAsig4 = 0
         wCodAsig5 = 0: wNomAsig5 = "": wTotAsig5 = 0: wDeuAsig5 = 0: wAdeAsig5 = 0: wNetAsig5 = 0
         wTotAport = wApo
         wCarF = CStr(IIf(IsNull(ADO8!carnetpnp2), wCar, ADO8!carnetpnp2))
'         wTotAdela = ADO8!adelanto
'         wTotDeuda = ADO8!deuda_pt2
                  
         wTotDeuda = SaldoFoto(wSoc, wMesOld)
         If wTotDeuda < 0 Then
            wTotAdela = -wTotDeuda
            wTotDeuda = 0
         End If
         
         
         If wTotAdela >= wApo Then
            wNetSocio = 0
         Else
            wNetSocio = wApo - wTotAdela
         End If
         If wTotDeuda > 0 And wTotDeuda < Round(6 * wApo, 2) Then
            If wTotDeuda > wApo Then
               wNetSocio = wNetSocio + wApo
            Else
               wNetSocio = wNetSocio + wTotDeuda
            End If
         End If
         
         wBen = "01"
         zz = Leerado7("SELECT * FROM CAJAMP_REPRE " _
                    & " WHERE    CODPER = '" + Format(wCar, "0000000000") + "' ")
         If zz > 0 Then
            wBen = ADO7!codbeni
            If (Len(Trim(wDoc)) = 0 Or Val(wDoc) = 0) And Len(Trim(ADO7!nrodoi)) > 0 Then
               wDoc = Trim(ADO7!nrodoi)
            
               Db.BeginTrans
               Db.Execute ("UPDATE MAESOCIO " _
               & " SET NUMDOC = '" + wDoc + "' " _
               & " WHERE CODSOCIO = " + Str(wSoc) + " ")
               Db.CommitTrans
            End If
         End If
         Set ADO7 = Nothing
         
         zz = Leerado7("SELECT D.LIN, D.CODHIJO, M.NOMBRE, M.ADELANTO, M.DEUDA_PT2 " _
                & " FROM MAEASIGNADO AS D INNER JOIN MAESOCIO AS M " _
                & "   ON D.CODHIJO = M.CODSOCIO " _
                 & " WHERE D.CODSOCIO = " + Str(wSoc) + " AND " _
                 & "         D.ESTADO = 'H'  " _
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
                    
'                    wDeuAsig1 = ADO7!deuda_pt2
'                    wAdeAsig1 = ADO7!adelanto
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
'                    wDeuAsig2 = ADO7!deuda_pt2
'                    wAdeAsig2 = ADO7!adelanto
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
'                    wDeuAsig3 = ADO7!deuda_pt2
'                    wAdeAsig3 = ADO7!adelanto
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
'                    wDeuAsig4 = ADO7!deuda_pt2
'                    wAdeAsig4 = ADO7!adelanto
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
'                    wDeuAsig5 = ADO7!deuda_pt2
'                    wAdeAsig5 = ADO7!adelanto
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
         wTotEnvio = wNetSocio + _
                     wNetAsig1 + wNetAsig2 + wNetAsig3 + wNetAsig4 + wNetAsig5

         If wTotEnvio > 0 Then
            Db.BeginTrans
            Db.Execute ("INSERT INTO TMP_CAJMPCAB " _
            & " (MES, CODSOCIO, CODIGO, INS, E_SOCIO, CARNETPNP, NUMDOC, CODBENI, NOMBRE, FECENV, FECDSC, " _
            & "   TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, " _
            & "   NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, " _
            & "   ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
            & "   DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
            & "   TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
            & "   CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
            & "   NOMASIG1, NOMASIG2, NOMASIG3, NOMASIG4, NOMASIG5, " _
            & "   TOTENVIO, DSCCAJMP, DSCDIFER, TIPCOB, USU ) " _
            & " VALUES " _
            & " ('" + wAno + wMes + "', " + Str(wSoc) + ", " + Str(wCod) + ", " + Str(wIns) + ", '" + ADO8!e_socio + "', '" + wCarF + "', " _
            & "  '" + wDoc + "', '" + wBen + "', '" + Trim(wNom) + "', " _
            & "  '" + Format(Date, "dd/mm/yyyy") + "', null, " _
            & "  " + Str(wTotAport) + ", " + Str(wTotDeuda) + ", " + Str(wTotAdela) + ", " + Str(wNetSocio) + ", " _
            & "  " + Str(wNetAsig1) + ", " + Str(wNetAsig2) + ", " + Str(wNetAsig3) + ", " + Str(wNetAsig4) + ", " + Str(wNetAsig5) + ",  " _
            & "  " + Str(wAdeAsig1) + ", " + Str(wAdeAsig2) + ", " + Str(wAdeAsig3) + ", " + Str(wAdeAsig4) + ", " + Str(wAdeAsig5) + ",  " _
            & "  " + Str(wDeuAsig1) + ", " + Str(wDeuAsig2) + ", " + Str(wDeuAsig3) + ", " + Str(wDeuAsig4) + ", " + Str(wDeuAsig5) + ",  " _
            & "  " + Str(wTotAsig1) + ", " + Str(wTotAsig2) + ", " + Str(wTotAsig3) + ", " + Str(wTotAsig4) + ", " + Str(wTotAsig5) + ",  " _
            & "  " + Str(wCodAsig1) + ", " + Str(wCodAsig2) + ", " + Str(wCodAsig3) + ", " + Str(wCodAsig4) + ", " + Str(wCodAsig5) + ",  " _
            & "  '" + Trim(wNomAsig1) + "', '" + Trim(wNomAsig2) + "', '" + Trim(wNomAsig3) + "', '" + Trim(wNomAsig4) + "', '" + Trim(wNomAsig5) + "', " _
            & "  " + Str(wTotEnvio) + ", 0, 0, '" + ADO8!tipcob + "', '" + wcodusu + "' ) ")
            Db.CommitTrans
         End If
   
         wRegAct = wRegAct + 1
         ADO8.MoveNext
      Loop
   End If

   lblMensaje.Caption = ""
   lblMensaje.Refresh

   zz = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, NETSOCIO, " _
                & "      NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, TOTENVIO, " _
                & "      TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
                & "      ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
                & "      DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
                & "      DSCCAJMP, DSCDIFER, TIPCOB, " _
                & "      CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
                & "      NOMASIG1, NOMASIG2, NOMASIG3, NOMASIG4, NOMASIG5, " _
                & "      NOMASIG1, NOMASIG2, NOMASIG3, NOMASIG4, NOMASIG5, CARNETPNP, NUMDOC, CODBENI, E_SOCIO " _
                & " FROM TMP_CAJMPCAB " _
                & " WHERE MES = '" + wAno + wMes + "' AND USU = '" + wcodusu + "' " _
                & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO2

'   LlenaCab
   LlenaCab1
   LabelCab
   TotalCab
End Sub

Private Sub cmdCrear_Click()
   Dim aa As Integer, wMes As String, wAno As String
   
   wAno = txtAnoCab.Text
   wMes = wanocia + "00"
   aa = Leerado8("SELECT MAX(MES) AS MES " _
                & " FROM CONTROL_ENVIO " _
                & " WHERE TIPO = '02' AND " _
                & "       MES LIKE '" + wanocia + "%' ")
   If aa > 0 Then
      wMes = IIf(IsNull(ADO8!mes), wanocia + "00", ADO8!mes)
   End If
   wMes = Format(Val(Right(wMes, 2)) + 1, "00")

   cmbMeses.AddItem wMes + " " + Trim(funnommes(wMes))

   Db.BeginTrans
   Db.Execute ("UPDATE CONTROL_ENVIO " _
   & " SET ESTADO = 'C' " _
   & " WHERE TIPO = '02' AND " _
   & "       MES < '" + wAno + wMes + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO CONTROL_ENVIO " _
   & " (TIPO, MES, ESTADO) " _
   & " VALUES " _
   & " ('02', '" + wAno + wMes + "', 'A') ")
   Db.CommitTrans

   cmbMeses.SetFocus
End Sub

Private Sub cmdCreaTXT_Click()
   Dim zz As Long, wRegAct As Long, wRegTot As Long, _
       wAno As String, wMes As String, wCod As Integer, wIns As Integer, _
       wDir As String, wDir2 As String, wFile As String, wSec As String, _
       wTotSoc As Currency, wTotAsi As Currency, wTotDsc As Currency
         
   wAno = txtAnoCab.Text
   wMes = Left(cmbMeses.Text, 2)
   wDir = xraizCAJAMP + wAno
   wDir2 = xraizCAJAMP + wAno + "\" + wAno + "-" + wMes
   wFile = wDir2 + "\ENVIO_APIP" + wMes + Right(wAno, 2) + ".TXT"
         
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
   
   zz = Leerado8("SELECT * FROM TMP_CAJMPCAB WHERE USU = '" + wcodusu + "' AND MES = '" + wAno + wMes + "' ORDER BY NOMBRE ")
   If zz > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wTotSoc = ADO8!netsocio
         wTotAsi = ADO8!netasig1 + ADO8!netasig2 + ADO8!netasig3 + ADO8!netasig4 + ADO8!netasig5
         wTotDsc = wTotSoc + wTotAsi
         wSec = IIf(IsNull(ADO8!codbeni), "", ADO8!codbeni)
         
         Print #1, "0019" + _
                   "1502" + _
                   "8" + _
                   Format(ADO8!carnetpnp, "0000000000") + _
                   wSec + _
                   "LE" + ADO8!numdoc + "  " + _
                   wAno + wMes + _
                   Format(wTotDsc, "0000000.00")
                   
         ADO8.MoveNext
      Loop
   End If
   
   Close #1

   MsgBox "Proceso Termino OK", vbExclamation

   MsgBox "El Archivo Creado se encontrar� en " + _
          App.Path + "\CAJAMP\" + wAno + "-" + wMes
End Sub

Private Sub cmdDetalle_Click()
   zDetaCambio = True
   zDetaCodSoc = ADO2!codsocio
   zDetaTipDsc = "02"
   zDetaAnoDsc = txtAnoCab.Text
   zDetaMesDsc = Left(cmbMeses.Text, 2)
   zDetaSw = False

   frmDIECODetalle.Show vbModal

   If zDetaSw = True Then
      optTodos.Value = True
   
      ADO2.Requery
      LlenaCab
      LlenaCab1
      LabelCab
      TotalCab
      ADO2.Find "CODSOCIO=" + Str(zDetaCodSoc) + ""
   End If
End Sub

Private Sub cmdEliminar_Click()
   Dim wAno As String, wMes As String
   
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
   
   If MsgBox("Esta Seguro de Querer Eliminar Calculo " + vbNewLine + _
             "del Mes " + Trim(funnommes(wMes)) + " " + wAno + _
             "???", vbYesNo + vbQuestion, "Eliminar Archivo CAJA MILITAR POLICIAL") = vbNo Then
      Exit Sub
   End If
   Set DataGrid1.DataSource = Nothing
   
   lblEnviado.Caption = ""
   lblRecibido.Caption = ""
   lblNoDscto.Caption = ""
   lblCanApo.Caption = ""
   lblCanAsi.Caption = ""

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_CAJMPCAB WHERE USU = '" + wcodusu + "' AND MES = '" + wAno + wMes + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM CAJMPCAB WHERE MES = '" + wAno + wMes + "' ")
   Db.CommitTrans

   MsgBox "Calculo de Mes " + Trim(funnommes(wMes)) + "-" + wAno + " Eliminado OK", vbExclamation
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(9) As String, wRegAct As Integer, wRegTot As Integer
   Dim wNom As String, wAno As String, wMes As String, _
       wTotSoc As Currency, wTotAsi As Currency, wTotDsc As Currency
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
   
   Heading(0) = "COD.ENTIDAD"
   Heading(1) = "CONCEPTO"
   Heading(2) = "INSTITUTO"
   Heading(3) = "COD.PERSONA"
   Heading(4) = "NOMBRE"
   Heading(5) = "NRO.SECUE"
   Heading(6) = "TIPO DOC.IDE"
   Heading(7) = "NRO.DOC.IDE"
   Heading(8) = "PERIODO"
   Heading(9) = "IMPORTE"
   aa = Leerado3("SELECT * FROM TMP_CAJMPCAB WHERE USU = '" + wcodusu + "' AND MES = '" + wAno + wMes + "'  ORDER BY NOMBRE ")
   If aa > 0 Then
      wRegAct = 1
      wRegTot = aa
      Set objExcel = New Excel.Application
      objExcel.SheetsInNewWorkbook = 1
      objExcel.Workbooks.Add
      With objExcel
           
           .Range(objExcel.Cells(1, 1), .Cells(1, 10)).Merge
           .Range(objExcel.Cells(1, 1), .Cells(1, 10)).HorizontalAlignment = xlCenter
           
           .Range(objExcel.Cells(2, 1), .Cells(2, 10)).Merge
           .Range(objExcel.Cells(2, 1), .Cells(2, 10)).HorizontalAlignment = xlCenter
           
           .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
           .Range(.Cells(3, 1), .Cells(3, 10)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 10)).Font.Bold = True
           .Cells(1, 1) = "REPORTE DE LOS SOCIOS AOPIP PARA EL DESCUENTO CAJA MILITAR POLICIAL HABERES " + funnommes(wMes) + "-" + wanocia
           .Cells(2, 1) = "POR CONCEPTO DE CUOTAS DE APORTACION MENSUAL"
           
           For I = 1 To 10 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           
           .Range("A1:I1").Merge
           .Range("A2:I2").Merge
           
'           .Range(objExcel.Cells(3, 1), .Cells(3, 8)).Range
'           .Range(objExcel.Cells(3, 1), .Cells(3, 8)).VerticalAlignment = xlCenter
           
           
           objExcel.Columns("A").ColumnWidth = 8
           objExcel.Columns("B").ColumnWidth = 8
           objExcel.Columns("C").ColumnWidth = 6
           objExcel.Columns("D").ColumnWidth = 13
           objExcel.Columns("E").ColumnWidth = 60
           objExcel.Columns("F").ColumnWidth = 7
           objExcel.Columns("G").ColumnWidth = 13
           objExcel.Columns("H").ColumnWidth = 10
           objExcel.Columns("I").ColumnWidth = 14
      End With
      V = 4
      H = 1
      wRegAct = 1
      Do While Not ADO3.EOF
         DoEvents
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + _
                              Format(wRegAct, "####0") + " / " + _
                              Format(wRegTot, "####0")
         lblMensaje.Refresh
         
         wTotSoc = ADO3!netsocio
         wTotAsi = ADO3!netasig1 + ADO3!netasig2 + ADO3!netasig3 + ADO3!netasig4 + ADO3!netasig5
         wTotDsc = wTotSoc + wTotAsi
         
         objExcel.Columns(H + 0).Select
         objExcel.Selection.NumberFormat = "@"
         
         objExcel.Columns(H + 1).Select
         objExcel.Selection.NumberFormat = "@"
         
         objExcel.Columns(H + 2).Select
         objExcel.Selection.NumberFormat = "@"
         
         objExcel.Columns(H + 3).Select
         objExcel.Selection.NumberFormat = "@"
         
         objExcel.Columns(H + 5).Select
         objExcel.Selection.NumberFormat = "@"
         
         objExcel.Columns(H + 7).Select
         objExcel.Selection.NumberFormat = "@"
         
         objExcel.Columns(H + 8).Select
         objExcel.Selection.NumberFormat = "@"
         
         objExcel.Range(objExcel.Cells(V, H + 9), objExcel.Cells(V, H + 9)).NumberFormat = "#####,##0.00"
         
         objExcel.Cells(V, H + 0) = "0019"
         objExcel.Cells(V, H + 1) = "1502"
         objExcel.Cells(V, H + 2) = "8"
         objExcel.Cells(V, H + 3) = ADO3!carnetpnp
         objExcel.Cells(V, H + 4) = ADO3!nombre
         objExcel.Cells(V, H + 5) = ADO3!codbeni
         objExcel.Cells(V, H + 6) = "LE"
         objExcel.Cells(V, H + 7) = ADO3!numdoc
         objExcel.Cells(V, H + 8) = wAno + wMes
         objExcel.Cells(V, H + 9) = wTotDsc
         
         wRegAct = wRegAct + 1
         V = V + 1
         ADO3.MoveNext
      Loop
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
   Db.Execute ("DELETE FROM CAJMPCAB WHERE MES = '" + zAno + zMes + "' ")
   Db.CommitTrans
 
   Db.BeginTrans
   Db.Execute ("INSERT INTO CAJMPCAB " _
   & " ( MES, CODSOCIO, CODIGO, INS, E_SOCIO, CARNETPNP, NUMDOC, CODBENI, FECENV, FECDSC, " _
   & "   TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, TOTENVIO, DSCCAJMP, DSCDIFER, TIPCOB, " _
   & "   CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
   & "   TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
   & "   DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
   & "   ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
   & "   NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, " _
   & "   DSCASIG1, DSCASIG2, DSCASIG3, DSCASIG4, DSCASIG5, " _
   & "   DIFASIG1, DIFASIG2, DIFASIG3, DIFASIG4, DIFASIG5 ) " _
   & " SELECT " _
   & "   MES, CODSOCIO, CODIGO, INS, E_SOCIO, CARNETPNP, NUMDOC, CODBENI, FECENV, FECDSC, " _
   & "   TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, TOTENVIO, DSCCAJMP, DSCDIFER, TIPCOB, " _
   & "   CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
   & "   TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
   & "   DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
   & "   ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
   & "   NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, " _
   & "   DSCASIG1, DSCASIG2, DSCASIG3, DSCASIG4, DSCASIG5, " _
   & "   DIFASIG1, DIFASIG2, DIFASIG3, DIFASIG4, DIFASIG5 " _
   & "   TOTENVIO " _
   & " FROM TMP_CAJMPCAB " _
   & " WHERE MES = '" + zAno + zMes + "' AND " _
   & "       USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE MAESOCIO " _
   & " SET ENV_540 = T.NETSOCIO, ENV_541 = NETASIG1 + NETASIG2 + NETASIG3 + NETASIG4 + NETASIG5 " _
   & " FROM MAESOCIO AS M INNER JOIN TMP_CAJMPCAB AS T " _
   & "   ON M.CODSOCIO = T.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' AND " _
   & "       T.MES = '" + zAno + zMes + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_MAESTRO " _
   & " SET ENV_540 = T.NETSOCIO, ENV_541 = NETASIG1 + NETASIG2 + NETASIG3 + NETASIG4 + NETASIG5 " _
   & " FROM MAESOCIO AS M INNER JOIN TMP_CAJMPCAB AS T " _
   & "   ON M.CODSOCIO = T.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' AND " _
   & "       T.MES = '" + zAno + zMes + "' ")
   Db.CommitTrans

   MsgBox "Proceso CAJA MILITAR POLICIAL Grabado OK", vbExclamation
End Sub

Private Sub cmdRecupera_Click()
   Dim wMes As String, wRegAct As Integer, wRegTot As Integer, aa As Integer, _
       wSoc As Integer, wCod As Long, wIns As Integer, wNom As String, wE_S As String, wTipCob As String, _
       file As String, file2 As String
   wMes = "201808"
   file = xraiz + "AAA_CAJAMP\2018\2018-08\ENVIO_APIP0818.TXT"
   file2 = xraiz + "AAA_CAJAMP\2018\2018-08\0019_DSCTO_201808.TXT"
   
   Db.BeginTrans
   Db.Execute ("UPDATE CAJMPCAB " _
   & " SET NETSOCIO = 0, DSCSOCIO = 0, DIFSOCIO = 0, " _
   & "     NETASIG1 = 0, DSCASIG1 = 0, DIFASIG1 = 0, " _
   & "     NETASIG2 = 0, DSCASIG2 = 0, DIFASIG2 = 0, " _
   & "     NETASIG3 = 0, DSCASIG3 = 0, DIFASIG3 = 0, " _
   & "     NETASIG4 = 0, DSCASIG4 = 0, DIFASIG4 = 0, " _
   & "     NETASIG5 = 0, DSCASIG5 = 0, DIFASIG5 = 0, " _
   & "     TOTENVIO = 0, DSCCAJMP = 0, DSCDIFER = 0 " _
   & " WHERE MES = '" + wMes + "' ")
   Db.CommitTrans

   Dim micadena As String, wCarPNP As Long, wcodben As String, wNumDoc As String, wTotDsc As Currency
    'configura
   Open file For Input As #1 ' Abre el archivo para recibir los datos.
    ' Repite el bucle hasta el final del archivo.
   Do While Not EOF(1)
      Line Input #1, micadena
      
      wCarPNP = Mid(micadena, 10, 10)
      wcodben = Mid(micadena, 20, 2)
      wNumDoc = Mid(micadena, 24, 8)
      wTotDsc = Mid(micadena, 40, 10)
      wSoc = 0: wCod = 0: wIns = 0
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE NUMDOC = '" + wNumDoc + "' AND CARNETPNP > 0  ")
      If aa = 0 Then
         MsgBox "DNI " + wNumDoc + " NO Existe"
         Exit Sub
      End If
      If aa > 1 Then
         MsgBox "DNI " + wNumDoc + " Esta Duplicado"
         Exit Sub
      End If
      wSoc = ADO8!codsocio
      wCod = ADO8!codigo
      wIns = ADO8!ins
      wNom = ADO8!nombre
      wE_S = ADO8!e_socio
      wTipCob = ADO8!tipcob
      Set ADO8 = Nothing
      
      aa = Leerado8("SELECT * FROM CAJMPCAB " _
                & " WHERE      MES = '" + wMes + "' AND " _
                & "       CODSOCIO = " + Str(wSoc) + " ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO CAJMPCAB " _
         & " (MES, CODSOCIO, CODIGO, INS, E_SOCIO, CARNETPNP, NUMDOC, CODBENI, FECENV, FECDSC, TIPCOB ) " _
         & " VALUES " _
         & " ('" + wMes + "', " + Str(wSoc) + ", " + Str(wCod) + ", " + Str(wIns) + ", '" + wE_S + "', " _
         & "  " + Str(wCarPNP) + ", '" + wNumDoc + "', '" + wcodben + "', '01/08/2018', '01/08/2018', '" + wTipCob + "' ) ")
         Db.CommitTrans
      End If
      Set ADO8 = Nothing
         
      Db.BeginTrans
      Db.Execute ("UPDATE CAJMPCAB " _
      & " SET TOTENVIO = " + Str(wTotDsc) + " " _
      & " WHERE      MES = '" + wMes + "' AND " _
      & "       CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
   
   Loop
   Close #1
    
   Open file2 For Input As #1 ' Abre el archivo para recibir los datos.
    ' Repite el bucle hasta el final del archivo.
   Do While Not EOF(1)
      Line Input #1, micadena
    
      wCarPNP = Mid(micadena, 10, 10)
      wcodben = Mid(micadena, 20, 2)
      wNumDoc = Mid(micadena, 24, 8)
      wTotDsc = Mid(micadena, 40, 10)
      wSoc = 0: wCod = 0: wIns = 0
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE NUMDOC = '" + wNumDoc + "' AND CARNETPNP > 0 ")
      If aa = 0 Then
         MsgBox "DNI " + wNumDoc + " NO Existe"
         Exit Sub
      End If
      If aa > 1 Then
         MsgBox "DNI " + wNumDoc + " Esta Duplicado"
         Exit Sub
      End If
      wSoc = ADO8!codsocio
      wCod = ADO8!codigo
      wIns = ADO8!ins
      wNom = ADO8!nombre
      wE_S = ADO8!e_socio
      wTipCob = ADO8!tipcob
      Set ADO8 = Nothing
      
      aa = Leerado8("SELECT * FROM CAJMPCAB " _
                & " WHERE      MES = '" + wMes + "' AND " _
                & "       CODSOCIO = " + Str(wSoc) + " ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO CAJMPCAB " _
         & " (MES, CODSOCIO, CODIGO, INS, E_SOCIO, CARNETPNP, NUMDOC, CODBENI, FECENV, FECDSC, TIPCOB ) " _
         & " VALUES " _
         & " ('" + wMes + "', " + Str(wSoc) + ", " + Str(wCod) + ", " + Str(wIns) + ", '" + wE_S + "', " _
         & "  " + Str(wCarPNP) + ", '" + wNumDoc + "', '" + wcodben + "', '01/08/2018', '01/08/2018', '" + wTipCob + "' ) ")
         Db.CommitTrans
      End If
      Set ADO8 = Nothing
         
      Db.BeginTrans
      Db.Execute ("UPDATE CAJMPCAB " _
      & " SET DSCCAJMP = " + Str(wTotDsc) + " " _
      & " WHERE      MES = '" + wMes + "' AND " _
      & "       CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans

   Loop
   Close #1

   Db.BeginTrans
   Db.Execute ("UPDATE CAJMPCAB " _
   & " SET DSCDIFER = TOTENVIO - DSCCAJMP " _
   & " WHERE      MES = '" + wMes + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE CAJMPCAB " _
   & " SET DSCSOCIO = DSCCAJMP, " _
   & "     DIFSOCIO = DSCDIFER " _
   & " WHERE      MES = '" + wMes + "' AND " _
   & "       CODASIG1 = 0 AND " _
   & "       CODASIG2 = 0 AND " _
   & "       CODASIG3 = 0 AND " _
   & "       CODASIG4 = 0 AND " _
   & "       CODASIG5 = 0 ")
   Db.CommitTrans

   Dim wDscCajMp As Currency, wDscDifer As Currency, wTotEnvio As Currency
   Dim wSocSocio As Integer, wSocAsig1 As Integer, wSocAsig2 As Integer, wSocAsig3 As Integer, wSocAsig4 As Integer, wSocAsig5 As Integer
   Dim wCodSocio As Long, wCodAsig1 As Long, wCodAsig2 As Long, wCodAsig3 As Long, wCodAsig4 As Long, wCodAsig5 As Long
   Dim wInsSocio As Integer, wInsAsig1 As Integer, wInsAsig2 As Integer, wInsAsig3 As Integer, wInsAsig4 As Integer, wInsAsig5 As Integer
   Dim wNetSocio As Currency, wNetAsig1 As Currency, wNetAsig2 As Currency, wNetAsig3 As Currency, wNetAsig4 As Currency, wNetAsig5 As Currency
   Dim wDscSocio As Currency, wDscAsig1 As Currency, wDscAsig2 As Currency, wDscAsig3 As Currency, wDscAsig4 As Currency, wDscAsig5 As Currency
   Dim wDifSocio As Currency, wDifAsig1 As Currency, wDifAsig2 As Currency, wDifAsig3 As Currency, wDifAsig4 As Currency, wDifAsig5 As Currency
   
   aa = Leerado8("SELECT * FROM CAJMPCAB " _
                & " WHERE (MES = '" + wMes + "') AND " _
                & "       (CODASIG1 <> 0 OR CODASIG2 <> 0 OR CODASIG3 <> 0 OR CODASIG4 <> 0 OR CODASIG5 <> 0) ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wTotEnvio = ADO8!totenvio
         wDscCajMp = ADO8!dsccajmp
         wDscDifer = ADO8!dscdifer
         wSocSocio = ADO8!codsocio
         wSocAsig1 = ADO8!codasig1
         wSocAsig2 = ADO8!codasig2
         wSocAsig3 = ADO8!codasig3
         wSocAsig4 = ADO8!codasig4
         wSocAsig5 = ADO8!codasig5
         wCodSocio = 0: wCodAsig1 = 0: wCodAsig2 = 0: wCodAsig3 = 0: wCodAsig4 = 0: wCodAsig5 = 0
         wInsSocio = 0: wInsAsig1 = 0: wInsAsig2 = 0: wInsAsig3 = 0: wInsAsig4 = 0: wInsAsig5 = 0
         wNetSocio = 0: wNetAsig1 = 0: wNetAsig2 = 0: wNetAsig3 = 0: wNetAsig4 = 0: wNetAsig5 = 0
         wDscSocio = 0: wDscAsig1 = 0: wDscAsig2 = 0: wDscAsig3 = 0: wDscAsig4 = 0: wDscAsig5 = 0
         wDifSocio = 0: wDifAsig1 = 0: wDifAsig2 = 0: wDifAsig3 = 0: wDifAsig4 = 0: wDifAsig5 = 0
         
         aa = Leerado7("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSocSocio) + " ")
         If aa > 0 Then
            wCodSocio = ADO7!codigo
            wInsSocio = ADO7!ins
         End If
         Set ADO7 = Nothing
   
         If wSocAsig1 <> 0 Then
            aa = Leerado7("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSocAsig1) + " ")
            If aa > 0 Then
               wCodAsig1 = ADO7!codigo
               wInsAsig1 = ADO7!ins
            End If
            Set ADO7 = Nothing
         End If
   
         If wSocAsig2 <> 0 Then
            aa = Leerado7("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSocAsig2) + " ")
            If aa > 0 Then
               wCodAsig2 = ADO7!codigo
               wInsAsig2 = ADO7!ins
            End If
            Set ADO7 = Nothing
         End If
   
         If wSocAsig3 <> 0 Then
            aa = Leerado7("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSocAsig3) + " ")
            If aa > 0 Then
               wCodAsig3 = ADO7!codigo
               wInsAsig3 = ADO7!ins
            End If
            Set ADO7 = Nothing
         End If
   
         If wSocAsig4 <> 0 Then
            aa = Leerado7("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSocAsig4) + " ")
            If aa > 0 Then
               wCodAsig4 = ADO7!codigo
               wInsAsig4 = ADO7!ins
            End If
            Set ADO7 = Nothing
         End If
   
         If wSocAsig5 <> 0 Then
            aa = Leerado7("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSocAsig5) + " ")
            If aa > 0 Then
               wCodAsig5 = ADO7!codigo
               wInsAsig5 = ADO7!ins
            End If
            Set ADO7 = Nothing
         End If
         
         If wCodSocio <> 0 Then
            aa = Leerado7("SELECT * FROM ZZZ_APOR_PLA " _
                    & " WHERE CODIGO = " + Str(wCodSocio) + " AND " _
                    & "          INS = " + Str(wInsSocio) + " AND " _
                    & "       CUOANO = '2018' ")
            If aa > 0 Then
               wNetSocio = ADO7!impo08
               wDscSocio = ADO7!impo08
            End If
         End If
         
         If wCodAsig1 <> 0 Then
            aa = Leerado7("SELECT * FROM ZZZ_APOR_PLA " _
                    & " WHERE CODIGO = " + Str(wCodAsig1) + " AND " _
                    & "          INS = " + Str(wInsAsig1) + " AND " _
                    & "       CUOANO = '2018' ")
            If aa > 0 Then
               wNetAsig1 = ADO7!impo08
               wDscAsig1 = ADO7!impo08
            End If
         End If
         
         If wCodAsig2 <> 0 Then
            aa = Leerado7("SELECT * FROM ZZZ_APOR_PLA " _
                    & " WHERE CODIGO = " + Str(wCodAsig2) + " AND " _
                    & "          INS = " + Str(wInsAsig2) + " AND " _
                    & "       CUOANO = '2018' ")
            If aa > 0 Then
               wNetAsig2 = ADO7!impo08
               wDscAsig2 = ADO7!impo08
            End If
         End If
         
         If wCodAsig3 <> 0 Then
            aa = Leerado7("SELECT * FROM ZZZ_APOR_PLA " _
                    & " WHERE CODIGO = " + Str(wCodAsig3) + " AND " _
                    & "          INS = " + Str(wInsAsig3) + " AND " _
                    & "       CUOANO = '2018' ")
            If aa > 0 Then
               wNetAsig3 = ADO7!impo08
               wDscAsig3 = ADO7!impo08
            End If
         End If
         
         If wCodAsig4 <> 0 Then
            aa = Leerado7("SELECT * FROM ZZZ_APOR_PLA " _
                    & " WHERE CODIGO = " + Str(wCodAsig4) + " AND " _
                    & "          INS = " + Str(wInsAsig4) + " AND " _
                    & "       CUOANO = '2018' ")
            If aa > 0 Then
               wNetAsig4 = ADO7!impo08
               wDscAsig4 = ADO7!impo08
            End If
         End If
         
         If wCodAsig5 <> 0 Then
            aa = Leerado7("SELECT * FROM ZZZ_APOR_PLA " _
                    & " WHERE CODIGO = " + Str(wCodAsig5) + " AND " _
                    & "          INS = " + Str(wInsAsig5) + " AND " _
                    & "       CUOANO = '2018' ")
            If aa > 0 Then
               wNetAsig5 = ADO7!impo08
               wDscAsig5 = ADO7!impo08
            End If
         End If
   
         If wNetSocio + wNetAsig1 + wNetAsig2 + wNetAsig3 + wNetAsig4 + wNetAsig5 <> wTotEnvio Or _
            wDscSocio + wDscAsig1 + wDscAsig2 + wDscAsig3 + wDscAsig4 + wDscAsig5 <> wDscCajMp Then
            
            
            
         Else
            Db.BeginTrans
            Db.Execute ("UPDATE CAJMPCAB " _
            & " SET NETSOCIO = " + Str(wNetSocio) + ", DSCSOCIO = " + Str(wDscSocio) + ", DIFSOCIO = " + Str(wDifSocio) + ", " _
            & "     NETASIG1 = " + Str(wNetAsig1) + ", DSCASIG1 = " + Str(wDscAsig1) + ", DIFASIG1 = " + Str(wDifAsig1) + ", " _
            & "     NETASIG2 = " + Str(wNetAsig2) + ", DSCASIG2 = " + Str(wDscAsig2) + ", DIFASIG2 = " + Str(wDifAsig2) + ", " _
            & "     NETASIG3 = " + Str(wNetAsig3) + ", DSCASIG3 = " + Str(wDscAsig3) + ", DIFASIG3 = " + Str(wDifAsig3) + ", " _
            & "     NETASIG4 = " + Str(wNetAsig4) + ", DSCASIG4 = " + Str(wDscAsig4) + ", DIFASIG4 = " + Str(wDifAsig4) + ", " _
            & "     NETASIG5 = " + Str(wNetAsig5) + ", DSCASIG5 = " + Str(wDscAsig5) + ", DIFASIG5 = " + Str(wDifAsig5) + " " _
            & " WHERE MES = '201808' AND " _
            & "       CODSOCIO = " + Str(wSocSocio) + " ")
            Db.CommitTrans
         End If
         
         ADO8.MoveNext
      Loop
   End If

   MsgBox "Proceso Termino OK", vbExclamation
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub DataGrid1_DblClick()
   zDetaCambio = False
   zDetaCodSoc = ADO2!codsocio
   zDetaTipDsc = "02"
   zDetaAnoDsc = txtAnoCab.Text
   zDetaMesDsc = Left(cmbMeses.Text, 2)
   zDetaSw = False

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
   frmMPEnvio.Left = (Screen.Width - Width) \ 2
   frmMPEnvio.Top = 0
   
   txtAnoCab.Text = wanocia
   txtAnoCab.Enabled = False
   
   Dim a As Integer
   cmbMeses.Clear
   a = Leerado("select * from CONTROL_ENVIO " _
            & " WHERE  MES LIKE '" + wanocia + "%' AND " _
            & "       TIPO = '02' " _
            & " ORDER BY MES ")
   If a > 0 Then
      ADO1.MoveFirst
      Do While Not ADO1.EOF
         cmbMeses.AddItem Right(ADO1!mes, 2) + " " + Trim(funnommes(Right(ADO1!mes, 2)))
         ADO1.MoveNext
      Loop
   End If
   cmbMeses.SetFocus
End Sub

Private Sub LlenaCab()
   Dim wAno As String, wMes As String, zz As Integer
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
      
   lblEnviado.Caption = ""
   lblRecibido.Caption = ""
   lblNoDscto.Caption = ""
   lblCanApo.Caption = ""
   lblCanAsi.Caption = ""
      
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_CAJMPCAB WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_CAJMPCAB " _
   & " (MES, CODSOCIO, CODIGO, INS, CARNETPNP, NUMDOC, CODBENI, NOMBRE, FECENV, FECDSC, " _
   & "   TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, TIPCOB, " _
   & "   NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, " _
   & "   ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
   & "   DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
   & "   TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
   & "   CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
   & "   NOMASIG1, NOMASIG2, NOMASIG3, NOMASIG4, NOMASIG5, " _
   & "   TOTENVIO, DSCCAJMP, DSCDIFER, USU ) " _
   & " SELECT " _
   & "  MES, CODSOCIO, CODIGO, INS, CARNETPNP, NUMDOC, CODBENI, '', FECENV, FECDSC, " _
   & "   TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, TIPCOB, " _
   & "   NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, " _
   & "   ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
   & "   DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
   & "   TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
   & "   CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
   & "   '', '', '', '', '', " _
   & "   TOTENVIO, DSCCAJMP, DSCDIFER, '" + wcodusu + "'  " _
   & " FROM CAJMPCAB " _
   & " WHERE MES = '" + wAno + wMes + "'  ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_CAJMPCAB " _
   & " SET NOMBRE = M.NOMBRE " _
   & " FROM TMP_CAJMPCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODSOCIO = M.CODSOCIO " _
   & " WHERE T.MES = '" + wAno + wMes + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_CAJMPCAB " _
   & " SET NOMASIG1 = M.NOMBRE " _
   & " FROM TMP_CAJMPCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG1 = M.CODSOCIO " _
   & " WHERE T.MES = '" + wAno + wMes + "' AND " _
   & "       T.CODASIG1 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_CAJMPCAB " _
   & " SET NOMASIG2 = M.NOMBRE " _
   & " FROM TMP_CAJMPCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG2 = M.CODSOCIO " _
   & " WHERE T.MES = '" + wAno + wMes + "' AND " _
   & "       T.CODASIG2 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_CAJMPCAB " _
   & " SET NOMASIG3 = M.NOMBRE " _
   & " FROM TMP_CAJMPCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG3 = M.CODSOCIO " _
   & " WHERE T.MES = '" + wAno + wMes + "' AND " _
   & "       T.CODASIG3 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_CAJMPCAB " _
   & " SET NOMASIG4 = M.NOMBRE " _
   & " FROM TMP_CAJMPCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG4 = M.CODSOCIO " _
   & " WHERE T.MES = '" + wAno + wMes + "' AND " _
   & "       T.CODASIG4 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_CAJMPCAB " _
   & " SET NOMASIG5 = M.NOMBRE " _
   & " FROM TMP_CAJMPCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG5 = M.CODSOCIO " _
   & " WHERE T.MES = '" + wAno + wMes + "' AND " _
   & "       T.CODASIG5 <> 0 ")
   Db.CommitTrans

   If wEstado = False Then
      cmdDetalle.Enabled = False
      cmdCalculo.Enabled = False
      cmdEliminar.Enabled = False
      cmdGrabar.Enabled = False
   Else
      cmdDetalle.Enabled = True
      cmdCalculo.Enabled = True
      cmdEliminar.Enabled = True
      cmdGrabar.Enabled = True
   End If

   zz = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, NETSOCIO, " _
                & "      NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, TOTENVIO, " _
                & "      ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
                & "      DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
                & "      TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
                & "      DSCCAJMP, DSCDIFER, TIPCOB, " _
                & "      CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
                & "      NOMASIG1, NOMASIG2, NOMASIG3, NOMASIG4, NOMASIG5, CARNETPNP, NUMDOC, CODBENI, E_SOCIO " _
                & " FROM TMP_CAJMPCAB " _
                & " WHERE MES = '" + wAno + wMes + "' AND USU = '" + wcodusu + "' " _
                & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO2
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
    
   DataGrid1.Columns(3).Width = 4000  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE ASOCIADO"
    
   DataGrid1.Columns(4).Width = 770    ' NETSOCIO
   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(4).Caption = "T.SOCIO"
    
   DataGrid1.Columns(5).Width = 770    ' NETASIG1
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(5).Caption = "ASIG 1"
    
   DataGrid1.Columns(6).Width = 770    ' NETASIG2
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(6).Caption = "ASIG 2"
    
   DataGrid1.Columns(7).Width = 770    ' NETASIG3
   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(7).Caption = "ASIG 3"
    
   DataGrid1.Columns(8).Width = 770    ' NETASIG4
   DataGrid1.Columns(8).Alignment = dbgRight
   DataGrid1.Columns(8).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(8).Caption = "ASIG 4"
    
   DataGrid1.Columns(9).Width = 770    ' NETASIG5
   DataGrid1.Columns(9).Alignment = dbgRight
   DataGrid1.Columns(9).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(9).Caption = "ASIG 5"
    
   DataGrid1.Columns(10).Width = 770    ' TOTENVIO
   DataGrid1.Columns(10).Alignment = dbgRight
   DataGrid1.Columns(10).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(10).Caption = "TOTAL"

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
'   DataGrid1.Columns(43).Visible = False
'   DataGrid1.Columns(44).Visible = False
'   DataGrid1.Columns(45).Visible = False
'   DataGrid1.Columns(46).Visible = False
'   DataGrid1.Columns(47).Visible = False
   
   DataGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
   DataGrid1.SetFocus
End Sub

Private Sub LabelCab()
   If Not ADO2.EOF Then
      lblAsig1.Caption = ADO2!nomasig1
      lblAsig2.Caption = ADO2!nomasig2
      lblAsig3.Caption = ADO2!nomasig3
      lblAsig4.Caption = ADO2!nomasig4
      lblAsig5.Caption = ADO2!nomasig5
   End If
End Sub

Private Sub TotalCab()
   Dim zz As Integer, _
       zAno As String, zMes As String, _
       zTotEnv As Currency, zTotDsc As Currency, zTotNoD As Currency, _
       zCanApo As Integer, zCanAsi As Integer
   
   zAno = wanocia
   zMes = Left(cmbMeses.Text, 2)
   
   zz = Leerado8("SELECT SUM(TOTENVIO) AS TOTENVIO, " _
                & "      SUM(DSCCAJMP) AS DSCCAJMP, " _
                & "      SUM(DSCDIFER) AS DSCDIFER, " _
                & "      COUNT(*) AS CAN " _
                & " FROM TMP_CAJMPCAB " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "       MES = '" + zAno + zMes + "' ")
   If zz > 0 Then
      zTotEnv = IIf(IsNull(ADO8!totenvio), 0, ADO8!totenvio)
      zTotDsc = IIf(IsNull(ADO8!dsccajmp), 0, ADO8!dsccajmp)
      zTotNoD = IIf(IsNull(ADO8!dscdifer), 0, ADO8!dscdifer)
      zCanApo = IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT COUNT(*) AS CAN " _
                & " FROM TMP_CAJMPCAB " _
                & " WHERE (USU = '" + wcodusu + "') AND " _
                & "       (MES = '" + zAno + zMes + "') AND " _
                & "       (TOTASIG1 <> 0) ")
   If zz > 0 Then
      zCanAsi = IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT COUNT(*) AS CAN " _
                & " FROM TMP_CAJMPCAB " _
                & " WHERE (USU = '" + wcodusu + "') AND " _
                & "       (MES = '" + zAno + zMes + "') AND " _
                & "       (TOTASIG2 <> 0) ")
   If zz > 0 Then
      zCanAsi = zCanAsi + IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT COUNT(*) AS CAN " _
                & " FROM TMP_CAJMPCAB " _
                & " WHERE (USU = '" + wcodusu + "') AND " _
                & "       (MES = '" + zAno + zMes + "') AND " _
                & "       (TOTASIG3 <> 0) ")
   If zz > 0 Then
      zCanAsi = zCanAsi + IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT COUNT(*) AS CAN " _
                & " FROM TMP_CAJMPCAB " _
                & " WHERE (USU = '" + wcodusu + "') AND " _
                & "       (MES = '" + zAno + zMes + "') AND " _
                & "       (TOTASIG4 <> 0) ")
   If zz > 0 Then
      zCanAsi = zCanAsi + IIf(IsNull(ADO8!can), 0, ADO8!can)
   End If
   Set ADO8 = Nothing
   
   zz = Leerado8("SELECT COUNT(*) AS CAN " _
                & " FROM TMP_CAJMPCAB " _
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
      LabelCab
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

Private Sub txtFiltrar_Change()
   Dim aa As Integer, wAno As String, wMes As String
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
   
   optFiltro.Value = True
   
   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, NETSOCIO, " _
                & "      NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, TOTENVIO, " _
                & "      ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
                & "      DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
                & "      TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
                & "      DSCCAJMP, DSCDIFER, TIPCOB, " _
                & "      CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
                & "      NOMASIG1, NOMASIG2, NOMASIG3, NOMASIG4, NOMASIG5, CARNETPNP, NUMDOC, CODBENI, E_SOCIO " _
                & " FROM TMP_CAJMPCAB " _
                & " WHERE MES = '" + wAno + wMes + "' AND " _
                & "       USU = '" + wcodusu + "' AND " _
                & "       NOMBRE LIKE '%" + Trim(txtFiltrar.Text) + "%' " _
                & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO2
   
   LlenaCab1
   LabelCab
   TotalCab
'   txtFiltrar.SelStart = Len(Trim(txtFiltrar.Text)) - 1
'   txtFiltrar.SelLength = 1
   
   
   txtFiltrar.SetFocus
End Sub

Private Sub txtFiltrar_GotFocus()
   If Len(Trim(txtFiltrar.Text)) > 0 Then
      txtFiltrar.SelStart = Len(Trim(txtFiltrar.Text))
   Else
      txtFiltrar.SelStart = 0
   End If
   txtFiltrar.SelLength = 1
End Sub

Private Sub txtFiltrar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
   
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub
