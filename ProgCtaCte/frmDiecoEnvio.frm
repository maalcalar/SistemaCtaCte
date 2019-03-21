VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDiecoEnvio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envio Dieco"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   14475
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
      Left            =   3720
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtro"
      Height          =   855
      Left            =   360
      TabIndex        =   30
      Top             =   7440
      Width           =   6495
      Begin VB.OptionButton optTodos 
         Caption         =   "Mostrar Todos"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   200
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Filtrar x Nombre"
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   500
         Width           =   1575
      End
      Begin VB.TextBox txtFiltrar 
         Height          =   285
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   31
         Top             =   500
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Modificar Envio Un Socio"
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
      Left            =   13080
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2280
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
      Left            =   6360
      TabIndex        =   17
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
      Left            =   11880
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7680
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
      Left            =   8640
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7680
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
      Left            =   5040
      TabIndex        =   13
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
      Left            =   7440
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7680
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
      Left            =   13200
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5415
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   9551
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
   Begin VB.TextBox txtAnoCab 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.ComboBox cmbMeses 
      Height          =   315
      ItemData        =   "frmDiecoEnvio.frx":0000
      Left            =   960
      List            =   "frmDiecoEnvio.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Cobrado"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   13080
      TabIndex        =   37
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblRecibido 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   13080
      TabIndex        =   36
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "No Cobrado"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   13080
      TabIndex        =   35
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblNoDscto 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   13080
      TabIndex        =   34
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ENVIA DESCUENTOS A DIECO"
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
      Left            =   3840
      TabIndex        =   28
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
      Left            =   6960
      TabIndex        =   27
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label lblAsig5 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7680
      TabIndex        =   26
      Top             =   6840
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
      Left            =   6960
      TabIndex        =   25
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label lblAsig4 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7680
      TabIndex        =   24
      Top             =   6600
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
      Left            =   360
      TabIndex        =   23
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label lblAsig3 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   22
      Top             =   7080
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
      Left            =   360
      TabIndex        =   21
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label lblAsig2 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   20
      Top             =   6840
      Width           =   5175
   End
   Begin VB.Label lblAsig1 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   6600
      Width           =   5175
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
      Left            =   360
      TabIndex        =   18
      Top             =   6600
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
      Left            =   7440
      TabIndex        =   16
      Top             =   7320
      Width           =   5295
   End
   Begin VB.Label Label4 
      Caption         =   "Total Envio S/."
      ForeColor       =   &H00FF0000&
      Height          =   200
      Left            =   13080
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Cant.Asignados"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   11760
      TabIndex        =   9
      Top             =   615
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Cant.Titulares"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   11760
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblCanAsi 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11760
      TabIndex        =   7
      Top             =   825
      Width           =   1095
   End
   Begin VB.Label lblEnviado 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   13080
      TabIndex        =   6
      Top             =   320
      Width           =   1095
   End
   Begin VB.Label lblCanApo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11760
      TabIndex        =   5
      Top             =   315
      Width           =   1095
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
Attribute VB_Name = "frmDiecoEnvio"
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
      
      wEstado = True
      
      
      zz = Leerado8("SELECT * FROM CONTROL_ENVIO " _
                    & " WHERE MES = '" + wanocia + wMes + "' AND " _
                    & "       TIPO = '01' ")
      If zz > 0 Then
         wEstado = IIf(ADO8!estado = "C", False, True)
      End If
      Set ADO8 = Nothing
      
      zz = Leerado2("SELECT * FROM DIECOCAB " _
                & "  WHERE MES = '" + wAno + wMes + "' ")
      If zz > 0 Then
      
         lblMensaje.Caption = "Trae Calculo DIECO - Mes " + Left(Trim(funnommes(wMes)), 3) + " " + wAno
         lblMensaje.Refresh
      
         LlenaCab
         LlenaCab1
         TotalCab
         ADO2.MoveFirst
         LabelCab
      
         lblMensaje.Caption = ""
         lblMensaje.Refresh
      
      Else
         cmdCalculo.SetFocus
      End If
   End If
End Sub

Private Sub cmdCalculo_Click()
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
       wNetSocio As Currency, wLin As Integer, wE_S As String, wCodPnp As Long, wInsPnp As Integer, _
       wMesOld As String
   
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
   
   If wMes > "01" Then
      wMesOld = wAno + Format(Val(wMes) - 1, "00")
   Else
      wMesOld = Format(Val(wAno) - 1, "00") + "12"
   End If
   
   wFecEnv = "01/" + wMes + "/" + wAno
   
   wGrabar = True
   
   zz = Leerado8("SELECT * FROM DIECOCAB " _
             & "  WHERE MES = '" + wAno + wMes + "' ")
   If zz > 0 Then
      If MsgBox("Ya Existe Proceso DIECO del Mes" + vbNewLine + _
                "Desea Volver a Crearlo???", vbYesNo + vbQuestion, "Crear Archivo Descuento DIECO") = vbNo Then
         Exit Sub
      End If
   End If
   Set ADO8 = Nothing
   
   Set DataGrid1.DataSource = Nothing
   
   lblMensaje.Caption = "Calculando Descuentos DIECO - Mes " + Trim(funnommes(wMes)) + " " + wAno
   lblMensaje.Refresh
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_DIECOCAB WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM DIECOCAB WHERE MES = '" + wAno + wMes + "' ")
   Db.CommitTrans
   
'' SE CONSIDERAN TIPO DE COBRO 01 DIECO Y 04 PNP NO ASOCIADO
   
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
         wE_S = ADO8!e_socio
         
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
                 & "         D.ESTADO = 'H' " _
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
            & "  " + Str(wCodPnp) + ", " + Str(wInsPnp) + ", '" + wE_S + "', '" + wcodusu + "' ) ")
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

Private Sub cmdCrear_Click()
   Dim aa As Integer, wMes As String, wAno As String
   
   wAno = txtAnoCab.Text
   wMes = wanocia + "00"
   aa = Leerado8("SELECT MAX(MES) AS MES " _
                & " FROM CONTROL_ENVIO " _
                & " WHERE TIPO = '01' AND " _
                & "       MES LIKE '" + wanocia + "%' ")
   If aa > 0 Then
      wMes = IIf(IsNull(ADO8!mes), wanocia + "00", ADO8!mes)
   End If
   wMes = Format(Val(Right(wMes, 2)) + 1, "00")

   cmbMeses.AddItem wMes + " " + Trim(funnommes(wMes))

   Db.BeginTrans
   Db.Execute ("UPDATE CONTROL_ENVIO " _
   & " SET ESTADO = 'C' " _
   & " WHERE TIPO = '01' AND " _
   & "       MES < '" + wAno + wMes + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO CONTROL_ENVIO " _
   & " (TIPO, MES, ESTADO) " _
   & " VALUES " _
   & " ('01', '" + wAno + wMes + "', 'A') ")
   Db.CommitTrans

   cmbMeses.SetFocus
End Sub

Private Sub cmdCreaTXT_Click()
   Dim zz As Long, wRegAct As Long, wRegTot As Long, _
       wAno As String, wMes As String, wCod As Integer, wIns As Integer, _
       wDir As String, wDir2 As String, wFile As String, _
       wTotSoc As Currency, wTotAsi As Currency, wTotDsc As Currency
         
   wAno = txtAnoCab.Text
   wMes = Left(cmbMeses.Text, 2)
   
   wDir = xraizDIECO + wAno
   wDir2 = xraizDIECO + wAno + "\" + wAno + "-" + wMes
   wFile = wDir2 + "\H15020001.TXT"
   
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
   
   Dim zTotDsc As String, zTotSoc As String, zTotAsi As String, _
       wUno As Currency, wDos As Currency
   
   zz = Leerado8("SELECT * FROM TMP_DIECOCAB WHERE USU = '" + wcodusu + "' ORDER BY CODIGO, INS ")
   If zz > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wTotSoc = ADO8!netsocio
         wTotAsi = ADO8!netasig1 + ADO8!netasig2 + ADO8!netasig3 + ADO8!netasig4 + ADO8!netasig5
         wTotDsc = wTotSoc + wTotAsi
   
   
         If wTotSoc >= 50 Then
            wUno = 50
            wDos = wTotDsc - 50
         Else
            wUno = wTotSoc
            wDos = wTotDsc - wTotSoc
         End If
   
         zTotDsc = Format(Int(wTotDsc), "0000000000") + Format((wTotDsc - Int(wTotDsc)) * 100, "00")
         zTotSoc = Format(Int(wUno), "0000000000") + Format((wUno - Int(wUno)) * 100, "00")
         zTotAsi = Format(Int(wDos), "0000000000") + Format((wDos - Int(wDos)) * 100, "00")
   
         Print #1, Format(ADO8!codigo, "00000000") + _
                   Format(ADO8!ins, "0") + _
                   "1502" + _
                   "0001" + _
                   zTotDsc + _
                   zTotSoc + _
                   zTotAsi
   
         ADO8.MoveNext
      Loop
   End If
   
   Close #1

   MsgBox "Proceso Termino OK", vbExclamation

   MsgBox "El Archivo Creado se encontrará en " + _
          App.Path + "DIECO\" + wAno + "-" + wMes

End Sub

Private Sub cmdDetalle_Click()
   zDetaCambio = True
   zDetaCodSoc = ADO2!codsocio
   zDetaTipDsc = "01"
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
   Dim wAno As String, wMes As String, zz As Integer
   
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
   
   If MsgBox("Esta Seguro de Querer Eliminar Calculo " + vbNewLine + _
             "del Mes " + Trim(funnommes(wMes)) + " " + wAno + _
             "???", vbYesNo + vbQuestion, "Eliminar Archivo DIECO") = vbNo Then
      Exit Sub
   End If
   Set DataGrid1.DataSource = Nothing
   
   lblEnviado.Caption = ""
   lblRecibido.Caption = ""
   lblNoDscto.Caption = ""
   lblCanApo.Caption = ""
   lblCanAsi.Caption = ""

   Db.BeginTrans
   Db.Execute ("DELETE FROM CONTROL_ENVIO " _
   & " WHERE TIPO = '01' AND " _
   & "        MES = '" + wAno + wMes + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_DIECOCAB WHERE USU = '" + wcodusu + "' AND MES = '" + wAno + wMes + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM DIECOCAB WHERE MES = '" + wAno + wMes + "' ")
   Db.CommitTrans

   MsgBox "Calculo de Mes " + Trim(funnommes(wMes)) + "-" + wAno + " Eliminado OK", vbExclamation
End Sub

Private Sub cmdExportar_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(7) As String, wRegAct As Integer, wRegTot As Integer
   Dim wNom As String, wMes As String, _
       wTotSoc As Currency, wTotAsi As Currency, wTotDsc As Currency, _
       wUno As Currency, wDos As Currency
   wMes = Left(cmbMeses.Text, 2)
   
   Heading(0) = "NRO"
   Heading(1) = "CODOFIN"
   Heading(2) = "NOMBRE"
   Heading(3) = "COD.DES"
   Heading(4) = "SUB COD"
   Heading(5) = "CUOTA NORMAL"
   Heading(6) = "CUOTA HIJOS"
   Heading(7) = "IMPORTE TOTAL"
   aa = Leerado3("SELECT * FROM TMP_DIECOCAB WHERE USU = '" + wcodusu + "' ORDER BY NOMBRE ")
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
           .Range(.Cells(3, 1), .Cells(3, 8)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 8)).Font.Bold = True
           .Cells(1, 1) = "REPORTE DE LOS SOCIOS AOPIP PARA EL DESCUENTO POR PLANILLA DE HABERES " + funnommes(wMes) + "-" + wanocia
           .Cells(2, 1) = "POR CONCEPTO DE CUOTAS DE APORTACION MENSUAL"
           
           For I = 1 To 8 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           
           .Range("A1:H1").Merge
           .Range("A2:H2").Merge
           
'           .Range(objExcel.Cells(3, 1), .Cells(3, 8)).Range
'           .Range(objExcel.Cells(3, 1), .Cells(3, 8)).VerticalAlignment = xlCenter
           
           
           objExcel.Columns("A").ColumnWidth = 8
           objExcel.Columns("B").ColumnWidth = 11
           objExcel.Columns("C").ColumnWidth = 50
           objExcel.Columns("D").ColumnWidth = 10
           objExcel.Columns("E").ColumnWidth = 10
           objExcel.Columns("F").ColumnWidth = 14
           objExcel.Columns("G").ColumnWidth = 14
           objExcel.Columns("H").ColumnWidth = 14
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
         
         If wTotSoc >= 50 Then
            wUno = 50
            wDos = wTotDsc - 50
         Else
            wUno = wTotSoc
            wDos = wTotDsc - wTotSoc
         End If
         
         objExcel.Range(objExcel.Cells(V, H + 5), objExcel.Cells(V, H + 7)).NumberFormat = "#####,##0.00"
         
         objExcel.Cells(V, H + 0) = wRegAct
         objExcel.Cells(V, H + 1) = Format(ADO3!codigo, "#######0") + Format(ADO3!ins, "0")
         objExcel.Cells(V, H + 2) = ADO3!nombre
         objExcel.Cells(V, H + 3) = "1502"
         objExcel.Cells(V, H + 4) = "0001"
         objExcel.Cells(V, H + 5) = wUno
         objExcel.Cells(V, H + 6) = wDos
         objExcel.Cells(V, H + 7) = wTotDsc
         
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
   Db.Execute ("DELETE FROM DIECOCAB WHERE MES = '" + zAno + zMes + "' ")
   Db.CommitTrans
 
   Db.BeginTrans
   Db.Execute ("INSERT INTO DIECOCAB " _
   & " (MES, CODSOCIO, CODIGO, INS, FECENV, FECDSC, E_SOCIO, TIPCOB, CODPNP, INSPNP, " _
   & "   TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, " _
   & "   NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, " _
   & "   ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
   & "   DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
   & "   TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
   & "   CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
   & "   TOTENVIO, DSCDIECO, DSCDIFER ) " _
   & " SELECT " _
   & "  MES, CODSOCIO, CODIGO, INS, FECENV, FECDSC, E_SOCIO, TIPCOB, CODPNP, INSPNP, " _
   & "   TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, " _
   & "   NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, " _
   & "   ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
   & "   DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
   & "   TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
   & "   CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
   & "   TOTENVIO, DSCDIECO, DSCDIFER " _
   & " FROM TMP_DIECOCAB " _
   & " WHERE MES = '" + zAno + zMes + "' AND " _
   & "       USU = '" + wcodusu + "' ")
   Db.CommitTrans

   zz = Leerado8("SELECT * FROM CONTROL_ENVIO " _
                & " WHERE  MES = '" + zAno + zMes + "' AND " _
                & "       TIPO = '01' ")
   If zz = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO CONTROL_ENVIO " _
      & " (TIPO, MES, ESTADO) " _
      & " VALUES " _
      & " ('01', '" + zAno + zMes + "', 'A' ) ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE CONTROL_ENVIO " _
      & " SET ESTADO = 'A' " _
      & " WHERE TIPO = '01' AND " _
      & "        MES = '" + zAno + zMes + "' ")
      Db.CommitTrans
   End If
   
   Db.BeginTrans
   Db.Execute ("UPDATE MAESOCIO " _
   & " SET ENV_540 = T.TOTAPORT, ENV_541 = T.TOTADELA + T.TOTDEUDA + T.NETASIG1 + T.NETASIG2 + T.NETASIG3 + T.NETASIG4 + T.NETASIG5 " _
   & " FROM MAESOCIO AS M INNER JOIN TMP_DIECOCAB AS T " _
   & "   ON M.CODSOCIO = T.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' AND " _
   & "       T.MES = '" + zAno + zMes + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_MAESTRO " _
   & " SET ENV_540 = T.TOTAPORT, ENV_541 = T.TOTADELA + T.TOTDEUDA + T.NETASIG1 + T.NETASIG2 + T.NETASIG3 + T.NETASIG4 + T.NETASIG5 " _
   & " FROM MAESOCIO AS M INNER JOIN TMP_DIECOCAB AS T " _
   & "   ON M.CODSOCIO = T.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' AND " _
   & "       T.MES = '" + zAno + zMes + "' ")
   Db.CommitTrans

   wGrabar = False

   MsgBox "Proceso DIECO Grabado OK", vbExclamation
End Sub

Private Sub cmdSalir_Click()
   If wGrabar = True Then
      If MsgBox("NO Se Grabó Proceso Actual" + vbNewLine + "Desea Salir Sin Hacerlo??", vbCritical + vbQuestion + vbYesNo, "Envio DIECO") = vbYes Then
         
         Unload Me
      Else
         Call cmdGrabar_Click
         
         Unload Me
      End If
   Else
      Unload Me
   End If
End Sub

Private Sub DataGrid1_DblClick()
   zDetaCambio = False
   zDetaCodSoc = ADO2!codsocio
   zDetaTipDsc = "01"
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
   frmDiecoEnvio.Left = (Screen.Width - Width) \ 2
   frmDiecoEnvio.Top = 0
   
   txtAnoCab.Text = wanocia
   txtAnoCab.Enabled = False
   wGrabar = False
   
   Dim a As Integer
   cmbMeses.Clear
   a = Leerado("select * from CONTROL_ENVIO " _
            & " WHERE  MES LIKE '" + wanocia + "%' AND " _
            & "       TIPO = '01' " _
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
   Db.Execute ("DELETE FROM TMP_DIECOCAB WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECOCAB " _
   & " (MES     , CODSOCIO, CODIGO  , INS     , E_SOCIO, NOMBRE  , FECENV  , " _
   & "  FECDSC  , TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCDIECO, DSCSOCIO, DSCDIFER, TOTENVIO, " _
   & "  NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, " _
   & "  ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
   & "  DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
   & "  TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
   & "  DSCASIG1, DSCASIG2, DSCASIG3, DSCASIG4, DSCASIG5, " _
   & "  DIFASIG1, DIFASIG2, DIFASIG3, DIFASIG4, DIFASIG5, " _
   & "  CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
   & "  NOMASIG1, NOMASIG2, NOMASIG3, NOMASIG4, NOMASIG5, " _
   & "  TIPCOB  , CODPNP  , INSPNP  , USU ) " _
   & " SELECT " _
   & "   D.MES, D.CODSOCIO, M.CODIGO, M.INS, D.E_SOCIO, M.NOMBRE, D.FECENV, " _
   & "   D.FECDSC, TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCDIECO, DSCSOCIO, DSCDIFER, TOTENVIO, " _
   & "   NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, " _
   & "   ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
   & "   DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
   & "   TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
   & "   DSCASIG1, DSCASIG2, DSCASIG3, DSCASIG4, DSCASIG5, " _
   & "   DIFASIG1, DIFASIG2, DIFASIG3, DIFASIG4, DIFASIG5, " _
   & "   CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
   & "   ''      , ''      , ''      , ''      , ''      , D.TIPCOB, D.CODPNP, INSPNP, '" + wcodusu + "'  " _
   & " FROM DIECOCAB AS D INNER JOIN MAESOCIO AS M  ON D.CODSOCIO = M.CODSOCIO " _
   & " WHERE D.MES = '" + wAno + wMes + "'  ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG1 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS D INNER JOIN MAESOCIO AS M " _
   & "   ON D.CODASIG1 = M.CODSOCIO " _
   & " WHERE D.CODASIG1 > 0 AND " _
   & "       D.USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG2 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS D INNER JOIN MAESOCIO AS M " _
   & "   ON D.CODASIG2 = M.CODSOCIO " _
   & " WHERE D.CODASIG2 > 0 AND " _
   & "       D.USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG3 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS D INNER JOIN MAESOCIO AS M " _
   & "   ON D.CODASIG3 = M.CODSOCIO " _
   & " WHERE D.CODASIG3 > 0 AND " _
   & "       D.USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG4 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS D INNER JOIN MAESOCIO AS M " _
   & "   ON D.CODASIG4 = M.CODSOCIO " _
   & " WHERE D.CODASIG4 > 0 AND " _
   & "       D.USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG5 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS D INNER JOIN MAESOCIO AS M " _
   & "   ON D.CODASIG5 = M.CODSOCIO " _
   & " WHERE D.CODASIG5 > 0 AND " _
   & "       D.USU = '" + wcodusu + "' ")
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

   zz = Leerado2("SELECT CODSOCIO, CODIGO  , INS     , NOMBRE  , NETSOCIO, " _
                & "      NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, " _
                & "      TOTENVIO, " _
                & "      USU     , MES     , TOTAPORT, TOTDEUDA, TOTADELA, " _
                & "      DSCDIECO, DSCSOCIO, DSCDIFER, TIPCOB  , CODPNP  , INSPNP , " _
                & "      TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
                & "      ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
                & "      DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
                & "      DSCASIG1, DSCASIG2, DSCASIG3, DSCASIG4, DSCASIG5, " _
                & "      CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
                & "      NOMASIG1, NOMASIG2, NOMASIG3, NOMASIG4, NOMASIG5 " _
                & " FROM TMP_DIECOCAB " _
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
    
   DataGrid1.Columns(3).Width = 4400  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE ASOCIADO"
    
   DataGrid1.Columns(4).Width = 770    ' TOTSOCIO
   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(4).Caption = "T.SOCIO"
    
   DataGrid1.Columns(5).Width = 770    ' TOTASIG1
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(5).Caption = "ASIG 1"
    
   DataGrid1.Columns(6).Width = 770    ' TOTASIG2
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(6).Caption = "ASIG 2"
    
   DataGrid1.Columns(7).Width = 770    ' TOTASIG3
   DataGrid1.Columns(7).Alignment = dbgRight
   DataGrid1.Columns(7).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(7).Caption = "ASIG 3"
    
   DataGrid1.Columns(8).Width = 770    ' TOTASIG4
   DataGrid1.Columns(8).Alignment = dbgRight
   DataGrid1.Columns(8).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(8).Caption = "ASIG 4"
    
   DataGrid1.Columns(9).Width = 770    ' TOTASIG5
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
   DataGrid1.Columns(43).Visible = False
   DataGrid1.Columns(44).Visible = False
   DataGrid1.Columns(45).Visible = False
   DataGrid1.Columns(46).Visible = False
   DataGrid1.Columns(47).Visible = False
   DataGrid1.Columns(48).Visible = False
   DataGrid1.Columns(49).Visible = False
   DataGrid1.Columns(50).Visible = False
   DataGrid1.Columns(51).Visible = False
   
   DataGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
'   DataGrid1.SetFocus
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
   
   aa = Leerado2("SELECT CODSOCIO, CODIGO  , INS     , NOMBRE  , NETSOCIO, " _
                & "      NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, " _
                & "      TOTENVIO, " _
                & "      USU     , MES     , TOTAPORT, TOTDEUDA, TOTADELA, " _
                & "      DSCDIECO, DSCSOCIO, DSCDIFER, TIPCOB  , CODPNP  , INSPNP, " _
                & "      TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
                & "      ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
                & "      DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, " _
                & "      DSCASIG1, DSCASIG2, DSCASIG3, DSCASIG4, DSCASIG5, " _
                & "      CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, " _
                & "      NOMASIG1, NOMASIG2, NOMASIG3, NOMASIG4, NOMASIG5 " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE MES = '" + wAno + wMes + "' AND " _
                & "       USU = '" + wcodusu + "' AND " _
                & "       NOMBRE LIKE '%" + Trim(txtFiltrar.Text) + "%' " _
                & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO2
   
   LlenaCab1
   LabelCab
   TotalCab
End Sub

