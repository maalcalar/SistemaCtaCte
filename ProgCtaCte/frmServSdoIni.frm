VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmServSdoIni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar Saldo Inicial Oct 2017"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   15375
   Begin VB.Frame Frame2 
      Caption         =   "Crear Nuevo Socio"
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   7320
      TabIndex        =   27
      Top             =   7800
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtNuevo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         MaxLength       =   9
         TabIndex        =   28
         Top             =   440
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Apellidos y Nombres"
         Height          =   195
         Left            =   1680
         TabIndex        =   31
         Top             =   240
         Width           =   3420
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Cod.Socio"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   255
         Width           =   900
      End
      Begin VB.Label lblNuevo 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   29
         Top             =   440
         Width           =   5175
      End
   End
   Begin VB.CommandButton cmdCrear 
      Caption         =   "&Traer Nuevo Socio"
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
      TabIndex        =   26
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Socio"
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   240
      TabIndex        =   14
      Top             =   6960
      Width           =   6975
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
         Left            =   5160
         TabIndex        =   25
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtCodofin 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   23
         Top             =   840
         Width           =   930
      End
      Begin VB.TextBox txtIns 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         MaxLength       =   1
         TabIndex        =   22
         Top             =   840
         Width           =   330
      End
      Begin VB.TextBox txtAdelan 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   20
         Top             =   1200
         Width           =   930
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   18
         Top             =   1200
         Width           =   930
      End
      Begin VB.TextBox txtCodSocio 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codofin"
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
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Adelanto"
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
         Left            =   3030
         TabIndex        =   21
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sdo x Cobrar"
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
         Index           =   13
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1140
      End
      Begin VB.Label lblCodSocio 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   16
         Top             =   480
         Width           =   1140
      End
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtro x Nombre"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   6240
      Width           =   8295
      Begin VB.OptionButton optTodos 
         Caption         =   "Mostrar Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Filtrar x Nombre"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtFiltrar 
         Height          =   285
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   8
         Top             =   240
         Width           =   5055
      End
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
      Left            =   14040
      TabIndex        =   6
      Top             =   8160
      Width           =   1215
   End
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmServSdoIni.frx":0000
      Left            =   120
      List            =   "frmServSdoIni.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   7215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4815
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   8493
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
   Begin MSMask.MaskEdBox txtMes 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   820
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblSdoNew 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   12720
      TabIndex        =   13
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lblCargos 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   10560
      TabIndex        =   12
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lblAbonos 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   11640
      TabIndex        =   11
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Mes Inicial"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   620
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Esta Opción Se Utiliza para Modificar Los Saldos Iniciales que se migran desde el Sistema FOXPRO. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   7560
      TabIndex        =   2
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label Label25 
      Caption         =   "Compañia"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmServSdoIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Limpiar()
   txtCodSocio.Text = ""
   lblCodSocio.Caption = ""
   txtCodofin.Text = ""
   txtIns.Text = ""
   txtSaldo.Text = ""
   txtAdelan.Text = ""
End Sub

Private Sub refrescar()
   txtCodSocio.Text = ADO2!codsocio
   lblCodSocio.Caption = ADO2!nombre
   txtCodofin.Text = ADO2!codigo
   txtIns.Text = ADO2!ins
   txtSaldo.Text = Format(ADO2!cargos, "#####0.00;;\ ")
   txtAdelan.Text = Format(ADO2!abonos, "#####0.00;;\ ")
End Sub

Private Sub cmdCrear_Click()
   Frame2.Visible = True
   
   txtNuevo.SetFocus
End Sub

Private Sub cmdGrabar_Click()
   Dim aa As Long, wMes As String, wSoc As Integer, _
       wCargos As Currency, wAbonos As Currency, wSdoNew As Currency, _
       wOldCar As Currency, wOldAbo As Currency, wOldSdo As Currency, _
       wMon As String, WE_S As String
   
   wMes = txtMes.Text
   wSoc = ADO2!codsocio
   wCargos = 0: wAbonos = 0
   wCargos = Val(txtSaldo.Text)
   wAbonos = Val(txtAdelan.Text)
   wSdoNew = wCargos - wAbonos

   If wCargos <> 0 And wAbonos <> 0 Then
      MsgBox "Saldo Inicial NO Puede Tener Saldo por Cobrar y Adelanto a la Vez", vbExclamation
      Exit Sub
   End If


   aa = Leerado7a("SELECT * FROM SDOINI " _
                & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                & "            MES = '" + wMes + "' AND " _
                & "       CONCEPTO = '01' ")
   If aa = 0 Then
      wMon = "S"
      WE_S = ""
      wOldCar = 0: wOldAbo = 0: wOldSdo = 0
      aa = Leerado6a("SELECT E_SOCIO, MONEDA, " _
                    & " CARGOS, ABONOS, SDONEW " _
                    & " FROM CTASXCAB " _
                    & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                    & "            MES = '" + wMes + "' AND " _
                    & "       CONCEPTO = '01' ")
      If aa > 0 Then
         WE_S = ADO6a!e_socio
         wMon = ADO6a!moneda
         wOldCar = ADO6a!cargos
         wOldAbo = ADO6a!abonos
         wOldSdo = ADO6a!sdonew
      End If
      Set ADO6a = Nothing
      
      Db.BeginTrans
      Db.Execute ("INSERT INTO SDOINI " _
      & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, " _
      & "  CARGOS, ABONOS, SDONEW, " _
      & "  CARGOSOLD, ABONOSOLD, SDONEWOLD) " _
      & " VALUES " _
      & " (" + Str(wSoc) + ", '" + wMes + "', '01', " _
      & "  '" + WE_S + "', '" + wMon + "', " _
      & "  " + Str(wCargos) + ", " + Str(wAbonos) + ", " + Str(wSdoNew) + ", " _
      & "  " + Str(wOldCar) + ", " + Str(wOldAbo) + ", " + Str(wOldSdo) + " ) ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE SDOINI " _
      & " SET CARGOS = " + Str(wCargos) + ", " _
      & "     ABONOS = " + Str(wAbonos) + ", " _
      & "     SDONEW = " + Str(wSdoNew) + " " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "            MES = '" + wMes + "' AND " _
      & "       CONCEPTO = '01' ")
      Db.CommitTrans
   End If
   
   
   Db.BeginTrans
   Db.Execute ("UPDATE CTASXCAB " _
   & " SET CARGOS = " + Str(wCargos) + ", " _
   & "     ABONOS = " + Str(wAbonos) + ", " _
   & "     SDONEW = " + Str(wCargos - wAbonos) + " " _
   & " WHERE      MES = '" + wMes + "' AND " _
   & "       CODSOCIO = " + Str(wSoc) + " AND " _
   & "       CONCEPTO = '01' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE CTASXDET " _
   & " SET CARGOS = " + Str(wCargos) + ", " _
   & "     ABONOS = " + Str(wAbonos) + ", " _
   & "     SDONEW = " + Str(wCargos - wAbonos) + " " _
   & " WHERE      MES = '" + wMes + "' AND " _
   & "       CODSOCIO = " + Str(wSoc) + " AND " _
   & "       CONCEPTO = '01' AND " _
   & "         TIPCOB = '00' ")
   Db.CommitTrans
   
   Call ActualizaSaldos(wSoc, wMes, "01")
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_CTASXCAB " _
   & " SET CARGOS = " + Str(wCargos) + ", " _
   & "     ABONOS = " + Str(wAbonos) + ", " _
   & "     SDONEW = " + Str(wSdoNew) + " " _
   & " WHERE      MES = '" + wMes + "' AND " _
   & "       CODSOCIO = " + Str(wSoc) + " AND " _
   & "       CONCEPTO = '01' AND " _
   & "            USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   
   ADO2.Requery
   LlenaCab1
   ADO2.Find "CODSOCIO = " + Str(wSoc) + " "
   
   MsgBox "Saldo Inicial Socio " + Trim(Str(wSoc)) + " Grabados OK", vbExclamation
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub DataGrid1_GotFocus()
   DataGrid1.col = 8
   DataGrid1.SelStart = 0
   If Len(Trim(DataGrid1.Text)) > 0 Then
      DataGrid1.SelLength = Len(Trim(DataGrid1.Text))
   End If
   DataGrid1.Refresh
   Limpiar
   refrescar
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim wvariable As String
    
    On Error GoTo err
    Select Case KeyCode
    Case 40  ' DOWN
            
         wvariable = DataGrid1.Text
         
         Select Case DataGrid1.col
         Case 8
              ADO2!cargos = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 9
              ADO2!abonos = IIf(IsNull(wvariable), 0, Val(wvariable))
         End Select
    
         Select Case DataGrid1.col
         Case 8
              DataGrid1.Text = IIf(IsNull(ADO2!cargos), 0, ADO2!cargos)
         Case 9
              DataGrid1.Text = IIf(IsNull(ADO2!abonos), 0, ADO2!abonos)
         End Select
            
    Case 37 ' Retroceder
         wvariable = DataGrid1.Text
         
         Select Case DataGrid1.col
         Case 8
              ADO2!cargos = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 9
              ADO2!abonos = IIf(IsNull(wvariable), 0, Val(wvariable))
         End Select
         
         If DataGrid1.col = 1 Then
            If DataGrid1.Row > 0 Then
               DataGrid1.Row = DataGrid1.Row - 1
            End If
            DataGrid1.col = 0
         End If
         
         Select Case DataGrid1.col
         Case 8
              DataGrid1.Text = IIf(IsNull(ADO2!cargos), 0, ADO2!cargos)
         Case 9
              DataGrid1.Text = IIf(IsNull(ADO2!abonos), 0, ADO2!abonos)
         End Select
         
    Case 38 ' Subir
         wvariable = DataGrid1.Text
         
         Select Case DataGrid1.col
         Case 8
              ADO2!cargos = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 9
              ADO2!abonos = IIf(IsNull(wvariable), 0, Val(wvariable))
         End Select
         
         Select Case DataGrid1.col
         Case 8
              DataGrid1.Text = IIf(IsNull(ADO2!cargos), 0, ADO2!cargos)
         Case 9
              DataGrid1.Text = IIf(IsNull(ADO2!abonos), 0, ADO2!abonos)
         End Select
    
    Case 39 ' Avanzar
         wvariable = DataGrid1.Text
         
         Select Case DataGrid1.col
         Case 8
              ADO2!cargos = IIf(IsNull(wvariable), 0, Val(wvariable))
         Case 9
              ADO2!abonos = IIf(IsNull(wvariable), 0, Val(wvariable))
         End Select
         
         If DataGrid1.col = 9 Then
            If Not ADO2.EOF Then
               DataGrid1.Row = DataGrid1.Row + 1
            End If
            DataGrid1.col = 0
         End If
          
         Select Case DataGrid1.col
         Case 8
              DataGrid1.Text = IIf(IsNull(ADO2!cargos), 0, ADO2!cargos)
         Case 9
              DataGrid1.Text = IIf(IsNull(ADO2!abonos), 0, ADO2!abonos)
         End Select
    
    End Select
    Exit Sub
err:
    MsgBox Format(err.Number, "00000000000") + " " + err.Description
    Resume Next

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    Dim c As Integer
    Dim wvariable As String, wvariable2 As String, wlll As Integer, wlll2 As Integer
    Dim wvariaold As String
    Dim wSoles As Currency, zTdo As String, zSer As String, zDoc As String, _
        wOld As Currency, wlinold As Integer, _
        waaa As String, wmmm As String
    
    On Error GoTo err
    Select Case KeyAscii
    Case 13
       Select Case DataGrid1.col
       Case 0  ' CodSocio
            DataGrid1.col = 8
       Case 1  ' Codigo
            DataGrid1.col = 8
       Case 2  ' INS
            DataGrid1.col = 8
       Case 3  ' Nombre
            DataGrid1.col = 8
       Case 4  ' MesCob
            DataGrid1.col = 8
       Case 5  ' NomCon
            DataGrid1.col = 8
       Case 6  ' E_Socio
            DataGrid1.col = 8
       Case 7  ' Moneda
            DataGrid1.col = 8
       Case 8  ' Cargos
            wvariable = Trim(DataGrid1.Text)
            DataGrid1.Text = wvariable
            ADO2!cargos = IIf(IsNull(wvariable) Or Len(Trim(wvariable)) = 0, 0, wvariable)
            ADO2!abonos = 0
            ADO2!sdonew = ADO2!cargos - ADO2!abonos
       
       Case 9  ' Abonos
            wvariable = Trim(DataGrid1.Text)
            DataGrid1.Text = wvariable
            ADO2!abonos = IIf(IsNull(wvariable) Or Len(Trim(wvariable)) = 0, 0, wvariable)
            ADO2!cargos = 0
            ADO2!sdonew = ADO2!cargos - ADO2!abonos
       
       Case 10 ' MesCob
            DataGrid1.col = 8
       
       End Select
       wvariable2 = IIf(IsNull(ADO2.Fields(DataGrid1.col)), "", Trim(ADO2.Fields(DataGrid1.col)))
       DataGrid1.Text = wvariable2
       ADO2.Update
       DataGrid1.Refresh
    End Select
    Exit Sub
err:
    MsgBox Format(err.Number, "00000000000") + " " + err.Description
    Resume Next
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim wvariable As String
    
   On Error GoTo err
   Select Case KeyCode
   Case 37  ' RETROCEDER
        If DataGrid1.col = 7 Then
           DataGrid1.col = 9
        End If
         
        Select Case DataGrid1.col
        Case 8
             DataGrid1.Text = IIf(IsNull(ADO2!cargos), 0, ADO2!cargos)
        Case 9
             DataGrid1.Text = IIf(IsNull(ADO2!abonos), 0, ADO2!abonos)
        End Select
   
   Case 38  ' UP
   
   Case 39  ' AVANZAR

        If DataGrid1.col = 10 Then
           DataGrid1.col = 8
        End If
          
        Select Case DataGrid1.col
        Case 8
             DataGrid1.Text = IIf(IsNull(ADO2!cargos), 0, ADO2!cargos)
        Case 9
             DataGrid1.Text = IIf(IsNull(ADO2!abonos), 0, ADO2!abonos)
        End Select
        
   Case 40  ' DOWN
   
   End Select
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   Limpiar
   refrescar
End Sub

Private Sub Form_Activate()
   Form_Initialize
End Sub

Private Sub Form_Initialize()
   frmServSdoIni.Left = (Screen.Width - Width) \ 2
   frmServSdoIni.Top = 0
   
   Dim a As Integer
   
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

   txtMes.Text = "2017/09"
   txtMes.Enabled = False
   
   LlenaCab
   LlenaCab1
   TotalCab
   Limpiar
   refrescar
   
   DataGrid1.AllowAddNew = False
   DataGrid1.AllowDelete = False
   DataGrid1.AllowUpdate = False
   DataGrid1.SetFocus
End Sub

Private Sub LlenaCab()
   Dim aa As Long, _
       wcon As String, WE_S As String, wMes As String, _
       wSoc As Integer, sw As String
   
   Set DataGrid1.DataSource = Nothing
   
   wMes = txtMes.Text
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_CTASXCAB WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   If Len(Trim(wMes)) <> 0 Then
      If Len(Trim(sw)) = 0 Then
         sw = "WHERE "
      End If
      sw = sw + "C.MES = '" + wMes + "'"
   End If
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_CTASXCAB " _
   & " (CODSOCIO, CODIGO, INS, MES, CONCEPTO, NOMBRE, NOMCON, E_SOCIO, MONEDA, " _
   & "  CARGOS, ABONOS, SDONEW, USU ) " _
   & " SELECT " _
   & "  D.CODSOCIO, S.CODIGO, S.INS, D.MES, D.CONCEPTO, s.nombre, m.nombre, s.E_SOCIO, " _
   & "  C.MONEDA, D.CARGOS, D.ABONOS, D.SDONEW, '" + wcodusu + "' " _
   & " from ctasxdet as D inner join MAESOCIO as S on d.CODSOCIO = S.CODSOCIO " _
   & "                    inner join MAECONCEPTO AS M on d.CONCEPTO = m.concepto " _
   & "                    inner join CTASXCAB as C ON D.CODSOCIO = C.CODSOCIO AND D.MES = C.MES AND D.CONCEPTO = C.CONCEPTO " _
   & " WHERE D.mes = '" + wMes + "' and d.concepto = '01' and D.TIPCOB ='00'")
   Db.CommitTrans

'   Db.BeginTrans
'   Db.Execute ("INSERT INTO TMP_CTASXCAB " _
'   & " (CODSOCIO, CODIGO, INS, MES, CONCEPTO, NOMBRE, NOMCON, E_SOCIO, MONEDA, " _
'   & "  CARGOS, ABONOS, SDONEW, USU ) " _
'   & " SELECT " _
'   & "  C.CODSOCIO, S.CODIGO, S.INS, C.MES, C.CONCEPTO, S.NOMBRE, M.NOMBRE, S.E_SOCIO, " _
'   & "  C.MONEDA, C.CARGOS, C.ABONOS, C.SDONEW, '" + wcodusu + "'  " _
'   & " FROM CTASXCAB AS C INNER JOIN MAECONCEPTO AS M ON C.CONCEPTO = M.CONCEPTO " _
'   & "                    INNER JOIN MAESOCIO    AS S ON C.CODSOCIO = S.CODSOCIO " _
'   & " " + sw + "")
'   Db.CommitTrans

   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, MES, NOMCON, E_SOCIO, " _
            & "          MONEDA, CARGOS, ABONOS, SDONEW, USU, " _
            & "          CONCEPTO, CODSOCIO " _
            & " FROM TMP_CTASXCAB " _
            & " WHERE USU = '" + wcodusu + "' " _
            & " ORDER BY NOMBRE, MES, CONCEPTO ")
   Set DataGrid1.DataSource = ADO2

   DataGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
End Sub
   
Private Sub LlenaCab1()
   DataGrid1.Columns(0).Width = 700   ' CODSOCIO
   DataGrid1.Columns(0).Alignment = dbgLeft
   DataGrid1.Columns(0).Caption = "COD.SOCIO"
    
   DataGrid1.Columns(1).Width = 1000  ' CODIGO
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Caption = "CODIGO"
    
   DataGrid1.Columns(2).Width = 400   ' INS
   DataGrid1.Columns(2).Alignment = dbgLeft
   DataGrid1.Columns(2).Caption = "INS"
    
   DataGrid1.Columns(3).Width = 4150  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE"
    
   DataGrid1.Columns(4).Width = 750   ' MESCOB
   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Caption = "MES"
    
   DataGrid1.Columns(5).Width = 2500  ' NOMCON
   DataGrid1.Columns(5).Alignment = dbgLeft
   DataGrid1.Columns(5).Caption = "CONCEPTO"
    
   DataGrid1.Columns(6).Width = 500   ' E_SOCIO
   DataGrid1.Columns(6).Alignment = dbgLeft
   DataGrid1.Columns(6).Caption = "E_SOCIO"
    
   DataGrid1.Columns(7).Width = 350   ' MONEDA
   DataGrid1.Columns(7).Alignment = dbgCenter
   DataGrid1.Columns(7).Caption = "MON"
    
   DataGrid1.Columns(8).Width = 850   ' CARGOS
   DataGrid1.Columns(8).Alignment = dbgRight
   DataGrid1.Columns(8).Caption = "CARGOS"
   DataGrid1.Columns(8).NumberFormat = "###,##0.00;-###,##0.00;\ "
    
   DataGrid1.Columns(9).Width = 850   ' ABONOS
   DataGrid1.Columns(9).Alignment = dbgRight
   DataGrid1.Columns(9).Caption = "ABONOS"
   DataGrid1.Columns(9).NumberFormat = "###,##0.00;-###,##0.00;\ "
    
   DataGrid1.Columns(10).Width = 850   ' SDONEW
   DataGrid1.Columns(10).Alignment = dbgRight
   DataGrid1.Columns(10).Caption = "SALDOS"
   DataGrid1.Columns(10).NumberFormat = "###,##0.00;-###,##0.00;###,##0.00"
    
   DataGrid1.Columns(11).Visible = False
   DataGrid1.Columns(12).Visible = False
   DataGrid1.Columns(13).Visible = False
  
   DataGrid1.Columns(0).Locked = True
   DataGrid1.Columns(1).Locked = True
   DataGrid1.Columns(2).Locked = True
   DataGrid1.Columns(3).Locked = True
   DataGrid1.Columns(4).Locked = True
   DataGrid1.Columns(5).Locked = True
   DataGrid1.Columns(6).Locked = True
   DataGrid1.Columns(7).Locked = True
   DataGrid1.Columns(10).Locked = True
End Sub

Private Sub TotalCab()
   Dim zz As Integer, _
       wCargos As Currency, wAbonos As Currency, wSdoNew As Currency
   
   wCargos = 0: wAbonos = 0: wSdoNew = 0
   zz = Leerado7a("SELECT SUM(CARGOS) AS CARGOS, SUM(ABONOS) AS ABONOS, " _
                & "       SUM(SDONEW) AS SDONEW " _
                & " FROM TMP_CTASXCAB " _
                & " WHERE USU = '" + wcodusu + "' ")
   If zz > 0 Then
      wCargos = IIf(IsNull(ADO7a!cargos), 0, ADO7a!cargos)
      wAbonos = IIf(IsNull(ADO7a!abonos), 0, ADO7a!abonos)
      wSdoNew = IIf(IsNull(ADO7a!sdonew), 0, ADO7a!sdonew)
   End If
   Set ADO7a = Nothing
   
   lblCargos.Caption = Format(wCargos, "####,##0.00;;\ ")
   lblAbonos.Caption = Format(wAbonos, "####,##0.00;;\ ")
   lblSdoNew.Caption = Format(wSdoNew, "####,##0.00;;\ ")
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

Private Sub txtAdelan_GotFocus()
   txtAdelan.SelStart = 0
   txtAdelan.SelLength = Len(Trim(txtAdelan.Text))
End Sub

Private Sub txtAdelan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtAdelan.Text) <> 0 Then
         txtAdelan.Text = Format(txtAdelan.Text, "#####0.00;;\ ")
         txtSaldo.Text = ""
      End If
      cmdGrabar.SetFocus
   Else
      If InStr(1, "0123456789." + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFiltrar_Change()
   Dim a As Long

   a = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, MES, NOMCON, E_SOCIO, " _
            & "          MONEDA, CARGOS, ABONOS, SDONEW, USU, " _
            & "          CONCEPTO, CODSOCIO " _
            & " FROM TMP_CTASXCAB " _
            & " WHERE USU = '" + wcodusu + "' AND " _
            & "       NOMBRE LIKE '" + Trim(txtFiltrar.Text) + "%' " _
            & " ORDER BY NOMBRE, MES, CONCEPTO ")
   Set DataGrid1.DataSource = ADO2
   LlenaCab1
End Sub

Private Sub txtFiltrar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ADO2.RecordCount > 0 Then
         DataGrid1.SetFocus
      End If
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtNuevo_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtNuevo.Text)) + " ")
   If aa > 0 Then
      lblNuevo.Caption = ADO6a!nombre
   Else
      lblNuevo.Caption = ""
   End If
   Set ADO6a = Nothing
End Sub

Private Sub txtNuevo_GotFocus()
   txtNuevo.SelStart = 0
   If Len(Trim(txtNuevo.Text)) > 0 Then
      txtNuevo.SelLength = Len(Trim(txtNuevo.Text))
   Else
      txtNuevo.SelLength = 8
   End If
End Sub

Private Sub txtNuevo_KeyDown(KeyCode As Integer, Shift As Integer)
   
   Dim KeyNew As Integer
   KeyNew = Asc(UCase(Chr(KeyCode)))
   
   Select Case KeyCode
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtNuevo.Text = xseleccion
        End If
   Case 65 To 90
        xlista = "1"
        xseleccion = Chr(KeyNew)
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtNuevo.Text = xseleccion
        End If
          
   End Select
End Sub

Private Sub txtNuevo_KeyPress(KeyAscii As Integer)
   Dim aa As Integer, wSoc As Integer
   If KeyAscii = 13 Then
      If Len(Trim(txtNuevo.Text)) = 0 Then
         MsgBox "Codigo Socio En Blanco", vbExclamation
         txtNuevo.Text = ""
         Exit Sub
      End If
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtNuevo.Text)) + " ")
      If aa = 0 Then
         MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
         txtNuevo.Text = ""
         Exit Sub
      End If
      Set ADO8 = Nothing
      
      aa = Leerado8("SELECT * FROM TMP_CTASXCAB WHERE USU = '" + wcodusu + "' AND CODSOCIO = " + Str(Val(txtNuevo.Text)) + " ")
      If aa > 0 Then
         MsgBox "Codigo Socio Ya Existe En Saldo Inicial", vbExclamation
         txtNuevo.Text = ""
         Exit Sub
      End If
      Set ADO8 = Nothing
      
      wSoc = Val(txtNuevo.Text)
      
      aa = Leerado8("SELECT * FROM TMP_CTASXCAB " _
                & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                & "            MES = '2017/09' AND " _
                & "       CONCEPTO = '01' AND " _
                & "            USU = '" + wcodusu + "' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO TMP_CTASXCAB " _
         & " (CODSOCIO, CODIGO, INS, MES, CONCEPTO, NOMBRE, NOMCON, E_SOCIO, MONEDA, " _
         & "  CARGOS, ABONOS, SDONEW, USU ) " _
         & " SELECT " _
         & "  M.CODSOCIO, M.CODIGO, M.INS, '2017/09', '01', M.NOMBRE, 'APORTACION MENSUAL', M.E_SOCIO, " _
         & "  E.MONEDA, 0, 0, 0, '" + wcodusu + "'  " _
         & " FROM MAESOCIO AS M INNER JOIN MAEE_SOCIO AS E " _
         & "   ON M.E_SOCIO = E.E_SOCIO " _
         & " WHERE M.CODSOCIO = " + Str(wSoc) + " ")
         Db.CommitTrans
      End If
      
      aa = Leerado8("SELECT * FROM CTASXCAB " _
                & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                & "            MES = '2017/09' AND " _
                & "       CONCEPTO = '01' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO CTASXCAB " _
         & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, " _
         & "  CARGOS, ABONOS, SDONEW ) " _
         & " SELECT " _
         & "  M.CODSOCIO, '2017/09', '01', M.E_SOCIO, " _
         & "  E.MONEDA, 0, 0, 0 " _
         & " FROM MAESOCIO AS M INNER JOIN MAEE_SOCIO AS E " _
         & "   ON M.E_SOCIO = E.E_SOCIO " _
         & " WHERE M.CODSOCIO = " + Str(wSoc) + " ")
         Db.CommitTrans
      End If
      
      aa = Leerado8("SELECT * FROM CTASXDET " _
                & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                & "            MES = '2017/09' AND " _
                & "       CONCEPTO = '01' AND " _
                & "         TIPCOB = '00' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO CTASXDET " _
         & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, TIPCAM, " _
         & "  DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW, OBS ) " _
         & " SELECT " _
         & "  CODSOCIO, '2017/09', '01', '00', '', '', '', '1', '01/09/2017', 0,  " _
         & "  0, 0, 0, 0, 0, 0, '' " _
         & " FROM MAESOCIO " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE CTASXDET " _
         & " SET CARGOS = " + Str(wSdoOld) + ", " _
         & "     ABONOS = " + Str(wAdelan) + " " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "            MES = '2017/09' AND " _
         & "       CONCEPTO = '01' AND " _
         & "         TIPCOB = '00' ")
         Db.CommitTrans
      End If
      
      Db.BeginTrans
      Db.Execute ("INSERT INTO SDOINI " _
      & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, CARGOS, ABONOS, SDONEW, " _
      & "  CARGOSOLD, ABONOSOLD, SDONEWOLD ) " _
      & " SELECT " _
      & "  M.CODSOCIO, '2017/09', '01', M.E_SOCIO, E.MONEDA, 0, 0, 0, 0, 0, 0   " _
      & " FROM MAESOCIO AS M INNER JOIN MAEE_SOCIO AS E " _
      & "   ON M.E_SOCIO = E.E_SOCIO " _
      & " WHERE M.CODSOCIO = " + Str(wSoc) + " ")
      Db.CommitTrans
      
      ADO2.Requery
      LlenaCab1
      TotalCab
      
      Frame2.Visible = False
      DataGrid1.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
'         KeyAscii = 0
'      End If
   End If
End Sub

Private Sub txtSaldo_GotFocus()
   txtSaldo.SelStart = 0
   txtSaldo.SelLength = Len(Trim(txtSaldo.Text))
End Sub

Private Sub txtSaldo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtSaldo.Text) <> 0 Then
         txtSaldo.Text = Format(txtSaldo.Text, "#####0.00;;\ ")
         txtAdelan.Text = ""
         cmdGrabar.SetFocus
      Else
         txtAdelan.SetFocus
      End If
   Else
      If InStr(1, "0123456789." + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub
