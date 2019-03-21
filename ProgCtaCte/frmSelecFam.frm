VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSelecFam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de Familiares"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   12765
   Begin VB.TextBox txtCodSocio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   720
      MaxLength       =   9
      TabIndex        =   4
      Top             =   300
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5295
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   9340
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
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
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
      TabIndex        =   1
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Seleccionar"
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
      Left            =   9000
      TabIndex        =   0
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de Socio"
      Height          =   195
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblCodSocio 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   300
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "Cod.Socio"
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   120
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
      Left            =   240
      TabIndex        =   3
      Top             =   6480
      Width           =   7215
   End
End
Attribute VB_Name = "frmSelecFam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer, wsql As String

Private Sub chkSaldos_Click()
   LlenaCab
End Sub

Private Sub cmdCancela_Click()
    xseleccion = ""
    Unload Me
End Sub

Private Sub cmdSelect_Click()
   On Error GoTo err
   
   xseleccion = ADOx!nombre
   Unload Me
   Exit Sub
err:
   MsgBox err.Description
   Resume Next
End Sub

Private Sub DataGrid1_Click()
   cmdSelect_Click
End Sub

Private Sub DataGrid1_DblClick()
'   cmdSelect_Click
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo err
    
   Select Case KeyCode
   Case 13
        cmdSelect_Click
        If Trim(xseleccion) <> "" Then
           Unload Me
           Exit Sub
        End If
   End Select
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub Form_Activate()
   frmSelecFam.Left = (Screen.Width - Width) \ 2
   frmSelecFam.Top = 0
   
   LlenaCab
End Sub

Private Sub LlenaCab()
   Dim aa As Integer, wLin As Integer, _
       wSoc As String, wCod As Long, wIns As Integer
   
   lblMensaje.Caption = "Buscando Articulos.... Espere"
   lblMensaje.Refresh
   
   wSoc = zFamSocio
   
   aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
   If aa > 0 Then
      wCod = ADO8!codigo
      wIns = ADO8!ins
   End If
   Set ADO8 = Nothing
   
   wsql = "SELECT S.CODSOCIO, S.CODIGO, S.INS, F.TIPOPARIENTE, F.NOMBRE  " _
            & " FROM MAEFAMILIA AS F INNER JOIN MAESOCIO AS S " _
            & "   ON F.CODSOCIO = S.CODSOCIO " _
            & " WHERE S.CODSOCIO = " + Str(zFamSocio) + " "
   
'   wsql = "SELECT CODOFIN, INS, COD_PIP, CON_PIP, NOM_FAM  " _
'           & " FROM ZZZ_FAMILIA " _
'           & " WHERE CODOFIN = " + Str(wCod) + " AND " _
'           & "           INS = " + Str(wIns) + " AND " _
'           & "       CON_PIP LIKE '" + Trim(zFamParie) + "%' " _
'           & " ORDER BY NOM_FAM"
   
   lblMensaje.Caption = ""
   lblMensaje.Refresh
   
   a = Leeradox(wsql)
   If a > 0 Then
      cmdSelect.Enabled = True
   Else
      cmdSelect.Enabled = False
   End If
   Set DataGrid1.DataSource = ADOx
   DataGrid1.SetFocus
   LlenaCab1
   
End Sub

Private Sub LlenaCab1()
   DataGrid1.Columns(0).Width = 900
   DataGrid1.Columns(0).Alignment = dbgCenter
   DataGrid1.Columns(0).Caption = "COD.SOCIO"
   
   DataGrid1.Columns(1).Width = 950
   DataGrid1.Columns(1).Alignment = dbgCenter
   DataGrid1.Columns(1).Caption = "CODOFIN"
    
   DataGrid1.Columns(2).Width = 450
   DataGrid1.Columns(2).Alignment = dbgCenter
   DataGrid1.Columns(2).Caption = "INS"
    
   DataGrid1.Columns(3).Width = 750
   DataGrid1.Columns(3).Alignment = dbgCenter
   DataGrid1.Columns(3).Caption = "TIPO"
    
   DataGrid1.Columns(4).Width = 4700
   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Caption = "NOMBRE"
End Sub

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
