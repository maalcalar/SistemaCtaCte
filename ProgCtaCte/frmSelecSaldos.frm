VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSelecSaldos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario de Ayuda"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   12765
   Begin VB.CheckBox chkSaldos 
      Caption         =   "Mostrar Solo Con Stock"
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
      Left            =   7920
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5175
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   9128
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox txtBuscar 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   6255
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
      TabIndex        =   6
      Top             =   6480
      Width           =   7215
   End
   Begin VB.Label Label1 
      Caption         =   "Filtro"
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
      Left            =   480
      TabIndex        =   3
      Top             =   300
      Width           =   855
   End
End
Attribute VB_Name = "frmSelecSaldos"
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
   xseleccion = ADOx!codigo
   xselecNombr = ADOx!nombre
'   xselecLote = ADOx!lote
   xselecUniMe = Trim(ADOx!um)
   xselecCanti = IIf(IsNull(ADOx!sdonew), 0, ADOx!sdonew)
   xselecCosto = IIf(IsNull(ADOx!cosnew), 0, ADOx!cosnew)
   Unload Me
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
   frmSelecSaldos.Left = (Screen.Width - Width) \ 2
   frmSelecSaldos.Top = 0
   
   LlenaCab
   txtBuscar.SetFocus
End Sub

Private Sub LlenaCab()
   Dim AA As Integer, wLin As Integer
   lblMensaje.Caption = "Buscando Articulos.... Espere"
   lblMensaje.Refresh
   
   Db.BeginTrans
   Db.Execute ("UPDATE MAEARTICULO " _
   & " SET ALMACEN = '" + wAyuAlm + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE MAEARTICULO " _
   & " SET SDONEW = S.SDONEW,  COSNEW = S.COSNEW, " _
   & "     TOTNEW = S.TOTNEW, ALMACEN = S.ALMACEN " _
   & " FROM MAEARTICULO AS M INNER JOIN SALDOS AS S " _
   & "   ON M.CODIGO = S.ARTICULO " _
   & " WHERE S.MES = '" + wAyuMes + "' AND " _
   & "       S.ALMACEN = '" + wAyuAlm + "' AND " _
   & "       S.ANO = '" + wanocia + "' AND " _
   & "       UPPER(M.NOMBRE) LIKE UPPER('%" + xseleccion + "%') ")
   Db.CommitTrans
         
   Db.BeginTrans
   Db.Execute ("UPDATE MAEARTICULO " _
   & " SET SALDOS = CONVERT(NUMERIC(10,3), SDONEW) " _
   & " WHERE (UPPER(NOMBRE) LIKE UPPER('%" + xseleccion + "%')) ")
   Db.CommitTrans
         
   Db.BeginTrans
   Db.Execute ("UPDATE MAEARTICULO " _
   & " SET NOMUM = U.NOMBRE " _
   & " FROM MAEARTICULO AS A INNER JOIN MAEUNIDAD AS U " _
   & "   ON A.UM = U.UNIDAD ")
   Db.CommitTrans
         
   If chkSaldos.Value = vbChecked Then
      If xseleccion = "" Then
         wsql = "SELECT ALMACEN, CODIGO, NOMBRE, SDONEW, NOMUM, COSNEW, TOTNEW, UM " _
                 & " FROM MAEARTICULO " _
                 & " WHERE SDONEW <> 0 AND ESTADO = 1 " _
                 & " ORDER BY NOMBRE"
      Else
         wsql = "SELECT ALMACEN, CODIGO, NOMBRE, SDONEW, NOMUM, COSNEW, TOTNEW, UM " _
                 & " FROM MAEARTICULO " _
                 & " WHERE UPPER(NOMBRE) LIKE UPPER('%" + xseleccion + "%') AND " _
                 & "       SDONEW <> 0 AND ESTADO = 1 " _
                 & " ORDER BY NOMBRE"
      End If
   Else
      If xseleccion = "" Then
         wsql = "SELECT ALMACEN, CODIGO, NOMBRE, SDONEW, NOMUM, COSNEW, TOTNEW, UM " _
                 & " FROM MAEARTICULO " _
                 & " WHERE  ESTADO = 1  " _
                 & " ORDER BY NOMBRE"
      Else
         wsql = "SELECT ALMACEN, CODIGO, NOMBRE, SDONEW, NOMUM, COSNEW, TOTNEW, UM " _
                 & " FROM MAEARTICULO " _
                 & " WHERE UPPER(NOMBRE) LIKE UPPER('%" + xseleccion + "%')  AND ESTADO = 1  " _
                 & " ORDER BY NOMBRE"
      End If
   End If
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
   DataGrid1.Columns(0).Width = 400
   DataGrid1.Columns(0).Alignment = dbgCenter
   DataGrid1.Columns(0).Caption = "ALM"
   
   DataGrid1.Columns(1).Width = 950
   DataGrid1.Columns(1).Alignment = dbgCenter
   DataGrid1.Columns(1).Caption = "ARTICULO"
    
   DataGrid1.Columns(2).Width = 5000
   DataGrid1.Columns(2).Alignment = dbgLeft
   DataGrid1.Columns(2).Caption = "NOMBRE ARTICULO"
    
   DataGrid1.Columns(3).Width = 1000
   DataGrid1.Columns(3).Alignment = dbgRight
   DataGrid1.Columns(3).Caption = "STOCK"
   DataGrid1.Columns(3).NumberFormat = "######0.000"
    
   DataGrid1.Columns(4).Width = 850
   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Caption = "UM"

   DataGrid1.Columns(5).Width = 1000
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).Caption = "COSTO"
   DataGrid1.Columns(5).NumberFormat = "######0.000"

   DataGrid1.Columns(6).Width = 1000
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).Caption = "TOTAL"
   DataGrid1.Columns(6).NumberFormat = "######0.00"

   DataGrid1.Columns(7).Visible = False
End Sub

Private Sub txtBuscar_Change()
   Dim AA As Integer
   
   If chkSaldos.Value = vbChecked Then
      If Len(Trim(txtBuscar.Text)) = 0 Then
         AA = Leeradox("SELECT ALMACEN, CODIGO, NOMBRE, SDONEW, NOMUM, COSNEW, TOTNEW, UM " _
                 & " FROM MAEARTICULO " _
                 & " WHERE SDONEW <> 0 " _
                 & " ORDER BY NOMBRE")
      Else
         AA = Leeradox("SELECT ALMACEN, CODIGO, NOMBRE, SDONEW, NOMUM, COSNEW, TOTNEW, UM " _
                 & " FROM MAEARTICULO " _
                 & " WHERE UPPER(NOMBRE) LIKE UPPER('%" + Trim(txtBuscar.Text) + "%') AND " _
                 & "       SDONEW <> 0 " _
                 & " ORDER BY NOMBRE")
      End If
   Else
      If Len(Trim(txtBuscar.Text)) > 0 Then
         AA = Leeradox("SELECT ALMACEN, CODIGO, NOMBRE, SDONEW, NOMUM, COSNEW, TOTNEW, UM " _
                & " FROM MAEARTICULO " _
                & " WHERE NOMBRE LIKE '%" + Trim(txtBuscar.Text) + "%' " _
                & " ORDER BY NOMBRE ")
      Else
         AA = Leeradox("SELECT ALMACEN, CODIGO, NOMBRE, SDONEW, NOMUM, COSNEW, TOTNEW, UM " _
                & " FROM MAEARTICULO " _
                & " ORDER BY NOMBRE ")
      End If
   End If
   Set DataGrid1.DataSource = ADOx
   LlenaCab1
End Sub

Private Sub txtBuscar_GotFocus()
   txtBuscar.SelStart = 0
   txtBuscar.SelLength = Len(Trim(txtBuscar.Text))
End Sub

Private Sub txtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 40
        DataGrid1.SetFocus
   End Select
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      DataGrid1.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

