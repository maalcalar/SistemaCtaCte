VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAsigFam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Familiares Para Descuento"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   11820
   Begin VB.Frame Frame2 
      Caption         =   "Consultas"
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
      Height          =   735
      Left            =   9000
      TabIndex        =   36
      Top             =   1920
      Width           =   2655
      Begin VB.CommandButton cmdExporta 
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
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtro x Nombre"
      Height          =   615
      Left            =   1680
      TabIndex        =   20
      Top             =   7800
      Width           =   8175
      Begin VB.OptionButton optTodos 
         Caption         =   "Mostrar Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Filtrar x Nombre"
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtFiltrar 
         Height          =   285
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   21
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.CommandButton cmdMigrar 
      Caption         =   "&Migrar"
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
      Left            =   360
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7920
      Width           =   975
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
      Left            =   10440
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   8040
      Width           =   975
   End
   Begin VB.Frame fraDesplaza 
      Caption         =   "Desplazamiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   9000
      TabIndex        =   12
      Top             =   2640
      Width           =   2655
      Begin VB.CommandButton cmdMover 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1920
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdMover 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdMover 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdMover 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraMantenimiento 
      Caption         =   "Mantenimiento"
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
      Height          =   1815
      Left            =   9000
      TabIndex        =   6
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton cmdDeshacer 
         Caption         =   "&Deshacer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   840
         Width           =   975
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
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraDetalle 
      Caption         =   "Detalle del Registro"
      ForeColor       =   &H00C00000&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VB.ComboBox cmbEstado 
         Height          =   315
         ItemData        =   "frmAsigFam.frx":0000
         Left            =   5640
         List            =   "frmAsigFam.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   880
         Width           =   2535
      End
      Begin VB.TextBox txtObserv 
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   30
         Top             =   1340
         Width           =   2655
      End
      Begin VB.TextBox txtCodHijo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   9
         TabIndex        =   26
         Top             =   880
         Width           =   975
      End
      Begin VB.TextBox txtLin 
         Height          =   285
         Left            =   5880
         MaxLength       =   2
         TabIndex        =   24
         Top             =   420
         Width           =   375
      End
      Begin VB.TextBox txtCodSocio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   9
         TabIndex        =   1
         Top             =   420
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtFectop 
         Height          =   285
         Left            =   2760
         TabIndex        =   32
         Top             =   1340
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Estado Asignado"
         Height          =   195
         Index           =   9
         Left            =   5865
         TabIndex        =   35
         Top             =   700
         Width           =   1200
      End
      Begin VB.Label Label14 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   2760
         TabIndex        =   33
         Top             =   1160
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1160
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Cod.Hijo"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   700
         Width           =   975
      End
      Begin VB.Label lblCodHijo 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   880
         Width           =   4575
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre del Hijo"
         Height          =   195
         Left            =   1200
         TabIndex        =   27
         Top             =   700
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Lin"
         Height          =   195
         Left            =   5880
         TabIndex        =   25
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre de Socio"
         Height          =   195
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblCodSocio 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   420
         Width           =   4575
      End
      Begin VB.Label Label3 
         Caption         =   "Cod.Socio"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "RELACION DE FAMILIARES  ASIGNADOS"
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
      Left            =   720
      TabIndex        =   19
      Top             =   8520
      Width           =   5775
   End
End
Attribute VB_Name = "frmAsigFam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Limpiar()
   txtCodSocio.Text = ""
   lblCodSocio.Caption = ""
   txtCodHijo.Text = ""
   lblCodHijo.Caption = ""
   txtLin.Text = ""
   txtObserv.Text = ""
   txtFectop.Text = "__/__/____"

   cmbEstado.ListIndex = 0

End Sub

Sub refrescar()
   If Not ADO1.BOF And Not ADO1.EOF Then
      txtCodSocio.Text = ADO1!codsocio
      txtCodHijo.Text = ADO1!codhijo
      
      txtLin.Text = ADO1!lin
      txtObserv.Text = ADO1!observ
      If IsDate(ADO1!fectop) Then
         txtFectop.Text = ADO1!fectop
      Else
         txtFectop.Text = "__/__/____"
      End If
   
      cmbEstado.ListIndex = BuscaEstadoAsignado(ADO1!estado)
   End If
End Sub

Sub grabar()
   On Error GoTo err
   
   Dim aa As Integer, wCod As Integer, wNom As String, wLin As String, _
       wNomEst As String, wEstado As String
   wCod = Val(txtCodSocio.Text)
   wNom = Trim(lblCodSocio.Caption)
   wLin = txtLin.Text
   wNomEst = cmbEstado.Text
   wEstado = BuscaCodEstadoAsignado(cmbEstado.List(cmbEstado.ListIndex))
   
   If validaAsi Then
      MsgBox "Asignación de Socio Esta Con Errores, No Se Graba", vbExclamation
      Exit Sub
   End If
   
   aa = Leerado8("SELECT * FROM MAEASIGNADO " _
                & " WHERE CODSOCIO = " + Str(wCod) + " AND LIN = '" + wLin + "' ")
   If aa = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO MAEASIGNADO " _
      & " (CODSOCIO, LIN, CODHIJO, ESTADO, OBSERV ) " _
      & " VALUES " _
      & " (" + Str(wCod) + ", '" + wLin + "', " + Str(Val(txtCodHijo.Text)) + ", " _
      & "  '" + wEstado + "', '" + txtObserv.Text + "'   ) ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAEASIGNADO " _
      & " SET CODHIJO = " + Str(Val(txtCodHijo.Text)) + ", " _
      & "     ESTADO = '" + wEstado + "', " _
      & "     OBSERV = '" + txtObserv.Text + "' " _
      & " WHERE CODSOCIO = " + Str(wCod) + " AND " _
      & "            LIN = '" + wLin + "'  ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFectop.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE MAEASIGNADO " _
      & " SET FECTOP = '" + Format(txtFectop, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(wCod) + " AND " _
      & "            LIN = '" + wLin + "'  ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAEASIGNADO " _
      & " SET FECTOP = null " _
      & " WHERE CODSOCIO = " + Str(wCod) + " AND " _
      & "            LIN = '" + wLin + "'  ")
      Db.CommitTrans
   End If
   
   aa = Leerado8("SELECT * FROM TMP_MAEASIGNADO " _
                & " WHERE CODSOCIO = " + Str(wCod) + " AND " _
                & "            LIN = '" + wLin + "' AND " _
                & "            USU = '" + wcodusu + "' ")
   If aa = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO TMP_MAEASIGNADO " _
      & " (USU, CODSOCIO, LIN, CODHIJO, NOMSOCIO, NOMHIJO, ESTADO, OBSERV ) " _
      & " VALUES " _
      & " ('" + wcodusu + "', " + Str(wCod) + ", '" + wLin + "', " _
      & "  " + Str(Val(txtCodHijo.Text)) + ", " _
      & "  '" + Trim(lblCodSocio.Caption) + "', " _
      & "  '" + Trim(lblCodHijo.Caption) + "', " _
      & "  '" + wEstado + "', '" + txtObserv.Text + "'   ) ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAEASIGNADO " _
      & " SET CODHIJO = " + Str(Val(txtCodHijo.Text)) + ", " _
      & "     ESTADO = '" + wEstado + "', " _
      & "     OBSERV = '" + txtObserv.Text + "', " _
      & "     NOMSOCIO = '" + Trim(lblCodSocio.Caption) + "', " _
      & "      NOMHIJO = '" + Trim(lblCodHijo.Caption) + "' " _
      & " WHERE CODSOCIO = " + Str(wCod) + " AND " _
      & "            LIN = '" + wLin + "' AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   If IsDate(txtFectop.Text) Then
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAEASIGNADO " _
      & " SET FECTOP = '" + Format(txtFectop, "dd/mm/yyyy") + "' " _
      & " WHERE CODSOCIO = " + Str(wCod) + " AND " _
      & "            LIN = '" + wLin + "' AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_MAEASIGNADO " _
      & " SET FECTOP = NULL " _
      & " WHERE CODSOCIO = " + Str(wCod) + " AND " _
      & "            LIN = '" + wLin + "' AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   End If
   
   lblMensaje.Caption = ""
   lblMensaje.Refresh
   
   LlenaCab
   LlenaCab1
   Limpiar
           
   ADO1.MoveFirst
   ADO1.Find "[CODSOCIO]=" + Str(wCod) + ""
   ADO1.Find "LIN = '" + wLin + "' "
   refrescar
   
   MsgBox "Asignación de Socios Esta Grabado OK", vbExclamation
   
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub editar(estado As Boolean)
   fraDetalle.Enabled = estado
   txtLin.Enabled = False
   
   cmdNuevo.Visible = Not estado
   cmdModificar.Visible = Not estado
   cmdEliminar.Visible = Not estado
   
   DataGrid1.Enabled = Not estado
   fraDesplaza.Enabled = Not estado
   
   cmdGrabar.Visible = estado
   cmdDeshacer.Visible = estado
   cmdSalir.Visible = Not estado
End Sub

Private Sub cmdCerrar_Click()

End Sub

Private Sub cmbEstado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtObserv.SetFocus
   End If
End Sub

Private Sub cmdDeshacer_Click()
   MsgBox "Los Cambios Efectuados Se Perderán", vbExclamation
   ACCION = 0
   
   editar (False)
   
   Limpiar
   refrescar
End Sub

Private Sub cmdEliminar_Click()
   On Error GoTo err
   
   Dim wCod As Integer, wNew As Integer, aa As Long, wLin As String, wLinNew As String
   wCod = ADO1!codsocio
   wLin = ADO1!lin
   wNew = 0
   wLinNew = ""
   ADO1.MoveNext
   If Not ADO1.EOF Then
      wNew = ADO1!codsocio
      wLinNew = ""
   Else
      ADO1.MovePrevious
      ADO1.MovePrevious
      If ADO1.BOF Then
         wNew = 0
         wLinNew = ""
      Else
         wNew = ADO1!codsocio
         wLinNew = ""
      End If
   End If
   If Not ADO1.BOF Or Not ADO1.EOF Then
      If MsgBox("¿Esta seguro de borrar Registro?", vbYesNo + vbDefaultButton2 + vbQuestion, "Advertencia") = vbYes Then
         
         Db.BeginTrans
         Db.Execute ("DELETE FROM MAEASIGNADO WHERE CODSOCIO = " + Str(wCod) + " AND LIN = '" + wLin + "' ")
         Db.CommitTrans
   
         Db.BeginTrans
         Db.Execute ("DELETE FROM TMP_MAEASIGNADO WHERE CODSOCIO = " + Str(wCod) + " AND LIN = '" + wLin + "' AND USU = '" + wcodusu + "' ")
         Db.CommitTrans
   
         DoEvents
         lblMensaje.Caption = ""
         lblMensaje.Refresh
         
         ADO1.Requery
         LlenaCab
         LlenaCab1
         Limpiar
         refrescar
      
         If wNew <> 0 Then
            ADO1.Find "CODSOCIO=" + Str(Val(wNew)) + ""
            ADO1.Find "LIN='" + wLinNew + "'"
         End If
          MsgBox "Asignado " + Str(wcon) + " " + wNom + vbNewLine + _
                "Eliminado OK", vbExclamation
         
      End If
   Else
      MsgBox "No Existe Registro a Eliminar", vbExclamation
   End If
   ACCION = 0
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdExporta_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(11) As String, wreg As Integer, wTot As Integer
   Dim wNom As String, wCod As Long, wIns As Integer, wCodHijo As Long, wInsHijo As Integer, _
       wFec As Date, wSoc As Integer
   Heading(0) = "SOCIO"
   Heading(1) = "CODOFIN"
   Heading(2) = "INS"
   Heading(3) = "NOMBRE"
   Heading(4) = "LIN"
   Heading(5) = "SOCIO HIJO"
   Heading(6) = "CODIGO HIJO"
   Heading(7) = "INS"
   Heading(8) = "NOMBRE HIJO"
   Heading(9) = "ESTADO"
   Heading(10) = "OBSERVAC"
   Heading(11) = "FECHA TOPE"
   
   Set objExcel = New Excel.Application
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add
   With objExcel
        .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
        .Range(.Cells(3, 1), .Cells(3, 12)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 12)).Font.Bold = True
        .Cells(1, 1) = wnomcia
        .Cells(2, 1) = "MAESTRO DE GRADOS"
        For I = 1 To 12 Step 1
            .Cells(3, I) = Heading(I - 1)
        Next
        objExcel.Columns("A").ColumnWidth = 9
        objExcel.Columns("B").ColumnWidth = 10
        objExcel.Columns("C").ColumnWidth = 4
        objExcel.Columns("D").ColumnWidth = 50
        objExcel.Columns("E").ColumnWidth = 5
        objExcel.Columns("F").ColumnWidth = 9
        objExcel.Columns("G").ColumnWidth = 10
        objExcel.Columns("H").ColumnWidth = 4
        objExcel.Columns("I").ColumnWidth = 50
        objExcel.Columns("J").ColumnWidth = 8
        objExcel.Columns("K").ColumnWidth = 20
        objExcel.Columns("L").ColumnWidth = 11
   End With
   
   
   aa = Leerado3("SELECT * FROM TMP_MAEASIGNADO ORDER BY NOMSOCIO, LIN ")
   If aa > 0 Then
      wTot = aa
      V = 4
      H = 1
      wreg = 1
      wSoc = ADO3!codsocio
      Do While Not ADO3.EOF
         DoEvents
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wreg, "####0") + " / " + Format(wTot, "####0")
         lblMensaje.Refresh
         
         If ADO3!codsocio <> wSoc Then
            V = V + 1
         End If
         
         wSoc = ADO3!codsocio
         wCod = 0: wIns = 0
         wCodHijo = 0: wInsHijo = 0
         aa = Leerado7("SELECT * FROM MAESOCIO " _
                    & " WHERE CODSOCIO = " + Str(ADO3!codsocio) + " ")
         If aa > 0 Then
            wCod = ADO7!codigo
            wIns = ADO7!ins
         End If
         Set ADO7 = Nothing
         
         aa = Leerado7("SELECT * FROM MAESOCIO " _
                    & " WHERE CODSOCIO = " + Str(ADO3!codhijo) + " ")
         If aa > 0 Then
            wCodHijo = ADO7!codigo
            wInsHijo = ADO7!ins
         End If
         Set ADO7 = Nothing
         
         objExcel.Cells(V, H + 0) = ADO3!codsocio
         objExcel.Cells(V, H + 1) = wCod
         objExcel.Cells(V, H + 2) = wIns
         objExcel.Cells(V, H + 3) = ADO3!nomsocio
         objExcel.Cells(V, H + 4) = ADO3!lin
         objExcel.Cells(V, H + 5) = ADO3!codhijo
         objExcel.Cells(V, H + 6) = wCodHijo
         objExcel.Cells(V, H + 7) = wInsHijo
         objExcel.Cells(V, H + 8) = ADO3!nomhijo
         objExcel.Cells(V, H + 9) = ADO3!estado
         objExcel.Cells(V, H + 10) = ADO3!observ
         If IsDate(ADO3!fectop) Then
            wFec = ADO3!fectop
            objExcel.Cells(V, H + 11) = wFec
         End If
         wreg = wreg + 1
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
   On Error GoTo err
   Dim wCod As Integer, wLin As String
   If ACCION = 1 Then
      wCod = Val(txtCodSocio.Text)
      wLin = txtLin.Text
      
      If Leerado7a("SELECT * FROM TMP_MAEASIGNADO " _
                & " WHERE CODSOCIO = " + Str(wCod) + " AND " _
                & "            LIN = '" + wLin + "' AND " _
                & "            USU = '" + wcodusu + "' ") > 0 Then
         MsgBox "Codigo de Socio Ya Existe", vbExclamation
         Limpiar
         txtCodSocio.SetFocus
         Exit Sub
      End If
   End If
   grabar
   editar False
   ACCION = 0
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdMigrar_Click()
   Dim aa As Integer, wRegAct As Integer, wRegTot As Integer, _
       wCodPadre As Long, wInsPadre As Long, wSocPadre As Long, wNomPadre As String, _
       wCodHijos As Long, wInsHijos As Long, wSocHijos As Long, wNomHijos As String, _
       wEstado As String, wMarca As String, wLin As Integer, wFecTop As Date

   Db.BeginTrans
   Db.Execute ("DELETE FROM MAEASIGNADO")
   Db.CommitTrans
   
   aa = Leerado8("SELECT * FROM ZZZ_M_ASIGNA ORDER BY COD_PADRE, INS_PADRE, CODIGO, INS ")
   If aa > 0 Then
      ADO8.MoveFirst
      wRegAct = 1
      wRegTot = aa
      Do While Not ADO8.EOF
         wCodPadre = ADO8!cod_padre
         wInsPadre = ADO8!ins_padre
         wSocPadre = 0
         wNomPadre = ""
         aa = Leerado7("SELECT * FROM zzz_MAESTRO " _
                        & " WHERE CODIGO = " + Str(wCodPadre) + " AND " _
                        & "          INS = " + Str(wInsPadre) + " ")
         If aa > 0 Then
            wSocPadre = ADO7!codsocio
            wNomPadre = ADO7!nombre
         End If
         Set ADO7 = Nothing
         wLin = 1
         
   '      Db.BeginTrans
   '      Db.Execute ("INSERT INTO ASIGFAMCAB " _
   '      & " (CODSOCIO, NOMBRE) " _
   '      & "  values " _
   '      & " (" + Str(wSocPadre) + ", '" + GlosaLibre(wNomPadre) + "')  ")
   '      Db.CommitTrans
      
         Do While ADO8!cod_padre = wCodPadre And ADO8!ins_padre = wInsPadre
            DoEvents
            lblMensaje.Caption = "Registro " + _
                                 Trim(Format(wRegAct, "####0")) + " / " + _
                                 Trim(Format(wRegTot, "####0"))
            lblMensaje.Refresh
            
            wCodHijos = ADO8!codigo
            wInsHijos = ADO8!ins
            wSocHijos = 0: wNomHijos = ""
            wEstado = ADO8!estado
            If IsDate(Trim(ADO8!marca)) Then
               wMarca = ""
               wFecTop = Format(Trim(ADO8!marca), "dd/mm/yyyy")
            Else
               wMarca = IIf(IsNull(ADO8!marca), "", ADO8!marca)
               wFecTop = Format("01/01/1900", "dd/mm/yyyy")
            End If
   
            aa = Leerado7("SELECT * FROM ZZZ_MAESTRO " _
                           & " WHERE CODIGO = " + Str(wCodHijos) + " AND " _
                           & "          INS = " + Str(wInsHijos) + " ")
            If aa > 0 Then
               wSocHijos = ADO7!codsocio
               wNomHijos = ADO7!nombre
            End If
            Set ADO7 = Nothing
      
            Db.BeginTrans
            Db.Execute ("INSERT INTO MAEASIGNADO " _
            & " (CODSOCIO, LIN, CODHIJO, NOMHIJO, ESTADO, OBSERV ) " _
            & " VALUES " _
            & " (" + Str(wSocPadre) + ", '" + Format(wLin, "00") + "', " _
            & "  " + Str(wSocHijos) + ", '" + GlosaLibre(wNomHijos) + "', " _
            & "  '" + wEstado + "', '" + wMarca + "' ) ")
            Db.CommitTrans
   
            If Format(wFecTop, "dd/mm/yyyy") <> "01/01/1900" Then
               Db.BeginTrans
               Db.Execute ("UPDATE MAEASIGNADO " _
               & " SET FECTOP = '" + Format(wFecTop, "dd/mm/yyyy") + "' " _
               & " WHERE CODSOCIO = " + Str(wSocPadre) + "  AND " _
               & "            LIN = '" + Format(wLin, "00") + "' ")
               Db.CommitTrans
            End If
   
            wRegAct = wRegAct + 1
            wLin = wLin + 1
            ADO8.MoveNext
            If ADO8.EOF Then
               Exit Do
            End If
         Loop
   
         If ADO8.EOF Then
            Exit Do
         End If
      Loop
   End If

   MsgBox "Proceso Termino OK", vbExclamation
End Sub

Private Sub cmdModificar_Click()
   If Not ADO1.BOF Or Not ADO1.EOF Then
      ACCION = 2
      editar True
      refrescar
      
      txtCodSocio.Enabled = False
      txtLin.Enabled = False
      
      txtCodHijo.SetFocus
   Else
      MsgBox "No Existe Registro a Modificar", vbExclamation
   End If
End Sub

Private Sub cmdMover_Click(Index As Integer)
    With ADO1
    If .BOF And .EOF Then
       Exit Sub
    End If
    Select Case Index
    Case 0
        .MoveFirst
    Case 1
        .MovePrevious
        If .BOF Then .MoveFirst
    Case 2
        .MoveNext
        If .EOF Then .MoveLast
    Case 3
        .MoveLast
    End Select
    End With
    refrescar
End Sub

Private Sub cmdNuevo_Click()
   Dim wNew As String, aa As Integer
   
   ACCION = 1
   editar True
   Limpiar
   
   txtCodSocio.SetFocus
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If Not ADO1.BOF = True And Not ADO1.EOF = True Then
      DataGrid1.Refresh
      
      Limpiar
      refrescar
   End If
End Sub

Private Sub Form_Activate()
   frmAsigFam.Left = (Screen.Width - Width) \ 2
   frmAsigFam.Top = 0
   
'   Llenacab
'   refrescar
'   DataGrid1.SetFocus
   
   ACCION = "0"
   
   Dim a As Integer, I As Integer
   a = Leerado8("SELECT * FROM MAEESTADOASIGNADO ORDER BY CODIGO ")
   If a > 0 Then
      ADO8.MoveFirst
      I = 0
      Do While Not ADO8.EOF
         cmbEstado.AddItem ADO8!nombre
         I = I + 1
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   
   editar (False)
   
   LlenaCab
   LlenaCab1
   Limpiar
   refrescar

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataGrid1.DataSource = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_MAEASIGNADO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim a As Integer
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_MAEASIGNADO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_MAEASIGNADO " _
   & " (CODSOCIO, LIN, CODHIJO, ESTADO, OBSERV, FECTOP, " _
   & "  NOMSOCIO, NOMHIJO, USU ) " _
   & " SELECT " _
   & "  A.CODSOCIO, A.LIN, A.CODHIJO, A.ESTADO, A.OBSERV, A.FECTOP, " _
   & "  M.NOMBRE, '', '" + wcodusu + "' " _
   & " FROM MAEASIGNADO AS A LEFT JOIN MAESOCIO AS M " _
   & "   ON A.CODSOCIO = M.CODSOCIO ")
   Db.CommitTrans
   
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_MAEASIGNADO " _
   & " SET NOMHIJO = M.NOMBRE " _
   & " FROM TMP_MAEASIGNADO AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODHIJO = M.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   a = Leerado("SELECT CODSOCIO, NOMSOCIO, LIN, CODHIJO, NOMHIJO, ESTADO, OBSERV, FECTOP " _
                & " FROM TMP_MAEASIGNADO " _
                & " WHERE USU = '" + wcodusu + "' " _
                & " ORDER BY NOMSOCIO, LIN ")
   Set DataGrid1.DataSource = ADO1
End Sub

Private Sub LlenaCab1()
   DataGrid1.Columns(0).Width = 700
   DataGrid1.Columns(0).Alignment = dbgRight
   DataGrid1.Columns(0).Caption = "CODIGO"
    
   DataGrid1.Columns(1).Width = 3500
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Caption = "NOMBRE"
    
   DataGrid1.Columns(2).Width = 300
   DataGrid1.Columns(2).Alignment = dbgLeft
   DataGrid1.Columns(2).Caption = "LIN"
    
   DataGrid1.Columns(3).Width = 700
   DataGrid1.Columns(3).Alignment = dbgRight
   DataGrid1.Columns(3).Caption = "COD.HIJO"
    
   DataGrid1.Columns(4).Width = 3500
   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Caption = "NOMBRE HIJO"
    
   DataGrid1.Columns(5).Width = 500
   DataGrid1.Columns(5).Alignment = dbgLeft
   DataGrid1.Columns(5).Caption = "ESTADO"
    
   DataGrid1.Columns(6).Width = 1800
   DataGrid1.Columns(6).Alignment = dbgLeft
   DataGrid1.Columns(6).Caption = "OBSERV"
    
   DataGrid1.Columns(7).Width = 1050
   DataGrid1.Columns(7).Alignment = dbgLeft
   DataGrid1.Columns(7).Caption = "FEC.TOPE"
   DataGrid1.Columns(7).NumberFormat = "dd/mm/yyyy"
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
      Limpiar
      refrescar
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
         ADO1.Filter = ""
         Set DataGrid1.DataSource = ADO1
         DataGrid1.SetFocus
      Else
         txtFiltrar.Enabled = True
         optFiltro.Value = True
         txtFiltrar.SetFocus
      End If
   End If
End Sub

Private Sub txtCodHijo_Change()
   Dim aa As Integer
   aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodHijo.Text)) + " ")
   If aa > 0 Then
      lblCodHijo.Caption = ADO8!nombre
   Else
      lblCodHijo.Caption = ""
   End If
   Set ADO8 = Nothing
End Sub

Private Sub txtCodHijo_GotFocus()
   txtCodHijo.SelStart = 0
   If Len(Trim(txtCodHijo.Text)) > 0 Then
      txtCodHijo.SelLength = Len(Trim(txtCodHijo.Text))
   Else
      txtCodHijo.SelLength = 8
   End If
End Sub

Private Sub txtCodHijo_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        If txtLin.Enabled = True Then
           txtLin.SetFocus
        End If
   Case 40
        cmbEstado.SetFocus
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtCodHijo.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtCodHijo_KeyPress(KeyAscii As Integer)
   Dim aa As Integer, wLin As String
   If KeyAscii = 13 Then
      If Len(Trim(txtCodHijo.Text)) = 0 Then
         MsgBox "Codigo Hijo En Blanco", vbExclamation
         txtCodHijo.Text = ""
         Exit Sub
      End If
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodHijo.Text)) + " ")
      If aa = 0 Then
         MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
         txtCodHijo.Text = ""
         Exit Sub
      End If
      lblCodHijo.Caption = ADO8!nombre
      
      cmbEstado.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtCodSocio_Change()
   Dim aa As Integer
   aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
   If aa > 0 Then
      lblCodSocio.Caption = ADO8!nombre
   Else
      lblCodSocio.Caption = ""
   End If
   Set ADO8 = Nothing
End Sub

Private Sub txtCodSocio_GotFocus()
   txtCodSocio.SelStart = 0
   If Len(Trim(txtCodSocio.Text)) > 0 Then
      txtCodSocio.SelLength = Len(Trim(txtCodSocio.Text))
   Else
      txtCodSocio.SelLength = 8
   End If
End Sub

Private Sub txtCodSocio_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 40
        If txtLin.Enabled = True Then
           txtLin.SetFocus
        End If
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtCodSocio.Text = xseleccion
        End If
   End Select
End Sub

Private Sub txtCodSocio_KeyPress(KeyAscii As Integer)
   Dim aa As Integer, wLin As String
   If KeyAscii = 13 Then
      If Len(Trim(txtCodSocio.Text)) = 0 Then
         MsgBox "Codigo Socio En Blanco", vbExclamation
         txtCodSocio.Text = ""
         Exit Sub
      End If
      aa = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
      If aa = 0 Then
         MsgBox "Codigo Socio Digitado NO Existe", vbExclamation
         txtCodSocio.Text = ""
         Exit Sub
      End If
      lblCodSocio.Caption = ADO8!nombre
      
      wLin = "00"
      aa = Leerado8("SELECT MAX(LIN) AS LIN FROM MAEASIGNADO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
      If aa > 0 Then
         wLin = IIf(IsNull(ADO8!lin), "00", ADO8!lin)
      End If
      Set ADO8 = Nothing
    
      wLin = Format(Val(wLin) + 1, "00")
      txtLin.Text = wLin
      txtLin.Enabled = False
      
      txtCodHijo.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtE_socio_Change()
   Dim aa As Integer
   aa = Leerado8("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + txtE_socio.Text + "' ")
   If aa > 0 Then
      lblE_socio.Caption = ADO8!nombre
   Else
      lblE_socio.Caption = ""
   End If
   Set ADO8 = Nothing
End Sub

Private Sub txtFectop_GotFocus()
   txtFectop.SelStart = 0
   txtFectop.SelLength = 10
End Sub

Private Sub txtFectop_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtObserv.SetFocus
   End Select
End Sub

Private Sub txtFectop_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdGrabar.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFiltrar_Change()
   Dim a As Long
   
   a = Leerado("SELECT CODSOCIO, NOMSOCIO, LIN, CODHIJO, NOMHIJO, ESTADO, OBSERV, FECTOP " _
                & " FROM TMP_MAEASIGNADO " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "       NOMSOCIO LIKE '%" + Trim(txtFiltrar.Text) + "%' " _
                & " ORDER BY NOMSOCIO, LIN ")
   Set DataGrid1.DataSource = ADO1
   
   LlenaCab1
   Limpiar
   refrescar
End Sub

Private Sub txtFiltrar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtGrado_Change()
   Dim aa As Integer
   aa = Leerado8("SELECT * FROM MAEGRADO WHERE GRADO = " + Str(Val(txtGrado.Text)) + " ")
   If aa > 0 Then
      lblGrado.Caption = ADO8!nombre
   Else
      lblGrado.Caption = ""
   End If
   Set ADO8 = Nothing
End Sub

Private Sub txtTipCob_Change()
   Dim aa As Integer
   aa = Leerado8("SELECT * FROM MAETIPCOB WHERE TIPCOB = '" + txtTipCob.Text + "' ")
   If aa > 0 Then
      lblTipCob.Caption = ADO8!nombre
   Else
      lblTipCob.Caption = ""
   End If
   Set ADO8 = Nothing
End Sub

Private Function validaAsi()
   On Error GoTo err
   Dim a As Integer, wTot As Currency
   
   If Len(Trim(txtCodSocio.Text)) = 0 Then
      MsgBox "Codigo de Socio En Blanco", vbExclamation
      txtCodSocio.SetFocus
      validaAsi = True
      Exit Function
   Else
      a = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO=" + Str(Val(txtCodSocio.Text)) + " ")
      If a = 0 Then
         MsgBox "Codigo de Socio No Existe", vbExclamation
         txtCodSocio.Text = ""
         txtCodSocio.SetFocus
         validaAsi = True
         Exit Function
      End If
   End If
   
   If Len(Trim(txtCodHijo.Text)) = 0 Then
      MsgBox "Codigo de Hijo En Blanco", vbExclamation
      txtCodHijo.SetFocus
      validaAsi = True
      Exit Function
   Else
      a = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO=" + Str(Val(txtCodHijo.Text)) + " ")
      If a = 0 Then
         MsgBox "Codigo de Hijo No Existe", vbExclamation
         txtCodHijo.Text = ""
         txtCodHijo.SetFocus
         validaAsi = True
         Exit Function
      End If
   End If
   
   validaAsi = False
   Exit Function
err:
   MsgBox Format(err.Number, "000000000000") + " " + err.Description
   Resume Next
End Function

Private Sub txtObserv_GotFocus()
   txtObserv.SelStart = 0
   txtObserv.SelLength = Len(Trim(txtObserv.Text))
End Sub

Private Sub txtObserv_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        cmbEstado.SetFocus
   Case 40
        txtFectop.SetFocus
   End Select
End Sub

Private Sub txtObserv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFectop.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub
