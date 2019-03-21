VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMaeGrado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Grados"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   13545
   Begin VB.CommandButton Command6 
      Caption         =   "Recibos Anulados"
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
      Left            =   10080
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Borra Duplicados"
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
      Left            =   8280
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmdAcciones 
      Caption         =   "Acciones Socios"
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
      Left            =   11400
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Actualiza COBRODET (2 Lineas)"
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
      Left            =   6600
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fraccion."
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
      Left            =   10080
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Socios ACTIVOS"
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
      Left            =   5160
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Socios DEL 01/10/2017"
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
      Left            =   3720
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Socios sin FECING"
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
      Left            =   2280
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Frame fraDesplaza 
      Caption         =   "Desplazamiento"
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
      Left            =   10920
      TabIndex        =   20
      Top             =   840
      Width           =   2295
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
         Left            =   1560
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
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
         Left            =   1080
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
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
         Left            =   600
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
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
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
   End
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
      Left            =   10920
      TabIndex        =   18
      Top             =   120
      Width           =   2295
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
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
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
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   8520
      TabIndex        =   12
      Top             =   120
      Width           =   2295
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
         Left            =   1200
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
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
         Left            =   1200
         TabIndex        =   16
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
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
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   840
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtro"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      Width           =   8775
      Begin VB.TextBox txtFiltrar 
         Height          =   285
         Left            =   4080
         MaxLength       =   50
         TabIndex        =   11
         Top             =   240
         Width           =   4335
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Filtrar x Nombre"
         Height          =   255
         Left            =   2280
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Mostrar Todos"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Detalles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8295
      Begin VB.ComboBox cmbGrupo 
         Height          =   315
         ItemData        =   "frmMaeGrado.frx":0000
         Left            =   1800
         List            =   "frmMaeGrado.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1280
         Width           =   2895
      End
      Begin VB.TextBox txtCodigo 
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
         Height          =   285
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   4
         Top             =   480
         Width           =   570
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   3
         Top             =   880
         Width           =   6015
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
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
         Index           =   2
         Left            =   1020
         TabIndex        =   26
         Top             =   1275
         Width           =   525
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   720
         TabIndex        =   6
         Top             =   880
         Width           =   840
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   720
         TabIndex        =   5
         Top             =   480
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
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
      Left            =   11160
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   480
      TabIndex        =   7
      Top             =   2040
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5953
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
      Caption         =   "TABLA DE GRADOS"
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
            LCID            =   3082
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
            LCID            =   3082
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
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   6240
      Width           =   7935
   End
End
Attribute VB_Name = "frmMaeGrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACCION As Byte
Dim marca As Variant, wcia As String
    
Sub Limpiar()
   txtCodigo.Text = ""
   txtNombre.Text = ""
   cmbGrupo.ListIndex = 0
End Sub

Sub refrescar()
   If ADO1.BOF Then Exit Sub
   If ADO1.EOF Then Exit Sub
   txtCodigo.Text = ADO1!grado
   txtNombre.Text = IIf(IsNull(ADO1!nombre), "", ADO1!nombre)
   cmbGrupo.ListIndex = ADO1!gradogrupo - 1
End Sub

Sub grabar()
   On Error GoTo err
   
   Dim aa As Integer, wCod As String, wNom As String, wGru As Integer
   wCod = txtCodigo.Text
   wNom = txtNombre.Text
   wGru = cmbGrupo.ListIndex + 1
   
   If Len(Trim(wCod)) = 0 Then
      MsgBox "Codigo En Blanco", vbExclamation
      Exit Sub
   End If
   
   If Len(Trim(wNom)) = 0 Then
      MsgBox "Nombre En Blanco", vbExclamation
      Exit Sub
   End If
   
   aa = Leerado8("SELECT * FROM MAEGRADO WHERE GRADO = " + Str(Val(wCod)) + " ")
   If aa = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO MAEGRADO " _
      & " (GRADO, NOMBRE, GRADOGRUPO) " _
      & " VALUES " _
      & " (" + Str(Val(wCod)) + ", '" + wNom + "' ) ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAEGRADO " _
      & " SET NOMBRE = '" + wNom + "', GRADOGRUPO = " + Str(wGru) + " " _
      & " WHERE GRADO = " + Str(Val(wCod)) + " ")
      Db.CommitTrans
   End If
   ADO1.Requery
   LlenaCab1
   ADO1.Find "GRADO = " + Str(Val(wCod)) + " "
   MsgBox "Grado " + wCod + " " + wNom + vbNewLine + _
          "Grabado OK", vbExclamation
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub editar(estado As Boolean)
   txtCodigo.Enabled = estado
   txtNombre.Enabled = estado
   cmbGrupo.Enabled = estado
   
   cmdNuevo.Visible = Not estado
   cmdModificar.Visible = Not estado
   cmdEliminar.Visible = Not estado
   
   DataGrid1.Enabled = Not estado
   fraDesplaza.Enabled = Not estado
   fraFiltro.Enabled = Not estado
   
   cmdGrabar.Visible = estado
   cmdDeshacer.Visible = estado
   cmdExporta.Visible = Not estado
   cmdCerrar.Visible = Not estado
End Sub

Private Sub cmbGrupo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdGrabar.SetFocus
   End If
End Sub

Private Sub cmdAcciones_Click()
   Dim aa As Long
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM MAESOCIO_ACCION ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO MAESOCIO_ACCION " _
   & " (CODSOCIO, LINEA, TIPO, FECHA) " _
   & " SELECT " _
   & "  CODSOCIO, '01', 'ING', FECING " _
   & " FROM MAESOCIO " _
   & " WHERE FECING IS NOT NULL ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO MAESOCIO_ACCION " _
   & " (CODSOCIO, LINEA, TIPO, FECHA) " _
   & " SELECT " _
   & "  CODSOCIO, '02', 'REN', FECRENU " _
   & " FROM MAESOCIO " _
   & " WHERE FECRENU IS NOT NULL ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO MAESOCIO_ACCION " _
   & " (CODSOCIO, LINEA, TIPO, FECHA) " _
   & " SELECT " _
   & "  CODSOCIO, '02', 'REI', FECREIN " _
   & " FROM MAESOCIO " _
   & " WHERE FECREIN IS NOT NULL AND " _
   & "       FECRENU IS NULL ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO MAESOCIO_ACCION " _
   & " (CODSOCIO, LINEA, TIPO, FECHA) " _
   & " SELECT " _
   & "  CODSOCIO, '03', 'REI', FECREIN " _
   & " FROM MAESOCIO " _
   & " WHERE FECREIN IS NOT NULL AND " _
   & "       FECRENU IS NOT NULL ")
   Db.CommitTrans

   Dim wSoc As Integer, wCod As Long, wIns As Integer, wFecRen As Date
   
   aa = Leerado8("SELECT CODSOCIO, CODIGO, INS, NOMBRE, FECRENU " _
                & " From MAESOCIO " _
                & " WHERE FECRENU > '21/10/2017' " _
                & " ORDER BY FECRENU")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wSoc = ADO8!codsocio
         wCod = ADO8!codigo
         wIns = ADO8!ins
         wFecRen = Format(ADO8!fecrenu, "dd/mm/yyyy")
   
   
   
   
         ADO8.MoveNext
      Loop
   End If
   


End Sub

Private Sub cmdCerrar_Click()
   Unload Me
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
   
   Dim wcon As Integer, wNom As String, wNew As Integer, aa As Integer
   wcon = ADO1!grado
   wNom = Trim(ADO1!nombre)
   wNew = ""
   ADO1.MoveNext
   If Not ADO1.EOF Then
      wNew = ADO1!grado
   Else
      ADO1.MovePrevious
      ADO1.MovePrevious
      If ADO1.BOF Then
         wNew = ""
      Else
         wNew = ADO1!grado
      End If
   End If
   
   aa = Leerado8("SELECT * FROM MAESOCIO WHERE GRA = " + Str(Val(wcon)) + " ")
   If aa > 0 Then
      MsgBox "Grado Tiene Movimiento, No Se Pude Borrar", vbExclamation
      Exit Sub
   End If
   
   If MsgBox("¿Esta seguro de borrar Codigo " + wcon + "?", vbYesNo + vbDefaultButton2 + vbQuestion, "Advertencia") = vbYes Then
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEGRADO WHERE GRADO = " + Str(wcon) + " ")
      Db.CommitTrans
      
      ADO1.Requery
      LlenaCab
      LlenaCab1
      Limpiar
      refrescar
      
      If wNew <> 0 Then
         ADO1.Find "GRADO=" + Str(Val(wNew)) + ""
      End If
      MsgBox "Grado " + Str(wcon) + " " + wNom + vbNewLine + _
             "Eliminado OK", vbExclamation
   End If
   ACCION = 0
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdExporta_Click()
   On Error GoTo err
   
   Dim aa As Integer, I As Integer, Heading(3) As String, wreg As Integer, wTot As Integer
   Dim wNom As String
   Heading(0) = "CODIGO"
   Heading(1) = "NOMBRE"
   Heading(2) = "GRUPO"
   Heading(3) = "NOMBRE GRUPO"
   aa = Leerado3("SELECT * FROM MAEGRADO ORDER BY GRADO ")
   If aa > 0 Then
      wTot = aa
      Set objExcel = New Excel.Application
      objExcel.SheetsInNewWorkbook = 1
      objExcel.Workbooks.Add
      With objExcel
           .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
           .Range(.Cells(3, 1), .Cells(3, 2)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 2)).Font.Bold = True
           .Cells(1, 1) = wnomcia
           .Cells(2, 1) = "MAESTRO DE GRADOS"
           For I = 1 To 2 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           objExcel.Columns("A").ColumnWidth = 15
           objExcel.Columns("B").ColumnWidth = 50
           objExcel.Columns("C").ColumnWidth = 15
           objExcel.Columns("D").ColumnWidth = 50
      End With
      V = 4
      H = 1
      wreg = 1
      Do While Not ADO3.EOF
         DoEvents
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wreg, "####0") + " / " + Format(wTot, "####0")
         lblMensaje.Refresh
         
         wNom = ""
         aa = Leerado7("SELECT * FROM MAEGRADOGRUPO " _
                    & " WHERE GRADOGRUPO = " + Str(ADO3!gradogrupo) + " ")
         If aa > 0 Then
            wNom = ADO7!nombre
         End If
         Set ADO7 = Nothing
         
         objExcel.Cells(V, H + 0) = ADO3!grado
         objExcel.Cells(V, H + 1) = IIf(IsNull(ADO3!nombre), "", ADO3!nombre)
         objExcel.Cells(V, H + 2) = ADO3!gradogrupo
         objExcel.Cells(V, H + 3) = wNom
         
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
   Dim wCod As String
   If ACCION = 1 Then
      wCod = txtCodigo.Text
      If Leerado2("SELECT * FROM MAEGRADO WHERE GRADO = " + Str(Val(wCod)) + " ") > 0 Then
         MsgBox "Codigo Ya Existe", vbExclamation
         Limpiar
         txtCodigo.SetFocus
         Exit Sub
      End If
   End If
   grabar
   ACCION = 0
   editar False
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub cmdModificar_Click()
   ACCION = 2
   editar True
   refrescar
   txtCodigo.Enabled = False
   txtNombre.SetFocus
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
   
   txtCodigo.SetFocus
End Sub

Private Sub Command1_Click()
5   Dim zz As Integer, wSoc As Integer, wFecIng As Date, _
       wMon As String, wApo As String, wE_S As String, _
       aa As Integer, II As Integer, wDesde As Integer, wHasta As Integer, _
       wAnoDes As Integer, wAnoHas As Integer, wMesDes As Integer, wMesHas As Integer, _
       wmmm As String, waaa As String, wFec As Date
   
   zz = Leerado8("SELECT M.CODSOCIO, M.CODIGO, M.INS, M.FECING, M.E_SOCIO, E.APORTE, E.MONEDA " _
                & " FROM MAESOCIO AS M INNER JOIN MAEE_SOCIO AS E ON M.E_SOCIO = E.E_SOCIO " _
                & " WHERE (FECING >= '01/10/2017') AND " _
                & "       (M.E_SOCIO = 'TIT' OR " _
                & "        M.E_SOCIO = 'VIU' OR " _
                & "        M.E_SOCIO = 'HIJ' OR " _
                & "        M.E_SOCIO = 'NIE' OR " _
                & "        M.E_SOCIO = 'HER' OR " _
                & "        M.E_SOCIO = 'CIV' OR " _
                & "        M.E_SOCIO = 'CI1' OR " _
                & "        M.E_SOCIO = 'ADH' OR " _
                & "        M.E_SOCIO = 'TRA' OR " _
                & "        M.E_SOCIO = 'PNP' ) " _
                & " ORDER BY M.CODSOCIO ")
   If zz > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wSoc = ADO8!codsocio
         wFecIng = ADO8!fecing
         wE_S = ADO8!e_socio
         wMon = ADO8!moneda
         wApo = ADO8!aporte
   
         If wSoc = 10302 Then
            MsgBox "10302"
         End If
   
         wAnoDes = Year(wFecIng)
         If Day(wFecIng) >= 20 Then
            
            waaa = Format(Year(wFecIng), "0000")
            wmmm = Format(Month(wFecIng), "00")
            
            Db.BeginTrans
            Db.Execute ("DELETE FROM CTASXCAB " _
            & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
            & "       CONCEPTO = '01' AND " _
            & "            MES = '" + waaa + "/" + wmmm + "'  ")
            Db.CommitTrans
            
            Db.BeginTrans
            Db.Execute ("DELETE FROM CTASXDET " _
            & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
            & "            MES = '" + waaa + "/" + wmmm + "' AND " _
            & "       CONCEPTO = '01' AND " _
            & "         TIPCOB = '00' AND " _
            & "         TIPMOV = '1' ")
            Db.CommitTrans
            
            wMesDes = Month(wFecIng) + 1
            If wMesDes > 12 Then
               wMesDes = 1
               wAnoDes = wAnoDes + 1
            End If
         Else
            wMesDes = Month(wFecIng)
         End If
             
         wAnoHas = 2018
         wMesHas = 12
             
         Db.BeginTrans
         Db.Execute ("DELETE FROM CTASXDET " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       MES >= '2017/10' AND MES < '" + Format(wAnoDes, "0000") + "/" + Format(wMesDes, "00") + "' AND " _
         & "       TIPCOB = '00' AND " _
         & "       TIPMOV = '1' ")
         Db.CommitTrans
             
         For aa = wAnoDes To wAnoHas
             waaa = Format(aa, "0000")
         
             If aa = wAnoDes Then
                wDesde = wMesDes
             Else
                wDesde = 1
             End If
         
             If aa = wAnoHas Then
                wHasta = wMesHas
             Else
                wHasta = 12
             End If
         
             For II = wDesde To wHasta
                 wmmm = Format(II, "00")
         
                 wFec = Format("01/" + wmmm + "/" + waaa, "dd/mm/yyyy")
                
                 zz = Leerado6a("SELECT * FROM CTASXCAB " _
                             & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                             & "       CONCEPTO = '01' AND " _
                             & "            MES = '" + waaa + "/" + wmmm + "'  ")
                 If zz = 0 Then
                    Db.BeginTrans
                    Db.Execute ("INSERT INTO CTASXCAB " _
                    & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, CARGOS, ABONOS, SDONEW ) " _
                    & " VALUES " _
                    & " (" + Str(wSoc) + ", '" + waaa + "/" + wmmm + "', '01', '" + wE_S + "', '" + wMon + "', " _
                    & "  " + Str(wApo) + ", 0, " + Str(wApo) + " ) ")
                    Db.CommitTrans
                 End If
                 Set ADO6a = Nothing
                
                 zz = Leerado6a("SELECT * FROM CTASXDET " _
                            & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                            & "            MES = '" + waaa + "/" + wmmm + "' AND " _
                            & "       CONCEPTO = '01' AND " _
                            & "         TIPCOB = '00' AND " _
                            & "         TIPMOV = '1' ")
                 If zz = 0 Then
                    Db.BeginTrans
                    Db.Execute ("INSERT INTO CTASXDET " _
                    & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
                    & "  DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW) " _
                    & " VALUES " _
                    & " (" + Str(wSoc) + ", '" + waaa + "/" + wmmm + "', '01', '00', '', '', '', '1', " _
                    & "  '" + Format(wFec, "dd/mm/yyyy") + "', 0, 0, 0, " + Str(wApo) + ", " _
                    & "  0, " + Str(wApo) + " ) ")
                    Db.CommitTrans
                 Else
                    Db.BeginTrans
                    Db.Execute ("UPDATE CTASXDET " _
                    & " SET FECHA = '" + Format(wFec, "dd/mm/yyyy") + "', " _
                    & "     CARGOS = " + Str(wApo) + ", SDONEW = " + Str(wApo) + " " _
                    & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                    & "            MES = '" + waaa + "/" + wmmm + "' AND " _
                    & "       CONCEPTO = '01' AND " _
                    & "         TIPCOB = '00' AND " _
                    & "         TIPMOV = '1' ")
                    Db.CommitTrans
                 End If
                 Set ADO6a = Nothing
         
             Next II
         
         Next aa
                   
         ADO8.MoveNext
      Loop
   End If
   
   MsgBox "Proceso Termino ok", vbExclamation
End Sub

Private Sub Command10_Click()
   Dim aa As Integer, wRegAct As Integer, wRegTot As Integer, _
       wFecIng As Date, wSoc As Integer, wCod As Long, wIns As Integer, _
       wDiaIng As String, wMesIng As String, wAnoIng As String, _
       wDesde As Integer, wHasta As Integer, II As Integer, wmmm As String, _
       wE_S As String, wMon As String, wApo As Currency, wFec As Date
   
   
   aa = Leerado8("SELECT * FROM MAESOCIO WHERE FECING IS NULL ORDER BY CODSOCIO")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         DoEvents
         lblMensaje.Caption = "Registro " + _
                              Trim(Format(wRegAct, "####0")) + " / " + _
                              Trim(Format(wRegTot, "####0"))
         lblMensaje.Refresh
         wSoc = ADO8!codsocio
         wCod = ADO8!codigo
         wIns = ADO8!ins
         wE_S = ADO8!e_socio
         
         aa = Leerado7("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + wE_S + "' ")
         If aa > 0 Then
            wMon = ADO7!moneda
            wApo = ADO7!aporte
         End If
         Set ADO7 = Nothing
   
         aa = Leerado7("SELECT * FROM ZZZ_MRECIBOS " _
                    & " WHERE CODIGO = " + Str(wCod) + " AND " _
                    & "          INS = " + Str(wIns) + " " _
                    & " ORDER BY FECHA_PAGO ")
         If aa > 0 Then
            ADO7.MoveFirst
            wFecIng = ADO7!fecha_pago
         End If
         Set ADO7 = Nothing
         
         wDiaIng = Format(Day(wFecIng), "00")
         wMesIng = Format(Month(wFecIng), "00")
         wAnoIng = Format(Year(wFecIng), "0000")
   
         If wDiaIng >= "20" Then
             wMesIng = Format(Val(wMesIng) + 1, "00")
             If wMesIng > "12" Then
                wMesIng = Format(Val(wMesIng) - 12, "00")
                wAnoIng = Format(Val(wAnoIng) + 1, "0000")
             End If
         End If
   
         If wAnoIng = "2017" Then
            If wAnoIng = "2017" Then
               wDesde = Val(wMesIng)
            End If
            
            For II = wDesde To 12
                wmmm = "2017/" + Format(II, "00")
                wFec = Format("01/" + Right(wmmm, 2) + "/" + Left(wmmm, 4), "dd/mm/yyyy")
                
                aa = Leerado6a("SELECT * FROM CTASXCAB " _
                            & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                            & "       CONCEPTO = '01' AND " _
                            & "            MES = '" + wmmm + "'  ")
                If aa = 0 Then
                   Db.BeginTrans
                   Db.Execute ("INSERT INTO CTASXCAB " _
                   & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, CARGOS, ABONOS, SDONEW ) " _
                   & " VALUES " _
                   & " (" + Str(wSoc) + ", '" + wmmm + "', '01', '" + wE_S + "', '" + wMon + "', " _
                   & "  " + Str(wApo) + ", 0, " + Str(wApo) + " ) ")
                   Db.CommitTrans
                End If
                
                aa = Leerado6a("SELECT * FROM CTASXDET " _
                           & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                           & "            MES = '" + wmmm + "' AND " _
                           & "       CONCEPTO = '01' AND " _
                           & "         TIPCOB = '00' AND " _
                           & "         TIPMOV = '1' ")
                If aa = 0 Then
                   Db.BeginTrans
                   Db.Execute ("INSERT INTO CTASXDET " _
                   & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
                   & "  DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW) " _
                   & " VALUES " _
                   & " (" + Str(wSoc) + ", '" + wmmm + "', '01', '00', '', '', '', '1', " _
                   & "  '" + Format(wFec, "dd/mm/yyyy") + "', 0, 0, 0, " + Str(wApo) + ", " _
                   & "  0, " + Str(wApo) + " ) ")
                   Db.CommitTrans
                Else
                   Db.BeginTrans
                   Db.Execute ("UPDATE CTASXDET " _
                   & " SET FECHA = '" + Format(wFec, "dd/mm/yyyy") + "', " _
                   & "     CARGOS = " + Str(wApo) + ", SDONEW = " + Str(wApo) + " " _
                   & " WHERE CODSOCIO = " + Str(wSoc) + " " _
                   & "            MES = '" + wmmm + "' AND " _
                   & "       CONCEPTO = '01' AND " _
                   & "         TIPCOB = '00' AND " _
                   & "         TIPMOV = '1' ")
                   Db.CommitTrans
                End If
            Next II
         End If
   
         If wAnoIng >= "2017" Then
            If wAnoIng = "2017" Then
               wDesde = 1
            Else
               wDesde = Val(wMesIng)
            End If
            
            For II = wDesde To 12
                wmmm = "2018/" + Format(II, "00")
                wFec = Format("01/" + Right(wmmm, 2) + "/" + Left(wmmm, 4), "dd/mm/yyyy")
                
                aa = Leerado6a("SELECT * FROM CTASXCAB " _
                            & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                            & "       CONCEPTO = '01' AND " _
                            & "            MES = '" + wmmm + "'  ")
                If aa = 0 Then
                   Db.BeginTrans
                   Db.Execute ("INSERT INTO CTASXCAB " _
                   & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, CARGOS, ABONOS, SDONEW ) " _
                   & " VALUES " _
                   & " (" + Str(wSoc) + ", '" + wmmm + "', '01', '" + wE_S + "', '" + wMon + "', " _
                   & "  " + Str(wApo) + ", 0, " + Str(wApo) + " ) ")
                   Db.CommitTrans
                End If
                
                aa = Leerado6a("SELECT * FROM CTASXDET " _
                           & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                           & "            MES = '" + wmmm + "' AND " _
                           & "       CONCEPTO = '01' AND " _
                           & "         TIPCOB = '00' AND " _
                           & "         TIPMOV = '1' ")
                If aa = 0 Then
                   Db.BeginTrans
                   Db.Execute ("INSERT INTO CTASXDET " _
                   & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
                   & "  DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW) " _
                   & " VALUES " _
                   & " (" + Str(wSoc) + ", '" + wmmm + "', '01', '00', '', '', '', '1', " _
                   & "  '" + Format(wFec, "dd/mm/yyyy") + "', 0, 0, 0, " + Str(wApo) + ", " _
                   & "  0, " + Str(wApo) + " ) ")
                   Db.CommitTrans
                Else
                   Db.BeginTrans
                   Db.Execute ("UPDATE CTASXDET " _
                   & " SET FECHA = '" + Format(wFec, "dd/mm/yyyy") + "', " _
                   & "     CARGOS = " + Str(wApo) + ", SDONEW = " + Str(wApo) + " " _
                   & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                   & "            MES = '" + wmmm + "' AND " _
                   & "       CONCEPTO = '01' AND " _
                   & "         TIPCOB = '00' AND " _
                   & "         TIPMOV = '1' ")
                   Db.CommitTrans
                End If
            Next II
         End If
   
         Db.BeginTrans
         Db.Execute ("UPDATE MAESOCIO " _
         & " SET FECING = '" + Format(wFecIng, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " ")
         Db.CommitTrans
   
         wRegAct = wRegAct + 1
         ADO8.MoveNext
      Loop
   End If


   MsgBox "Proceso Termino OK", vbExclamation
End Sub

Private Sub Command2_Click()
   Dim zz As Integer, wSoc As Integer, wFecIng As Date, _
       wMon As String, wApo As String, wE_S As String, _
       aa As Integer, II As Integer, wDesde As Integer, wHasta As Integer, _
       wAnoDes As Integer, wAnoHas As Integer, wMesDes As Integer, wMesHas As Integer, _
       wmmm As String, waaa As String, wFec As Date
   
   zz = Leerado8("SELECT M.CODSOCIO, M.CODIGO, M.INS, M.FECING, M.E_SOCIO, E.APORTE, E.MONEDA " _
                & " FROM MAESOCIO AS M INNER JOIN MAEE_SOCIO AS E ON M.E_SOCIO = E.E_SOCIO " _
                & " WHERE (FECING < '01/10/2017') AND " _
                & "       (M.E_SOCIO = 'TIT' OR " _
                & "        M.E_SOCIO = 'VIU' OR " _
                & "        M.E_SOCIO = 'HIJ' OR " _
                & "        M.E_SOCIO = 'NIE' OR " _
                & "        M.E_SOCIO = 'HER' OR " _
                & "        M.E_SOCIO = 'CIV' OR " _
                & "        M.E_SOCIO = 'CI1' OR " _
                & "        M.E_SOCIO = 'ADH' OR " _
                & "        M.E_SOCIO = 'TRA') " _
                & " ORDER BY M.CODSOCIO ")
   If zz > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wSoc = ADO8!codsocio
         wFecIng = ADO8!fecing
         wE_S = ADO8!e_socio
         wMon = ADO8!moneda
         wApo = ADO8!aporte
   
         wAnoDes = 2017
         wMesDes = 10
             
         wAnoHas = 2018
         wMesHas = 12
             
         For aa = wAnoDes To wAnoHas
             waaa = Format(aa, "0000")
         
             If aa = wAnoDes Then
                wDesde = wMesDes
             Else
                wDesde = 1
             End If
         
             If aa = wAnoHas Then
                wHasta = wMesHas
             Else
                wHasta = 12
             End If
         
             For II = wDesde To wHasta
                 wmmm = Format(II, "00")
         
                 wFec = Format("01/" + wmmm + "/" + waaa, "dd/mm/yyyy")
                
                 zz = Leerado6a("SELECT * FROM CTASXCAB " _
                             & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                             & "       CONCEPTO = '01' AND " _
                             & "            MES = '" + waaa + "/" + wmmm + "'  ")
                 If zz = 0 Then
                    Db.BeginTrans
                    Db.Execute ("INSERT INTO CTASXCAB " _
                    & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, CARGOS, ABONOS, SDONEW ) " _
                    & " VALUES " _
                    & " (" + Str(wSoc) + ", '" + waaa + "/" + wmmm + "', '01', '" + wE_S + "', '" + wMon + "', " _
                    & "  " + Str(wApo) + ", 0, " + Str(wApo) + " ) ")
                    Db.CommitTrans
                 End If
                 Set ADO6a = Nothing
                
                 zz = Leerado6a("SELECT * FROM CTASXDET " _
                            & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                            & "            MES = '" + waaa + "/" + wmmm + "' AND " _
                            & "       CONCEPTO = '01' AND " _
                            & "         TIPCOB = '00' AND " _
                            & "         TIPMOV = '1' ")
                 If zz = 0 Then
                    Db.BeginTrans
                    Db.Execute ("INSERT INTO CTASXDET " _
                    & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
                    & "  DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW) " _
                    & " VALUES " _
                    & " (" + Str(wSoc) + ", '" + waaa + "/" + wmmm + "', '01', '00', '', '', '', '1', " _
                    & "  '" + Format(wFec, "dd/mm/yyyy") + "', 0, 0, 0, " + Str(wApo) + ", " _
                    & "  0, " + Str(wApo) + " ) ")
                    Db.CommitTrans
                 End If
                 Set ADO6a = Nothing
         
             Next II
         
         Next aa
                   
         ADO8.MoveNext
      Loop
   End If
   
   MsgBox "Proceso Termino ok", vbExclamation
End Sub

Private Sub Command3_Click()
   Dim aa As Integer, wFra As String, wFec As Date, _
       wSoc As Integer, wCod As Long, wIns As Integer, _
       wCuoIni As Currency, wCuoMes As Currency, _
       wAnoCob As String, wMesCob As String, wTipCob As String, _
       wSerCob As String, wNumCob As String, wLinCob As String, wFecCob As Date, _
       wLinFra As Integer

   Db.BeginTrans
   Db.Execute ("UPDATE COBRODET " _
   & " SET NUMFRA = '', LINFRA = '', NUMOPE = '' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE FRACDET " _
   & " SET SERCOB = '', NUMCOB = '', LINCOB = '', FECCOB = NULL, ABONOS = 0 ")
   Db.CommitTrans

   aa = Leerado8("SELECT * FROM FRACCAB ORDER BY NUMERO ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wFra = ADO8!numero
         wFec = Format(ADO8!fecha, "dd/mm/yyyy")
         wSoc = ADO8!codsocio
         wCuoIni = ADO8!cuoini
         wCuoMes = ADO8!cuomes
   
         aa = Leerado7("SELECT D.ANO, D.MES, D.TIPCOB, D.SERCOB, D.NUMCOB, D.LINCOB, C.FECHA, " _
                    & "        C.CODSOCIO, M.CODIGO, M.INS, M.NOMBRE, C.MONEDA, C.GLOSA, D.CONPAGO, " _
                    & "        D.MESCOB, D.DOLARE, D.SOLESS, D.NUMFRA, D.LINFRA " _
                    & " FROM COBRODET AS D INNER JOIN COBROCAB AS C " _
                    & "   ON D.ANO = C.ANO AND D.MES = C.MES AND D.TIPCOB = C.TIPCOB AND D.SERCOB = C.SERCOB AND D.NUMCOB = C.NUMCOB " _
                    & "                    INNER JOIN MAESOCIO AS M ON C.CODSOCIO = M.CODSOCIO " _
                    & " WHERE C.FECHA >= '" + Format(wFec, "dd/mm/yyyy") + "' AND " _
                    & "       M.CODSOCIO = " + Str(wSoc) + " AND D.SERCOB = '004' AND " _
                    & "       (CONPAGO = '100' OR CONPAGO = '128') " _
                    & " ORDER BY C.FECHA")
         If aa > 0 Then
            ADO7.MoveFirst
            
            If ADO7!soless = wCuoIni Then
               wAnoCob = ADO7!ano
               wMesCob = ADO7!mes
               wTipCob = ADO7!tipcob
               wSerCob = ADO7!sercob
               wNumCob = ADO7!numcob
               wLinCob = ADO7!lincob
               wFecCob = ADO7!fecha
            
               Db.BeginTrans
               Db.Execute ("UPDATE COBRODET " _
               & " SET CONPAGO = '128', MESCOB = '" + Format(wFec, "yyyy/mm") + "', " _
               & "     NUMFRA = '" + wFra + "', LINFRA = ' 0' " _
               & " WHERE    ANO = '" + wAnoCob + "' AND " _
               & "          MES = '" + wMesCob + "' AND " _
               & "       TIPCOB = '" + wTipCob + "' AND " _
               & "       SERCOB = '" + wSerCob + "' AND " _
               & "       NUMCOB = '" + wNumCob + "' AND " _
               & "       LINCOB = '" + wLinCob + "' ")
               Db.CommitTrans
            
               Db.BeginTrans
               Db.Execute ("UPDATE FRACDET " _
               & " SET ABONOS = " + Str(wCuoIni) + ", " _
               & "     SDONEW = CARGOS - " + Str(wCuoIni) + ", " _
               & "     SERCOB = '" + wSerCob + "', " _
               & "     NUMCOB = '" + wNumCob + "', " _
               & "     LINCOB = '" + wLinCob + "', " _
               & "     FECCOB = '" + Format(wFecCob, "dd/mm/yyyy") + "' " _
               & " WHERE NUMERO = '" + wFra + "' AND LINEA = ' 0' ")
               Db.CommitTrans
            End If
               
            wLinFra = 1
            ADO7.MoveNext
            If Not ADO7.EOF Then
               Do While Not ADO7.EOF
            
                  If ADO7!soless = wCuoMes Then
                     wAnoCob = ADO7!ano
                     wMesCob = ADO7!mes
                     wTipCob = ADO7!tipcob
                     wSerCob = ADO7!sercob
                     wNumCob = ADO7!numcob
                     wLinCob = ADO7!lincob
                     wFecCob = ADO7!fecha
                  
                     Db.BeginTrans
                     Db.Execute ("UPDATE COBRODET " _
                     & " SET CONPAGO = '128', MESCOB = '" + Format(wFecCob, "yyyy/mm") + "', " _
                     & "     NUMFRA = '" + wFra + "', LINFRA = ' " + Format(wLinFra, "0") + "' " _
                     & " WHERE    ANO = '" + wAnoCob + "' AND " _
                     & "          MES = '" + wMesCob + "' AND " _
                     & "       TIPCOB = '" + wTipCob + "' AND " _
                     & "       SERCOB = '" + wSerCob + "' AND " _
                     & "       NUMCOB = '" + wNumCob + "' AND " _
                     & "       LINCOB = '" + wLinCob + "' ")
                     Db.CommitTrans
                  
                     Db.BeginTrans
                     Db.Execute ("UPDATE FRACDET " _
                     & " SET ABONOS = " + Str(wCuoMes) + ", " _
                     & "     SDONEW = CARGOS - " + Str(wCuoMes) + ", " _
                     & "     SERCOB = '" + wSerCob + "', " _
                     & "     NUMCOB = '" + wNumCob + "', " _
                     & "     LINCOB = '" + wLinCob + "', " _
                     & "     FECCOB = '" + Format(wFecCob, "dd/mm/yyyy") + "' " _
                     & " WHERE NUMERO = '" + wFra + "' AND LINEA = ' " + Format(wLinFra, "0") + "' ")
                     Db.CommitTrans
                     
                     wLinFra = wLinFra + 1
                  End If
            
                  ADO7.MoveNext
               Loop
         
            End If
         
         End If
   
         ADO8.MoveNext
      Loop
   End If

   MsgBox "Proceso Termino OK", vbExclamation
End Sub

Private Sub Command4_Click()
   Dim aa As Integer, waaa As String, wmmm As String, _
       wTip As String, wSer As String, wNum As String, wLin As String

   Dim wImp As Currency, wMes As String, wccc As String, wcon As String
   Dim wFec As Date, wCam As Currency, wSoc As Integer, wCod As Long, wIns As Integer, _
       wGlo As String, wDeu As Currency, wAde As Currency, wqqq As Currency, wMon As String


   aa = Leerado8("SELECT * FROM COBRODET " _
                & " WHERE SERCOB = '004' " _
                & " ORDER BY ANO, MES, TIPCOB, SERCOB, NUMCOB, LINCOB ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         waaa = ADO8!ano
         wmmm = ADO8!mes
         wTip = ADO8!tipcob
         wSer = ADO8!sercob
         wNum = ADO8!numcob
         wLin = ADO8!lincob
         wMes = ADO8!mescob
         wcon = ADO8!conpago
         wccc = ADO8!concepto
         wMes = ADO8!mescob
         wImp = ADO8!importe
      
'         If wLin = "02" Then
'            MsgBox "Linea 02"
'         End If
      
         aa = Leerado7("SELECT * FROM COBROCAB " _
                    & " WHERE    ANO = '" + waaa + "' AND " _
                    & "          MES = '" + wmmm + "' AND " _
                    & "       TIPCOB = '" + wTip + "' AND " _
                    & "       SERCOB = '" + wSer + "' AND " _
                    & "       NUMCOB = '" + wNum + "' ")
         If aa = 0 Then
            MsgBox "Cobro Sin Cabecera", vbExclamation
            Exit Sub
         End If
         wFec = Format(ADO7!fecha, "dd/mm/yyyy")
         wGlo = ADO7!glosa
         wCam = ADO7!tipcam
         wSoc = ADO7!codsocio
         wMon = ADO7!moneda
         Set ADO7 = Nothing

         aa = Leerado7("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSoc) + " ")
         If aa > 0 Then
            wCod = ADO7!codigo
            wIns = ADO7!ins
         End If
         Set ADO7 = Nothing

         
         If Len(Trim(wccc)) > 0 Then
            Db.BeginTrans
            Db.Execute ("DELETE FROM CTASXDET " _
            & " WHERE   TIPMOV = '2' AND " _
            & "         TIPCOB = '03' AND " _
            & "         SERCOB = '" + wSer + "' AND " _
            & "         NUMCOB = '" + wNum + "' AND " _
            & "         LINCOB = '" + wLin + "' ")
            Db.CommitTrans
         
            aa = Leerado7a("SELECT * FROM CTASXCAB " _
                         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                         & "            MES = '" + wMes + "' AND " _
                         & "       CONCEPTO = '" + wccc + "' ")
            If aa = 0 Then
               wqqq = CreaAporteMes(wSoc, wMes, wccc, 1)
            End If
               
            Db.BeginTrans
            Db.Execute ("INSERT INTO CTASXDET " _
            & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
            & "  TIPCAM, DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW ) " _
            & " VALUES " _
            & " (" + Str(wSoc) + ", '" + wMes + "', '" + wccc + "', " _
            & "  '03', '" + wSer + "', '" + wNum + "', '" + wLin + "', " _
            & "  '2', '" + Format(wFec, "dd/mm/yyyy") + "', " + Str(wCam) + ", " _
            & "  " + Str(ADO8!dolare) + ", " + Str(ADO8!soless) + ", " _
            & "  0, 0, " + Str(ADO8!abonos) + ", 0 ) ")
            Db.CommitTrans
         
            Call ActualizaSaldos(wSoc, wMes, wccc)
         End If
         
         Db.BeginTrans
         Db.Execute ("DELETE FROM ZZZ_MRECIBOS " _
         & " WHERE YEAR(FECHA_PAGO) = " + Str(Year(wFec)) + " AND " _
         & "       SERIE = '" + wSer + "' AND " _
         & "       NRO_COMP = " + Str(Val(wNum)) + " and " _
         & "       LINCOB = '" + wLin + "' ")
         Db.CommitTrans
      
         
         aa = Leerado7a("SELECT * FROM ZZZ_MRECIBOS " _
                   & " WHERE    SERIE = '" + wSer + "' AND " _
                   & "       NRO_COMP = " + Str(Val(wNum)) + " AND " _
                   & "       YEAR(FECHA_PAGO) = " + Str(Val(wFec)) + " AND " _
                   & "      LINCOB = '" + wLin + "' ")
         If aa = 0 Then
            Db.BeginTrans
            Db.Execute ("INSERT INTO ZZZ_MRECIBOS " _
            & " (CODIGO, INS, CONCEPTO, SERIE, AUXILIAR, NRO_COMP, MONTO, MONEDA, T_CAMBIO, " _
            & "  FECHA_PAGO, FECHA_CADU, OBS, D_IMPOR, DEUDA_PT2, DINS_CER, ADELANTO, " _
            & "  MARCA1, MARCA2, MARCA3, MARCA4, OBS1, LINCOB ) " _
            & " VALUES " _
            & " (" + Str(wCod) + ", " + Str(wIns) + ", " + Str(Val(wcon)) + ", '" + wSer + "', 0, " _
            & "  " + Str(Val(wNum)) + ", " + Str(wImp) + ", " _
            & "  '" + wMon + "', " + Str(wCam) + ", " _
            & "  '" + Format(wFec, "dd/mm/yyyy") + "', null, '" + Left(wGlo, 50) + "', " _
            & "  '', " + Str(wDeu) + ", 0, " + Str(wAde) + ", " _
            & "  '" + Format(Date, "dd/mm/yyyy") + "', 'N', '" + wcodusu + "', " _
            & "  '" + Format(Time, "hh:mm:ss") + "', '', '" + wLin + "' ) ")
            Db.CommitTrans
         Else
            Db.BeginTrans
            Db.Execute ("UPDATE ZZZ_MRECIBOS " _
            & " SET CODIGO = " + Str(wCod) + ", INS = " + Str(wIns) + ", CONCEPTO = " + Str(Val(wcon)) + ", " _
            & "     AUXILIAR = 0, MONTO = " + Str(wImp) + ", " _
            & "     MONEDA = '" + wMon + "', " _
            & "     T_CAMBIO = " + Str(wCam) + ", FECHA_PAGO = '" + Format(wFec, "dd/mm/yyyy") + "', " _
            & "     FECHA_CADU = null, OBS = '" + Left(wGlo, 50) + "', D_IMPOR = '', " _
            & "     DEUDA_PT2 = " + Str(wDeu) + ", DINS_CER = 0, ADELANTO = " + Str(wAde) + ", " _
            & "     MARCA1 = '" + Format(Date, "dd/mm/yyyy") + "', MARCA2 = 'N', " _
            & "     MARCA3 = '" + wcodusu + "', MARCA4 = '" + Format(Time, "hh:mm:ss") + "', " _
            & "     OBS1 = '' " _
            & " WHERE    SERIE = '" + wSer + "' AND " _
            & "       NRO_COMP = " + Str(Val(wNum)) + " AND " _
            & "       YEAR(FECHA_PAGO) = " + Str(Val(wanocia)) + " AND " _
            & "       LINCOB = '" + wLin + "' ")
            Db.CommitTrans
         End If
         Set ADO7a = Nothing
         
         ADO8.MoveNext
      Loop
   End If


   MsgBox "Proceso Termino OK", vbExclamation
End Sub

Private Sub Command5_Click()
   Dim aa As Long, _
       wSoc As Integer, wConcep As String, wTipCob As String, wSerCob As String, wNumCob As String, wLinCob As String, _
       wTipMov As String, wFec As Date, wAbono As Currency

   aa = Leerado8("select CODSOCIO, concepto, tipcob, numcob, lincob, TIPMOV, fecha, abonos, COUNT(*) as cant " _
                & " From ctasxdet " _
                & " where TIPCOB = '03' " _
                & " group by CODSOCIO, concepto, tipcob, numcob, lincob, TIPMOV, fecha, abonos " _
                & " having COUNT(*) = 2 " _
                & " order by CODSOCIO, concepto, tipcob, numcob, lincob, TIPMOV, fecha, abonos")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wSoc = ADO8!codsocio
         wConcep = ADO8!concepto
         wTipCob = ADO8!tipcob
         wNumCob = ADO8!numcob
         wLinCob = ADO8!lincob
         wTipMov = ADO8!tipmov
         wFec = Format(ADO8!fecha, "dd/mm/yyyy")
         wAbono = ADO8!abonos
   
         aa = Leerado7("SELECT *  " _
                    & " FROM CTASXDET " _
                    & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                    & "       CONCEPTO = '" + wConcep + "' AND " _
                    & "       TIPCOB = '" + wTipCob + "' AND " _
                    & "       SERCOB = '' AND " _
                    & "       NUMCOB = '" + wNumCob + "' AND " _
                    & "       LINCOB = '" + wLinCob + "' AND " _
                    & "       TIPMOV = '" + wTipMov + "' AND " _
                    & "       FECHA = '" + Format(wFec, "dd/mm/yyyy") + "' AND " _
                    & "       ABONOS = " + Str(wAbono) + " ")
         If aa = 1 Then
            Db.BeginTrans
            Db.Execute ("DELETE FROM CTASXDET " _
            & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
            & "       CONCEPTO = '" + wConcep + "' AND " _
            & "       TIPCOB = '" + wTipCob + "' AND " _
            & "       SERCOB = '' AND " _
            & "       NUMCOB = '" + wNumCob + "' AND " _
            & "       LINCOB = '" + wLinCob + "' AND " _
            & "       TIPMOV = '" + wTipMov + "' AND " _
            & "       FECHA = '" + Format(wFec, "dd/mm/yyyy") + "' AND " _
            & "       ABONOS = " + Str(wAbono) + " ")
            Db.CommitTrans
         End If
   
         ADO8.MoveNext
      Loop
   End If
   MsgBox "Proceso Termino OK", vbExclamation
End Sub

Private Sub Command6_Click()
   Dim aa As Integer, II As Integer, wFec As Date, _
       wAno As String, wMes As String, wTip As String, wSer As String, wNum As String
   
   wAno = "2018"
   wMes = "12"
   wTip = "2"
   wSer = "004"
   wFec = Format("02/12/2018", "dd/mm/yyyy")
   
   For II = 8052 To 8228
       wNum = Format(II, "0000000000")
   
       aa = Leerado8("SELECT * FROM COBROCAB " _
                    & " WHERE    ANO = '2018' AND " _
                    & "          MES = '12' AND " _
                    & "       TIPCOB = '2' AND " _
                    & "       SERCOB = '004' AND " _
                    & "       NUMCOB = '" + wNum + "' ")
       If aa = 0 Then
          Db.BeginTrans
          Db.Execute ("INSERT INTO COBROCAB " _
          & " (ANO, MES, TIPCOB, SERCOB, NUMCOB, FECHA, MONEDA, IMPORTE, GLOSA, CODSOCIO, TIPCAM, DOLARE, " _
          & "  SOLESS, FORPAG, USU ) " _
          & " VALUES " _
          & " ('" + wAno + "', '" + wMes + "', '" + wTip + "', '" + wSer + "', '" + wNum + "', " _
          & "  '" + Format(wFec, "dd/mm/yyyy") + "', 'S', 0, 'DOCUMENTO ANULADO', " _
          & "  0, 0, 0, 0, '', '" + wcodusu + "'  ) ")
          Db.CommitTrans
       End If
   
   
       aa = Leerado8("SELECT * FROM COBRODET " _
                    & " WHERE    ANO = '2018' AND " _
                    & "          MES = '12' AND " _
                    & "       TIPCOB = '2' AND " _
                    & "       SERCOB = '004' AND " _
                    & "       NUMCOB = '" + wNum + "' ")
       If aa = 0 Then
          Db.BeginTrans
          Db.Execute ("INSERT INTO COBRODET " _
          & " (ANO, MES, TIPCOB, SERCOB, NUMCOB, LINCOB, MESCOB, CONPAGO, DOLARE, SOLESS, MONDOC, " _
          & "  SDOOLD, CARGOS, ABONOS, SDONEW, IMPORTE, CONCEPTO, PARIENTE, LINPARIE, NOMBRE, " _
          & "  NUMFRA, LINFRA, NUMOPE) " _
          & " VALUES " _
          & " ('" + wAno + "', '" + wMes + "', '" + wTip + "', '" + wSer + "', '" + wNum + "', " _
          & "  '01', '', '', 0, 0, '', 0, 0, 0, 0, 0, '', '', '', '', '', '', '' ) ")
          Db.CommitTrans
       End If
   
   Next II
   
   MsgBox "Proceso Termino OK", vbExclamation
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   Select Case ColIndex
   Case 0
        ADO1.Sort = "GRADO"
   Case 1
        ADO1.Sort = "NOMBRE"
   End Select
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If ACCION = 0 Then
      refrescar
   End If
End Sub

Private Sub Form_Activate()
   frmMaeGrado.Left = (Screen.Width - Width) \ 2
   frmMaeGrado.Top = 0
   optTodos.Value = True
   
   Dim a As Integer
   a = Leerado8("SELECT * FROM MAEGRADOGRUPO ORDER BY GRADOGRUPO ")
   If a > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
   
         cmbGrupo.AddItem Format(ADO8!gradogrupo, "@@@") + " " + ADO8!nombre
   
         ADO8.MoveNext
      Loop
   End If
   Set ADO8 = Nothing
   
   LlenaCab
   LlenaCab1
   Limpiar
   refrescar
   editar False
   
   DataGrid1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set ADO1 = Nothing
End Sub

Private Sub LlenaCab()
   Dim a As Integer
   
'   optTodos.Value = True
   txtFiltrar.Text = ""
   
   a = Leerado("SELECT GRADO, NOMBRE, GRADOGRUPO " _
                & " FROM MAEGRADO " _
                & " ORDER BY GRADO ")
   Set DataGrid1.DataSource = ADO1
End Sub

Private Sub LlenaCab1()
   DataGrid1.Columns(0).Width = 500
   DataGrid1.Columns(0).Alignment = dbgRight
   DataGrid1.Columns(0).Caption = "CODIGO"
    
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Width = 6000
   DataGrid1.Columns(1).Caption = "NOMBRE"

   DataGrid1.Columns(2).Width = 500
   DataGrid1.Columns(2).Alignment = dbgRight
   DataGrid1.Columns(2).Caption = "GRUPO"
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

Private Sub SSTab1_Click(PreviousTab As Integer)
   If PreviousTab = 0 Then
      If ACCION = 0 Then
      End If
   Else
      ACCION = 0
   End If
End Sub

Private Sub txtCodigo_GotFocus()
    txtCodigo.SelStart = 0
    txtCodigo.SelLength = Len(Trim(txtCodigo))
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
   If KeyAscii = 13 Then
      If txtCodigo = "" Then
         MsgBox "Codigo En Blanco", vbExclamation
         Exit Sub
      End If
      aa = Leerado8("SELECT * FROM MAEGRADO WHERE GRADO = " + Str(Val(txtCodigo.Text)) + " ")
      If aa > 0 Then
         MsgBox "Codigo de Grado Ya Existe", vbExclamation
         txtCodigo.Text = ""
         Exit Sub
      End If
      txtNombre.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFiltrar_GotFocus()
   txtFiltrar.SelStart = 0
   txtFiltrar.SelLength = Len(Trim(txtFiltrar.Text))
End Sub

Private Sub txtFiltrar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFiltrar.Text = "" Then
         MsgBox "Filtro En Blanco", vbExclamation
         Exit Sub
      End If
      ADO1.Filter = "NOMBRE LIKE '%" & Trim(txtFiltrar) & "%' "
      refrescar
      DataGrid1.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtNombre_GotFocus()
    txtNombre.SelStart = 0
    txtNombre.SelLength = Len(Trim(txtNombre))
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtNombre = "" Then
         MsgBox "Nombre En Blanco", vbExclamation
         Exit Sub
      End If
      cmbGrupo.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

