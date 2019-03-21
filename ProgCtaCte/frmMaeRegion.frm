VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmMaeRegion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Grupo de Regiones Policiales"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   10950
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
      Left            =   8520
      TabIndex        =   20
      Top             =   3000
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
      Left            =   8520
      TabIndex        =   18
      Top             =   2160
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
      Top             =   240
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
      Left            =   240
      TabIndex        =   8
      Top             =   7320
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
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8295
      Begin VB.ComboBox cmbGrupo 
         Height          =   315
         ItemData        =   "frmMaeRegion.frx":0000
         Left            =   1800
         List            =   "frmMaeRegion.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1320
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
         Left            =   1155
         TabIndex        =   26
         Top             =   1320
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
         Left            =   840
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
         Width           =   960
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
      Left            =   9480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1095
   End
   Begin Crystal.CrystalReport Crys1 
      Left            =   10320
      Top             =   4080
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
      Height          =   3375
      Left            =   480
      TabIndex        =   7
      Top             =   3840
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
      Caption         =   "TABLA DE REGIONES"
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
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   8040
      Width           =   8295
   End
End
Attribute VB_Name = "frmMaeRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACCION As Byte
Dim Marca As Variant, wcia As String
    
Sub Limpiar()
   txtCodigo.Text = ""
   txtNombre.Text = ""
   cmbGrupo.ListIndex = 0
End Sub

Sub refrescar()
   If ADO1.BOF Then Exit Sub
   If ADO1.EOF Then Exit Sub
   txtCodigo.Text = ADO1!region
   txtNombre.Text = IIf(IsNull(ADO1!Nombre), "", ADO1!Nombre)
   cmbGrupo.ListIndex = ADO1!regiongrupo - 1
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
   
   aa = Leerado8("SELECT * FROM MAEREGION WHERE REGION = '" + wCod + "' ")
   If aa = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO MAEREGION " _
      & " (REGION, NOMBRE) " _
      & " VALUES " _
      & " ('" + wCod + "', '" + wNom + "' ) ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE MAEREGION " _
      & " SET NOMBRE = '" + wNom + "' " _
      & " WHERE REGION = '" + wCod + "' ")
      Db.CommitTrans
   End If
   ADO1.Requery
   LlenaCab1
   ADO1.Find "REGION = '" + wCod + "' "
   MsgBox "REGION '" + wCod + " " + wNom + "'" + vbNewLine + _
          "Grabado OK", vbExclamation
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub editar(estado As Boolean)
   txtCodigo.Enabled = estado
   txtNombre.Enabled = estado
   
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
   
   Dim wcon As String, wNom As String, wNew As String, aa As Integer
   wcon = ADO1!region
   wNom = Trim(ADO1!Nombre)
   wNew = ""
   ADO1.MoveNext
   If Not ADO1.EOF Then
      wNew = ADO1!region
   Else
      ADO1.MovePrevious
      ADO1.MovePrevious
      If ADO1.BOF Then
         wNew = ""
      Else
         wNew = ADO1!region
      End If
   End If
   
   aa = Leerado8("SELECT * FROM MAESOCIO WHERE REGION = '" + wcon + "' ")
   If aa > 0 Then
      MsgBox "Región Tiene Movimiento, No Se Pude Borrar", vbExclamation
      Exit Sub
   End If
   
   If MsgBox("¿Esta seguro de borrar Codigo '" + wcon + "'?", vbYesNo + vbDefaultButton2 + vbQuestion, "Advertencia") = vbYes Then
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEREGION WHERE REGION = '" + wcon + "' ")
      Db.CommitTrans
      
      ADO1.Requery
      LlenaCab
      LlenaCab1
      Limpiar
      refrescar
      
      If Len(Trim(wNew)) <> 0 Then
         ADO1.Find "REGION='" + wNew + "' "
      End If
      MsgBox "Región '" + wcon + "' " + wNom + vbNewLine + _
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
   
   Dim aa As Integer, I As Integer, Heading(2) As String, wreg As Integer, wTot As Integer
   Dim wNom As String
   Heading(0) = "CODIGO"
   Heading(1) = "NOMBRE"
   Heading(2) = "GRUPO"
   aa = Leerado3("SELECT * FROM MAEREGION ORDER BY REGION ")
   If aa > 0 Then
      wTot = aa
      Set objExcel = New Excel.Application
      objExcel.SheetsInNewWorkbook = 1
      objExcel.Workbooks.Add
      With objExcel
           .Range(.Cells(1, 1), .Cells(2, 1)).Font.Bold = True
           .Range(.Cells(3, 1), .Cells(3, 3)).Borders.LineStyle = xlContinuous
           .Range(.Cells(3, 1), .Cells(3, 3)).Font.Bold = True
           .Cells(1, 1) = wnomcia
           .Cells(2, 1) = "MAESTRO DE REGIONES"
           For I = 1 To 3 Step 1
               .Cells(3, I) = Heading(I - 1)
           Next
           objExcel.Columns("A").ColumnWidth = 15
           objExcel.Columns("B").ColumnWidth = 50
           objExcel.Columns("C").ColumnWidth = 15
      End With
      V = 4
      H = 1
      wreg = 1
      Do While Not ADO3.EOF
         lblMensaje.Caption = "Traslando a EXCEL - Registro " + Format(wreg, "####0") + " / " + Format(wTot, "####0")
         lblMensaje.Refresh
         objExcel.Cells(V, H + 0) = ADO3!region
         objExcel.Cells(V, H + 1) = IIf(IsNull(ADO3!Nombre), "", ADO3!Nombre)
         objExcel.Cells(V, H + 2) = ADO3!regiongrupo
         
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
      If Leerado2("SELECT * FROM MAEREGION WHERE REGION = '" + wCod + "' ") > 0 Then
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

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   Select Case ColIndex
   Case 0
        ADO1.Sort = "REGION"
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
   frmMaeRegion.Left = (Screen.Width - Width) \ 2
   frmMaeRegion.Top = 0
   optTodos.Value = True
   Dim a As Integer
   
   a = Leerado8("SELECT * FROM MAEREGIONGRUPO ORDER BY REGIONGRUPO ")
   If a > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
   
         cmbGrupo.AddItem Format(ADO8!regiongrupo, "@@@") + " " + ADO8!Nombre
   
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
   
   a = Leerado("SELECT REGION, NOMBRE, REGIONGRUPO " _
                & " FROM MAEREGION " _
                & " ORDER BY REGION ")
   Set DataGrid1.DataSource = ADO1
End Sub

Private Sub LlenaCab1()
   DataGrid1.Columns(0).Width = 500
   DataGrid1.Columns(0).Alignment = dbgCenter
   DataGrid1.Columns(0).Caption = "CODIGO"
    
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Width = 6000
   DataGrid1.Columns(1).Caption = "NOMBRE"

   DataGrid1.Columns(2).Width = 1200
   DataGrid1.Columns(2).Alignment = dbgCenter
   DataGrid1.Columns(2).Caption = "GRUPO REGION"
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
      aa = Leerado8("SELECT * FROM MAEREGION WHERE REGION = '" + txtCodigo.Text + "' ")
      If aa > 0 Then
         MsgBox "Codigo de Región Ya Existe", vbExclamation
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

