VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmServGlosas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar Glosas de Tesoreria"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   16185
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
      Left            =   12840
      TabIndex        =   15
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cndOtro 
      Caption         =   "&Otra Consulta"
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
      Left            =   10800
      TabIndex        =   14
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
      Height          =   495
      Left            =   14160
      TabIndex        =   13
      Top             =   7080
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   3615
      Left            =   12240
      TabIndex        =   11
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6376
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
      Caption         =   "DIECO"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6135
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   10821
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
      Caption         =   "TESORERIA"
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
   Begin VB.TextBox txtNumdoc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8160
      MaxLength       =   8
      TabIndex        =   3
      Top             =   300
      Width           =   975
   End
   Begin VB.TextBox txtIns 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7800
      MaxLength       =   1
      TabIndex        =   2
      Top             =   300
      Width           =   375
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   6840
      MaxLength       =   8
      TabIndex        =   1
      Top             =   300
      Width           =   975
   End
   Begin VB.TextBox txtCodSocio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      MaxLength       =   9
      TabIndex        =   0
      Top             =   300
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   3255
      Left            =   12240
      TabIndex        =   12
      Top             =   3720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5741
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
      Caption         =   "CAJA MP"
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
   Begin VB.Label lblConcepto 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   17
      Top             =   6840
      Width           =   5175
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Concepto"
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
      TabIndex        =   16
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "D.N.I."
      Height          =   195
      Left            =   8160
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Ins"
      Height          =   195
      Left            =   7800
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Codofin"
      Height          =   195
      Left            =   6840
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de Socio"
      Height          =   195
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblCodSocio 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   300
      Width           =   5775
   End
   Begin VB.Label Label3 
      Caption         =   "Cod.Socio"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmServGlosas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGrabar_Click()
   Dim aa As Long, wSoc As Integer, wCod As Long, wIns As Integer, _
       wFec As Date, wSer As String, wDoc As Long, wGlo As String, wItem As Long
   
   wSoc = Val(txtCodSocio.Text)
   wCod = Val(txtCodigo.Text)
   wIns = Val(txtIns.Text)
   
   aa = Leerado8("SELECT * FROM TMP_TESOR WHERE USU = '" + wcodusu + "' ORDER BY FECHA_PAGO, SERIE, NRO_COMP ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wFec = Format(ADO8!fecha_pago, "dd/mm/yyyy")
         wSer = ADO8!serie
         wDoc = ADO8!nro_comp
         wGlo = Trim(ADO8!obs)
         wItem = ADO8!Item
         
         Db.BeginTrans
         Db.Execute ("UPDATE ZZZ_MRECIBOS " _
         & " SET OBS = '" + wGlo + "' " _
         & " WHERE     CODIGO = " + Str(wCod) + " AND " _
         & "              INS = " + Str(wIns) + " AND " _
         & "             ITEM = " + Str(wItem) + " ")
         Db.CommitTrans
         
         ADO8.MoveNext
      Loop
   End If

   MsgBox "Grabado OK"
   Unload Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cndOtro_Click()

   txtCodSocio.Text = ""
   txtCodigo.Text = ""
   txtIns.Text = ""
   txtNumDoc.Text = ""
   
   Call Limpiar
   
   txtCodSocio.SetFocus
End Sub

Private Sub DataGrid1_GotFocus()
   DataGrid1.Row = 0
   DataGrid1.col = 5
   DataGrid1.SelStart = 0
   DataGrid1.SelLength = 3
   
'   If Len(Trim(DataGrid1.Text)) > 0 Then
'      DataGrid1.SelLength = Len(Trim(DataGrid1.Text))
'   End If
   DataGrid1.Refresh
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim wvariable As String
    
    On Error GoTo err
    Select Case KeyCode
    Case 116  '' F5
         Select Case DataGrid1.col
         Case 5
              xlista = "CO2"
              xseleccion = ""
              zSerCaj = ADO2!serie
              zMonCaj = IIf(ADO2!moneda = "S/.", "S", "D")
              frmSeleccion.Show 1
              If xseleccion <> "" Then
                 DataGrid1.col = 5
                 DataGrid1.Text = xseleccion
                 ADO2!conpago = xseleccion
                 ADO2.Update
              End If
         End Select
    Case 40  ' DOWN
         If ACCION = 1 Or ACCION = 2 Then
            
         wvariable = DataGrid1.Text
         
         Select Case DataGrid1.col
         Case 0
              ADO2!fecha_pago = IIf(IsNull(wvariable), Null, wvariable)
         Case 1
              ADO2!serie = IIf(IsNull(wvariable), "", wvariable)
         Case 2
              ADO2!nro_comp = IIf(IsNull(wvariable), 0, wvariable)
         Case 3
              ADO2!moneda = IIf(IsNull(wvariable), "", Val(wvariable))
         Case 4
              ADO2!monto = IIf(IsNull(wvariable), 0, wvariable)
         Case 5
              ADO2!conpago = IIf(IsNull(wvariable), "", wvariable)
         Case 6
              ADO2!obs = IIf(IsNull(wvariable), "", wvariable)
         End Select
    
         Select Case DataGrid1.col
         Case 0
              DataGrid1.Text = ADO2!fecha_pago
         Case 1
              DataGrid1.Text = IIf(IsNull(ADO2!serie), "", ADO2!serie)
         Case 2
              DataGrid1.Text = IIf(IsNull(ADO2!nro_comp), "", ADO2!nro_comp)
         Case 3
              DataGrid1.Text = IIf(IsNull(ADO2!moneda), "", ADO2!moneda)
         Case 4
              DataGrid1.Text = IIf(IsNull(ADO2!monto), 0, ADO2!monto)
         Case 5
              DataGrid1.Text = IIf(IsNull(ADO2!conpago), "", ADO2!conpago)
         Case 6
              DataGrid1.Text = IIf(IsNull(ADO2!obs), "", ADO2!obs)
         End Select
            
'         If ADO2.AbsolutePosition = ADO2.RecordCount And Len(Trim(ADO2!fecha_pago)) > 0 Then
'            creadet
'            totaldet
'         End If
         End If
    Case 37 ' Retroceder
         wvariable = DataGrid1.Text
         
         Select Case DataGrid1.col
         Case 0
              ADO2!fecha_pago = IIf(IsNull(wvariable), Null, wvariable)
         Case 1
              ADO2!serie = IIf(IsNull(wvariable), "", wvariable)
         Case 2
              ADO2!nro_comp = IIf(IsNull(wvariable), 0, wvariable)
         Case 3
              ADO2!moneda = IIf(IsNull(wvariable), "", Val(wvariable))
         Case 4
              ADO2!monto = IIf(IsNull(wvariable), 0, wvariable)
         Case 5
              ADO2!conpago = IIf(IsNull(wvariable), "", wvariable)
         Case 6
              ADO2!obs = IIf(IsNull(wvariable), "", wvariable)
         End Select
         
         If DataGrid1.col = 1 Then
            If DataGrid1.Row > 0 Then
               DataGrid1.Row = DataGrid1.Row - 1
            End If
            DataGrid1.col = 0
         End If
         
         Select Case DataGrid1.col
         Case 0
              DataGrid1.Text = ADO2!fecha_pago
         Case 1
              DataGrid1.Text = IIf(IsNull(ADO2!serie), "", ADO2!serie)
         Case 2
              DataGrid1.Text = IIf(IsNull(ADO2!nro_comp), "", ADO2!nro_comp)
         Case 3
              DataGrid1.Text = IIf(IsNull(ADO2!moneda), "", ADO2!moneda)
         Case 4
              DataGrid1.Text = IIf(IsNull(ADO2!monto), 0, ADO2!monto)
         Case 5
              DataGrid1.Text = IIf(IsNull(ADO2!conpago), "", ADO2!conpago)
         Case 6
              DataGrid1.Text = IIf(IsNull(ADO2!obs), "", ADO2!obs)
         End Select
         
    Case 38 ' Subir
         wvariable = DataGrid1.Text
         
         Select Case DataGrid1.col
         Case 0
              ADO2!fecha_pago = IIf(IsNull(wvariable), Null, wvariable)
         Case 1
              ADO2!serie = IIf(IsNull(wvariable), "", wvariable)
         Case 2
              ADO2!nro_comp = IIf(IsNull(wvariable), 0, wvariable)
         Case 3
              ADO2!moneda = IIf(IsNull(wvariable), "", Val(wvariable))
         Case 4
              ADO2!monto = IIf(IsNull(wvariable), 0, wvariable)
         Case 5
              ADO2!conpago = IIf(IsNull(wvariable), "", wvariable)
         Case 6
              ADO2!obs = IIf(IsNull(wvariable), "", wvariable)
         End Select
    
         Select Case DataGrid1.col
         Case 0
              DataGrid1.Text = ADO2!fecha_pago
         Case 1
              DataGrid1.Text = IIf(IsNull(ADO2!serie), "", ADO2!serie)
         Case 2
              DataGrid1.Text = IIf(IsNull(ADO2!nro_comp), "", ADO2!nro_comp)
         Case 3
              DataGrid1.Text = IIf(IsNull(ADO2!moneda), "", ADO2!moneda)
         Case 4
              DataGrid1.Text = IIf(IsNull(ADO2!monto), 0, ADO2!monto)
         Case 5
              DataGrid1.Text = IIf(IsNull(ADO2!conpago), "", ADO2!conpago)
         Case 6
              DataGrid1.Text = IIf(IsNull(ADO2!obs), "", ADO2!obs)
         End Select
    
    Case 39 ' Avanzar
         wvariable = DataGrid1.Text
         
         Select Case DataGrid1.col
         Case 0
              ADO2!fecha_pago = IIf(IsNull(wvariable), Null, wvariable)
         Case 1
              ADO2!serie = IIf(IsNull(wvariable), "", wvariable)
         Case 2
              ADO2!nro_comp = IIf(IsNull(wvariable), 0, wvariable)
         Case 3
              ADO2!moneda = IIf(IsNull(wvariable), "", Val(wvariable))
         Case 4
              ADO2!monto = IIf(IsNull(wvariable), 0, wvariable)
         Case 5
              ADO2!conpago = IIf(IsNull(wvariable), "", wvariable)
         Case 6
              ADO2!obs = IIf(IsNull(wvariable), "", wvariable)
         End Select
         
'         If DataGrid2.col = 5 Then
'            If Val(ADO2!lincob) < ADO2.RecordCount Then
'               DataGrid2.Row = DataGrid2.Row + 1
'            End If
'            DataGrid2.col = 0
'         End If
          
         Select Case DataGrid1.col
         Case 0
              DataGrid1.Text = ADO2!fecha_pago
         Case 1
              DataGrid1.Text = IIf(IsNull(ADO2!serie), "", ADO2!serie)
         Case 2
              DataGrid1.Text = IIf(IsNull(ADO2!nro_comp), "", ADO2!nro_comp)
         Case 3
              DataGrid1.Text = IIf(IsNull(ADO2!moneda), "", ADO2!moneda)
         Case 4
              DataGrid1.Text = IIf(IsNull(ADO2!monto), 0, ADO2!monto)
         Case 5
              DataGrid1.Text = IIf(IsNull(ADO2!conpago), "", ADO2!conpago)
         Case 6
              DataGrid1.Text = IIf(IsNull(ADO2!obs), "", ADO2!obs)
         End Select
    
'    Case 45 ' Insertar
'         If Len(Trim(ADO2!conpago)) > 0 Then
'            insertlinea ADO2!lincob
'            totaldet
'            DataGrid2.col = 1
'            DataGrid2.SelStart = 0
'            DataGrid2.SelLength = Len(Trim(DataGrid2.Text))
'         End If
    End Select
    Exit Sub
err:
    MsgBox Format(err.Number, "00000000000") + " " + err.Description
    Resume Next
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    Dim c As Integer
    Dim wvariable As String, wvariable2 As String
    
    On Error GoTo err
    Select Case KeyAscii
    Case 13
       Select Case DataGrid1.col
       Case 0  ' Linea
            DataGrid1.col = 5
       Case 1  ' Linea
            DataGrid1.col = 5
       Case 2  ' Linea
            DataGrid1.col = 5
       Case 3  ' Linea
            DataGrid1.col = 5
       Case 4  ' Linea
            DataGrid1.col = 5
       Case 5  ' ConPago
            DataGrid1.col = 5
       Case 6  ' Obs
            wvariable = Trim(DataGrid1.Text)
            DataGrid1.Text = wvariable
            ADO2!obs = Trim(Left(wvariable, 100))
            DataGrid1.col = 1
       End Select
       wvariable2 = IIf(IsNull(ADO2.Fields(DataGrid1.col)), "", Trim(ADO2.Fields(DataGrid1.col)))
       DataGrid1.Text = wvariable2
       ADO2.Update
       DataGrid1.Refresh
    Case Else
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
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
         If DataGrid1.col = 0 Then
            DataGrid1.col = 6
         End If
         
         Select Case DataGrid1.col
         Case 0
              DataGrid1.Text = ADO2!fecha_pago
         Case 1
              DataGrid1.Text = IIf(IsNull(ADO2!serie), "", ADO2!serie)
         Case 2
              DataGrid1.Text = IIf(IsNull(ADO2!nro_comp), "", ADO2!nro_comp)
         Case 3
              DataGrid1.Text = IIf(IsNull(ADO2!moneda), "", ADO2!moneda)
         Case 4
              DataGrid1.Text = IIf(IsNull(ADO2!monto), 0, ADO2!monto)
         Case 5
              DataGrid1.Text = IIf(IsNull(ADO2!conpago), "", ADO2!conpago)
         Case 6
              DataGrid1.Text = IIf(IsNull(ADO2!obs), "", ADO2!obs)
         End Select
   Case 38  ' UP
   
   Case 39  ' AVANZAR

        If DataGrid1.col = 0 Then
           DataGrid1.col = 1
        End If
          
         Select Case DataGrid1.col
         Case 0
              DataGrid1.Text = ADO2!fecha_pago
         Case 1
              DataGrid1.Text = IIf(IsNull(ADO2!serie), "", ADO2!serie)
         Case 2
              DataGrid1.Text = IIf(IsNull(ADO2!nro_comp), "", ADO2!nro_comp)
         Case 3
              DataGrid1.Text = IIf(IsNull(ADO2!moneda), "", ADO2!moneda)
         Case 4
              DataGrid1.Text = IIf(IsNull(ADO2!monto), 0, ADO2!monto)
         Case 5
              DataGrid1.Text = IIf(IsNull(ADO2!conpago), "", ADO2!conpago)
         Case 6
              DataGrid1.Text = IIf(IsNull(ADO2!obs), "", ADO2!obs)
         End Select
        
   Case 40  ' DOWN
   
   End Select
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   LabelCab
End Sub

Private Sub Form_Activate()
   frmServGlosas.Left = (Screen.Width - Width) \ 2
   frmServGlosas.Top = 0
   
   txtCodSocio.Text = ""
   txtCodigo.Text = ""
   txtIns.Text = ""
   txtNumDoc.Text = ""
   
   Limpiar
   txtCodSocio.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Limpiar
End Sub

Private Sub Limpiar()
   Set DataGrid1.DataSource = Nothing
   Set DataGrid2.DataSource = Nothing
   Set DataGrid3.DataSource = Nothing
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_DIECO WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_CAJMP WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_TESOR WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
End Sub

Private Sub LlenaCab()
   Dim aa As Integer, wSoc As Integer, wCod As Long, wIns As Integer
   wSoc = Val(txtCodSocio.Text)
   wCod = Val(txtCodigo.Text)
   wIns = Val(txtIns.Text)
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECO " _
   & " (CODSOCIO, MES, NETSOCIO, DSCSOCIO, DIFSOCIO, USU ) " _
   & " SELECT " _
   & "  CODSOCIO, MES, NETSOCIO, DSCSOCIO, DIFSOCIO, '" + wcodusu + "' " _
   & " FROM DIECOCAB " _
   & " WHERE CODSOCIO = " + Str(wSoc) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECO " _
   & " (CODSOCIO, MES, NETSOCIO, DSCSOCIO, DIFSOCIO, USU ) " _
   & " SELECT " _
   & "  CODASIG1, MES, NETASIG1, DSCASIG1, DIFASIG1, '" + wcodusu + "' " _
   & " FROM DIECOCAB " _
   & " WHERE CODASIG1 = " + Str(wSoc) + " AND CODASIG1 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECO " _
   & " (CODSOCIO, MES, NETSOCIO, DSCSOCIO, DIFSOCIO, USU ) " _
   & " SELECT " _
   & "  CODASIG2, MES, NETASIG2, DSCASIG2, DIFASIG2, '" + wcodusu + "' " _
   & " FROM DIECOCAB " _
   & " WHERE CODASIG2 = " + Str(wSoc) + " AND CODASIG2 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECO " _
   & " (CODSOCIO, MES, NETSOCIO, DSCSOCIO, DIFSOCIO, USU ) " _
   & " SELECT " _
   & "  CODASIG3, MES, NETASIG3, DSCASIG3, DIFASIG3, '" + wcodusu + "' " _
   & " FROM DIECOCAB " _
   & " WHERE CODASIG3 = " + Str(wSoc) + " AND CODASIG3 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECO " _
   & " (CODSOCIO, MES, NETSOCIO, DSCSOCIO, DIFSOCIO, USU ) " _
   & " SELECT " _
   & "  CODASIG4, MES, NETASIG4, DSCASIG4, DIFASIG4, '" + wcodusu + "' " _
   & " FROM DIECOCAB " _
   & " WHERE CODASIG4 = " + Str(wSoc) + " AND CODASIG4 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECO " _
   & " (CODSOCIO, MES, NETSOCIO, DSCSOCIO, DIFSOCIO, USU ) " _
   & " SELECT " _
   & "  CODASIG5, MES, NETASIG5, DSCASIG5, DIFASIG5, '" + wcodusu + "' " _
   & " FROM DIECOCAB " _
   & " WHERE CODASIG5 = " + Str(wSoc) + " AND CODASIG5 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_CAJMP " _
   & " (CODSOCIO, MES, NETSOCIO, DSCSOCIO, DIFSOCIO, USU ) " _
   & " SELECT " _
   & "  CODSOCIO, MES, NETSOCIO, DSCSOCIO, DIFSOCIO, '" + wcodusu + "' " _
   & " FROM CAJMPCAB " _
   & " WHERE CODSOCIO = " + Str(wSoc) + " ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_CAJMP " _
   & " (CODSOCIO, MES, NETSOCIO, DSCSOCIO, DIFSOCIO, USU ) " _
   & " SELECT " _
   & "  CODASIG1, MES, NETASIG1, DSCASIG1, DIFASIG1, '" + wcodusu + "' " _
   & " FROM CAJMPCAB " _
   & " WHERE CODASIG1 = " + Str(wSoc) + " AND CODASIG1 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_CAJMP " _
   & " (CODSOCIO, MES, NETSOCIO, DSCSOCIO, DIFSOCIO, USU ) " _
   & " SELECT " _
   & "  CODASIG2, MES, NETASIG2, DSCASIG2, DIFASIG2, '" + wcodusu + "' " _
   & " FROM CAJMPCAB " _
   & " WHERE CODASIG2 = " + Str(wSoc) + " AND CODASIG2 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_CAJMP " _
   & " (CODSOCIO, MES, NETSOCIO, DSCSOCIO, DIFSOCIO, USU ) " _
   & " SELECT " _
   & "  CODASIG3, MES, NETASIG3, DSCASIG3, DIFASIG3, '" + wcodusu + "' " _
   & " FROM CAJMPCAB " _
   & " WHERE CODASIG3 = " + Str(wSoc) + " AND CODASIG3 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_CAJMP " _
   & " (CODSOCIO, MES, NETSOCIO, DSCSOCIO, DIFSOCIO, USU ) " _
   & " SELECT " _
   & "  CODASIG3, MES, NETASIG4, DSCASIG4, DIFASIG4, '" + wcodusu + "' " _
   & " FROM CAJMPCAB " _
   & " WHERE CODASIG4 = " + Str(wSoc) + " AND CODASIG4 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_CAJMP " _
   & " (CODSOCIO, MES, NETSOCIO, DSCSOCIO, DIFSOCIO, USU ) " _
   & " SELECT " _
   & "  CODASIG5, MES, NETASIG5, DSCASIG5, DIFASIG5, '" + wcodusu + "' " _
   & " FROM CAJMPCAB " _
   & " WHERE CODASIG5 = " + Str(wSoc) + " AND CODASIG5 <> 0 ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_TESOR " _
   & " (CODSOCIO, CODIGO, INS, SERIE, NRO_COMP, FECHA_PAGO, MONEDA, " _
   & "  MONTO, OBS, CONCEPTO, CONPAGO, NOMCON, ITEM, USU ) " _
   & " SELECT " _
   & "  " + Str(wSoc) + ", R.CODIGO, R.INS, R.SERIE, R.NRO_COMP, R.FECHA_PAGO, " _
   & "  R.MONEDA, R.MONTO, R.OBS, C.CCONCE, C.CONCEPTO, C.DESCONCE, R.ITEM, '" + wcodusu + "' " _
   & " FROM ZZZ_MRECIBOS AS R INNER JOIN ZZZ_CONCEPTO AS C " _
   & "   ON R.CONCEPTO = C.CCONCE " _
   & " WHERE R.CODIGO = " + Str(wCod) + " AND " _
   & "          R.INS = " + Str(wIns) + " AND " _
   & "        R.MONTO > 0 ")
   Db.CommitTrans

   aa = Leerado2("SELECT FECHA_PAGO, SERIE, NRO_COMP, MONEDA, " _
                & "      MONTO, CONPAGO, OBS, USU, CONCEPTO, NOMCON, ITEM " _
                & " FROM TMP_TESOR " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "       CODIGO = " + Str(wCod) + " AND " _
                & "          INS = " + Str(wIns) + " " _
                & " ORDER BY FECHA_PAGO ")
   Set DataGrid1.DataSource = ADO2

   DataGrid1.Columns(0).Width = 1080   ' FECHA
   DataGrid1.Columns(0).Alignment = dbgLeft
   DataGrid1.Columns(0).Caption = "FECHA"
   DataGrid1.Columns(0).NumberFormat = "dd/mm/yyyy"
    
   DataGrid1.Columns(1).Width = 430   ' SERIE
   DataGrid1.Columns(1).Alignment = dbgLeft
   DataGrid1.Columns(1).Caption = "SERIE"
    
   DataGrid1.Columns(2).Width = 1000  ' NRO_COMP
   DataGrid1.Columns(2).Alignment = dbgRight
   DataGrid1.Columns(2).Caption = "RECIBO"
   DataGrid1.Columns(2).NumberFormat = "####0"
    
   DataGrid1.Columns(3).Width = 390   ' MONEDA
   DataGrid1.Columns(3).Alignment = dbgCenter
   DataGrid1.Columns(3).Caption = "MON"
    
   DataGrid1.Columns(4).Width = 900   ' MONTO
   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).Caption = "MONTO"
   DataGrid1.Columns(4).NumberFormat = "####0.00"
    
   DataGrid1.Columns(5).Width = 500   ' CONCEPTO
   DataGrid1.Columns(5).Alignment = dbgLeft
   DataGrid1.Columns(5).Caption = "CONCEP"
    
   DataGrid1.Columns(6).Width = 6000  ' OBS
   DataGrid1.Columns(6).Alignment = dbgLeft
   DataGrid1.Columns(6).Caption = "OBSERV"
   
   DataGrid1.Columns(7).Visible = False
   DataGrid1.Columns(8).Visible = False
   DataGrid1.Columns(9).Visible = False
   DataGrid1.Columns(10).Visible = False
   
   DataGrid1.Columns(0).Locked = True
   DataGrid1.Columns(1).Locked = True
   DataGrid1.Columns(2).Locked = True
   DataGrid1.Columns(3).Locked = True
   DataGrid1.Columns(4).Locked = True
   DataGrid1.Columns(5).Locked = True

   aa = Leerado3("SELECT MES, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
                & "      CODSOCIO, USU " _
                & " FROM TMP_DIECO " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "       CODSOCIO = " + Str(wSoc) + " " _
                & " ORDER BY MES DESC ")
   Set DataGrid2.DataSource = ADO3

   DataGrid2.Columns(0).Width = 750   ' MES
   DataGrid2.Columns(0).Alignment = dbgLeft
   DataGrid2.Columns(0).Caption = "MES"
    
   DataGrid2.Columns(1).Width = 800   ' NETSOCIO
   DataGrid2.Columns(1).Alignment = dbgRight
   DataGrid2.Columns(1).Caption = "ENVIO"
   DataGrid2.Columns(1).NumberFormat = "####0.00"
    
   DataGrid2.Columns(2).Width = 800   ' DSCSOCIO
   DataGrid2.Columns(2).Alignment = dbgRight
   DataGrid2.Columns(2).Caption = "RETORNO"
   DataGrid2.Columns(2).NumberFormat = "####0.00"
    
   DataGrid2.Columns(3).Width = 800   ' DIFSOCIO
   DataGrid2.Columns(3).Alignment = dbgRight
   DataGrid2.Columns(3).Caption = "DIFER"
   DataGrid2.Columns(3).NumberFormat = "####0.00"
    
   DataGrid2.Columns(4).Visible = False
   DataGrid2.Columns(5).Visible = False

   aa = Leerado4("SELECT MES, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
                & "      CODSOCIO, USU " _
                & " FROM TMP_CAJMP " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "       CODSOCIO = " + Str(wSoc) + " " _
                & " ORDER BY MES DESC ")
   Set DataGrid3.DataSource = ADO4

   DataGrid3.Columns(0).Width = 750   ' MES
   DataGrid3.Columns(0).Alignment = dbgLeft
   DataGrid3.Columns(0).Caption = "MES"
    
   DataGrid3.Columns(1).Width = 800   ' NETSOCIO
   DataGrid3.Columns(1).Alignment = dbgRight
   DataGrid3.Columns(1).Caption = "ENVIO"
   DataGrid3.Columns(1).NumberFormat = "####0.00"
    
   DataGrid3.Columns(2).Width = 800   ' DSCSOCIO
   DataGrid3.Columns(2).Alignment = dbgRight
   DataGrid3.Columns(2).Caption = "RETORNO"
   DataGrid3.Columns(2).NumberFormat = "####0.00"
    
   DataGrid3.Columns(3).Width = 800   ' DIFSOCIO
   DataGrid3.Columns(3).Alignment = dbgRight
   DataGrid3.Columns(3).Caption = "DIFER"
   DataGrid3.Columns(3).NumberFormat = "####0.00"
    
   DataGrid3.Columns(4).Visible = False
   DataGrid3.Columns(5).Visible = False
End Sub

Private Sub LabelCab()
   If Not ADO2.BOF And Not ADO2.EOF Then
      lblConcepto = IIf(IsNull(ADO2!nomcon), "", ADO2!nomcon)
   End If
End Sub

Private Sub txtCodSocio_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
   If aa > 0 Then
      lblCodSocio.Caption = ADO6a!nombre
   Else
      lblCodSocio.Caption = ""
      Limpiar
   End If
   Set ADO6a = Nothing
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
   
   Dim KeyNew As Integer
   KeyNew = Asc(UCase(Chr(KeyCode)))
   
   Select Case KeyCode
   Case 116
        xlista = "1"
        xseleccion = ""
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtCodSocio.Text = xseleccion
        End If
   Case 65 To 90
        xlista = "1"
        xseleccion = Chr(KeyNew)
        frmSelecSocio.Show 1
        If xseleccion <> "" Then
           txtCodSocio.Text = xseleccion
        End If
          
   End Select
End Sub

Private Sub txtCodSocio_KeyPress(KeyAscii As Integer)
   Dim aa As Integer
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
      Limpiar
      
      txtCodigo.Text = ADO8!codigo
      txtIns.Text = ADO8!ins
      txtNumDoc.Text = ADO8!numdoc
      
      LlenaCab
      LabelCab
      
      If ADO2.RecordCount > 0 Then
         DataGrid1.SetFocus
      Else
         If ADO3.RecordCount > 0 Then
            DataGrid2.SetFocus
         Else
            If ADO4.RecordCount > 0 Then
               DataGrid3.SetFocus
            End If
         End If
      End If
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
'         KeyAscii = 0
'      End If
   End If
End Sub


