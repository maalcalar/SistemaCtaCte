VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmServRenovac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresar Saldo al 2018 por Renovaciones"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   14040
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "&Actualiza Saldos"
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
      Left            =   8880
      TabIndex        =   6
      Top             =   7800
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
      Left            =   12360
      TabIndex        =   5
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCrear 
      Caption         =   "Crear Saldos"
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
      Left            =   3000
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox cmbCia 
      Height          =   315
      ItemData        =   "frmServRenovac.frx":0000
      Left            =   240
      List            =   "frmServRenovac.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   300
      Width           =   7215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   10610
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
   Begin VB.Label Label25 
      Caption         =   "Compañia"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Esta Opción Sirve Para Ingresar Los Saldos Pendientes de Cobro Por renovaciones de Asociados Tipo TRA al año 2018"
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
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmServRenovac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActualiza_Click()
   Dim aa As Integer, _
       wSoc As Integer, wCod As Long, wIns As Integer, _
       wMesCob As String, wE_S As String, wTot As Currency
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM CTASXCAB WHERE CONCEPTO = '02' AND MES LIKE '2018%' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("DELETE FROM CTASXDET WHERE CONCEPTO = '02' AND MES LIKE '2018%' AND TIPMOV = '1' ")
   Db.CommitTrans
   
   aa = Leerado8("SELECT * FROM SDOINI_RENOV WHERE TOTAL > 0 ORDER BY NOMBRE")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wSoc = ADO8!codsocio
         wCod = ADO8!codigo
         wIns = ADO8!ins
         wE_S = ADO8!e_socio
         wMesCob = ADO8!mescob
         wTot = ADO8!Total
   
         Db.BeginTrans
         Db.Execute ("INSERT INTO CTASXCAB " _
         & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, CARGOS, ABONOS, SDONEW) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", '" + wMesCob + "', '02', '" + wE_S + "', " _
         & "  'D', " + Str(wTot) + ", 0, " + Str(wTot) + " ) ")
         Db.CommitTrans
   
         Db.BeginTrans
         Db.Execute ("INSERT INTO CTASXDET " _
         & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
         & "  TIPCAM, DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW, OBS ) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", '" + wMesCob + "', '02', '', '', '', '', '1', " _
         & "  '" + Format(Date, "dd/mm/yyyy") + "', 0, 0, 0, 0, " + Str(wTot) + ", " _
         & "  0, " + Str(wTot) + ", '' ) ")
         Db.CommitTrans
   
         Call ActualizaSaldos(wSoc, wMesCob, "02")
   
         ADO8.MoveNext
      Loop
   End If
   MsgBox "Proceso Termino OK", vbExclamation
End Sub

Private Sub cmdCrear_Click()
   Dim aa As Integer, wSoc As Integer, wCod As Long, wIns As Integer, wFecIng As Date, wMesCob As String

   Db.BeginTrans
   Db.Execute ("DELETE FROM SDOINI_RENOV ")
   Db.CommitTrans

   aa = Leerado8("SELECT * FROM MAESOCIO WHERE E_SOCIO = 'TRA' ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wSoc = ADO8!codsocio
         wCod = ADO8!codigo
         wIns = ADO8!ins
         
         wFecIng = Format(ADO8!fecing, "dd/mm/yyyy")
         wMesCob = "2018/02"
   
         Db.BeginTrans
         Db.Execute ("INSERT INTO SDOINI_RENOV " _
         & " (CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, FECING, MESCOB, MONEDA, TOTAL) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", " + Str(wCod) + ", " + Str(wIns) + ", " _
         & "  '" + ADO8!nombre + "', '" + ADO8!e_socio + "', '" + Format(wFecIng, "dd/mm/yyyy") + "', " _
         & "  '" + wMesCob + "',  'D', 50 ) ")
         Db.CommitTrans
   
         ADO8.MoveNext
      Loop
   End If

   LlenaCab
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub DataGrid1_GotFocus()
   DataGrid1.col = 6
   DataGrid1.SelStart = 0
   If Len(Trim(DataGrid1.Text)) > 0 Then
      DataGrid1.SelLength = Len(Trim(DataGrid1.Text))
   End If
   DataGrid1.Refresh
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   Select Case ColIndex
   Case 0
        ADO2.Sort = "CODSOCIO"
   Case 1
        ADO2.Sort = "CODIGO"
   Case 2
        ADO2.Sort = "INS"
   Case 3
        ADO2.Sort = "NOMBRE"
   End Select
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim wvariable As String
    
    On Error GoTo err
    Select Case KeyCode
    Case 40  ' DOWN
            
         wvariable = DataGrid1.Text
         
         Select Case DataGrid1.col
         Case 6
              ADO2!mescob = IIf(IsNull(wvariable), "", wvariable)
         Case 8
              ADO2!Total = IIf(IsNull(wvariable), 0, Val(wvariable))
         End Select
    
         Select Case DataGrid1.col
         Case 6
              DataGrid1.Text = IIf(IsNull(ADO2!mescob), "2018/02", ADO2!mescob)
         Case 8
              DataGrid1.Text = IIf(IsNull(ADO2!Total), 0, ADO2!Total)
         End Select
            
    Case 37 ' Retroceder
         wvariable = DataGrid1.Text
         
         Select Case DataGrid1.col
         Case 6
              ADO2!mescob = IIf(IsNull(wvariable), "", wvariable)
         Case 8
              ADO2!Total = IIf(IsNull(wvariable), 0, Val(wvariable))
         End Select
         
         If DataGrid1.col = 1 Then
            If DataGrid1.Row > 0 Then
               DataGrid1.Row = DataGrid1.Row - 1
            End If
            DataGrid1.col = 0
         End If
         
         Select Case DataGrid1.col
         Case 6
              DataGrid1.Text = IIf(IsNull(ADO2!mescob), "", ADO2!mescob)
         Case 8
              DataGrid1.Text = IIf(IsNull(ADO2!Total), 0, ADO2!Total)
         End Select
         
    Case 38 ' Subir
         wvariable = DataGrid1.Text
         
         Select Case DataGrid1.col
         Case 6
              ADO2!mescob = IIf(IsNull(wvariable), "", wvariable)
         Case 8
              ADO2!Total = IIf(IsNull(wvariable), 0, Val(wvariable))
         End Select
         
         Select Case DataGrid1.col
         Case 6
              DataGrid1.Text = IIf(IsNull(ADO2!mescob), "", ADO2!mescob)
         Case 8
              DataGrid1.Text = IIf(IsNull(ADO2!Total), 0, ADO2!Total)
         End Select
    
    Case 39 ' Avanzar
         wvariable = DataGrid1.Text
         
         Select Case DataGrid1.col
         Case 6
              ADO2!mescob = IIf(IsNull(wvariable), "", wvariable)
         Case 8
              ADO2!Total = IIf(IsNull(wvariable), 0, Val(wvariable))
         End Select
         
         If DataGrid1.col = 8 Then
            If Not ADO2.EOF Then
               DataGrid1.Row = DataGrid1.Row + 1
            End If
            DataGrid1.col = 6
         End If
          
         Select Case DataGrid1.col
         Case 6
              DataGrid1.Text = IIf(IsNull(ADO2!mescob), "", ADO2!mescob)
         Case 8
              DataGrid1.Text = IIf(IsNull(ADO2!Total), 0, ADO2!Total)
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
    Dim waaa As String, wmmm As String
    
    On Error GoTo err
    Select Case KeyAscii
    Case 13
       Select Case DataGrid1.col
       Case 0  ' CodSocio
            DataGrid1.col = 6
       Case 1  ' Codigo
            DataGrid1.col = 6
       Case 2  ' INS
            DataGrid1.col = 6
       Case 3  ' Nombre
            DataGrid1.col = 6
       Case 4  ' E_SOCIO
            DataGrid1.col = 6
       Case 5  ' FecIng
            DataGrid1.col = 6
       Case 6  ' mescob
            wvariable = Trim(DataGrid1.Text)
            waaa = Left(wvariable, 4)
            wmmm = Right(wvariable, 2)
            If wmmm <> "01" And wmmm <> "02" And wmmm <> "03" And wmmm <> "04" And _
               wmmm <> "05" And wmmm <> "06" And wmmm <> "07" And wmmm <> "08" And _
               wmmm <> "09" And wmmm <> "10" And wmmm <> "11" And wmmm <> "12" Then
               MsgBox "Mes Digitado Invalido", vbExclamation
               ADO2!mescob = "2018/02"
               Exit Sub
            End If
            If waaa < "2017" Or waaa > "2018" Then
               MsgBox "Año Digitado Fuera de Rango", vbExclamation
               ADO2!mescob = "2018/02"
               Exit Sub
            End If
            DataGrid1.Text = wvariable
            ADO2!mescob = wvariable
            DataGrid1.col = 8
       Case 7  ' Moneda
            DataGrid1.col = 8
       Case 8  ' Cargos
            wvariable = Trim(DataGrid1.Text)
            If Not IsNumeric(wvariable) Then
               MsgBox "Importe Digitado Es Invalido", vbExclamation
               ADO2!Total = "0"
               Exit Sub
            End If
            DataGrid1.Text = wvariable
            ADO2!Total = IIf(IsNull(wvariable) Or Len(Trim(wvariable)) = 0, 0, wvariable)
            DataGrid1.col = 6
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
           DataGrid1.col = 8
        End If
         
        Select Case DataGrid1.col
        Case 6
             DataGrid1.Text = IIf(IsNull(ADO2!mescob), "", ADO2!mescob)
        Case 8
             DataGrid1.Text = IIf(IsNull(ADO2!Total), 0, ADO2!Total)
        End Select
   
   Case 38  ' UP
   
   Case 39  ' AVANZAR

        If DataGrid1.col = 9 Then
           DataGrid1.col = 5
        End If
          
        Select Case DataGrid1.col
        Case 6
             DataGrid1.Text = IIf(IsNull(ADO2!mescob), "", ADO2!mescob)
        Case 8
             DataGrid1.Text = IIf(IsNull(ADO2!Total), 0, ADO2!Total)
        End Select
        
   Case 40  ' DOWN
   
   End Select
   Exit Sub
err:
   MsgBox Format(err.Number, "00000000000") + " " + err.Description
   Resume Next
End Sub

Private Sub Form_Activate()
   Form_Initialize
End Sub

Private Sub Form_Initialize()
   frmServRenovac.Left = (Screen.Width - Width) \ 2
   frmServRenovac.Top = 0
   
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

   LlenaCab
   
   DataGrid1.AllowAddNew = False
   DataGrid1.AllowDelete = False
   DataGrid1.AllowUpdate = True
   
   cmdCrear.SetFocus
End Sub

Private Sub LlenaCab()
   Dim aa As Integer

   aa = Leerado2("SELECT CODSOCIO, CODIGO, INS, NOMBRE, E_SOCIO, FECING, MESCOB, MONEDA, TOTAL " _
            & " FROM SDOINI_RENOV " _
            & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO2
   
   DataGrid1.Columns(0).Width = 700   ' CODSOCIO
   DataGrid1.Columns(0).Alignment = dbgRight
   DataGrid1.Columns(0).Caption = "COD.SOCIO"
    
   DataGrid1.Columns(1).Width = 1000  ' CODIGO
   DataGrid1.Columns(1).Alignment = dbgRight
   DataGrid1.Columns(1).Caption = "CODIGO"
    
   DataGrid1.Columns(2).Width = 400   ' INS
   DataGrid1.Columns(2).Alignment = dbgCenter
   DataGrid1.Columns(2).Caption = "INS"
    
   DataGrid1.Columns(3).Width = 4850  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE"
    
   DataGrid1.Columns(4).Width = 550   ' E_SOCIO
   DataGrid1.Columns(4).Alignment = dbgLeft
   DataGrid1.Columns(4).Caption = "EST"
    
   DataGrid1.Columns(5).Width = 1100  ' FECING
   DataGrid1.Columns(5).Alignment = dbgLeft
   DataGrid1.Columns(5).Caption = "FEC.ING"
   DataGrid1.Columns(5).NumberFormat = "dd/mm/yyyy"
    
   DataGrid1.Columns(6).Width = 800   ' MESCOB
   DataGrid1.Columns(6).Alignment = dbgCenter
   DataGrid1.Columns(6).Caption = "MES.COB"
   DataGrid1.Columns(6).NumberFormat = "yyyy/mm"
    
   DataGrid1.Columns(7).Width = 550   ' MONEDA
   DataGrid1.Columns(7).Alignment = dbgCenter
   DataGrid1.Columns(7).Caption = "MON"
    
   DataGrid1.Columns(8).Width = 1000  ' TOTAL
   DataGrid1.Columns(8).Alignment = dbgRight
   DataGrid1.Columns(8).Caption = "TOTAL"
   DataGrid1.Columns(8).NumberFormat = "######0.00;;\ "
End Sub
