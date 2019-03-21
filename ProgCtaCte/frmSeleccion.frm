VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSeleccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario de Ayuda"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   11160
   Begin VB.CommandButton cmbBorrar 
      Caption         =   "&Borrar Filtro"
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
      Left            =   6120
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   4545
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   8017
      _Version        =   393216
      OLEDropMode     =   1
      ListField       =   "registro"
      BoundColumn     =   "codigo"
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
      Left            =   9600
      TabIndex        =   3
      Top             =   5760
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
      Left            =   8160
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Filtrar"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtBuscar 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3375
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
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmSeleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer, wsql As String

Private Sub cmbBorrar_Click()
    txtBuscar.Text = ""
    Call Form_Activate
End Sub

Private Sub cmdBuscar_Click()
    If txtBuscar.Text <> "" Then
    Select Case xlista
    Case "1"  ' LINEA DE PRODUCTO
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAELINEA " _
              & " WHERE NOMBRE LIKE '" + Trim(txtBuscar.Text) + "%' ORDER BY CODIGO"
    Case "PT"  ' ARTICULO VENTA
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO " _
                & " FROM MAEARTVTA " _
                & " WHERE NOMBRE LIKE '%" + Trim(txtBuscar.Text) + "%' " _
                & " ORDER BY NOMBRE"
    Case "2"  ' ARTICULOS
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEARTICULO " _
              & " WHERE NOMBRE LIKE '%" + Trim(txtBuscar.Text) + "%' ORDER BY NOMBRE"
    Case "3"  ' CLIENTES
         wsql = "SELECT CODIGO AS CODIGO,CODIGO+' '+NOMBRE AS REGISTRO FROM MAECLIENTE " _
              & " WHERE NOMBRE LIKE '%" + Trim(txtBuscar.Text) + "%' ORDER BY NOMBRE"
    Case "4"  ' TIPDOC
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEDCMTO WHERE CODIGO='01' OR CODIGO='03' ORDER BY CODIGO "
    Case "5"  ' TIPDOC
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEDCMTO WHERE CODIGO='07' OR CODIGO='08' ORDER BY CODIGO "
    Case "6"  ' TRANSACCION
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAETRANSAC " _
              & " WHERE NOMBRE LIKE '" + Trim(txtBuscar.Text) + "%' ORDER BY NOMBRE"
    Case "7"  ' ALMACEN
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEALMACEN " _
              & " WHERE NOMBRE LIKE '" + Trim(txtBuscar.Text) + "%' ORDER BY NOMBRE"
    Case "8"  ' CENCOS
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAECENCOS " _
              & " WHERE NOMBRE LIKE '" + Trim(txtBuscar.Text) + "%' ORDER BY NOMBRE"
    Case "9"  ' ESTADO ASIGNADO
         wsql = "SELECT CODIGO,CODIGO+' '+NOMBRE AS REGISTRO FROM MAEESTADOASIGNADO ORDER BY NOMBRE"
    Case "10"  ' COSECHA
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAECOSECHA " _
              & " WHERE NOMBRE LIKE '" + Trim(txtBuscar.Text) + "%' ORDER BY NOMBRE"
    Case "11"  ' UBICACION LETRA
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEUBILET  " _
              & " WHERE NOMBRE LIKE '" + Trim(txtBuscar.Text) + "%' ORDER BY NOMBRE"
    Case "CO  ' PROVEEDOR"
         wsql = "SELECT CONCEPTO AS CODIGO, " _
         & "            CONCEPTO+' '+DESCONCE AS REGISTRO " _
         & " FROM ZZZ_CONCEPTO " _
         & " WHERE DESCONCE LIKE '%" + UCase(txtBuscar.Text) + "%' " _
         & " ORDER BY DESCONCE"
    Case "TP"  ' TIPO PARIENTE
         wsql = "SELECT TIPOPARIENTE AS CODIGO, " _
         & "            TIPOPARIENTE+' '+NOMBRE AS REGISTRO " _
         & " FROM MAETIPOPARIENTE " _
         & " ORDER BY TIPOPARIENTE"
    Case "B"  ' BANCO
         wsql = "SELECT CODIGO,CODIGO+' '+NOMBRE AS REGISTRO FROM MAEBANCO ORDER BY CODIGO"
    Case "C"  ' CUENTAS
         wsql = "SELECT CUENTA AS CODIGO,CUENTA+'     '+NOMBRE AS REGISTRO FROM MAECUENTA " _
                & "  WHERE TIPO_CTA='D' AND CIA = '01' " _
                & "        CUENTA LIKE '" + Trim(txtBuscar.Text) + "%' " _
                & " ORDER BY CUENTA"
    Case "D"  ' TIPCOB
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAETIPCOB ORDER BY CODIGO "
    Case "12" ' VENDEDORES
         wsql = "SELECT CODIGO,CODIGO+' '+NOMBRE AS REGISTRO FROM MAEVENDEDOR ORDER BY CODIGO"
    Case "13"  ' UM
         wsql = "SELECT UNIDAD AS CODIGO,UNIDAD+'     '+NOMBRE AS REGISTRO FROM MAEUNIDAD ORDER BY NOMBRE"
    Case "14"  ' TIPO TRANSACCION
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAETIPOTRANSAC ORDER BY NOMBRE"
    Case "15"  ' CEO
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAECEO ORDER BY NOMBRE"
    Case "U"  ' UBIGEO
         If xseleccion <> "" Then
            wsql = "SELECT CODIGO,CODIGO+'     '+DESC AS REGISTRO FROM MAEUBIGEO WHERE DESC LIKE '%" + Trim(txtBuscar.Text) + "%' ORDER BY CODIGO"
         Else
            wsql = "SELECT CODIGO,CODIGO+'     '+DESC AS REGISTRO FROM MAEUBIGEO ORDER BY CODIGO"
         End If
    Case "FP"  ' FORMA DE PAGO
         wsql = "SELECT FORPAG AS CODIGO,FORPAG+'     '+NOMBRE AS REGISTRO FROM MAEFORPAG ORDER BY NOMBRE"
    End Select
    If xlista = "C" Then
       a = LeeradoCb1(wsql)
       If a > 0 Then
          Set DataList1.RowSource = ADOCb1
          DataList1.SetFocus
       Else
          MsgBox "No Existen Registros"
'          cmdBuscar_Click
       End If
    Else
       a = Leeradox(wsql)
       If a < 1 Then
          MsgBox "No Existen Registros"
          txtBuscar.SetFocus
          Exit Sub
       Else
          Set DataList1.RowSource = ADOx
          DataList1.SetFocus
       End If
    End If
    End If
End Sub

Private Sub cmdCancela_Click()
    xseleccion = ""
    Unload Me
End Sub

Private Sub cmdSelect_Click()
   Select Case xlista
   Case "U1"
        xseleccion = Mid(DataList1.BoundText, 1, 2)
   Case "U2"
        xseleccion = Mid(DataList1.BoundText, 3, 2)
   Case "U3"
        xseleccion = Mid(DataList1.BoundText, 5, 2)
   Case Else
        xseleccion = DataList1.BoundText
   End Select
   Unload Me
End Sub

Private Sub DataList1_DblClick()
   Call DataList1_KeyPress(13)
End Sub

Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub DataList1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      xseleccion = DataList1.BoundText
      Unload Me
   End If
End Sub

Private Sub Form_Activate()
    Dim aa As Integer, wLin As Integer
    Select Case xlista
    Case "1"  ' SOCIO
         If xseleccion = "" Then
            wsql = "SELECT CODSOCIO AS CODIGO,STR(CODSOCIO)+'     '+NOMBRE AS REGISTRO FROM MAESOCIO ORDER BY NOMBRE"
         Else
            wsql = "SELECT CODSOCIO AS CODIGO,STR(CODSOCIO)+'     '+NOMBRE AS REGISTRO FROM MAESOCIO WHERE NOMBRE LIKE '%" + xseleccion + "%' ORDER BY NOMBRE"
         End If
    Case "PT"  ' ARTICULO VENTA
         If xseleccion = "" Then
            wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEARTVTA ORDER BY NOMBRE"
         Else
            wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEARTVTA WHERE NOMBRE LIKE '%" + xseleccion + "%' ORDER BY NOMBRE"
         End If
    Case "2"  ' ARTICULO
         If xseleccion = "" Then
            wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEARTICULO ORDER BY NOMBRE"
         Else
            wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEARTICULO WHERE NOMBRE LIKE '%" + xseleccion + "%' ORDER BY NOMBRE"
         End If
    Case "2C"  ' ARTICULO
         If xseleccion = "" Then
            wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEARTICULO WHERE SWPOR = 0 ORDER BY NOMBRE"
         Else
            wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEARTICULO WHERE SWPOR = 0 AND NOMBRE LIKE '%" + xseleccion + "%' ORDER BY NOMBRE"
         End If
    Case "3"  ' CLIENTES
         If xseleccion = "" Then
            wsql = "SELECT CODIGO AS CODIGO, " _
            & "            CODIGO+' '+NOMBRE AS REGISTRO " _
            & "  FROM MAECLIENTE ORDER BY NOMBRE"
         Else
            wsql = "SELECT CODIGO AS CODIGO, " _
            & "            CODIGO+' '+NOMBRE AS REGISTRO " _
            & "  FROM MAECLIENTE WHERE NOMBRE LIKE '%" + xseleccion + "%' ORDER BY NOMBRE"
         End If
    Case "4"  ' TIPDOC
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEDCMTO WHERE CODIGO='01' OR CODIGO='03' ORDER BY NOMBRE"
    Case "5"  ' TIPDOC
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEDCMTO WHERE CODIGO='07' OR CODIGO='08' ORDER BY NOMBRE"
    Case "5a"  ' TIPDOC
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEDCMTO ORDER BY NOMBRE"
    Case "6"  ' TRANSACCION
         wsql = "SELECT CODIGO,CODIGO+' '+NOMBRE AS REGISTRO FROM MAETRANSAC WHERE DIGITA = 1 ORDER BY NOMBRE"
    Case "6I"  ' TRANSACCION
         wsql = "SELECT CODIGO,CODIGO+' '+NOMBRE AS REGISTRO FROM MAETRANSAC WHERE DIGITA = 1 AND CLAMOV = '1' ORDER BY NOMBRE"
    Case "6S"  ' TRANSACCION
         wsql = "SELECT CODIGO,CODIGO+' '+NOMBRE AS REGISTRO FROM MAETRANSAC WHERE DIGITA = 1 AND CLAMOV = '2' ORDER BY NOMBRE"
    Case "7"  ' ALMACEN
         wsql = "SELECT CODIGO,CODIGO+' '+NOMBRE AS REGISTRO FROM MAEALMACEN ORDER BY CODIGO"
    Case "8"  ' CENCOS
         wsql = "SELECT CODIGO,CODIGO+' '+NOMBRE AS REGISTRO FROM MAECENCOS ORDER BY NOMBRE"
    Case "9"  ' ESTADO ASIGNADO
         wsql = "SELECT CODIGO,CODIGO+' '+NOMBRE AS REGISTRO FROM MAEESTADOASIGNADO ORDER BY NOMBRE"
    Case "CO"  ' zzz_concepto
         If xseleccion = "" Then
            wsql = "SELECT CONCEPTO AS CODIGO, " _
            & "            CONCEPTO+' '+DESCONCE AS REGISTRO " _
            & " FROM ZZZ_CONCEPTO " _
            & " WHERE SW = 1 AND SERCAJ = '" + zSerCaj + "' AND MONCAJ = '" + zMonCaj + "' " _
            & " ORDER BY CONCEPTO"
         Else
            wsql = "SELECT CONCEPTO AS CODIGO, " _
            & "            CONCEPTO+' '+DESCONCE AS REGISTRO " _
            & " FROM ZZZ_CONCEPTO " _
            & " WHERE DESCONCE LIKE '%" + xseleccion + "%' AND SERCAJ = '" + zSerCaj + "' AND MONCAJ = '" + zMonCaj + "' " _
            & " ORDER BY DESCONCE"
         End If
    Case "CO2"  ' zzz_concepto
         wsql = "SELECT CONCEPTO AS CODIGO, " _
         & "            CONCEPTO+' '+DESCONCE AS REGISTRO " _
         & " FROM ZZZ_CONCEPTO " _
         & " WHERE SERCAJ = '" + zSerCaj + "' AND MONCAJ = '" + zMonCaj + "' " _
         & " ORDER BY CONCEPTO"
    Case "TP"  ' TIPO PARIENTE
         wsql = "SELECT TIPOPARIENTE AS CODIGO, " _
         & "            TIPOPARIENTE+' '+NOMBRE AS REGISTRO " _
         & " FROM MAETIPOPARIENTE " _
         & " ORDER BY TIPOPARIENTE"
    Case "B"  ' BANCO
         wsql = "SELECT CODIGO,CODIGO+' '+NOMBRE AS REGISTRO FROM MAEBANCO ORDER BY CODIGO"
    Case "D"  ' TIPCOB
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAETIPCOB ORDER BY CODIGO "
    Case "E"  ' VENDEDOR
         If xseleccion = "" Then
            wsql = "SELECT CODIGO AS CODIGO, " _
            & "            CODIGO+' '+NOMBRE AS REGISTRO " _
            & "  FROM MAEVENDEDOR ORDER BY NOMBRE "
         Else
            wsql = "SELECT CODIGO AS CODIGO, " _
            & "            CODIGO+' '+NOMBRE AS REGISTRO " _
            & "  FROM MAEVENDEDOR WHERE NOMBRE LIKE '%" + xseleccion + "%' ORDER BY NOMBRE"
         End If
    Case "11"  ' UBICACION LETRA
         wsql = "SELECT CODIGO,CODIGO+' '+NOMBRE AS REGISTRO FROM MAEUBILET  ORDER BY NOMBRE"
    Case "12"  ' ZONA
         wsql = "SELECT CODIGO,CODIGO+' '+NOMBRE AS REGISTRO FROM MAEZONA  ORDER BY NOMBRE"
    Case "13"  ' UM
         wsql = "SELECT UNIDAD AS CODIGO,UNIDAD+'     '+NOMBRE AS REGISTRO FROM MAEUNIDAD ORDER BY NOMBRE"
    Case "14"  ' TIPO TRANSACCION
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAETIPOTRANSAC ORDER BY NOMBRE"
    Case "15"  ' CEO
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAECEO ORDER BY NOMBRE"
    Case "U1"  ' DEPARTAMENTO
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEUBIGEO WHERE CODIGO LIKE '%0000' ORDER BY CODIGO"
    Case "U2"  ' PROVINCIA
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEUBIGEO WHERE CODIGO LIKE '" + Trim(xseleccion) + "[0-9][0-9]00' ORDER BY CODIGO"
    Case "U3"  ' DISTRITO
         wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEUBIGEO WHERE CODIGO LIKE '" + Trim(xseleccion) + "[0-9][0-9]' ORDER BY CODIGO"
    Case "U"  ' UBIGEO
         If xseleccion <> "" Then
            wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEUBIGEO WHERE NOMBRE LIKE '%" + Trim(xseleccion) + "%' ORDER BY CODIGO"
         Else
            wsql = "SELECT CODIGO,CODIGO+'     '+NOMBRE AS REGISTRO FROM MAEUBIGEO ORDER BY CODIGO"
         End If
    Case "FP"  ' FORMA DE PAGO
         wsql = "SELECT FORPAG AS CODIGO,FORPAG+'     '+NOMBRE AS REGISTRO FROM MAEFORPAG ORDER BY NOMBRE"
    Case "DV"  ' DIRECTIVO
         wsql = "SELECT DIRECTIVO AS CODIGO,DIRECTIVO+'     '+NOMBRE AS REGISTRO FROM MAEDIRECTIVO ORDER BY NOMBRE"
    Case "PR"  ' PROMOCION
         wsql = "SELECT PROMOCION AS CODIGO,PROMOCION+'     '+NOMBRE AS REGISTRO FROM MAEPROMOCION ORDER BY NOMBRE"
    End Select
    
    If xlista = "C" Then
       a = LeeradoCb1(wsql)
       If a > 0 Then
          Set DataList1.RowSource = ADOCb1
          DataList1.SetFocus
       Else
          MsgBox "No Existen Registros"
          cmdBuscar_Click
       End If
    Else
       a = Leeradox(wsql)
       If a > 0 Then
          Set DataList1.RowSource = ADOx
'          If xlista = "3" And xseleccion = "" Then
'             txtBuscar.SetFocus
'          Else
             DataList1.SetFocus
'          End If
       Else
          MsgBox "No Existen Registros"
          cmdBuscar_Click
       End If
    End If
'   If Len(Trim(xseleccion)) <> 0 Then
'      DataList1.SetFocus
'   Else
'      txtBuscar.SetFocus
'   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If xlista = "C" Then
      DbCb.Close
   End If
   xlista = ""
End Sub

Private Sub txtBuscar_GotFocus()
   txtBuscar.SelStart = 0
   txtBuscar.SelLength = Len(Trim(txtBuscar.Text))
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtBuscar.Text = "" Then
         MsgBox "Codigo a Filtrar En Blanco", vbInformation
         Exit Sub
      End If
      cmdBuscar.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub
