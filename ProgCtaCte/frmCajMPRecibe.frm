VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMPRecibe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibir Descuento Caja Militar Policial"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   12345
   Begin VB.CommandButton cmdExtorna 
      Caption         =   "Extornar Descuento"
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
      Left            =   240
      TabIndex        =   19
      TabStop         =   0   'False
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
      Left            =   10320
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Actualiza Descuento"
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4935
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   8705
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
   Begin VB.CommandButton cmdRecibir 
      Caption         =   "Recibir Archivo"
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
      Left            =   3960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox cmbMeses 
      Height          =   315
      ItemData        =   "frmCajMPRecibe.frx":0000
      Left            =   1080
      List            =   "frmCajMPRecibe.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtAnoCab 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblCanApo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9600
      TabIndex        =   30
      Top             =   435
      Width           =   1095
   End
   Begin VB.Label lblEnviado 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10920
      TabIndex        =   29
      Top             =   435
      Width           =   1095
   End
   Begin VB.Label lblCanAsi 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9600
      TabIndex        =   28
      Top             =   945
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Cant.Titulares"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   9600
      TabIndex        =   27
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Cant.Asignados"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   9600
      TabIndex        =   26
      Top             =   735
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Total Envio S/."
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10920
      TabIndex        =   25
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblNoDscto 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10920
      TabIndex        =   24
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "No Cobrado"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10920
      TabIndex        =   23
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblRecibido 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10920
      TabIndex        =   22
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Cobrado"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10920
      TabIndex        =   21
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RECIBIR DESCUENTO DE CAJA MILITAR POLICIAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2040
      TabIndex        =   20
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Label11 
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
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label lblAsig1 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   6240
      Width           =   4215
   End
   Begin VB.Label lblAsig2 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   6480
      Width           =   4215
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
      TabIndex        =   15
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblAsig3 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   6720
      Width           =   4215
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
      TabIndex        =   13
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label lblAsig4 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6000
      TabIndex        =   12
      Top             =   6240
      Width           =   4215
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
      Left            =   5280
      TabIndex        =   11
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label lblAsig5 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6000
      TabIndex        =   10
      Top             =   6480
      Width           =   4215
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
      Left            =   5280
      TabIndex        =   9
      Top             =   6480
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
      Left            =   1560
      TabIndex        =   8
      Top             =   7200
      Width           =   4935
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
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   495
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
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "frmMPRecibe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Limpiar()
   lblCanApo.Caption = ""
   lblCanAsi.Caption = ""
   lblEnviado.Caption = ""
   lblRecibido.Caption = ""
   lblNoDscto.Caption = ""
End Sub

Private Sub cmbMeses_Click()
   cmbMeses_KeyPress (13)
End Sub

Private Sub cmbMeses_KeyPress(KeyAscii As Integer)
   Dim zz As Integer, wAno As String, wMes As String
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
   Limpiar
   If KeyAscii = 13 Then
      Set DataGrid1.DataSource = Nothing
      
      zz = Leerado2("SELECT * FROM DIECOCAB " _
                & "  WHERE MES = '" + wAno + wMes + "' ")
      If zz > 0 Then
      
         lblMensaje.Caption = "Trae Calculo DIECO - Mes " + Left(Trim(funnommes(wMes)), 3) + " " + wAno
         lblMensaje.Refresh
      
         LlenaCab
         LlenaCab1
         TotalCab
         ADO2.MoveFirst
      
         zz = Leerado5a("SELECT SUM(DSCDIECO) AS DSCDIECO " _
                        & " FROM TMP_DIECOCAB " _
                        & " WHERE MES = '" + wAno + wMes + "' AND " _
                        & "       USU = '" + wcodusu + "' ")
         If zz > 0 Then
            If ADO5a!dscdieco > 0 Then
               cmdExtorna.Enabled = True
               cmdGrabar.Enabled = False
            Else
               cmdExtorna.Enabled = False
               cmdGrabar.Enabled = True
            End If
         Else
            cmdGrabar.Enabled = True
            cmdGrabar.SetFocus
         End If
         
         lblMensaje.Caption = ""
         lblMensaje.Refresh
      Else
         If cmdGrabar.Enabled = True Then
            cmdGrabar.SetFocus
         End If
      End If
   End If
End Sub

Private Sub cmdDetalle_Click()
   zDetaCodSoc = ADO2!codsocio
   zDetaTipDsc = "01"
   zDetaAnoDsc = txtAnoCab.Text
   zDetaMesDsc = Left(cmbMeses.Text, 2)

   frmDIECODetalle.Show vbModal
End Sub

Private Sub cmdGrabar_Click()
   Dim zz As Integer, wAno As String, wMes As String, _
      wRegAct As Integer, wRegTot As Integer, _
      wSoc As Integer, wCod As Long, wIns As Integer, wNom As String, wSit As Integer, wEsp As Integer, _
      wDscDieco As Currency, wTotEnvio As Currency, wDscDifer As Currency, wSdoxDist As Currency, _
      wDscSocio As Currency, wDscAsig1 As Currency, wDscAsig2 As Currency, _
      wDscAsig3 As Currency, wDscAsig4 As Currency, wDscAsig5 As Currency, _
      wCodAsig1 As Long, wCodAsig2 As Long, wCodAsig3 As Long, wCodAsig4 As Long, wCodAsig5 As Long, _
      wInsAsig1 As Integer, wInsAsig2 As Integer, wInsAsig3 As Integer, wInsAsig4 As Integer, wInsAsig5 As Integer, _
      wSocAsig1 As Integer, wSocAsig2 As Integer, wSocAsig3 As Integer, wSocAsig4 As Integer, wSocAsig5 As Integer, _
      wTotAsig1 As Currency, wTotAsig2 As Currency, wTotAsig3 As Currency, wTotAsig4 As Currency, wTotAsig5 As Currency, _
      wNomAsig1 As String, wNomAsig2 As String, wNomAsig3 As String, wNomAsig4 As String, wNomAsig5 As String
   
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
   
   zz = Leerado8("SELECT * FROM TMP_DIECOCAB " _
                & " WHERE USU = '" + wcodusu + "' AND " _
                & "       MES = '" + wAno + wMes + "' " _
                & " ORDER BY CODSOCIO ")
   If zz > 0 Then
      wRegAct = 1
      wRegTot = zz
      ADO8.MoveFirst
      Do While Not ADO8.EOF
         wSoc = ADO8!codsocio
         wCod = ADO8!codigo
         wIns = ADO8!ins
         wNom = Trim(ADO8!nombre)
         wSit = ADO8!situ
         wEsp = ADO8!situesp
         wTotEnvio = ADO8!totenvio
         wDscDieco = ADO8!dscdieco
         wDscDifer = ADO8!dscdifer
         wSdoxDist = ADO8!dscdieco
         wDscSocio = 0: wDscAsig1 = 0: wDscAsig2 = 0: wDscAsig3 = 0: wDscAsig4 = 0: wDscAsig5 = 0
         wSocAsig1 = ADO8!codasig1: wTotAsig1 = ADO8!totasig1
         wSocAsig2 = ADO8!codasig2: wTotAsig2 = ADO8!totasig2
         wSocAsig3 = ADO8!codasig3: wTotAsig3 = ADO8!totasig3
         wSocAsig4 = ADO8!codasig4: wTotAsig4 = ADO8!totasig4
         wSocAsig5 = ADO8!codasig5: wTotAsig5 = ADO8!totasig5
         wCodAsig1 = 0: wCodAsig2 = 0: wCodAsig3 = 0: wCodAsig4 = 0: wCodAsig5 = 0
         wInsAsig1 = 0: wInsAsig2 = 0: wInsAsig3 = 0: wInsAsig4 = 0: wInsAsig5 = 0
         wNomAsig1 = "": wNomAsig2 = "": wNomAsig3 = "": wNomAsig4 = "": wNomAsig5 = ""
     
         Db.BeginTrans
         Db.Execute ("UPDATE MAESOCIO " _
         & " SET SITU = " + Str(wSit) + ", SITUESP = " + Str(wEsp) + " " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " ")
         Db.CommitTrans
     
         If wSdoxDist > 0 Then
            If ADO8!netsocio > 0 Then
               If wSdoxDist >= ADO8!netsocio Then
                  wDscSocio = ADO8!netsocio
               Else
                  wDscSocio = wSdoxDist
               End If
               wSdoxDist = wSdoxDist - wDscSocio
            End If
         End If
                  
         If wSdoxDist > 0 Then
            If ADO8!totasig1 > 0 Then
               If wSdoxDist >= wTotAsig1 Then
                  wDscAsig1 = wTotAsig1
               Else
                  wDscAsig1 = wSdoxDist
               End If
               wSdoxDist = wSdoxDist - wDscAsig1
            End If
         End If
         
         If wSdoxDist > 0 Then
            If ADO8!totasig2 > 0 Then
               If wSdoxDist >= wTotAsig2 Then
                  wDscAsig2 = wTotAsig1
               Else
                  wDscAsig2 = wSdoxDist
               End If
               wSdoxDist = wSdoxDist - wDscAsig2
            End If
         End If
                      
         If wSdoxDist > 0 Then
            If ADO8!totasig3 > 0 Then
               If wSdoxDist >= wTotAsig3 Then
                  wDscAsig3 = wTotAsig3
               Else
                  wDscAsig3 = wSdoxDist
               End If
               wSdoxDist = wSdoxDist - wDscAsig3
            End If
         End If
             
         If wSdoxDist > 0 Then
            If ADO8!totasig4 > 0 Then
               If wSdoxDist >= wTotAsig4 Then
                  wDscAsig4 = wTotAsig4
               Else
                  wDscAsig4 = wSdoxDist
               End If
               wSdoxDist = wSdoxDist - wDscAsig4
            End If
         End If
   
         If wSdoxDist > 0 Then
            If ADO8!totasig5 > 0 Then
               If wSdoxDist >= wTotAsig5 Then
                  wDscAsig5 = wTotAsig5
               Else
                  wDscAsig5 = wSdoxDist
               End If
               wSdoxDist = wSdoxDist - wDscAsig5
            End If
         End If
   
         Db.BeginTrans
         Db.Execute ("UPDATE TMP_DIECOCAB " _
         & " SET DSCSOCIO = " + Str(wDscSocio) + ", DSCASIG1 = " + Str(wDscAsig1) + ", " _
         & "     DSCASIG2 = " + Str(wDscAsig2) + ", DSCASIG3 = " + Str(wDscAsig3) + ", " _
         & "     DSCASIG4 = " + Str(wDscAsig4) + ", DSCASIG5 = " + Str(wDscAsig5) + ", " _
         & "     DSCDIFER = " + Str(wSdoxDist) + " " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "            MES = '" + wAno + wMes + "' AND " _
         & "            USU = '" + wcodusu + "' ")
         Db.CommitTrans
   
         Db.BeginTrans
         Db.Execute ("UPDATE DIECOCAB " _
         & " SET DSCSOCIO = " + Str(wDscSocio) + ", DSCASIG1 = " + Str(wDscAsig1) + ", " _
         & "     DSCASIG2 = " + Str(wDscAsig2) + ", DSCASIG3 = " + Str(wDscAsig3) + ", " _
         & "     DSCASIG4 = " + Str(wDscAsig4) + ", DSCASIG5 = " + Str(wDscAsig5) + ", " _
         & "     DSCDIFER = " + Str(wSdoxDist) + " " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "            MES = '" + wAno + wMes + "' ")
         Db.CommitTrans
   
         If wDscSocio > 0 Then
            zz = Leerado7a("SELECT * FROM ZZZ_APOR_PLA " _
                    & " WHERE  CODIGO = " + Str(wCod) + " AND " _
                    & "           INS = " + Str(wIns) + " AND " _
                    & "        CUOANO = '" + wAno + "' AND " _
                    & "       TIPAPOR = '1' ")
            If zz = 0 Then
               Db.BeginTrans
               Db.Execute ("INSERT INTO ZZZ_APOR_PLA " _
               & " (CODIGO, INS, NOMBRE, CUOANO, TIPAPOR, " _
               & "  IMPO01, IMPO02, IMPO03, IMPO04, IMPO05, IMPO06, IMPO07, IMPO08, IMPO09, IMPO10, IMPO11, IMPO12, " _
               & "  TOTIMPO, DEUDA_PT2, PASA_PLA ) " _
               & " VALUES " _
               & " (" + Str(wCod) + ", " + Str(wIns) + ", '" + wNom + "', '" + wAno + "', " _
               & "  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '' ) ")
               Db.CommitTrans
            End If
               
            Db.BeginTrans
            Db.Execute ("UPDATE ZZZ_APOR_PLA " _
            & " SET IMPO" + wMes + " = " + Str(wDscSocio) + " " _
            & " WHERE  CODIGO = " + Str(wCod) + " AND " _
            & "           INS = " + Str(wIns) + " AND " _
            & "        CUOANO = '" + wAno + "' AND " _
            & "       TIPAPOR = '1' ")
            Db.CommitTrans
            
            Db.BeginTrans
            Db.Execute ("UPDATE ZZZ_APOR_PLA " _
            & " SET TOTIMPO = IMPO01 + IMPO02 + IMPO03 + IMPO04 + IMPO05 + IMPO06 + " _
            & "               IMPO07 + IMPO08 + IMPO09 + IMPO10 + IMPO11 + IMPO12 " _
            & " WHERE  CODIGO = " + Str(wCod) + " AND " _
            & "           INS = " + Str(wIns) + " AND " _
            & "        CUOANO = '" + wAno + "' AND " _
            & "       TIPAPOR = '1' ")
            Db.CommitTrans
         
            Call CreaCtasxDet_Dieco(wSoc, wAno, wMes, "1", wDscSocio)
         End If
         
         If wDscAsig1 > 0 Then
            zz = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSocAsig1) + " ")
            If zz > 0 Then
               wCodAsig1 = ADO6a!codigo
               wInsAsig1 = ADO6a!ins
               wNomAsig1 = ADO6a!nombre
            End If
            Set ADO6a = Nothing
            
            zz = Leerado7a("SELECT * FROM ZZZ_APOR_PLA " _
                    & " WHERE  CODIGO = " + Str(wCodAsig1) + " AND " _
                    & "           INS = " + Str(wInsAsig1) + " AND " _
                    & "        CUOANO = '" + wAno + "' AND " _
                    & "       TIPAPOR = '1' ")
            If zz = 0 Then
               Db.BeginTrans
               Db.Execute ("INSERT INTO ZZZ_APOR_PLA " _
               & " (CODIGO, INS, NOMBRE, CUOANO, TIPAPOR, " _
               & "  IMPO01, IMPO02, IMPO03, IMPO04, IMPO05, IMPO06, IMPO07, IMPO08, IMPO09, IMPO10, IMPO11, IMPO12, " _
               & "  TOTIMPO, DEUDA_PT2, PASA_PLA ) " _
               & " VALUES " _
               & " (" + Str(wCodAsig1) + ", " + Str(wInsAsig1) + ", '" + wNomAsig1 + "', '" + wAno + "', " _
               & "  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '' ) ")
               Db.CommitTrans
            End If
         
            Db.BeginTrans
            Db.Execute ("UPDATE ZZZ_APOR_PLA " _
            & " SET IMPO" + wMes + " = " + Str(wDscAsig1) + " " _
            & " WHERE  CODIGO = " + Str(wCodAsig1) + " AND " _
            & "           INS = " + Str(wInsAsig1) + " AND " _
            & "        CUOANO = '" + wAno + "' AND " _
            & "       TIPAPOR = '1' ")
            Db.CommitTrans
         
            Db.BeginTrans
            Db.Execute ("UPDATE ZZZ_APOR_PLA " _
            & " SET TOTIMPO = IMPO01 + IMPO02 + IMPO03 + IMPO04 + IMPO05 + IMPO06 + " _
            & "               IMPO07 + IMPO08 + IMPO09 + IMPO10 + IMPO11 + IMPO12 " _
            & " WHERE  CODIGO = " + Str(wCodAsig1) + " AND " _
            & "           INS = " + Str(wInsAsig1) + " AND " _
            & "        CUOANO = '" + wAno + "' AND " _
            & "       TIPAPOR = '1' ")
            Db.CommitTrans
         
            Call CreaCtasxDet_Dieco(wSocAsig1, wAno, wMes, "1", wDscAsig1)
         
         End If
         
         If wDscAsig2 > 0 Then
            zz = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSocAsig2) + " ")
            If zz > 0 Then
               wCodAsig2 = ADO6a!codigo
               wInsAsig2 = ADO6a!ins
               wNomAsig2 = ADO6a!nombre
            End If
            Set ADO6a = Nothing
            
            zz = Leerado7a("SELECT * FROM ZZZ_APOR_PLA " _
                    & " WHERE  CODIGO = " + Str(wCodAsig2) + " AND " _
                    & "           INS = " + Str(wInsAsig2) + " AND " _
                    & "        CUOANO = '" + wAno + "' AND " _
                    & "       TIPAPOR = '1' ")
            If zz = 0 Then
               Db.BeginTrans
               Db.Execute ("INSERT INTO ZZZ_APOR_PLA " _
               & " (CODIGO, INS, NOMBRE, CUOANO, TIPAPOR, " _
               & "  IMPO01, IMPO02, IMPO03, IMPO04, IMPO05, IMPO06, IMPO07, IMPO08, IMPO09, IMPO10, IMPO11, IMPO12, " _
               & "  TOTIMPO, DEUDA_PT2, PASA_PLA ) " _
               & " VALUES " _
               & " (" + Str(wCodAsig2) + ", " + Str(wInsAsig2) + ", '" + wNomAsig2 + "', '" + wAno + "', " _
               & "  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '' ) ")
               Db.CommitTrans
            End If
         
            Db.BeginTrans
            Db.Execute ("UPDATE ZZZ_APOR_PLA " _
            & " SET IMPO" + wMes + " = " + Str(wDscAsig2) + " " _
            & " WHERE  CODIGO = " + Str(wCodAsig2) + " AND " _
            & "           INS = " + Str(wInsAsig2) + " AND " _
            & "        CUOANO = '" + wAno + "' AND " _
            & "       TIPAPOR = '1' ")
            Db.CommitTrans
         
            Db.BeginTrans
            Db.Execute ("UPDATE ZZZ_APOR_PLA " _
            & " SET TOTIMPO = IMPO01 + IMPO02 + IMPO03 + IMPO04 + IMPO05 + IMPO06 + " _
            & "               IMPO07 + IMPO08 + IMPO09 + IMPO10 + IMPO11 + IMPO12 " _
            & " WHERE  CODIGO = " + Str(wCodAsig2) + " AND " _
            & "           INS = " + Str(wInsAsig2) + " AND " _
            & "        CUOANO = '" + wAno + "' AND " _
            & "       TIPAPOR = '1' ")
            Db.CommitTrans
         
            Call CreaCtasxDet_Dieco(wSocAsig2, wAno, wMes, "1", wDscAsig2)
         
         End If
   
         If wDscAsig3 > 0 Then
            zz = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSocAsig3) + " ")
            If zz > 0 Then
               wCodAsig3 = ADO6a!codigo
               wInsAsig3 = ADO6a!ins
               wNomAsig3 = ADO6a!nombre
            End If
            Set ADO6a = Nothing
            
            zz = Leerado7a("SELECT * FROM ZZZ_APOR_PLA " _
                    & " WHERE  CODIGO = " + Str(wCodAsig3) + " AND " _
                    & "           INS = " + Str(wInsAsig3) + " AND " _
                    & "        CUOANO = '" + wAno + "' AND " _
                    & "       TIPAPOR = '1' ")
            If zz = 0 Then
               Db.BeginTrans
               Db.Execute ("INSERT INTO ZZZ_APOR_PLA " _
               & " (CODIGO, INS, NOMBRE, CUOANO, TIPAPOR, " _
               & "  IMPO01, IMPO02, IMPO03, IMPO04, IMPO05, IMPO06, IMPO07, IMPO08, IMPO09, IMPO10, IMPO11, IMPO12, " _
               & "  TOTIMPO, DEUDA_PT2, PASA_PLA ) " _
               & " VALUES " _
               & " (" + Str(wCodAsig3) + ", " + Str(wInsAsig3) + ", '" + wNomAsig3 + "', '" + wAno + "', " _
               & "  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '' ) ")
               Db.CommitTrans
            End If
         
            Db.BeginTrans
            Db.Execute ("UPDATE ZZZ_APOR_PLA " _
            & " SET IMPO" + wMes + " = " + Str(wDscAsig3) + " " _
            & " WHERE  CODIGO = " + Str(wCodAsig3) + " AND " _
            & "           INS = " + Str(wInsAsig3) + " AND " _
            & "        CUOANO = '" + wAno + "' AND " _
            & "       TIPAPOR = '1' ")
            Db.CommitTrans
         
            Db.BeginTrans
            Db.Execute ("UPDATE ZZZ_APOR_PLA " _
            & " SET TOTIMPO = IMPO01 + IMPO02 + IMPO03 + IMPO04 + IMPO05 + IMPO06 + " _
            & "               IMPO07 + IMPO08 + IMPO09 + IMPO10 + IMPO11 + IMPO12 " _
            & " WHERE  CODIGO = " + Str(wCodAsig3) + " AND " _
            & "           INS = " + Str(wInsAsig3) + " AND " _
            & "        CUOANO = '" + wAno + "' AND " _
            & "       TIPAPOR = '1' ")
            Db.CommitTrans
         
            Call CreaCtasxDet_Dieco(wSocAsig3, wAno, wMes, "1", wDscAsig3)
         
         End If
   
         If wDscAsig4 > 0 Then
            zz = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSocAsig4) + " ")
            If zz > 0 Then
               wCodAsig4 = ADO6a!codigo
               wInsAsig4 = ADO6a!ins
               wNomAsig4 = ADO6a!nombre
            End If
            Set ADO6a = Nothing
            
            zz = Leerado7a("SELECT * FROM ZZZ_APOR_PLA " _
                    & " WHERE  CODIGO = " + Str(wCodAsig4) + " AND " _
                    & "           INS = " + Str(wInsAsig4) + " AND " _
                    & "        CUOANO = '" + wAno + "' AND " _
                    & "       TIPAPOR = '1' ")
            If zz = 0 Then
               Db.BeginTrans
               Db.Execute ("INSERT INTO ZZZ_APOR_PLA " _
               & " (CODIGO, INS, NOMBRE, CUOANO, TIPAPOR, " _
               & "  IMPO01, IMPO02, IMPO03, IMPO04, IMPO05, IMPO06, IMPO07, IMPO08, IMPO09, IMPO10, IMPO11, IMPO12, " _
               & "  TOTIMPO, DEUDA_PT2, PASA_PLA ) " _
               & " VALUES " _
               & " (" + Str(wCodAsig4) + ", " + Str(wInsAsig4) + ", '" + wNomAsig4 + "', '" + wAno + "', " _
               & "  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '' ) ")
               Db.CommitTrans
            End If
         
            Db.BeginTrans
            Db.Execute ("UPDATE ZZZ_APOR_PLA " _
            & " SET IMPO" + wMes + " = " + Str(wDscAsig4) + " " _
            & " WHERE  CODIGO = " + Str(wCodAsig4) + " AND " _
            & "           INS = " + Str(wInsAsig4) + " AND " _
            & "        CUOANO = '" + wAno + "' AND " _
            & "       TIPAPOR = '1' ")
            Db.CommitTrans
         
            Db.BeginTrans
            Db.Execute ("UPDATE ZZZ_APOR_PLA " _
            & " SET TOTIMPO = IMPO01 + IMPO02 + IMPO03 + IMPO04 + IMPO05 + IMPO06 + " _
            & "               IMPO07 + IMPO08 + IMPO09 + IMPO10 + IMPO11 + IMPO12 " _
            & " WHERE  CODIGO = " + Str(wCodAsig4) + " AND " _
            & "           INS = " + Str(wInsAsig4) + " AND " _
            & "        CUOANO = '" + wAno + "' AND " _
            & "       TIPAPOR = '1' ")
            Db.CommitTrans
         
            Call CreaCtasxDet_Dieco(wSocAsig4, wAno, wMes, "1", wDscAsig4)
         
         End If
   
         If wDscAsig5 > 0 Then
            zz = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(wSocAsig5) + " ")
            If zz > 0 Then
               wCodAsig5 = ADO6a!codigo
               wInsAsig5 = ADO6a!ins
               wNomAsig5 = ADO6a!nombre
            End If
            Set ADO6a = Nothing
            
            zz = Leerado7a("SELECT * FROM ZZZ_APOR_PLA " _
                    & " WHERE  CODIGO = " + Str(wCodAsig5) + " AND " _
                    & "           INS = " + Str(wInsAsig5) + " AND " _
                    & "        CUOANO = '" + wAno + "' AND " _
                    & "       TIPAPOR = '1' ")
            If zz = 0 Then
               Db.BeginTrans
               Db.Execute ("INSERT INTO ZZZ_APOR_PLA " _
               & " (CODIGO, INS, NOMBRE, CUOANO, TIPAPOR, " _
               & "  IMPO01, IMPO02, IMPO03, IMPO04, IMPO05, IMPO06, IMPO07, IMPO08, IMPO09, IMPO10, IMPO11, IMPO12, " _
               & "  TOTIMPO, DEUDA_PT2, PASA_PLA ) " _
               & " VALUES " _
               & " (" + Str(wCodAsig5) + ", " + Str(wInsAsig5) + ", '" + wNomAsig5 + "', '" + wAno + "', " _
               & "  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '' ) ")
               Db.CommitTrans
            End If
         
            Db.BeginTrans
            Db.Execute ("UPDATE ZZZ_APOR_PLA " _
            & " SET IMPO" + wMes + " = " + Str(wDscAsig5) + " " _
            & " WHERE  CODIGO = " + Str(wCodAsig5) + " AND " _
            & "           INS = " + Str(wInsAsig5) + " AND " _
            & "        CUOANO = '" + wAno + "' AND " _
            & "       TIPAPOR = '1' ")
            Db.CommitTrans
         
            Db.BeginTrans
            Db.Execute ("UPDATE ZZZ_APOR_PLA " _
            & " SET TOTIMPO = IMPO01 + IMPO02 + IMPO03 + IMPO04 + IMPO05 + IMPO06 + " _
            & "               IMPO07 + IMPO08 + IMPO09 + IMPO10 + IMPO11 + IMPO12 " _
            & " WHERE  CODIGO = " + Str(wCodAsig5) + " AND " _
            & "           INS = " + Str(wInsAsig5) + " AND " _
            & "        CUOANO = '" + wAno + "' AND " _
            & "       TIPAPOR = '1' ")
            Db.CommitTrans
         
            Call CreaCtasxDet_Dieco(wSocAsig5, wAno, wMes, "1", wDscAsig5)
         
         End If
   
   
         wRegAct = wRegAct + 1
         ADO8.MoveNext
      Loop
   End If

   MsgBox "Descuento DIECO Actualizado OK", vbExclamation
End Sub

Private Sub cmdRecibir_Click()
   Dim wAno As String, wMes As String, wRuta As String, _
       wCod As Long, wIns As Integer, wSit As Integer, wEsp As Integer, wCip As Long, _
       wImp As Currency, wSin As Currency, wCom As Currency, Cadena As String, _
       zz As Integer, zRegAct As Integer, zRegTot As Integer, wNom As String, _
       wTotEnvio As Currency, wDscDieco As Currency, wDscDifer As Currency, _
       wNetSocio As Currency, wNetAsig1 As Currency, wNetAsig2 As Currency, wNetAsig3 As Currency, wNetAsig4 As Currency, wNetAsig5 As Currency, _
       wDscSocio As Currency, wDscAsig1 As Currency, wDscAsig2 As Currency, wDscAsig3 As Currency, wDscAsig4 As Currency, wDscAsig5 As Currency, _
       wDifSocio As Currency, wDifAsig1 As Currency, wDifAsig2 As Currency, wDifAsig3 As Currency, wDifAsig4 As Currency, wDifAsig5 As Currency

   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
   wRuta = App.Path + "\DIECO\" + wAno + "-" + wMes + "\15020001.txt"

'   MsgBox "El Archivo a Buscar->" + wRuta

   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_DIECOCAB WHERE MES = '" + wAno + wMes + "' AND USU = '" + wcodusu + "' ")
   Db.CommitTrans

   zRegAct = 1
   zRegTot = ADO2.RecordCount
   Open wRuta For Input As #1 ' Abre el archivo para recibir los datos.
    ' Repite el bucle hasta el final del archivo.
   Do While Not EOF(1)
      DoEvents
      lblMensaje.Caption = "Registro " + _
                           Trim(Format(zRegAct, "####0")) + " / " + _
                           Trim(Format(zRegTot, "####0"))
      lblMensaje.Refresh
      
      Line Input #1, Cadena
      
      wCod = Val(Trim(Mid(Cadena, 1, 8)))
      wIns = Val(Mid(Cadena, 9, 1))
      wSit = Val(Mid(Cadena, 18, 1))
      wImp = Format(Val(Mid(Cadena, 19, 10) + "." + Mid(Cadena, 29, 2)), "#####0.00")
      wSin = Format(Val(Mid(Cadena, 31, 10) + "." + Mid(Cadena, 41, 2)), "#####0.00")
      wCip = Val(Mid(Cadena, 43, 8))
      wEsp = Val(Mid(Cadena, 51, 1))
      wCom = Format(Val(Mid(Cadena, 52, 4) + "." + Mid(Cadena, 56, 2)), "#####0.00")
      wNom = ""
      
      zz = Leerado8("SELECT * FROM MAESOCIO " _
                    & " WHERE CODIGO = " + Str(wCod) + " AND " _
                    & "          INS = " + Str(wIns) + " ")
      If zz > 0 Then
         wNom = Trim(ADO8!nombre)
      End If
      Set ADO8 = Nothing
      
      zz = Leerado8("SELECT * FROM TMP_DIECOCAB " _
                    & " WHERE CODIGO = " + Str(wCod) + " AND " _
                    & "          INS = " + Str(wIns) + " AND " _
                    & "          USU = '" + wcodusu + "' ")
      If zz = 0 Then
         MsgBox "Descuento DIECO Sin Socio " + vbNewLine + Str(Trim(wCod)) + "-" + Str(wIns) + " Enviado", vbExclamation
      End If
      wTotEnvio = ADO8!totenvio
      wDscDieco = wImp
      wDscDifer = wTotEnvio - wDscDieco
      
      
      
      
      
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_DIECOCAB " _
      & " SET DSCDIECO = " + Str(wImp) + ", DSCDIFER = " + Str(wSin) + ", " _
      & "         SITU = " + Str(wSit) + ",  SITUESP = " + Str(wEsp) + "  " _
      & " WHERE    USU = '" + wcodusu + "' AND " _
      & "       CODIGO = " + Str(wCod) + " AND " _
      & "          INS = " + Str(wIns) + " AND " _
      & "          MES = '" + wAno + wMes + "' ")
      Db.CommitTrans
      
      
      
      
      Db.BeginTrans
      Db.Execute ("UPDATE DIECOCAB " _
      & " SET DSCDIECO = " + Str(wImp) + ", DSCDIFER = " + Str(wSin) + ", " _
      & "         SITU = " + Str(WSITU) + ", SITUESP = " + Str(wEsp) + "  " _
      & " WHERE CODIGO = " + Str(wCod) + " AND " _
      & "          INS = " + Str(wIns) + " AND " _
      & "          MES = '" + wAno + wMes + "' ")
      Db.CommitTrans
      
      zRegAct = zRegAct + 1
   Loop
   Close #1
   lblMensaje.Caption = ""
   lblMensaje.Refresh

'   DataGrid1.Refresh
   Limpiar
   LlenaCab
   LlenaCab1
   TotalCab
   cmdGrabar.Enabled = True
   cmdGrabar.SetFocus
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub DataGrid1_DblClick()
   zDetaCambio = False
   zDetaCodSoc = ADO2!codsocio
   zDetaTipDsc = "01"
   zDetaAnoDsc = txtAnoCab.Text
   zDetaMesDsc = Left(cmbMeses.Text, 2)

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
   frmDiecoRecibe.Left = (Screen.Width - Width) \ 2
   frmDiecoRecibe.Top = 0
   
   txtAnoCab.Text = wanocia
   
   Dim a As Integer
   cmbMeses.Clear
   a = Leerado("select * from MAEMESES " _
            & " WHERE ANO = '" + wanocia + "' AND " _
            & "       MES >= '01' AND " _
            & "       MES <= '12' " _
            & " ORDER BY MES ")
   ADO1.MoveFirst
   Do While Not ADO1.EOF
      cmbMeses.AddItem ADO1!mes + " " + Trim(funnommes(ADO1!mes))
       ADO1.MoveNext
   Loop
   cmdGrabar.Enabled = False
   cmdExtorna.Enabled = False
   cmbMeses.SetFocus
End Sub

Private Sub LlenaCab()
   Dim wAno As String, wMes As String, zz As Integer
   wAno = wanocia
   wMes = Left(cmbMeses.Text, 2)
      
   Db.BeginTrans
   Db.Execute ("DELETE FROM TMP_DIECOCAB WHERE USU = '" + wcodusu + "' ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO TMP_DIECOCAB " _
   & " (MES, CODSOCIO, CODIGO, INS, E_SOCIO, NOMBRE, FECENV, FECDSC, " _
   & "  TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "  TOTENVIO, DSCDIECO, DSCDIFER, SITU    , SITUESP , TIPCOB  , " _
   & "  CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
   & "  DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
   & "  NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, DSCASIG1, DSCASIG2, DSCASIG3, DSCASIG4, DSCASIG5, " _
   & "  DIFASIG1, DIFASIG2, DIFASIG3, DIFASIG4, DIFASIG5, " _
   & "  NOMASIG1, NOMASIG2, NOMASIG3, NOMASIG4, NOMASIG5, USU ) " _
   & " SELECT " _
   & "  D.MES, D.CODSOCIO, M.CODIGO, M.INS, M.E_SOCIO, M.NOMBRE, D.FECENV, D.FECDSC, " _
   & "  TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
   & "  TOTENVIO, DSCDIECO, DSCDIFER, D.SITU  , D.SITUESP, D.TIPCOB  , " _
   & "  CODASIG1, CODASIG2, CODASIG3, CODASIG4, CODASIG5, TOTASIG1, TOTASIG2, TOTASIG3, TOTASIG4, TOTASIG5, " _
   & "  DEUASIG1, DEUASIG2, DEUASIG3, DEUASIG4, DEUASIG5, ADEASIG1, ADEASIG2, ADEASIG3, ADEASIG4, ADEASIG5, " _
   & "  NETASIG1, NETASIG2, NETASIG3, NETASIG4, NETASIG5, DSCASIG1, DSCASIG2, DSCASIG3, DSCASIG4, DSCASIG5, " _
   & "  DIFASIG1, DIFASIG2, DIFASIG3, DIFASIG4, DIFASIG5, " _
   & "  ''      , ''      , ''      , ''      , ''      , '" + wcodusu + "'  " _
   & " FROM DIECOCAB AS D INNER JOIN MAESOCIO AS M ON D.CODSOCIO = M.CODSOCIO " _
   & " WHERE D.MES = '" + wAno + wMes + "'  ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG1 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG1 = M.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' AND " _
   & "       T.CODASIG1 <> 0 ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG2 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG2 = M.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' AND " _
   & "       T.CODASIG2 <> 0 ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG3 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG3 = M.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' AND " _
   & "       T.CODASIG3 <> 0 ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG4 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG4 = M.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' AND " _
   & "       T.CODASIG4 <> 0 ")
   Db.CommitTrans
   
   Db.BeginTrans
   Db.Execute ("UPDATE TMP_DIECOCAB " _
   & " SET NOMASIG5 = M.NOMBRE " _
   & " FROM TMP_DIECOCAB AS T INNER JOIN MAESOCIO AS M " _
   & "   ON T.CODASIG5 = M.CODSOCIO " _
   & " WHERE T.USU = '" + wcodusu + "' AND " _
   & "       T.CODASIG5 <> 0 ")
   Db.CommitTrans
   
   zz = Leerado2("SELECT CODSOCIO, CODIGO  , INS     , NOMBRE  , " _
                & "      TOTENVIO, DSCDIECO, DSCDIFER " _
                & "      TOTAPORT, TOTDEUDA, TOTADELA, NETSOCIO, DSCSOCIO, DIFSOCIO, " _
                & "      CODASIG1, TOTASIG1, DEUASIG1, ADEASIG1, NETASIG1, DSCASIG1, DIFASIG1, " _
                & "      CODASIG2, TOTASIG2, DEUASIG2, ADEASIG2, NETASIG2, DSCASIG2, DIFASIG2, " _
                & "      CODASIG3, TOTASIG3, DEUASIG3, ADEASIG3, NETASIG3, DSCASIG3, DIFASIG3, " _
                & "      CODASIG4, TOTASIG4, DEUASIG4, ADEASIG4, NETASIG4, DSCASIG4, DIFASIG4, " _
                & "      CODASIG5, TOTASIG5, DEUASIG5, ADEASIG5, NETASIG5, DSCASIG5, DIFASIG5, " _
                & "      NOMASIG1, NOMASIG2, NOMASIG3, NOMASIG4, NOMASIG5, MES, " _
                & "      SITU    , SITUESP , TIPCOB  , FECENV  , FECDSC  , E_SOCIO " _
                & " FROM TMP_DIECOCAB " _
                & " WHERE MES = '" + wAno + wMes + "' AND USU = '" + wcodusu + "' " _
                & " ORDER BY NOMBRE ")
   Set DataGrid1.DataSource = ADO2
   If zz = 0 Then
      MsgBox "No Existe Envio a DIECO del Mes " + Trim(funnommes(wMes)) + "-" + wAno
   End If
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
    
   DataGrid1.Columns(3).Width = 5500  ' NOMBRE
   DataGrid1.Columns(3).Alignment = dbgLeft
   DataGrid1.Columns(3).Caption = "NOMBRE ASOCIADO"
    
   DataGrid1.Columns(4).Width = 800    ' TOTENVIO
   DataGrid1.Columns(4).Alignment = dbgRight
   DataGrid1.Columns(4).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(4).Caption = "T.ENVIO"
    
   DataGrid1.Columns(5).Width = 800    ' TOTDIECO
   DataGrid1.Columns(5).Alignment = dbgRight
   DataGrid1.Columns(5).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(5).Caption = "DESCTO"
    
   DataGrid1.Columns(6).Width = 800    ' TOTDIFER
   DataGrid1.Columns(6).Alignment = dbgRight
   DataGrid1.Columns(6).NumberFormat = "####0.00;;\ "
   DataGrid1.Columns(6).Caption = "NO DSCTO"
    
   DataGrid1.Columns(7).Visible = False
   DataGrid1.Columns(8).Visible = False
   DataGrid1.Columns(9).Visible = False
   DataGrid1.Columns(10).Visible = False
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
   DataGrid1.Columns(52).Visible = False
   DataGrid1.Columns(53).Visible = False
   DataGrid1.Columns(54).Visible = False
   DataGrid1.Columns(55).Visible = False
   DataGrid1.Columns(56).Visible = False
   DataGrid1.Columns(57).Visible = False
   DataGrid1.Columns(58).Visible = False
'   DataGrid1.Columns(59).Visible = False
'   DataGrid1.Columns(60).Visible = False
   
   DataGrid1.MarqueeStyle = dbgHighlightRowRaiseCell
   DataGrid1.SetFocus
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



