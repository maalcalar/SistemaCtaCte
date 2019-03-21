VERSION 5.00
Begin VB.Form frmImpresora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Impresora"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmMatriz 
      Caption         =   "Impresora Matriz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2055
      Left            =   3120
      TabIndex        =   4
      Top             =   600
      Width           =   3255
      Begin VB.TextBox txtHasta 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtDesde 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton optRango 
         Caption         =   "Un Rango de Páginas"
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
         Left            =   480
         TabIndex        =   6
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton optTodo 
         Caption         =   "Todas Las Páginas"
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
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblHasta 
         Caption         =   "Hasta"
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
         Left            =   720
         TabIndex        =   10
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblDesde 
         Caption         =   "Desde"
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
         Left            =   720
         TabIndex        =   9
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.OptionButton OptMatriz 
      Caption         =   "Impresora Matriz"
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
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Width           =   2415
   End
   Begin VB.OptionButton OptLaser 
      Caption         =   "Impresora Inyectora o Laser"
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
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   3015
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
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   3960
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
End
Attribute VB_Name = "frmImpresora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OK As Boolean

Private Sub cmdAceptar_Click()
   If OptMatriz.Value = True Then
      tipoprint = "2"
   Else
      tipoprint = "1"
   End If
   OK = True
   Unload Me
End Sub

Private Sub cmdSalir_Click()
   OK = False
   todoprint = True
   txtDesde.Text = ""
   txtHasta.Text = ""
   Unload Me
End Sub

Private Sub Form_Activate()
   optTodo.Value = True
   optRango.Value = False
   txtDesde.Text = ""
   txtHasta.Text = ""
   
   optLaser.Value = True
   OptMatriz.Value = False
   
   optTodo.Enabled = False
   optRango.Enabled = False
   txtDesde.Enabled = False
   txtHasta.Enabled = False
   lblDesde.Enabled = False
   lblHasta.Enabled = False
   frmMatriz.Enabled = False
      
   cmdAceptar.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If optTodo.Value = True Then
      todoprint = True
      desdeprint = 1
      hastaprint = 9999
   Else
      todoprint = False
      desdeprint = Val(txtDesde.Text)
      hastaprint = Val(txtHasta.Text)
   End If
End Sub

Private Sub OptLaser_Click()
   optTodo.Enabled = False
   optRango.Enabled = False
   txtDesde.Enabled = False
   txtHasta.Enabled = False
   lblDesde.Enabled = False
   lblHasta.Enabled = False
   
   frmMatriz.Enabled = False

End Sub

Private Sub OptMatriz_Click()
   optTodo.Enabled = True
   optRango.Enabled = True
   txtDesde.Enabled = True
   txtHasta.Enabled = True
   lblDesde.Enabled = True
   lblHasta.Enabled = True
   frmMatriz.Enabled = True
   optTodo.SetFocus
End Sub

Private Sub optRango_Click()
   optTodo.Value = False
   txtDesde.Enabled = True
   txtHasta.Enabled = True
   lblDesde.Enabled = True
   lblHasta.Enabled = True
      
   txtDesde.Text = 1
   txtHasta.Text = 9999
   txtDesde.SetFocus
End Sub

Private Sub optTodo_Click()
   optRango.Value = False
   txtDesde.Text = ""
   txtHasta.Text = ""
   txtDesde.Enabled = False
   txtHasta.Enabled = False
   lblDesde.Enabled = False
   lblHasta.Enabled = False
   
   cmdAceptar.SetFocus
End Sub

Private Sub optTodo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdAceptar.SetFocus
   End If
End Sub

Private Sub txtDesde_GotFocus()
   optRango.Value = True
   txtDesde.SelStart = 0
   txtDesde.SelLength = Len(Trim(txtDesde.Text))
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtDesde.Text = "" Then
         MsgBox "Pagina Inicial En Blanco", vbInformation
         Exit Sub
      End If
      If Not IsNumeric(txtDesde.Text) Then
         MsgBox "Campo Digitado No Es Numerico", vbInformation
         txtDesde.Text = ""
         Exit Sub
      End If
      txtHasta.SetFocus
   End If
End Sub

Private Sub txtHasta_GotFocus()
   txtHasta.SelStart = 0
   txtHasta.SelLength = Len(Trim(txtHasta.Text))
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtHasta.Text = "" Then
         MsgBox "Pagina Final En Blanco", vbInformation
         Exit Sub
      End If
      If Not IsNumeric(txtHasta.Text) Then
         MsgBox "Campo Digitado No Es Numerico", vbInformation
         txtHasta.Text = ""
         Exit Sub
      End If
      cmdAceptar.SetFocus
   End If
End Sub
