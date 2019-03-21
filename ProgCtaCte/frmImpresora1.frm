VERSION 5.00
Begin VB.Form frmImpresora1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Impresora"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
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
      Left            =   720
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
      Left            =   720
      TabIndex        =   2
      Top             =   360
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
      Left            =   3360
      TabIndex        =   1
      Top             =   1800
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
      Left            =   2040
      TabIndex        =   0
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "frmImpresora1"
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
   tipoprint = "0"
   todoprint = True
   Unload Me
End Sub

Private Sub Form_Activate()
   frmImpresora1.Left = (Screen.Width - Width) \ 2
   frmImpresora1.Top = 0
   
   optLaser.Value = False
   OptMatriz.Value = True
      
   cmdAceptar.SetFocus
End Sub

Private Sub OptLaser_Click()
   cmdAceptar.SetFocus
End Sub


