VERSION 5.00
Begin VB.Form frmDIECODetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Envio a DIECO"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMes 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      MaxLength       =   6
      TabIndex        =   103
      Top             =   200
      Width           =   735
   End
   Begin VB.TextBox txtTotSocio 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      MaxLength       =   9
      TabIndex        =   91
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtDeuSocio 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      MaxLength       =   9
      TabIndex        =   90
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtAdeSocio 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      MaxLength       =   9
      TabIndex        =   89
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtNetSocio 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   8400
      MaxLength       =   9
      TabIndex        =   88
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtDscSocio 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   9240
      MaxLength       =   9
      TabIndex        =   87
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtDifSocio 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   10080
      MaxLength       =   9
      TabIndex        =   86
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtDifAsig5 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   10080
      MaxLength       =   9
      TabIndex        =   78
      Top             =   3300
      Width           =   855
   End
   Begin VB.TextBox txtDscAsig5 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   9240
      MaxLength       =   9
      TabIndex        =   77
      Top             =   3300
      Width           =   855
   End
   Begin VB.TextBox txtNetAsig5 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   8400
      MaxLength       =   9
      TabIndex        =   76
      Top             =   3300
      Width           =   855
   End
   Begin VB.TextBox txtAdeAsig5 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      MaxLength       =   9
      TabIndex        =   75
      Top             =   3300
      Width           =   855
   End
   Begin VB.TextBox txtDeuAsig5 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      MaxLength       =   9
      TabIndex        =   74
      Top             =   3300
      Width           =   855
   End
   Begin VB.TextBox txtTotAsig5 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      MaxLength       =   9
      TabIndex        =   73
      Top             =   3300
      Width           =   855
   End
   Begin VB.TextBox txtDifAsig4 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   10080
      MaxLength       =   9
      TabIndex        =   72
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtDscAsig4 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   9240
      MaxLength       =   9
      TabIndex        =   71
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtNetAsig4 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   8400
      MaxLength       =   9
      TabIndex        =   70
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtAdeAsig4 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      MaxLength       =   9
      TabIndex        =   69
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtDeuAsig4 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      MaxLength       =   9
      TabIndex        =   68
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtTotAsig4 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      MaxLength       =   9
      TabIndex        =   67
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtDifAsig3 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   10080
      MaxLength       =   9
      TabIndex        =   66
      Top             =   2700
      Width           =   855
   End
   Begin VB.TextBox txtDscAsig3 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   9240
      MaxLength       =   9
      TabIndex        =   65
      Top             =   2700
      Width           =   855
   End
   Begin VB.TextBox txtNetAsig3 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   8400
      MaxLength       =   9
      TabIndex        =   64
      Top             =   2700
      Width           =   855
   End
   Begin VB.TextBox txtAdeAsig3 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      MaxLength       =   9
      TabIndex        =   63
      Top             =   2700
      Width           =   855
   End
   Begin VB.TextBox txtDeuAsig3 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      MaxLength       =   9
      TabIndex        =   62
      Top             =   2700
      Width           =   855
   End
   Begin VB.TextBox txtTotAsig3 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      MaxLength       =   9
      TabIndex        =   61
      Top             =   2700
      Width           =   855
   End
   Begin VB.TextBox txtDifAsig2 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   10080
      MaxLength       =   9
      TabIndex        =   60
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtDscAsig2 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   9240
      MaxLength       =   9
      TabIndex        =   59
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtNetAsig2 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   8400
      MaxLength       =   9
      TabIndex        =   58
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtAdeAsig2 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      MaxLength       =   9
      TabIndex        =   57
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtDeuAsig2 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      MaxLength       =   9
      TabIndex        =   56
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtTotAsig2 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      MaxLength       =   9
      TabIndex        =   55
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtDifAsig1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   10080
      MaxLength       =   9
      TabIndex        =   54
      Top             =   2100
      Width           =   855
   End
   Begin VB.TextBox txtDscAsig1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   9240
      MaxLength       =   9
      TabIndex        =   53
      Top             =   2100
      Width           =   855
   End
   Begin VB.TextBox txtNetAsig1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   8400
      MaxLength       =   9
      TabIndex        =   52
      Top             =   2100
      Width           =   855
   End
   Begin VB.TextBox txtAdeAsig1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      MaxLength       =   9
      TabIndex        =   51
      Top             =   2100
      Width           =   855
   End
   Begin VB.TextBox txtDeuAsig1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      MaxLength       =   9
      TabIndex        =   50
      Top             =   2100
      Width           =   855
   End
   Begin VB.TextBox txtTotAsig1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      MaxLength       =   9
      TabIndex        =   49
      Top             =   2100
      Width           =   855
   End
   Begin VB.TextBox txtCodSocio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      MaxLength       =   9
      TabIndex        =   48
      Top             =   200
      Width           =   735
   End
   Begin VB.TextBox txtSocio2 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      MaxLength       =   8
      TabIndex        =   22
      Top             =   2400
      Width           =   570
   End
   Begin VB.TextBox txtSocio1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      MaxLength       =   8
      TabIndex        =   21
      Top             =   2100
      Width           =   570
   End
   Begin VB.TextBox txtSocio3 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      MaxLength       =   8
      TabIndex        =   20
      Top             =   2700
      Width           =   570
   End
   Begin VB.TextBox txtSocio4 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      MaxLength       =   8
      TabIndex        =   19
      Top             =   3000
      Width           =   570
   End
   Begin VB.TextBox txtSocio5 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      MaxLength       =   8
      TabIndex        =   18
      Top             =   3300
      Width           =   570
   End
   Begin VB.TextBox txtTipCob 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7320
      MaxLength       =   3
      TabIndex        =   12
      Top             =   640
      Width           =   495
   End
   Begin VB.TextBox txtGrado 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      MaxLength       =   3
      TabIndex        =   11
      Top             =   640
      Width           =   495
   End
   Begin VB.TextBox txtE_socio 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   6
      Top             =   640
      Width           =   495
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      MaxLength       =   8
      TabIndex        =   5
      Top             =   640
      Width           =   975
   End
   Begin VB.TextBox txtIns 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   4
      Top             =   640
      Width           =   375
   End
   Begin VB.TextBox txtNumdoc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   3
      Top             =   640
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
      Left            =   9840
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Mes"
      Height          =   195
      Left            =   360
      TabIndex        =   104
      Top             =   200
      Width           =   375
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1800
      TabIndex        =   102
      Top             =   3840
      Width           =   3735
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Aporte Mes"
      Height          =   195
      Index           =   11
      Left            =   5880
      TabIndex        =   101
      Top             =   1920
      Width           =   810
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Deudas"
      Height          =   195
      Index           =   10
      Left            =   6825
      TabIndex        =   100
      Top             =   1920
      Width           =   555
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Adelantos"
      Height          =   195
      Index           =   9
      Left            =   7680
      TabIndex        =   99
      Top             =   1920
      Width           =   705
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Envio"
      Height          =   195
      Index           =   8
      Left            =   8415
      TabIndex        =   98
      Top             =   1920
      Width           =   765
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Retorno"
      Height          =   195
      Index           =   7
      Left            =   9330
      TabIndex        =   97
      Top             =   1920
      Width           =   570
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "No Dscto"
      Height          =   195
      Index           =   6
      Left            =   10200
      TabIndex        =   96
      Top             =   1920
      Width           =   675
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   7320
      TabIndex        =   95
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblEnvio 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   8400
      TabIndex        =   94
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblDscto 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   9240
      TabIndex        =   93
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblDifer 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   10080
      TabIndex        =   92
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Importes del Titular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   85
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "No Dscto"
      Height          =   195
      Index           =   5
      Left            =   10200
      TabIndex        =   84
      Top             =   1140
      Width           =   675
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Retorno"
      Height          =   195
      Index           =   4
      Left            =   9330
      TabIndex        =   83
      Top             =   1140
      Width           =   570
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Envio"
      Height          =   195
      Index           =   3
      Left            =   8415
      TabIndex        =   82
      Top             =   1140
      Width           =   765
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Adelantos"
      Height          =   195
      Index           =   2
      Left            =   7680
      TabIndex        =   81
      Top             =   1140
      Width           =   705
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Deudas"
      Height          =   195
      Index           =   1
      Left            =   6825
      TabIndex        =   80
      Top             =   1140
      Width           =   555
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Aporte Mes"
      Height          =   195
      Index           =   0
      Left            =   5880
      TabIndex        =   79
      Top             =   1140
      Width           =   810
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   960
      TabIndex        =   47
      Top             =   740
      Width           =   855
   End
   Begin VB.Label lblCodigo2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   960
      TabIndex        =   46
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblIns2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1800
      TabIndex        =   45
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label lblSocio2 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2160
      TabIndex        =   44
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nombre del Asociado"
      Height          =   195
      Index           =   25
      Left            =   2640
      TabIndex        =   43
      Top             =   1920
      Width           =   1515
   End
   Begin VB.Label lblCodigo1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   960
      TabIndex        =   42
      Top             =   2100
      Width           =   855
   End
   Begin VB.Label lblIns1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1800
      TabIndex        =   41
      Top             =   2100
      Width           =   375
   End
   Begin VB.Label lblSocio1 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2160
      TabIndex        =   40
      Top             =   2100
      Width           =   3735
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "1.-"
      Height          =   195
      Index           =   26
      Left            =   120
      TabIndex        =   39
      Top             =   2100
      Width           =   180
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Codigo"
      Height          =   195
      Index           =   27
      Left            =   465
      TabIndex        =   38
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Codofin"
      Height          =   195
      Index           =   28
      Left            =   1035
      TabIndex        =   37
      Top             =   1920
      Width           =   660
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Ins"
      Height          =   195
      Index           =   29
      Left            =   1800
      TabIndex        =   36
      Top             =   1920
      Width           =   210
   End
   Begin VB.Label lblCodigo3 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   960
      TabIndex        =   35
      Top             =   2700
      Width           =   855
   End
   Begin VB.Label lblIns3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1800
      TabIndex        =   34
      Top             =   2700
      Width           =   375
   End
   Begin VB.Label lblSocio3 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2160
      TabIndex        =   33
      Top             =   2700
      Width           =   3735
   End
   Begin VB.Label lblCodigo4 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   960
      TabIndex        =   32
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblIns4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1800
      TabIndex        =   31
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblSocio4 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2160
      TabIndex        =   30
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label lblCodigo5 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   960
      TabIndex        =   29
      Top             =   3300
      Width           =   855
   End
   Begin VB.Label lblIns5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1800
      TabIndex        =   28
      Top             =   3300
      Width           =   375
   End
   Begin VB.Label lblSocio5 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2160
      TabIndex        =   27
      Top             =   3300
      Width           =   3735
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "2.-"
      Height          =   195
      Index           =   24
      Left            =   120
      TabIndex        =   26
      Top             =   2400
      Width           =   180
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "3.-"
      Height          =   195
      Index           =   30
      Left            =   120
      TabIndex        =   25
      Top             =   2700
      Width           =   180
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "4.-"
      Height          =   195
      Index           =   31
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "5.-"
      Height          =   195
      Index           =   32
      Left            =   120
      TabIndex        =   23
      Top             =   3300
      Width           =   180
   End
   Begin VB.Label Label1 
      Caption         =   "Familiares Asignados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblTipCob 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7800
      TabIndex        =   16
      Top             =   645
      Width           =   1935
   End
   Begin VB.Label Label11 
      Caption         =   "Tipo de Cobro"
      Height          =   195
      Left            =   7560
      TabIndex        =   15
      Top             =   460
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "Grado"
      Height          =   195
      Left            =   5160
      TabIndex        =   14
      Top             =   465
      Width           =   1335
   End
   Begin VB.Label lblGrado 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5400
      TabIndex        =   13
      Top             =   645
      Width           =   1935
   End
   Begin VB.Label lblE_socio 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   645
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Estado del Socio"
      Height          =   195
      Left            =   2400
      TabIndex        =   9
      Top             =   460
      Width           =   1335
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Ins"
      Height          =   195
      Left            =   1080
      TabIndex        =   8
      Top             =   460
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "D.N.I."
      Height          =   195
      Left            =   1440
      TabIndex        =   7
      Top             =   460
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Cod.Socio"
      Height          =   195
      Left            =   2160
      TabIndex        =   2
      Top             =   195
      Width           =   855
   End
   Begin VB.Label lblCodSocio 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   195
      Width           =   5895
   End
End
Attribute VB_Name = "frmDIECODetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalir_Click()
   Dim wTotAport As Currency, wTotDeuda As Currency, wTotAdela As Currency, wNetSocio As Currency, wDscSocio As Currency, wDifSocio As Currency, _
       wTotAsig1 As Currency, wDeuAsig1 As Currency, wAdeAsig1 As Currency, wNetAsig1 As Currency, wDscAsig1 As Currency, wDifAsig1 As Currency, _
       wTotAsig2 As Currency, wDeuAsig2 As Currency, wAdeAsig2 As Currency, wNetAsig2 As Currency, wDscAsig2 As Currency, wDifAsig2 As Currency, _
       wTotAsig3 As Currency, wDeuAsig3 As Currency, wAdeAsig3 As Currency, wNetAsig3 As Currency, wDscAsig3 As Currency, wDifAsig3 As Currency, _
       wTotAsig4 As Currency, wDeuAsig4 As Currency, wAdeAsig4 As Currency, wNetAsig4 As Currency, wDscAsig4 As Currency, wDifAsig4 As Currency, _
       wTotAsig5 As Currency, wDeuAsig5 As Currency, wAdeAsig5 As Currency, wNetAsig5 As Currency, wDscAsig5 As Currency, wDifAsig5 As Currency, _
       wSoc As Integer, wCod As Long, wIns As Integer, _
       wTotEnvio As Currency, wTotDieco As Currency, wDscDifer As Currency, _
       wCodAsig1 As Integer, wCodAsig2 As Integer, wCodAsig3 As Integer, wCodAsig4 As Integer, wCodAsig5 As Integer
   
   If cmdSalir.Caption = "Grabar" Then
      zDetaSw = True
      
      wSoc = Val(txtCodSocio.Text)
      wCodAsig1 = Val(txtSocio1.Text)
      wCodAsig2 = Val(txtSocio2.Text)
      wCodAsig3 = Val(txtSocio3.Text)
      wCodAsig4 = Val(txtSocio4.Text)
      wCodAsig5 = Val(txtSocio5.Text)
      
      wTotAport = Val(txtTotSocio.Text)
      wTotDeuda = Val(txtDeuSocio.Text)
      wTotAdela = Val(txtAdeSocio.Text)
      wNetSocio = Val(txtNetSocio.Text)
      wDscSocio = Val(txtDscSocio.Text)
      wDifSocio = Val(txtDifSocio.Text)
   
      wTotAsig1 = Val(txtTotAsig1.Text)
      wTotAsig1 = Val(txtDeuAsig1.Text)
      wTotAsig1 = Val(txtAdeAsig1.Text)
      wNetAsig1 = Val(txtNetAsig1.Text)
      wDscAsig1 = Val(txtDscAsig1.Text)
      wDifAsig1 = Val(txtDifAsig1.Text)
   
      wTotAsig2 = Val(txtTotAsig2.Text)
      wTotAsig2 = Val(txtDeuAsig2.Text)
      wTotAsig2 = Val(txtAdeAsig2.Text)
      wNetAsig2 = Val(txtNetAsig2.Text)
      wDscAsig2 = Val(txtDscAsig2.Text)
      wDifAsig2 = Val(txtDifAsig2.Text)
   
      wTotAsig3 = Val(txtTotAsig3.Text)
      wTotAsig3 = Val(txtDeuAsig3.Text)
      wTotAsig3 = Val(txtAdeAsig3.Text)
      wNetAsig3 = Val(txtNetAsig3.Text)
      wDscAsig3 = Val(txtDscAsig3.Text)
      wDifAsig3 = Val(txtDifAsig3.Text)
   
      wTotAsig4 = Val(txtTotAsig4.Text)
      wTotAsig4 = Val(txtDeuAsig4.Text)
      wTotAsig4 = Val(txtAdeAsig4.Text)
      wNetAsig4 = Val(txtNetAsig4.Text)
      wDscAsig4 = Val(txtDscAsig4.Text)
      wDifAsig4 = Val(txtDifAsig4.Text)
   
      wTotAsig5 = Val(txtTotAsig5.Text)
      wTotAsig5 = Val(txtDeuAsig5.Text)
      wTotAsig5 = Val(txtAdeAsig5.Text)
      wNetAsig5 = Val(txtNetAsig5.Text)
      wDscAsig5 = Val(txtDscAsig5.Text)
      wDifAsig5 = Val(txtDifAsig5.Text)
   
      wTotEnvio = wNetSocio + wNetAsig1 + wNetAsig2 + wNetAsig3 + wNetAsig4 + wNetAsig5
      wTotDieco = wDscSocio + wDscAsig1 + wDscAsig2 + wDscAsig3 + wDscAsig4 + wDscAsig5
      wDscDifer = wDifSocio + wDifAsig1 + wDifAsig2 + wDifAsig3 + wDifAsig4 + wDifAsig5
   
      Db.BeginTrans
      Db.Execute ("UPDATE TMP_DIECOCAB " _
      & " SET TOTAPORT=" + Str(wTotAport) + ", TOTDEUDA=" + Str(wTotDeuda) + ", TOTADELA=" + Str(wTotAdela) + ", " _
      & "     NETSOCIO=" + Str(wNetSocio) + ", DSCSOCIO=" + Str(wDscSocio) + ", DIFSOCIO=" + Str(wDifSocio) + ", " _
      & "     TOTASIG1=" + Str(wTotAsig1) + ", DEUASIG1=" + Str(wDeuAsig1) + ", ADEASIG1=" + Str(wAdeAsig1) + ", " _
      & "     NETASIG1=" + Str(wNetAsig1) + ", DSCASIG1=" + Str(wDscAsig1) + ", DIFASIG1=" + Str(wDifAsig1) + ", " _
      & "     TOTASIG2=" + Str(wTotAsig2) + ", DEUASIG2=" + Str(wDeuAsig2) + ", ADEASIG2=" + Str(wAdeAsig2) + ", " _
      & "     NETASIG2=" + Str(wNetAsig2) + ", DSCASIG2=" + Str(wDscAsig2) + ", DIFASIG2=" + Str(wDifAsig2) + ", " _
      & "     TOTASIG3=" + Str(wTotAsig3) + ", DEUASIG3=" + Str(wDeuAsig3) + ", ADEASIG3=" + Str(wAdeAsig3) + ", " _
      & "     NETASIG3=" + Str(wNetAsig3) + ", DSCASIG3=" + Str(wDscAsig3) + ", DIFASIG3=" + Str(wDifAsig3) + ", " _
      & "     TOTASIG4=" + Str(wTotAsig4) + ", DEUASIG4=" + Str(wDeuAsig4) + ", ADEASIG4=" + Str(wAdeAsig4) + ", " _
      & "     NETASIG4=" + Str(wNetAsig4) + ", DSCASIG4=" + Str(wDscAsig4) + ", DIFASIG4=" + Str(wDifAsig4) + ", " _
      & "     TOTASIG5=" + Str(wTotAsig5) + ", DEUASIG5=" + Str(wDeuAsig5) + ", ADEASIG5=" + Str(wAdeAsig5) + ", " _
      & "     NETASIG5=" + Str(wNetAsig5) + ", DSCASIG5=" + Str(wDscAsig5) + ", DIFASIG5=" + Str(wDifAsig5) + ", " _
      & "     CODASIG1=" + Str(wCodAsig1) + ", CODASIG2=" + Str(wCodAsig2) + ", CODASIG3=" + Str(wCodAsig3) + ", " _
      & "     CODASIG4=" + Str(wCodAsig4) + ", CODASIG5=" + Str(wCodAsig5) + ", " _
      & "     TOTENVIO=" + Str(wTotEnvio) + ", DSCDIECO=" + Str(wTotDieco) + ", DSCDIFER=" + Str(wDscDifer) + " " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "            MES = '" + zDetaAnoDsc + zDetaMesDsc + "' AND " _
      & "            USU = '" + wcodusu + "' ")
      Db.CommitTrans
   
      Db.BeginTrans
      Db.Execute ("UPDATE DIECOCAB " _
      & " SET TOTAPORT=" + Str(wTotAport) + ", TOTDEUDA=" + Str(wTotDeuda) + ", TOTADELA=" + Str(wTotAdela) + ", " _
      & "     NETSOCIO=" + Str(wNetSocio) + ", DSCSOCIO=" + Str(wDscSocio) + ", DIFSOCIO=" + Str(wDifSocio) + ", " _
      & "     TOTASIG1=" + Str(wTotAsig1) + ", DEUASIG1=" + Str(wDeuAsig1) + ", ADEASIG1=" + Str(wAdeAsig1) + ", " _
      & "     NETASIG1=" + Str(wNetAsig1) + ", DSCASIG1=" + Str(wDscAsig1) + ", DIFASIG1=" + Str(wDifAsig1) + ", " _
      & "     TOTASIG2=" + Str(wTotAsig2) + ", DEUASIG2=" + Str(wDeuAsig2) + ", ADEASIG2=" + Str(wAdeAsig2) + ", " _
      & "     NETASIG2=" + Str(wNetAsig2) + ", DSCASIG2=" + Str(wDscAsig2) + ", DIFASIG2=" + Str(wDifAsig2) + ", " _
      & "     TOTASIG3=" + Str(wTotAsig3) + ", DEUASIG3=" + Str(wDeuAsig3) + ", ADEASIG3=" + Str(wAdeAsig3) + ", " _
      & "     NETASIG3=" + Str(wNetAsig3) + ", DSCASIG3=" + Str(wDscAsig3) + ", DIFASIG3=" + Str(wDifAsig3) + ", " _
      & "     TOTASIG4=" + Str(wTotAsig4) + ", DEUASIG4=" + Str(wDeuAsig4) + ", ADEASIG4=" + Str(wAdeAsig4) + ", " _
      & "     NETASIG4=" + Str(wNetAsig4) + ", DSCASIG4=" + Str(wDscAsig4) + ", DIFASIG4=" + Str(wDifAsig4) + ", " _
      & "     TOTASIG5=" + Str(wTotAsig5) + ", DEUASIG5=" + Str(wDeuAsig5) + ", ADEASIG5=" + Str(wAdeAsig5) + ", " _
      & "     NETASIG5=" + Str(wNetAsig5) + ", DSCASIG5=" + Str(wDscAsig5) + ", DIFASIG5=" + Str(wDifAsig5) + ", " _
      & "     CODASIG1=" + Str(wCodAsig1) + ", CODASIG2=" + Str(wCodAsig2) + ", CODASIG3=" + Str(wCodAsig3) + ", " _
      & "     CODASIG4=" + Str(wCodAsig4) + ", CODASIG5=" + Str(wCodAsig5) + ", " _
      & "     TOTENVIO=" + Str(wTotEnvio) + ", DSCDIECO=" + Str(wTotDieco) + ", DSCDIFER=" + Str(wDscDifer) + " " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "            MES = '" + zDetaAnoDsc + zDetaMesDsc + "' ")
      Db.CommitTrans
   
      MsgBox "Registro Grabado OK", vbExclamation
   End If
   
   Unload Me
End Sub

Private Sub Form_Activate()
   frmDIECODetalle.Left = (Screen.Width - Width) \ 2
   frmDIECODetalle.Top = 0
   
   Limpiar

   If zDetaCambio = True Then
      lblMensaje.Caption = "MODIFICAR"
      cmdSalir.Caption = "Grabar"
   Else
      lblMensaje.Caption = "CONSULTAR"
      cmdSalir.Caption = "Salir"
   End If

   txtCodSocio.Text = zDetaCodSoc
   LlenaCab
   TotalCab

   If zDetaCambio = True Then
      txtTotSocio.Enabled = True
      txtDeuSocio.Enabled = True
      txtAdeSocio.Enabled = True
   
      If Len(Trim(txtSocio1.Text)) > 0 Then
         txtTotAsig1.Enabled = True
         txtDeuAsig1.Enabled = True
         txtAdeAsig1.Enabled = True
      End If
   
      If Len(Trim(txtSocio2.Text)) > 0 Then
         txtTotAsig2.Enabled = True
         txtDeuAsig2.Enabled = True
         txtAdeAsig2.Enabled = True
      End If
   
      If Len(Trim(txtSocio3.Text)) > 0 Then
         txtTotAsig3.Enabled = True
         txtDeuAsig3.Enabled = True
         txtAdeAsig3.Enabled = True
      End If
   
      If Len(Trim(txtSocio4.Text)) > 0 Then
         txtTotAsig4.Enabled = True
         txtDeuAsig4.Enabled = True
         txtAdeAsig4.Enabled = True
      End If
   
      If Len(Trim(txtSocio5.Text)) > 0 Then
         txtTotAsig5.Enabled = True
         txtDeuAsig5.Enabled = True
         txtAdeAsig5.Enabled = True
      End If
   
      txtTotSocio.SetFocus
   Else
      cmdSalir.SetFocus
   End If
End Sub

Private Sub LlenaCab()
   Dim aa As Integer

   aa = Leerado5a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(zDetaCodSoc) + " ")
   If aa > 0 Then
      txtCodigo.Text = ADO5a!codigo
      txtIns.Text = ADO5a!ins
      txtNumdoc.Text = ADO5a!numdoc
      txtE_socio.Text = ADO5a!e_socio
      txtGrado.Text = ADO5a!grado
      txtTipCob.Text = ADO5a!tipcob
   End If

   If zDetaTipDsc = "01" Then
      aa = Leerado6a("SELECT * FROM DIECOCAB " _
                & " WHERE      MES = '" + zDetaAnoDsc + zDetaMesDsc + "' AND " _
                & "       CODSOCIO = " + Str(zDetaCodSoc) + "  ")
   Else
      aa = Leerado6a("SELECT * FROM CAJMPCAB " _
                & " WHERE      MES = '" + zDetaAnoDsc + zDetaMesDsc + "' AND " _
                & "       CODSOCIO = " + Str(zDetaCodSoc) + "  ")
   End If
   If aa = 0 Then
      MsgBox "Primero se Debe Grabar El Calculo"
      Exit Sub
   End If
   txtTotSocio.Text = Format(ADO6a!totaport, "###,##0.00;;\ ")
   txtDeuSocio.Text = Format(ADO6a!totdeuda, "###,##0.00;;\ ")
   txtAdeSocio.Text = Format(ADO6a!totadela, "###,##0.00;;\ ")
   txtNetSocio.Text = Format(ADO6a!netsocio, "###,##0.00;;\ ")
   txtDscSocio.Text = Format(ADO6a!dscsocio, "###,##0.00;;\ ")
   txtDifSocio.Text = Format(ADO6a!difsocio, "###,##0.00;;\ ")
   
   txtSocio1.Text = Format(ADO6a!codasig1, "####0;;\ ")
   txtTotAsig1.Text = Format(ADO6a!totasig1, "###,##0.00;;\ ")
   txtDeuAsig1.Text = Format(ADO6a!deuasig1, "###,##0.00;;\ ")
   txtAdeAsig1.Text = Format(ADO6a!adeasig1, "###,##0.00;;\ ")
   txtNetAsig1.Text = Format(ADO6a!netasig1, "###,##0.00;;\ ")
   txtDscAsig1.Text = Format(ADO6a!dscasig1, "###,##0.00;;\ ")
   txtDifAsig1.Text = Format(ADO6a!difasig1, "###,##0.00;;\ ")

   txtSocio2.Text = Format(ADO6a!codasig2, "####0;;\ ")
   txtTotAsig2.Text = Format(ADO6a!totasig2, "###,##0.00;;\ ")
   txtDeuAsig2.Text = Format(ADO6a!deuasig2, "###,##0.00;;\ ")
   txtAdeAsig2.Text = Format(ADO6a!adeasig2, "###,##0.00;;\ ")
   txtNetAsig2.Text = Format(ADO6a!netasig2, "###,##0.00;;\ ")
   txtDscAsig2.Text = Format(ADO6a!dscasig2, "###,##0.00;;\ ")
   txtDifAsig2.Text = Format(ADO6a!difasig2, "###,##0.00;;\ ")
   
   txtSocio3.Text = Format(ADO6a!codasig3, "####0;;\ ")
   txtTotAsig3.Text = Format(ADO6a!totasig3, "###,##0.00;;\ ")
   txtDeuAsig3.Text = Format(ADO6a!deuasig3, "###,##0.00;;\ ")
   txtAdeAsig3.Text = Format(ADO6a!adeasig3, "###,##0.00;;\ ")
   txtNetAsig3.Text = Format(ADO6a!netasig3, "###,##0.00;;\ ")
   txtDscAsig3.Text = Format(ADO6a!dscasig3, "###,##0.00;;\ ")
   txtDifAsig3.Text = Format(ADO6a!difasig3, "###,##0.00;;\ ")
   
   txtSocio4.Text = Format(ADO6a!codasig4, "####0;;\ ")
   txtTotAsig4.Text = Format(ADO6a!totasig4, "###,##0.00;;\ ")
   txtDeuAsig4.Text = Format(ADO6a!deuasig4, "###,##0.00;;\ ")
   txtAdeAsig4.Text = Format(ADO6a!adeasig4, "###,##0.00;;\ ")
   txtNetAsig4.Text = Format(ADO6a!netasig4, "###,##0.00;;\ ")
   txtDscAsig4.Text = Format(ADO6a!dscasig4, "###,##0.00;;\ ")
   txtDifAsig4.Text = Format(ADO6a!difasig4, "###,##0.00;;\ ")
  
   txtSocio5.Text = Format(ADO6a!codasig5, "####0;;\ ")
   txtTotAsig5.Text = Format(ADO6a!totasig5, "###,##0.00;;\ ")
   txtDeuAsig5.Text = Format(ADO6a!deuasig5, "###,##0.00;;\ ")
   txtAdeAsig5.Text = Format(ADO6a!adeasig5, "###,##0.00;;\ ")
   txtNetAsig5.Text = Format(ADO6a!netasig5, "###,##0.00;;\ ")
   txtDscAsig5.Text = Format(ADO6a!dscasig5, "###,##0.00;;\ ")
   txtDifAsig5.Text = Format(ADO6a!difasig5, "###,##0.00;;\ ")

   cmdSalir.SetFocus
End Sub

Private Sub TotalCab()
   Dim aa As Integer, wTotEnvio As Currency, wTotDscto As Currency, wTotDifer As Currency
   
   wTotEnvio = Val(txtNetSocio.Text) + Val(txtNetAsig1.Text) + _
               Val(txtNetAsig2.Text) + Val(txtNetAsig3.Text) + _
               Val(txtNetAsig4.Text) + Val(txtNetAsig5.Text)

   wTotDscto = Val(txtDscSocio.Text) + Val(txtDscAsig1.Text) + _
               Val(txtDscAsig2.Text) + Val(txtDscAsig3.Text) + _
               Val(txtDscAsig4.Text) + Val(txtDscAsig5.Text)

   wTotDifer = Val(txtDifSocio.Text) + Val(txtDifAsig1.Text) + _
               Val(txtDifAsig2.Text) + Val(txtDifAsig3.Text) + _
               Val(txtDifAsig4.Text) + Val(txtDifAsig5.Text)

   lblEnvio.Caption = Format(wTotEnvio, "###,##0.00")
   lblDscto.Caption = Format(wTotDscto, "###,##0.00")
   lblDifer.Caption = Format(wTotDifer, "###,##0.00")
End Sub

Private Sub Limpiar()
   txtCodSocio.Text = ""
   txtCodigo.Text = ""
   txtIns.Text = ""
   txtNumdoc.Text = ""
   txtE_socio.Text = ""
   txtGrado.Text = ""
   txtTipCob.Text = ""
End Sub

Private Sub txtAdeAsig1_GotFocus()
   txtAdeAsig1.SelStart = 0
   txtAdeAsig1.SelLength = Len(Trim(txtAdeAsig1.Text))
End Sub

Private Sub txtAdeAsig1_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDeuAsig1.SetFocus
   Case 40
        If txtAdeAsig2.Enabled = True Then
           txtAdeAsig2.SetFocus
        End If
   End Select
End Sub

Private Sub txtAdeAsig1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtAdeAsig1.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtAdeAsig1.Text = ""
         Exit Sub
      End If
      txtAdeAsig1.Text = Format(txtAdeAsig1.Text, "###,##0.00;;\ ")
      Recal
      If txtDeuAsig2.Enabled = True Then
         txtDeuAsig2.SetFocus
      Else
         cmdSalir.SetFocus
      End If
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtAdeAsig2_GotFocus()
   txtAdeAsig2.SelStart = 0
   txtAdeAsig2.SelLength = Len(Trim(txtAdeAsig2.Text))
End Sub

Private Sub txtAdeAsig2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDeuAsig2.SetFocus
   Case 40
        If txtAdeAsig3.Enabled = True Then
           txtAdeAsig3.SetFocus
        End If
   End Select
End Sub

Private Sub txtAdeAsig2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtAdeAsig2.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtAdeAsig2.Text = ""
         Exit Sub
      End If
      txtAdeAsig2.Text = Format(txtAdeAsig2.Text, "###,##0.00;;\ ")
      Recal
      If txtDeuAsig3.Enabled = True Then
         txtDeuAsig3.SetFocus
      Else
         cmdSalir.SetFocus
      End If
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtAdeAsig3_GotFocus()
   txtAdeAsig3.SelStart = 0
   txtAdeAsig3.SelLength = Len(Trim(txtAdeAsig3.Text))
End Sub

Private Sub txtAdeAsig3_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDeuAsig3.SetFocus
   Case 40
        If txtAdeAsig4.Enabled = True Then
           txtAdeAsig4.SetFocus
        End If
   End Select
End Sub

Private Sub txtAdeAsig3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtAdeAsig3.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtAdeAsig3.Text = ""
         Exit Sub
      End If
      txtAdeAsig3.Text = Format(txtAdeAsig3.Text, "###,##0.00;;\ ")
      Recal
      If txtDeuAsig4.Enabled = True Then
         txtDeuAsig4.SetFocus
      Else
         cmdSalir.SetFocus
      End If
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtAdeAsig4_GotFocus()
   txtAdeAsig4.SelStart = 0
   txtAdeAsig4.SelLength = Len(Trim(txtAdeAsig4.Text))
End Sub

Private Sub txtAdeAsig4_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDeuAsig4.SetFocus
   Case 40
        If txtAdeAsig5.Enabled = True Then
           txtAdeAsig5.SetFocus
        End If
   End Select
End Sub

Private Sub txtAdeAsig4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtAdeAsig4.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtAdeAsig4.Text = ""
         Exit Sub
      End If
      txtAdeAsig4.Text = Format(txtAdeAsig4.Text, "###,##0.00;;\ ")
      Recal
      If txtDeuAsig5.Enabled = True Then
         txtDeuAsig5.SetFocus
      Else
         cmdSalir.SetFocus
      End If
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtAdeAsig5_GotFocus()
   txtAdeAsig5.SelStart = 0
   txtAdeAsig5.SelLength = Len(Trim(txtAdeAsig5.Text))
End Sub

Private Sub txtAdeAsig5_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDeuAsig5.SetFocus
   Case 40
   
   End Select
End Sub

Private Sub txtAdeAsig5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtAdeAsig5.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtAdeAsig5.Text = ""
         Exit Sub
      End If
      txtAdeAsig5.Text = Format(txtAdeAsig5.Text, "###,##0.00;;\ ")
      Recal
      
      cmdSalir.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtAdeSocio_GotFocus()
   txtAdeSocio.SelStart = 0
   txtAdeSocio.SelLength = Len(Trim(txtAdeSocio.Text))
End Sub

Private Sub txtAdeSocio_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDeuSocio.SetFocus
   Case 40
        txtTotAsig1.SetFocus
   End Select
End Sub

Private Sub txtAdeSocio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtAdeSocio.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtAdeSocio.Text = ""
         Exit Sub
      End If
      txtAdeSocio.Text = Format(txtAdeSocio.Text, "###,##0.00;;\ ")
      Recal
      If txtTotAsig1.Enabled = True Then
         txtTotAsig1.SetFocus
      End If
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtCodSocio_Change()
   Dim aa As Integer
   aa = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " ")
   If aa > 0 Then
      lblCodSocio.Caption = ADO6a!nombre
      txtCodigo.Text = ADO6a!codigo
      txtIns.Text = ADO6a!ins
      txtNumdoc.Text = ADO6a!numdoc
      txtE_socio.Text = ADO6a!e_socio
      txtGrado.Text = ADO6a!grado
      txtTipCob.Text = ADO6a!tipcob
   Else
      lblCodSocio.Caption = ""
      Limpiar
   End If
   Set ADO6a = Nothing
End Sub

Private Sub txtDeuAsig1_GotFocus()
   txtDeuAsig1.SelStart = 0
   txtDeuAsig1.SelLength = Len(Trim(txtDeuAsig1.Text))
End Sub

Private Sub txtDeuAsig1_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtAdeSocio.SetFocus
   Case 40
        If txtAdeAsig1.Enabled = True Then
           txtAdeAsig1.SetFocus
        End If
   End Select
End Sub

Private Sub txtDeuAsig1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtDeuAsig1.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtDeuAsig1.Text = ""
         Exit Sub
      End If
      txtDeuAsig1.Text = Format(txtDeuAsig1.Text, "###,##0.00;;\ ")
      Recal
      If txtAdeAsig1.Enabled = True Then
         txtAdeAsig1.SetFocus
      Else
         cmdSalir.SetFocus
      End If
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDeuAsig2_GotFocus()
   txtDeuAsig2.SelStart = 0
   txtDeuAsig2.SelLength = Len(Trim(txtDeuAsig2.Text))
End Sub

Private Sub txtDeuAsig2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtAdeAsig1.SetFocus
   Case 40
        If txtAdeAsig2.Enabled = True Then
           txtAdeAsig2.SetFocus
        End If
   End Select
End Sub

Private Sub txtDeuAsig2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtDeuAsig2.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtDeuAsig2.Text = ""
         Exit Sub
      End If
      txtDeuAsig2.Text = Format(txtDeuAsig2.Text, "###,##0.00;;\ ")
      Recal
      If txtAdeAsig2.Enabled = True Then
         txtAdeAsig2.SetFocus
      Else
         cmdSalir.SetFocus
      End If
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDeuAsig3_GotFocus()
   txtDeuAsig3.SelStart = 0
   txtDeuAsig3.SelLength = Len(Trim(txtDeuAsig3.Text))
End Sub

Private Sub txtDeuAsig3_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtAdeAsig2.SetFocus
   Case 40
        If txtAdeAsig3.Enabled = True Then
           txtAdeAsig3.SetFocus
        End If
   End Select
End Sub

Private Sub txtDeuAsig3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtDeuAsig3.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtDeuAsig3.Text = ""
         Exit Sub
      End If
      txtDeuAsig3.Text = Format(txtDeuAsig3.Text, "###,##0.00;;\ ")
      Recal
      If txtAdeAsig3.Enabled = True Then
         txtAdeAsig3.SetFocus
      Else
         cmdSalir.SetFocus
      End If
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDeuAsig4_GotFocus()
   txtDeuAsig4.SelStart = 0
   txtDeuAsig4.SelLength = Len(Trim(txtDeuAsig4.Text))
End Sub

Private Sub txtDeuAsig4_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtAdeAsig3.SetFocus
   Case 40
        If txtAdeAsig4.Enabled = True Then
           txtAdeAsig4.SetFocus
        End If
   End Select
End Sub

Private Sub txtDeuAsig4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtDeuAsig4.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtDeuAsig4.Text = ""
         Exit Sub
      End If
      txtDeuAsig4.Text = Format(txtDeuAsig4.Text, "###,##0.00;;\ ")
      Recal
      If txtAdeAsig4.Enabled = True Then
         txtAdeAsig4.SetFocus
      Else
         cmdSalir.SetFocus
      End If
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDeuAsig5_GotFocus()
   txtDeuAsig5.SelStart = 0
   txtDeuAsig5.SelLength = Len(Trim(txtDeuAsig5.Text))
End Sub

Private Sub txtDeuAsig5_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtAdeAsig4.SetFocus
   Case 40
        If txtAdeAsig5.Enabled = True Then
           txtAdeAsig5.SetFocus
        End If
   End Select
End Sub

Private Sub txtDeuAsig5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtDeuAsig5.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtDeuAsig5.Text = ""
         Exit Sub
      End If
      txtDeuAsig5.Text = Format(txtDeuAsig5.Text, "###,##0.00;;\ ")
      Recal
      If txtAdeAsig5.Enabled = True Then
         txtAdeAsig5.SetFocus
      Else
         cmdSalir.SetFocus
      End If
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDeuSocio_GotFocus()
   txtDeuSocio.SelStart = 0
   txtDeuSocio.SelLength = Len(Trim(txtDeuSocio.Text))
End Sub

Private Sub txtDeuSocio_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtTotSocio.SetFocus
   Case 40
        txtAdeSocio.SetFocus
   End Select
End Sub

Private Sub txtDeuSocio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtDeuSocio.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtDeuSocio.Text = ""
         Exit Sub
      End If
      txtDeuSocio.Text = Format(txtDeuSocio.Text, "###,##0.00;;\ ")
      Recal
      txtAdeSocio.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtE_socio_Change()
   Dim aa As Integer
   aa = Leerado8a("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + txtE_socio.Text + "' ")
   If aa > 0 Then
      lblE_socio.Caption = ADO8a!nombre
   Else
      lblE_socio.Caption = ""
   End If
   Set ADO8a = Nothing
End Sub

Private Sub txtGrado_Change()
   Dim aa As Integer
   aa = Leerado8a("SELECT * FROM MAEGRADO WHERE GRADO = " + Str(Val(txtGrado.Text)) + " ")
   If aa > 0 Then
      lblGrado.Caption = ADO8a!nombre
   Else
      lblGrado.Caption = ""
   End If
   Set ADO8a = Nothing
End Sub

Private Sub txtSocio1_Change()
   Dim aa As Integer
   aa = Leerado8a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio1.Text)) + " ")
   If aa > 0 Then
      lblCodigo1.Caption = ADO8a!codigo
      lblIns1.Caption = ADO8a!ins
      lblSocio1.Caption = ADO8a!nombre
   Else
      lblCodigo1.Caption = ""
      lblIns1.Caption = ""
      lblSocio1.Caption = ""
   End If
   Set ADO8a = Nothing
End Sub

Private Sub txtSocio2_Change()
   Dim aa As Integer
   aa = Leerado8a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio2.Text)) + " ")
   If aa > 0 Then
      lblCodigo2.Caption = ADO8a!codigo
      lblIns2.Caption = ADO8a!ins
      lblSocio2.Caption = ADO8a!nombre
   Else
      lblCodigo2.Caption = ""
      lblIns2.Caption = ""
      lblSocio2.Caption = ""
   End If
   Set ADO8a = Nothing
End Sub

Private Sub txtSocio3_Change()
   Dim aa As Integer
   aa = Leerado8a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio3.Text)) + " ")
   If aa > 0 Then
      lblCodigo3.Caption = ADO8a!codigo
      lblIns3.Caption = ADO8a!ins
      lblSocio3.Caption = ADO8a!nombre
   Else
      lblCodigo3.Caption = ""
      lblIns3.Caption = ""
      lblSocio3.Caption = ""
   End If
   Set ADO8a = Nothing
End Sub

Private Sub txtSocio4_Change()
   Dim aa As Integer
   aa = Leerado8a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio4.Text)) + " ")
   If aa > 0 Then
      lblCodigo4.Caption = ADO8a!codigo
      lblIns4.Caption = ADO8a!ins
      lblSocio4.Caption = ADO8a!nombre
   Else
      lblCodigo4.Caption = ""
      lblIns4.Caption = ""
      lblSocio4.Caption = ""
   End If
   Set ADO8a = Nothing
End Sub

Private Sub txtSocio5_Change()
   Dim aa As Integer
   aa = Leerado8a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(Val(txtSocio5.Text)) + " ")
   If aa > 0 Then
      lblCodigo5.Caption = ADO8a!codigo
      lblIns5.Caption = ADO8a!ins
      lblSocio5.Caption = ADO8a!nombre
   Else
      lblCodigo5.Caption = ""
      lblIns5.Caption = ""
      lblSocio5.Caption = ""
   End If
   Set ADO8a = Nothing
End Sub

Private Sub txtTipCob_Change()
   Dim aa As Integer
   aa = Leerado8a("SELECT * FROM MAETIPCOB WHERE TIPCOB = '" + txtTipCob.Text + "' ")
   If aa > 0 Then
      lblTipCob.Caption = ADO8a!nombre
   Else
      lblTipCob.Caption = ""
   End If
   Set ADO8a = Nothing
End Sub

Private Sub Recal()
   Dim aa As Integer, _
       wDeuSocio As Currency, wAdeSocio As Currency, wNetSocio As Currency, wDscSocio As Currency, wDifSocio As Currency, _
       wDeuAsig1 As Currency, wAdeAsig1 As Currency, wNetAsig1 As Currency, wDscAsig1 As Currency, wDifAsig1 As Currency, _
       wDeuAsig2 As Currency, wAdeAsig2 As Currency, wNetAsig2 As Currency, wDscAsig2 As Currency, wDifAsig2 As Currency, _
       wDeuAsig3 As Currency, wAdeAsig3 As Currency, wNetAsig3 As Currency, wDscAsig3 As Currency, wDifAsig3 As Currency, _
       wDeuAsig4 As Currency, wAdeAsig4 As Currency, wNetAsig4 As Currency, wDscAsig4 As Currency, wDifAsig4 As Currency, _
       wDeuAsig5 As Currency, wAdeAsig5 As Currency, wNetAsig5 As Currency, wDscAsig5 As Currency, wDifAsig5 As Currency, _
       wTotEnvio As Currency, wTotDscto As Currency, wTotDifer As Currency

   wTotEnvio = Val(lblEnvio.Caption)
   wTotDscto = Val(lblDscto.Caption)
   wTotDifer = Val(lblDifer.Caption)

   wTotSocio = Val(txtTotSocio.Text)
   wDeuSocio = Val(txtDeuSocio.Text)
   wAdeSocio = Val(txtAdeSocio.Text)
   wDscSocio = Val(txtDscSocio.Text)
   wDifSocio = Val(txtDifSocio.Text)

   wTotAsig1 = Val(txtTotAsig1.Text)
   wDeuAsig1 = Val(txtDeuAsig1.Text)
   wAdeAsig1 = Val(txtAdeAsig1.Text)
   wDscAsig1 = Val(txtDscAsig1.Text)
   wDifAsig1 = Val(txtDifAsig1.Text)

   wTotAsig2 = Val(txtTotAsig2.Text)
   wDeuAsig2 = Val(txtDeuAsig2.Text)
   wAdeAsig2 = Val(txtAdeAsig2.Text)
   wDscAsig2 = Val(txtDscAsig2.Text)
   wDifAsig2 = Val(txtDifAsig2.Text)

   wTotAsig3 = Val(txtTotAsig3.Text)
   wDeuAsig3 = Val(txtDeuAsig3.Text)
   wAdeAsig3 = Val(txtAdeAsig3.Text)
   wDscAsig3 = Val(txtDscAsig3.Text)
   wDifAsig3 = Val(txtDifAsig3.Text)

   wTotAsig4 = Val(txtTotAsig4.Text)
   wDeuAsig4 = Val(txtDeuAsig4.Text)
   wAdeAsig4 = Val(txtAdeAsig4.Text)
   wDscAsig4 = Val(txtDscAsig4.Text)
   wDifAsig4 = Val(txtDifAsig4.Text)

   wTotAsig5 = Val(txtTotAsig5.Text)
   wDeuAsig5 = Val(txtDeuAsig5.Text)
   wAdeAsig5 = Val(txtAdeAsig5.Text)
   wDscAsig5 = Val(txtDscAsig5.Text)
   wDifAsig5 = Val(txtDifAsig5.Text)

   wNetSocio = wTotSocio + wDeuSocio - wAdeSocio
   wNetAsig1 = wTotAsig1 + wDeuAsig1 - wAdeAsig1
   wNetAsig2 = wTotAsig2 + wDeuAsig2 - wAdeAsig2
   wNetAsig3 = wTotAsig3 + wDeuAsig3 - wAdeAsig3
   wNetAsig4 = wTotAsig4 + wDeuAsig4 - wAdeAsig4
   wNetAsig5 = wTotAsig5 + wDeuAsig5 - wAdeAsig5

   If wTotDscto > 0 Or wTotDifer > 0 Then
      wDifSocio = wNetSocio - wDscSocio
      wDifAsig1 = wNetAsig1 - wDscAsig1
      wDifAsig2 = wNetAsig2 - wDscAsig2
      wDifAsig3 = wNetAsig3 - wDscAsig3
      wDifAsig4 = wNetAsig4 - wDscAsig4
      wDifAsig5 = wNetAsig5 - wDscAsig5
   Else
      wDifSocio = 0: wDifAsig1 = 0: wDifAsig2 = 0: wDifAsig3 = 0: wDifAsig4 = 0: wDifAsig5 = 0
   End If

   wTotEnvio = wNetSocio + wNetAsig1 + wNetAsig2 + wNetAsig3 + wNetAsig4 + wNetAsig5
   wTotDscto = wDscSocio + wDscAsig1 + wDscAsig2 + wDscAsig3 + wDscAsig4 + wDscAsig5
   wTotDifer = wDifSocio + wDifAsig1 + wDifAsig2 + wDifAsig3 + wDifAsig4 + wDifAsig5

   txtNetSocio.Text = Format(wNetSocio, "##,##0.00;;\ ")
   txtNetAsig1.Text = Format(wNetAsig1, "##,##0.00;;\ ")
   txtNetAsig2.Text = Format(wNetAsig2, "##,##0.00;;\ ")
   txtNetAsig3.Text = Format(wNetAsig3, "##,##0.00;;\ ")
   txtNetAsig4.Text = Format(wNetAsig4, "##,##0.00;;\ ")
   txtNetAsig5.Text = Format(wNetAsig5, "##,##0.00;;\ ")

   txtDifSocio.Text = Format(wDifSocio, "##,##0.00;;\ ")
   txtDifAsig1.Text = Format(wDifAsig1, "##,##0.00;;\ ")
   txtDifAsig2.Text = Format(wDifAsig2, "##,##0.00;;\ ")
   txtDifAsig3.Text = Format(wDifAsig3, "##,##0.00;;\ ")
   txtDifAsig4.Text = Format(wDifAsig4, "##,##0.00;;\ ")
   txtDifAsig5.Text = Format(wDifAsig5, "##,##0.00;;\ ")

   lblEnvio.Caption = Format(wTotEnvio, "##,##0.00;;\ ")
   lblDscto.Caption = Format(wTotDscto, "##,##0.00;;\ ")
   lblDifer.Caption = Format(wTotDifer, "##,##0.00;;\ ")
End Sub

Private Sub txtTotAsig1_GotFocus()
   txtTotAsig1.SelStart = 0
   txtTotAsig1.SelLength = Len(Trim(txtTotAsig1.Text))
End Sub

Private Sub txtTotAsig1_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtAdeSocio.SetFocus
   Case 40
        txtDeuAsig1.SetFocus
   End Select
End Sub

Private Sub txtTotAsig1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtTotAsig1.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtTotAsig1.Text = ""
         Exit Sub
      End If
      txtTotAsig1.Text = Format(txtTotAsig1.Text, "###,##0.00;;\ ")
      Recal
      txtDeuAsig1.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtTotAsig2_GotFocus()
   txtTotAsig2.SelStart = 0
   txtTotAsig2.SelLength = Len(Trim(txtTotAsig2.Text))
End Sub

Private Sub txtTotAsig2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtAdeAsig1.SetFocus
   Case 40
        txtDeuAsig2.SetFocus
   End Select
End Sub

Private Sub txtTotAsig2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtTotAsig2.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtTotAsig2.Text = ""
         Exit Sub
      End If
      txtTotAsig2.Text = Format(txtTotAsig2.Text, "###,##0.00;;\ ")
      Recal
      txtDeuAsig2.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtTotAsig3_GotFocus()
   txtTotAsig3.SelStart = 0
   txtTotAsig3.SelLength = Len(Trim(txtTotAsig3.Text))
End Sub

Private Sub txtTotAsig3_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtAdeAsig2.SetFocus
   Case 40
        txtDeuAsig3.SetFocus
   End Select
End Sub

Private Sub txtTotAsig3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtTotAsig3.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtTotAsig3.Text = ""
         Exit Sub
      End If
      txtTotAsig3.Text = Format(txtTotAsig3.Text, "###,##0.00;;\ ")
      Recal
      txtDeuAsig3.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtTotAsig4_GotFocus()
   txtTotAsig4.SelStart = 0
   txtTotAsig4.SelLength = Len(Trim(txtTotAsig4.Text))
End Sub

Private Sub txtTotAsig4_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtAdeAsig3.SetFocus
   Case 40
        txtDeuAsig4.SetFocus
   End Select
End Sub

Private Sub txtTotAsig4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtTotAsig4.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtTotAsig4.Text = ""
         Exit Sub
      End If
      txtTotAsig4.Text = Format(txtTotAsig4.Text, "###,##0.00;;\ ")
      Recal
      txtDeuAsig4.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtTotAsig5_GotFocus()
   txtTotAsig5.SelStart = 0
   txtTotAsig5.SelLength = Len(Trim(txtTotAsig5.Text))
End Sub

Private Sub txtTotAsig5_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtAdeAsig4.SetFocus
   Case 40
        txtDeuAsig5.SetFocus
   End Select
End Sub

Private Sub txtTotAsig5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtTotAsig5.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtTotAsig5.Text = ""
         Exit Sub
      End If
      txtTotAsig5.Text = Format(txtTotAsig5.Text, "###,##0.00;;\ ")
      Recal
      txtDeuAsig5.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtTotSocio_GotFocus()
   txtTotSocio.SelStart = 0
   txtTotSocio.SelLength = Len(Trim(txtTotSocio.Text))
End Sub

Private Sub txtTotSocio_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
   
   Case 40
        txtDeuSocio.SetFocus
   End Select
End Sub

Private Sub txtTotSocio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(txtTotSocio.Text) < 0 Then
         MsgBox "Importe Es Negativo", vbExclamation
         txtTotSocio.Text = ""
         Exit Sub
      End If
      txtTotSocio.Text = Format(txtTotSocio.Text, "###,##0.00;;\ ")
      Recal
      txtDeuSocio.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub
