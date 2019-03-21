VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMaeFamilia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Familiares"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   12675
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNiet07 
      Height          =   285
      Left            =   6360
      MaxLength       =   50
      TabIndex        =   80
      Top             =   4700
      Width           =   4215
   End
   Begin VB.TextBox txtNiet06 
      Height          =   285
      Left            =   6360
      MaxLength       =   50
      TabIndex        =   79
      Top             =   4210
      Width           =   4215
   End
   Begin VB.TextBox txtNiet05 
      Height          =   285
      Left            =   6360
      MaxLength       =   50
      TabIndex        =   78
      Top             =   3760
      Width           =   4215
   End
   Begin VB.TextBox txtNiet04 
      Height          =   285
      Left            =   6360
      MaxLength       =   50
      TabIndex        =   77
      Top             =   3280
      Width           =   4215
   End
   Begin VB.TextBox txtNiet03 
      Height          =   285
      Left            =   6360
      MaxLength       =   50
      TabIndex        =   76
      Top             =   2840
      Width           =   4215
   End
   Begin VB.TextBox txtNiet02 
      Height          =   285
      Left            =   6360
      MaxLength       =   50
      TabIndex        =   75
      Top             =   2360
      Width           =   4215
   End
   Begin VB.TextBox txtNiet01 
      Height          =   285
      Left            =   6360
      MaxLength       =   50
      TabIndex        =   74
      Top             =   1900
      Width           =   4215
   End
   Begin VB.TextBox txtDniN02 
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
      Left            =   11520
      MaxLength       =   8
      TabIndex        =   73
      Top             =   2360
      Width           =   930
   End
   Begin VB.TextBox txtDniN04 
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
      Left            =   11520
      MaxLength       =   8
      TabIndex        =   72
      Top             =   3280
      Width           =   930
   End
   Begin VB.TextBox txtDniN06 
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
      Left            =   11520
      MaxLength       =   8
      TabIndex        =   71
      Top             =   4210
      Width           =   930
   End
   Begin VB.TextBox txtDniN01 
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
      Left            =   11520
      MaxLength       =   8
      TabIndex        =   70
      Top             =   1900
      Width           =   930
   End
   Begin VB.TextBox txtDniN03 
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
      Left            =   11520
      MaxLength       =   8
      TabIndex        =   69
      Top             =   2840
      Width           =   930
   End
   Begin VB.TextBox txtDniN05 
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
      Left            =   11520
      MaxLength       =   8
      TabIndex        =   68
      Top             =   3760
      Width           =   930
   End
   Begin VB.TextBox txtDniN07 
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
      Left            =   11520
      MaxLength       =   8
      TabIndex        =   67
      Top             =   4700
      Width           =   930
   End
   Begin VB.TextBox txtDniN08 
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
      Left            =   11520
      MaxLength       =   8
      TabIndex        =   66
      Top             =   5160
      Width           =   930
   End
   Begin VB.TextBox txtNiet08 
      Height          =   285
      Left            =   6360
      MaxLength       =   50
      TabIndex        =   65
      Top             =   5160
      Width           =   4215
   End
   Begin VB.TextBox txtHijo08 
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   62
      Top             =   5160
      Width           =   4215
   End
   Begin VB.TextBox txtDniH08 
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
      Left            =   5400
      MaxLength       =   8
      TabIndex        =   61
      Top             =   5160
      Width           =   930
   End
   Begin VB.TextBox txtDniH07 
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
      Left            =   5400
      MaxLength       =   8
      TabIndex        =   60
      Top             =   4700
      Width           =   930
   End
   Begin VB.TextBox txtDniH05 
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
      Left            =   5400
      MaxLength       =   8
      TabIndex        =   59
      Top             =   3760
      Width           =   930
   End
   Begin VB.TextBox txtDniH03 
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
      Left            =   5400
      MaxLength       =   8
      TabIndex        =   58
      Top             =   2840
      Width           =   930
   End
   Begin VB.TextBox txtDniH01 
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
      Left            =   5400
      MaxLength       =   8
      TabIndex        =   57
      Top             =   1900
      Width           =   930
   End
   Begin VB.TextBox txtDniH06 
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
      Left            =   5400
      MaxLength       =   8
      TabIndex        =   56
      Top             =   4210
      Width           =   930
   End
   Begin VB.TextBox txtDniH04 
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
      Left            =   5400
      MaxLength       =   8
      TabIndex        =   55
      Top             =   3280
      Width           =   930
   End
   Begin VB.TextBox txtDniH02 
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
      Left            =   5400
      MaxLength       =   8
      TabIndex        =   54
      Top             =   2360
      Width           =   930
   End
   Begin VB.TextBox txtDniMad 
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
      Left            =   5400
      MaxLength       =   8
      TabIndex        =   53
      Top             =   1420
      Width           =   930
   End
   Begin VB.TextBox txtDniPad 
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
      Left            =   11400
      MaxLength       =   8
      TabIndex        =   51
      Top             =   980
      Width           =   930
   End
   Begin VB.TextBox txtDniEsp 
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
      Left            =   5400
      MaxLength       =   8
      TabIndex        =   49
      Top             =   980
      Width           =   930
   End
   Begin VB.TextBox txtEsposa 
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   25
      Top             =   980
      Width           =   4215
   End
   Begin VB.TextBox txtPadre 
      Height          =   285
      Left            =   6360
      MaxLength       =   50
      TabIndex        =   24
      Top             =   980
      Width           =   4095
   End
   Begin VB.TextBox txtMadre 
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   23
      Top             =   1420
      Width           =   4215
   End
   Begin VB.TextBox txtHijo01 
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   22
      Top             =   1900
      Width           =   4215
   End
   Begin VB.TextBox txtHijo02 
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   21
      Top             =   2360
      Width           =   4215
   End
   Begin VB.TextBox txtHijo03 
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   20
      Top             =   2840
      Width           =   4215
   End
   Begin VB.TextBox txtHijo04 
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   19
      Top             =   3280
      Width           =   4215
   End
   Begin VB.TextBox txtHijo05 
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   18
      Top             =   3760
      Width           =   4215
   End
   Begin VB.TextBox txtHijo06 
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   17
      Top             =   4210
      Width           =   4215
   End
   Begin VB.TextBox txtHijo07 
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   16
      Top             =   4700
      Width           =   4215
   End
   Begin VB.TextBox txtCodSocio 
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
      Height          =   285
      Left            =   240
      MaxLength       =   8
      TabIndex        =   14
      Top             =   300
      Width           =   690
   End
   Begin VB.TextBox txtGrado 
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
      Left            =   7680
      MaxLength       =   3
      TabIndex        =   10
      Top             =   300
      Width           =   450
   End
   Begin VB.TextBox txtESocio 
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
      Left            =   10080
      MaxLength       =   3
      TabIndex        =   8
      Top             =   300
      Width           =   450
   End
   Begin VB.TextBox txtNombre 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   4
      Top             =   300
      Width           =   5520
   End
   Begin VB.TextBox txtCodigo 
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
      Left            =   960
      MaxLength       =   8
      TabIndex        =   3
      Top             =   300
      Width           =   810
   End
   Begin VB.TextBox txtIns 
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
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   2
      Top             =   300
      Width           =   330
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
      Height          =   495
      Left            =   9960
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6120
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
      Left            =   11400
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1095
   End
   Begin MSMask.MaskEdBox txtFnEsposa 
      Height          =   285
      Left            =   4440
      TabIndex        =   26
      Top             =   980
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnPadre 
      Height          =   285
      Left            =   10440
      TabIndex        =   27
      Top             =   980
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnMadre 
      Height          =   285
      Left            =   4440
      TabIndex        =   28
      Top             =   1420
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnHijo02 
      Height          =   285
      Left            =   4440
      TabIndex        =   29
      Top             =   2360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnHijo04 
      Height          =   285
      Left            =   4440
      TabIndex        =   30
      Top             =   3280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnHijo06 
      Height          =   285
      Left            =   4440
      TabIndex        =   31
      Top             =   4210
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnHijo01 
      Height          =   285
      Left            =   4440
      TabIndex        =   32
      Top             =   1900
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnHijo03 
      Height          =   285
      Left            =   4440
      TabIndex        =   33
      Top             =   2840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnHijo05 
      Height          =   285
      Left            =   4440
      TabIndex        =   34
      Top             =   3760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnHijo07 
      Height          =   285
      Left            =   4440
      TabIndex        =   35
      Top             =   4700
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnHijo08 
      Height          =   285
      Left            =   4440
      TabIndex        =   63
      Top             =   5160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnNiet02 
      Height          =   285
      Left            =   10560
      TabIndex        =   81
      Top             =   2360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnNiet04 
      Height          =   285
      Left            =   10560
      TabIndex        =   82
      Top             =   3280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnNiet06 
      Height          =   285
      Left            =   10560
      TabIndex        =   83
      Top             =   4210
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnNiet01 
      Height          =   285
      Left            =   10560
      TabIndex        =   84
      Top             =   1900
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnNiet03 
      Height          =   285
      Left            =   10560
      TabIndex        =   85
      Top             =   2840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnNiet05 
      Height          =   285
      Left            =   10560
      TabIndex        =   86
      Top             =   3760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnNiet07 
      Height          =   285
      Left            =   10560
      TabIndex        =   87
      Top             =   4700
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFnNiet08 
      Height          =   285
      Left            =   10560
      TabIndex        =   88
      Top             =   5160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nieto 07"
      Height          =   195
      Index           =   31
      Left            =   6360
      TabIndex        =   96
      Top             =   4520
      Width           =   600
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nieto 06"
      Height          =   195
      Index           =   30
      Left            =   6360
      TabIndex        =   95
      Top             =   4030
      Width           =   600
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nieto 05"
      Height          =   195
      Index           =   29
      Left            =   6360
      TabIndex        =   94
      Top             =   3590
      Width           =   600
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nieto 04"
      Height          =   195
      Index           =   28
      Left            =   6360
      TabIndex        =   93
      Top             =   3100
      Width           =   600
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nieto 03"
      Height          =   195
      Index           =   27
      Left            =   6360
      TabIndex        =   92
      Top             =   2660
      Width           =   600
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nieto 02"
      Height          =   195
      Index           =   26
      Left            =   6360
      TabIndex        =   91
      Top             =   2180
      Width           =   600
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nieto 01"
      Height          =   195
      Index           =   22
      Left            =   6360
      TabIndex        =   90
      Top             =   1730
      Width           =   600
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nieto 08"
      Height          =   195
      Index           =   10
      Left            =   6360
      TabIndex        =   89
      Top             =   4980
      Width           =   600
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Hijo 08"
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   64
      Top             =   4980
      Width           =   495
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "D.N.I."
      Height          =   195
      Index           =   6
      Left            =   11595
      TabIndex        =   52
      Top             =   800
      Width           =   420
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "D.N.I."
      Height          =   195
      Index           =   5
      Left            =   5595
      TabIndex        =   50
      Top             =   800
      Width           =   420
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Esposa"
      Height          =   195
      Index           =   11
      Left            =   240
      TabIndex        =   48
      Top             =   800
      Width           =   525
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Padre"
      Height          =   195
      Index           =   12
      Left            =   6360
      TabIndex        =   47
      Top             =   800
      Width           =   1860
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Familiares"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   240
      TabIndex        =   46
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Madre"
      Height          =   195
      Index           =   14
      Left            =   240
      TabIndex        =   45
      Top             =   1240
      Width           =   450
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Hijo 01"
      Height          =   195
      Index           =   15
      Left            =   240
      TabIndex        =   44
      Top             =   1730
      Width           =   495
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Hijo 02"
      Height          =   195
      Index           =   16
      Left            =   240
      TabIndex        =   43
      Top             =   2180
      Width           =   495
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Hijo 03"
      Height          =   195
      Index           =   17
      Left            =   240
      TabIndex        =   42
      Top             =   2660
      Width           =   495
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Hijo 04"
      Height          =   195
      Index           =   18
      Left            =   240
      TabIndex        =   41
      Top             =   3100
      Width           =   495
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Hijo 05"
      Height          =   195
      Index           =   19
      Left            =   240
      TabIndex        =   40
      Top             =   3590
      Width           =   495
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Hijo 06"
      Height          =   195
      Index           =   20
      Left            =   240
      TabIndex        =   39
      Top             =   4030
      Width           =   495
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Hijo 07"
      Height          =   195
      Index           =   21
      Left            =   240
      TabIndex        =   38
      Top             =   4520
      Width           =   495
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Fec.Nac."
      Height          =   195
      Index           =   23
      Left            =   4440
      TabIndex        =   37
      Top             =   800
      Width           =   900
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Fec.Nac."
      Height          =   195
      Index           =   24
      Left            =   10440
      TabIndex        =   36
      Top             =   800
      Width           =   900
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Codigo"
      Height          =   195
      Index           =   4
      Left            =   345
      TabIndex        =   15
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblESocio 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   10560
      TabIndex        =   13
      Top             =   300
      Width           =   1935
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Grado"
      Height          =   195
      Index           =   2
      Left            =   7680
      TabIndex        =   12
      Top             =   120
      Width           =   435
   End
   Begin VB.Label lblGrado 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8160
      TabIndex        =   11
      Top             =   300
      Width           =   1935
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "E-Socio"
      Height          =   195
      Index           =   25
      Left            =   10080
      TabIndex        =   9
      Top             =   120
      Width           =   555
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Codofin"
      Height          =   195
      Index           =   0
      Left            =   975
      TabIndex        =   7
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Nombre"
      Height          =   195
      Index           =   3
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   555
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Instit"
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   330
   End
End
Attribute VB_Name = "frmMaeFamilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGrabar_Click()
   Dim aa As Integer, _
       wNomPad As String, wNomEsp As String, wNomMad As String, _
       wNomHi1 As String, wNomHi2 As String, wNomHi3 As String, _
       wNomHi4 As String, wNomHi5 As String, wNomHi6 As String, _
       wNomHi7 As String, wNomHi8 As String, _
       wNomNi1 As String, wNomNi2 As String, wNomNi3 As String, _
       wNomNi4 As String, wNomNi5 As String, wNomNi6 As String, _
       wNomNi7 As String, wNomNi8 As String, _
       wDniPad As String, wDniEsp As String, wDniMad As String, _
       wDniHi1 As String, wDniHi2 As String, wDniHi3 As String, _
       wDniHi4 As String, wDniHi5 As String, wDniHi6 As String, _
       wDniHi7 As String, wDniHi8 As String, _
       wDniNi1 As String, wDniNi2 As String, wDniNi3 As String, _
       wDniNi4 As String, wDniNi5 As String, wDniNi6 As String, _
       wDniNi7 As String, wDniNi8 As String, _
       wFecPad As Date, wFecMad As Date, wFecEsp As Date, _
       wFecHi1 As Date, wFecHi2 As Date, wFecHi3 As Date, wFecHi4 As Date, _
       wFecHi5 As Date, wFecHi6 As Date, wFecHi7 As Date, wFecHi8 As Date, _
       wFecNi1 As Date, wFecNi2 As Date, wFecNi3 As Date, wFecNi4 As Date, _
       wFecNi5 As Date, wFecNi6 As Date, wFecNi7 As Date, wFecNi8 As Date, _
       wSoc As Long

   wSoc = Val(txtCodSocio.Text)
   wNomEsp = txtEsposa.Text
   wNomPad = txtPadre.Text
   wNomMad = txtMadre.Text
   wNomHi1 = txtHijo01.Text
   wNomHi2 = txtHijo02.Text
   wNomHi3 = txtHijo03.Text
   wNomHi4 = txtHijo04.Text
   wNomHi5 = txtHijo05.Text
   wNomHi6 = txtHijo06.Text
   wNomHi7 = txtHijo07.Text
   wNomHi8 = txtHijo08.Text
   wNomNi1 = txtNiet01.Text
   wNomNi2 = txtNiet02.Text
   wNomNi3 = txtNiet03.Text
   wNomNi4 = txtNiet04.Text
   wNomNi5 = txtNiet05.Text
   wNomNi6 = txtNiet06.Text
   wNomNi7 = txtNiet07.Text
   wNomNi8 = txtNiet08.Text

   wDniEsp = txtDniEsp.Text
   wDniPad = txtDniPad.Text
   wDniMad = txtDniMad.Text
   wDniHi1 = txtDniH01.Text
   wDniHi2 = txtDniH02.Text
   wDniHi3 = txtDniH03.Text
   wDniHi4 = txtDniH04.Text
   wDniHi5 = txtDniH05.Text
   wDniHi6 = txtDniH06.Text
   wDniHi7 = txtDniH07.Text
   wDniHi8 = txtDniH08.Text
   wDniNi1 = txtDniN01.Text
   wDniNi2 = txtDniN02.Text
   wDniNi3 = txtDniN03.Text
   wDniNi4 = txtDniN04.Text
   wDniNi5 = txtDniN05.Text
   wDniNi6 = txtDniN06.Text
   wDniNi7 = txtDniN07.Text
   wDniNi8 = txtDniN08.Text

   If IsDate(txtFnEsposa.Text) Then
      wFecEsp = Format(txtFnEsposa.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFnPadre.Text) Then
      wFecPad = Format(txtFnPadre.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFnMadre.Text) Then
      wFecMad = Format(txtFnMadre.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFnHijo01.Text) Then
      wFecHi1 = Format(txtFnHijo01.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFnHijo02.Text) Then
      wFecHi2 = Format(txtFnHijo02.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFnHijo03.Text) Then
      wFecHi3 = Format(txtFnHijo03.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFnHijo04.Text) Then
      wFecHi4 = Format(txtFnHijo04.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFnHijo05.Text) Then
      wFecHi5 = Format(txtFnHijo05.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFnHijo06.Text) Then
      wFecHi6 = Format(txtFnHijo06.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFnHijo07.Text) Then
      wFecHi7 = Format(txtFnHijo07.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFnHijo08.Text) Then
      wFecHi8 = Format(txtFnHijo08.Text, "dd/mm/yyyy")
   End If

   If IsDate(txtFnNiet01.Text) Then
      wFecNi1 = Format(txtFnNiet01.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFnNiet02.Text) Then
      wFecNi8 = Format(txtFnNiet02.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFnNiet03.Text) Then
      wFecNi3 = Format(txtFnNiet03.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFnNiet04.Text) Then
      wFecNi4 = Format(txtFnNiet04.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFnNiet05.Text) Then
      wFecNi5 = Format(txtFnNiet05.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFnNiet06.Text) Then
      wFecNi6 = Format(txtFnNiet06.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFnNiet07.Text) Then
      wFecNi7 = Format(txtFnNiet07.Text, "dd/mm/yyyy")
   End If
   If IsDate(txtFnNiet08.Text) Then
      wFecNi8 = Format(txtFnNiet08.Text, "dd/mm/yyyy")
   End If

   If Len(Trim(wNomEsp)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'E' AND " _
                  & "       LIN = '01' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'E', '01', '" + wDniEsp + "', " _
         & "  '" + wNomEsp + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniEsp + "', NOMBRE = '" + wNomEsp + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'E' AND " _
         & "       LIN = '01' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecEsp) And Format(wFecEsp, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecEsp, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'E' AND " _
         & "       LIN = '01' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'E' AND " _
         & "       LIN = '01' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'E' AND " _
      & "       LIN = '01' ")
      Db.CommitTrans
   End If

   If Len(Trim(wNomPad)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'P' AND " _
                  & "       LIN = '01' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'P', '01', '" + wDniPad + "', " _
         & "  '" + wNomPad + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniPad + "', NOMBRE = '" + wNomPad + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'P' AND " _
         & "       LIN = '01' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecPad) And Format(wFecPad, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecPad, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'P' AND " _
         & "       LIN = '01' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'P' AND " _
         & "       LIN = '01' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'P' AND " _
      & "       LIN = '01' ")
      Db.CommitTrans
   End If

   If Len(Trim(wNomMad)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'M' AND " _
                  & "       LIN = '01' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'M', '01', '" + wDniMad + "', " _
         & "  '" + wNomMad + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniMad + "', NOMBRE = '" + wNomMad + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'M' AND " _
         & "       LIN = '01' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecMad) And Format(wFecMad, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecMad, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'M' AND " _
         & "       LIN = '01' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'M' AND " _
         & "       LIN = '01' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'M' AND " _
      & "       LIN = '01' ")
      Db.CommitTrans
   End If

   If Len(Trim(wNomHi1)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'H' AND " _
                  & "       LIN = '01' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'H', '01', '" + wDniHi1 + "', " _
         & "  '" + wNomHi1 + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniHi1 + "', NOMBRE = '" + wNomHi1 + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '01' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecHi1) And Format(wFecHi1, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecHi1, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '01' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '01' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'H' AND " _
      & "       LIN = '01' ")
      Db.CommitTrans
   End If

   If Len(Trim(wNomHi2)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'H' AND " _
                  & "       LIN = '02' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'H', '02', '" + wDniHi2 + "', " _
         & "  '" + wNomHi2 + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniHi2 + "', NOMBRE = '" + wNomHi2 + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '02' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecHi2) And Format(wFecHi2, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecHi2, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '02' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '02' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'H' AND " _
      & "       LIN = '02' ")
      Db.CommitTrans
   End If

   If Len(Trim(wNomHi3)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'H' AND " _
                  & "       LIN = '03' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'H', '03', '" + wDniHi3 + "', " _
         & "  '" + wNomHi3 + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniHi3 + "', NOMBRE = '" + wNomHi3 + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '03' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecHi3) And Format(wFecHi3, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecHi3, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '03' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '03' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'H' AND " _
      & "       LIN = '03' ")
      Db.CommitTrans
   End If

   If Len(Trim(wNomHi4)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'H' AND " _
                  & "       LIN = '04' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'H', '04', '" + wDniHi4 + "', " _
         & "  '" + wNomHi4 + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniHi4 + "', NOMBRE = '" + wNomHi4 + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '04' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecHi4) And Format(wFecHi4, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecHi4, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '04' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '04' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'H' AND " _
      & "       LIN = '04' ")
      Db.CommitTrans
   End If

   If Len(Trim(wNomHi5)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'H' AND " _
                  & "       LIN = '05' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'H', '05', '" + wDniHi5 + "', " _
         & "  '" + wNomHi5 + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniHi5 + "', NOMBRE = '" + wNomHi5 + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '05' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecHi5) And Format(wFecHi5, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecHi5, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '05' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '05' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'H' AND " _
      & "       LIN = '05' ")
      Db.CommitTrans
   End If

   If Len(Trim(wNomHi6)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'H' AND " _
                  & "       LIN = '06' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'H', '06', '" + wDniHi6 + "', " _
         & "  '" + wNomHi6 + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniHi6 + "', NOMBRE = '" + wNomHi6 + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '06' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecHi6) And Format(wFecHi6, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecHi6, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '06' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '06' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'H' AND " _
      & "       LIN = '06' ")
      Db.CommitTrans
   End If

   If Len(Trim(wNomHi7)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'H' AND " _
                  & "       LIN = '07' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'H', '07', '" + wDniHi7 + "', " _
         & "  '" + wNomHi7 + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniHi7 + "', NOMBRE = '" + wNomHi7 + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '07' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecHi7) And Format(wFecHi7, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecHi7, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '07' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '07' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'H' AND " _
      & "       LIN = '07' ")
      Db.CommitTrans
   End If

   If Len(Trim(wNomHi8)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'H' AND " _
                  & "       LIN = '08' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'H', '08', '" + wDniHi8 + "', " _
         & "  '" + wNomHi8 + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniHi8 + "', NOMBRE = '" + wNomHi8 + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '08' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecHi8) And Format(wFecHi8, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecHi8, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '08' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'H' AND " _
         & "       LIN = '08' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'H' AND " _
      & "       LIN = '08' ")
      Db.CommitTrans
   End If

   
   
   
   
   
   
   If Len(Trim(wNomNi1)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'N' AND " _
                  & "       LIN = '01' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'N', '01', '" + wDniNi1 + "', " _
         & "  '" + wNomNi1 + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniNi1 + "', NOMBRE = '" + wNomNi1 + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '01' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecNi1) And Format(wFecNi1, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecNi1, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '01' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '01' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'N' AND " _
      & "       LIN = '01' ")
      Db.CommitTrans
   End If
   
   If Len(Trim(wNomNi2)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'N' AND " _
                  & "       LIN = '02' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'N', '02', '" + wDniNi2 + "', " _
         & "  '" + wNomNi2 + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniNi2 + "', NOMBRE = '" + wNomNi2 + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '02' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecNi2) And Format(wFecNi2, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecNi2, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '02' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '02' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'N' AND " _
      & "       LIN = '02' ")
      Db.CommitTrans
   End If
   
   If Len(Trim(wNomNi3)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'N' AND " _
                  & "       LIN = '03' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'N', '03', '" + wDniNi3 + "', " _
         & "  '" + wNomNi3 + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniNi3 + "', NOMBRE = '" + wNomNi3 + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '03' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecNi3) And Format(wFecNi3, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecNi3, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '03' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '03' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'N' AND " _
      & "       LIN = '03' ")
      Db.CommitTrans
   End If
   
   If Len(Trim(wNomNi4)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'N' AND " _
                  & "       LIN = '04' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'N', '04', '" + wDniNi4 + "', " _
         & "  '" + wNomNi4 + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniNi4 + "', NOMBRE = '" + wNomNi4 + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '04' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecNi4) And Format(wFecNi4, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecNi4, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '04' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '04' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'N' AND " _
      & "       LIN = '04' ")
      Db.CommitTrans
   End If
   
   If Len(Trim(wNomNi5)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'N' AND " _
                  & "       LIN = '05' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'N', '05', '" + wDniNi5 + "', " _
         & "  '" + wNomNi5 + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniNi5 + "', NOMBRE = '" + wNomNi5 + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '05' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecNi5) And Format(wFecNi5, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecNi5, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '05' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '05' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'N' AND " _
      & "       LIN = '05' ")
      Db.CommitTrans
   End If
   
   If Len(Trim(wNomNi6)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'N' AND " _
                  & "       LIN = '06' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'N', '06', '" + wDniNi6 + "', " _
         & "  '" + wNomNi6 + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniNi6 + "', NOMBRE = '" + wNomNi6 + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '06' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecNi6) And Format(wFecNi6, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecNi6, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '06' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '06' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'N' AND " _
      & "       LIN = '06' ")
      Db.CommitTrans
   End If
   
   
   If Len(Trim(wNomNi7)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'N' AND " _
                  & "       LIN = '07' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'N', '07', '" + wDniNi7 + "', " _
         & "  '" + wNomNi7 + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniNi7 + "', NOMBRE = '" + wNomNi7 + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '07' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecNi7) And Format(wFecNi7, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecNi7, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '07' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '07' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'N' AND " _
      & "       LIN = '07' ")
      Db.CommitTrans
   End If
   
   If Len(Trim(wNomNi8)) > 0 Then
      aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                  & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
                  & "       TIPOPARIENTE = 'N' AND " _
                  & "       LIN = '08' ")
      If aa = 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO MAEFAMILIA " _
         & " (CODSOCIO, TIPOPARIENTE, LIN, NUMDOC, NOMBRE, FECNAC, " _
         & "  FECVCM, FECREC, NUMREC, SERREC) " _
         & " VALUES " _
         & " (" + Str(wSoc) + ", 'N', '08', '" + wDniNi8 + "', " _
         & "  '" + wNomNi8 + "', NULL, NULL, NULL, '', '' ) ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET NUMDOC = '" + wDniNi8 + "', NOMBRE = '" + wNomNi8 + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '08' ")
         Db.CommitTrans
      End If
   
      If IsDate(wFecNi8) And Format(wFecNi8, "dd/mm/yyyy") <> "30/12/1899" Then
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = '" + Format(wFecNi8, "dd/mm/yyyy") + "' " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '08' ")
         Db.CommitTrans
      Else
         Db.BeginTrans
         Db.Execute ("UPDATE MAEFAMILIA " _
         & " SET FECNAC = NULL " _
         & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
         & "       TIPOPARIENTE = 'N' AND " _
         & "       LIN = '08' ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("DELETE FROM MAEFAMILIA " _
      & " WHERE CODSOCIO = " + Str(wSoc) + " AND " _
      & "       TIPOPARIENTE = 'N' AND " _
      & "       LIN = '08' ")
      Db.CommitTrans
   End If

   Call abrirEVENTO
   
   aa = LeeradoEvento1("SELECT * FROM MAESOCIO " _
                    & " WHERE CODIGO = '" + Format(wCod, "00000000") + "' AND " _
                    & "          INS = '" + Format(wIns, "0") + "' ")
   If aa > 0 Then
      DbEvento.BeginTrans
      DbEvento.Execute ("UPDATE MAESOCIO " _
      & " SET ESPOSA = '" + wNomEsp + "',  PADRE = '" + wNomPad + "',  MADRE = '" + wNomMad + "', " _
      & "     HIJO01 = '" + wNomHi1 + "', HIJO02 = '" + wNomHi2 + "', HIJO03 = '" + wNomHi3 + "', " _
      & "     HIJO04 = '" + wNomHi4 + "', HIJO05 = '" + wNomHi5 + "', HIJO06 = '" + wNomHi6 + "', " _
      & "     HIJO07 = '" + wNomHi7 + "' " _
      & " WHERE CODIGO = '" + Format(wCod, "00000000") + "' AND " _
      & "          INS = '" + Format(wIns, "0") + "' ")
      DbEvento.CommitTrans
   
      If IsDate(txtFnEsposa.Text) Then
         wFecEsp = Format(txtFnEsposa.Text, "dd/mm/yyyy")
      
         DbEvento.BeginTrans
         DbEvento.Execute ("UPDATE MAESOCIO " _
         & " SET FNESPOSA = '" + Format(wFecEsp, "dd/mm/yyyy") + "' " _
         & " WHERE CODIGO = '" + Format(wCod, "00000000") + "' AND " _
         & "          INS = '" + Format(wIns, "0") + "' ")
         DbEvento.CommitTrans
      End If
   
      If IsDate(txtFnHijo01.Text) Then
         wFecHi1 = Format(txtFnHijo01.Text, "dd/mm/yyyy")
      
         DbEvento.BeginTrans
         DbEvento.Execute ("UPDATE MAESOCIO " _
         & " SET FNHIJO01 = '" + Format(wFecHi1, "dd/mm/yyyy") + "' " _
         & " WHERE CODIGO = '" + Format(wCod, "00000000") + "' AND " _
         & "          INS = '" + Format(wIns, "0") + "' ")
         DbEvento.CommitTrans
      End If
      If IsDate(txtFnHijo02.Text) Then
         wFecHi2 = Format(txtFnHijo02.Text, "dd/mm/yyyy")
      
         DbEvento.BeginTrans
         DbEvento.Execute ("UPDATE MAESOCIO " _
         & " SET FNHIJO02 = '" + Format(wFecHi2, "dd/mm/yyyy") + "' " _
         & " WHERE CODIGO = '" + Format(wCod, "00000000") + "' AND " _
         & "          INS = '" + Format(wIns, "0") + "' ")
         DbEvento.CommitTrans
      End If
      If IsDate(txtFnHijo03.Text) Then
         wFecHi3 = Format(txtFnHijo03.Text, "dd/mm/yyyy")
      
         DbEvento.BeginTrans
         DbEvento.Execute ("UPDATE MAESOCIO " _
         & " SET FNHIJO03 = '" + Format(wFecHi3, "dd/mm/yyyy") + "' " _
         & " WHERE CODIGO = '" + Format(wCod, "00000000") + "' AND " _
         & "          INS = '" + Format(wIns, "0") + "' ")
         DbEvento.CommitTrans
      End If
      If IsDate(txtFnHijo04.Text) Then
         wFecHi4 = Format(txtFnHijo04.Text, "dd/mm/yyyy")
      
         DbEvento.BeginTrans
         DbEvento.Execute ("UPDATE MAESOCIO " _
         & " SET FNHIJO04 = '" + Format(wFecHi4, "dd/mm/yyyy") + "' " _
         & " WHERE CODIGO = '" + Format(wCod, "00000000") + "' AND " _
         & "          INS = '" + Format(wIns, "0") + "' ")
         DbEvento.CommitTrans
      End If
      If IsDate(txtFnHijo05.Text) Then
         wFecHi5 = Format(txtFnHijo05.Text, "dd/mm/yyyy")
      
         DbEvento.BeginTrans
         DbEvento.Execute ("UPDATE MAESOCIO " _
         & " SET FNHIJO05 = '" + Format(wFecHi5, "dd/mm/yyyy") + "' " _
         & " WHERE CODIGO = '" + Format(wCod, "00000000") + "' AND " _
         & "          INS = '" + Format(wIns, "0") + "' ")
         DbEvento.CommitTrans
      End If
      If IsDate(txtFnHijo06.Text) Then
         wFecHi6 = Format(txtFnHijo06.Text, "dd/mm/yyyy")
      
         DbEvento.BeginTrans
         DbEvento.Execute ("UPDATE MAESOCIO " _
         & " SET FNHIJO06 = '" + Format(wFecHi6, "dd/mm/yyyy") + "' " _
         & " WHERE CODIGO = '" + Format(wCod, "00000000") + "' AND " _
         & "          INS = '" + Format(wIns, "0") + "' ")
         DbEvento.CommitTrans
      End If
      If IsDate(txtFnHijo07.Text) Then
         wFecHi7 = Format(txtFnHijo07.Text, "dd/mm/yyyy")
      
         DbEvento.BeginTrans
         DbEvento.Execute ("UPDATE MAESOCIO " _
         & " SET FNHIJO07 = '" + Format(wFecHi7, "dd/mm/yyyy") + "' " _
         & " WHERE CODIGO = '" + Format(wCod, "00000000") + "' AND " _
         & "          INS = '" + Format(wIns, "0") + "' ")
         DbEvento.CommitTrans
      End If
   End If
   Set ADOEvento1 = Nothing
   
   DbEvento.Close

   MsgBox "Familiares Grabados OK", vbExclamation
   Unload Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   frmMaeFamilia.Left = (Screen.Width - Width) \ 2
   frmMaeFamilia.Top = 0
   
   txtCodSocio.Text = zSocio

   LlenaCab
   txtEsposa.SetFocus
End Sub

Private Sub LlenaCab()
   Dim aa As Integer, _
       wNomPad As String, wNomEsp As String, wNomMad As String, _
       wNomHi1 As String, wNomHi2 As String, wNomHi3 As String, _
       wNomHi4 As String, wNomHi5 As String, wNomHi6 As String, _
       wNomHi7 As String, wNomHi8 As String, _
       wNomNi1 As String, wNomNi2 As String, wNomNi3 As String, _
       wNomNi4 As String, wNomNi5 As String, wNomNi6 As String, _
       wNomNi7 As String, wNomNi8 As String, _
       wDniPad As String, wDniEsp As String, wDniMad As String, _
       wDniHi1 As String, wDniHi2 As String, wDniHi3 As String, _
       wDniHi4 As String, wDniHi5 As String, wDniHi6 As String, _
       wDniHi7 As String, wDniHi8 As String, _
       wDniNi1 As String, wDniNi2 As String, wDniNi3 As String, _
       wDniNi4 As String, wDniNi5 As String, wDniNi6 As String, _
       wDniNi7 As String, wDniNi8 As String, _
       wFecPad As Date, wFecMad As Date, wFecEsp As Date, _
       wFecHi1 As Date, wFecHi2 As Date, wFecHi3 As Date, wFecHi4 As Date, wFecHi5 As Date, _
       wFecHi6 As Date, wFecHi7 As Date, wFecHi8 As Date, wFecHi9 As Date, _
       wFecNi1 As Date, wFecNi2 As Date, wFecNi3 As Date, wFecNi4 As Date, wFecNi5 As Date, _
       wFecNi6 As Date, wFecNi7 As Date, wFecNi8 As Date, wFecNi9 As Date
   
   wNomEsp = "": wNomPad = "": wNomMad = ""
   wNomHi1 = "": wNomHi2 = "": wNomHi3 = "": wNomHi4 = ""
   wNomHi5 = "": wNomHi6 = "": wNomHi7 = "": wNomHi8 = ""
   wNomNi1 = "": wNomNi2 = "": wNomNi3 = "": wNomNi4 = ""
   wNomNi5 = "": wNomNi6 = "": wNomNi7 = "": wNomNi8 = ""
   
   wDniEsp = "": wDniPad = "": wDniMad = ""
   wDniHi1 = "": wDniHi2 = "": wDniHi3 = "": wDniHi4 = ""
   wDniHi5 = "": wDniHi6 = "": wDniHi7 = "": wDniHi8 = ""
   wDniNi1 = "": wDniNi2 = "": wDniNi3 = "": wDniNi4 = ""
   wDniNi5 = "": wDniNi6 = "": wDniNi7 = "": wDniNi8 = ""
   
   aa = Leerado8("SELECT * FROM MAEFAMILIA " _
                & " WHERE CODSOCIO = " + Str(Val(txtCodSocio.Text)) + " " _
                & " ORDER BY TIPOPARIENTE, LIN ")
   If aa > 0 Then
      ADO8.MoveFirst
      Do While Not ADO8.EOF
   
         Select Case ADO8!tipopariente
         Case "E"
              wNomEsp = ADO8!nombre
              wDniEsp = ADO8!numdoc
              If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                 wFecEsp = Format(ADO8!fecnac, "dd/mm/yyyy")
              End If
         Case "P"
              wNomPad = ADO8!nombre
              wDniPad = ADO8!numdoc
              If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                 wFecPad = Format(ADO8!fecnac, "dd/mm/yyyy")
              End If
         Case "M"
              wNomMad = ADO8!nombre
              wDniMad = ADO8!numdoc
              If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                 wFecMad = Format(ADO8!fecnac, "dd/mm/yyyy")
              End If
         Case "H"
              Select Case ADO8!lin
              Case "01"
                   wNomHi1 = ADO8!nombre
                   wDniHi1 = ADO8!numdoc
                   If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                      wFecHi1 = Format(ADO8!fecnac, "dd/mm/yyyy")
                   End If
              Case "02"
                   wNomHi2 = ADO8!nombre
                   wDniHi2 = ADO8!numdoc
                   If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                      wFecHi2 = Format(ADO8!fecnac, "dd/mm/yyyy")
                   End If
              Case "03"
                   wNomHi3 = ADO8!nombre
                   wDniHi3 = ADO8!numdoc
                   If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                      wFecHi3 = Format(ADO8!fecnac, "dd/mm/yyyy")
                   End If
              Case "04"
                   wNomHi4 = ADO8!nombre
                   wDniHi4 = ADO8!numdoc
                   If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                      wFecHi4 = Format(ADO8!fecnac, "dd/mm/yyyy")
                   End If
              Case "05"
                   wNomHi5 = ADO8!nombre
                   wDniHi5 = ADO8!numdoc
                   If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                      wFecHi5 = Format(ADO8!fecnac, "dd/mm/yyyy")
                   End If
              Case "06"
                   wNomHi6 = ADO8!nombre
                   wDniHi6 = ADO8!numdoc
                   If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                      wFecHi6 = Format(ADO8!fecnac, "dd/mm/yyyy")
                   End If
              Case "07"
                   wNomHi7 = ADO8!nombre
                   wDniHi7 = ADO8!numdoc
                   If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                      wFecHi7 = Format(ADO8!fecnac, "dd/mm/yyyy")
                   End If
              Case "08"
                   wNomHi8 = ADO8!nombre
                   wDniHi8 = ADO8!numdoc
                   If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                      wFecHi8 = Format(ADO8!fecnac, "dd/mm/yyyy")
                   End If
              End Select
         Case "N"
              Select Case ADO8!lin
              Case "01"
                   wNomNi1 = ADO8!nombre
                   wDniNi1 = ADO8!numdoc
                   If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                      wFecNi1 = Format(ADO8!fecnac, "dd/mm/yyyy")
                   End If
              Case "02"
                   wNomNi2 = ADO8!nombre
                   wDniNi2 = ADO8!numdoc
                   If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                      wFecNi2 = Format(ADO8!fecnac, "dd/mm/yyyy")
                   End If
              Case "03"
                   wNomNi3 = ADO8!nombre
                   wDniNi3 = ADO8!numdoc
                   If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                      wFecNi3 = Format(ADO8!fecnac, "dd/mm/yyyy")
                   End If
              Case "04"
                   wNomNi4 = ADO8!nombre
                   wDniNi4 = ADO8!numdoc
                   If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                      wFecNi4 = Format(ADO8!fecnac, "dd/mm/yyyy")
                   End If
              Case "05"
                   wNomNi5 = ADO8!nombre
                   wDniNi5 = ADO8!numdoc
                   If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                      wFecNi5 = Format(ADO8!fecnac, "dd/mm/yyyy")
                   End If
              Case "06"
                   wNomNi6 = ADO8!nombre
                   wDniNi6 = ADO8!numdoc
                   If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                      wFecNi6 = Format(ADO8!fecnac, "dd/mm/yyyy")
                   End If
              Case "07"
                   wNomNi7 = ADO8!nombre
                   wDniNi7 = ADO8!numdoc
                   If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                      wFecNi7 = Format(ADO8!fecnac, "dd/mm/yyyy")
                   End If
              Case "08"
                   wNomNi8 = ADO8!nombre
                   wDniNi8 = ADO8!numdoc
                   If IsDate(ADO8!fecnac) And Format(ADO8!fecnac, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
                      wFecNi8 = Format(ADO8!fecnac, "dd/mm/yyyy")
                   End If
              End Select
         End Select
   
         ADO8.MoveNext
      Loop
   End If
   
   txtEsposa.Text = wNomEsp
   txtPadre.Text = wNomPad
   txtMadre.Text = wNomMad
   txtHijo01.Text = wNomHi1
   txtHijo02.Text = wNomHi2
   txtHijo03.Text = wNomHi3
   txtHijo04.Text = wNomHi4
   txtHijo05.Text = wNomHi5
   txtHijo06.Text = wNomHi6
   txtHijo07.Text = wNomHi7
   txtHijo08.Text = wNomHi8
   txtNiet01.Text = wNomNi1
   txtNiet02.Text = wNomNi2
   txtNiet03.Text = wNomNi3
   txtNiet04.Text = wNomNi4
   txtNiet05.Text = wNomNi5
   txtNiet06.Text = wNomNi6
   txtNiet07.Text = wNomNi7
   txtNiet08.Text = wNomNi8
   
   txtDniEsp.Text = wDniEsp
   txtDniPad.Text = wDniPad
   txtDniMad.Text = wDniMad
   txtDniH01.Text = wDniHi1
   txtDniH02.Text = wDniHi2
   txtDniH03.Text = wDniHi3
   txtDniH04.Text = wDniHi4
   txtDniH05.Text = wDniHi5
   txtDniH06.Text = wDniHi6
   txtDniH07.Text = wDniHi7
   txtDniH08.Text = wDniHi8
   txtDniN01.Text = wDniNi1
   txtDniN02.Text = wDniNi2
   txtDniN03.Text = wDniNi3
   txtDniN04.Text = wDniNi4
   txtDniN05.Text = wDniNi5
   txtDniN06.Text = wDniNi6
   txtDniN07.Text = wDniNi7
   txtDniN08.Text = wDniNi8
   
   If IsDate(wFecEsp) And Format(wFecEsp, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnEsposa.Text = Format(wFecEsp, "dd/mm/yyyy")
   Else
      txtFnEsposa.Text = "__/__/____"
   End If
   If IsDate(wFecPad) And Format(wFecPad, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnPadre.Text = Format(wFecPad, "dd/mm/yyyy")
   Else
      txtFnPadre.Text = "__/__/____"
   End If
   If IsDate(wFecMad) And Format(wFecMad, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnMadre.Text = Format(wFecMad, "dd/mm/yyyy")
   Else
      txtFnMadre.Text = "__/__/____"
   End If
   If IsDate(wFecHi1) And Format(wFecHi1, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnHijo01.Text = Format(wFecHi1, "dd/mm/yyyy")
   Else
      txtFnHijo01.Text = "__/__/____"
   End If
   If IsDate(wFecHi2) And Format(wFecHi2, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnHijo02.Text = Format(wFecHi2, "dd/mm/yyyy")
   Else
      txtFnHijo02.Text = "__/__/____"
   End If
   If IsDate(wFecHi3) And Format(wFecHi3, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnHijo03.Text = Format(wFecHi3, "dd/mm/yyyy")
   Else
      txtFnHijo03.Text = "__/__/____"
   End If
   If IsDate(wFecHi4) And Format(wFecHi4, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnHijo04.Text = Format(wFecHi4, "dd/mm/yyyy")
   Else
      txtFnHijo04.Text = "__/__/____"
   End If
   If IsDate(wFecHi5) And Format(wFecHi5, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnHijo05.Text = Format(wFecHi5, "dd/mm/yyyy")
   Else
      txtFnHijo05.Text = "__/__/____"
   End If
   If IsDate(wFecHi6) And Format(wFecHi6, "dd/mm/yyyy") <> Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnHijo06.Text = Format(wFecHi6, "dd/mm/yyyy")
   Else
      txtFnHijo06.Text = "__/__/____"
   End If
   If IsDate(wFecHi7) And Format(wFecHi7, "dd/mm/yyyy") > Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnHijo07.Text = Format(wFecHi7, "dd/mm/yyyy")
   Else
      txtFnHijo07.Text = "__/__/____"
   End If
   If IsDate(wFecHi8) And Format(wFecHi8, "dd/mm/yyyy") > Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnHijo08.Text = Format(wFecHi8, "dd/mm/yyyy")
   Else
      txtFnHijo08.Text = "__/__/____"
   End If

   If IsDate(wFecNi1) And Format(wFecNi1, "dd/mm/yyyy") > Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnNiet01.Text = Format(wFecNi1, "dd/mm/yyyy")
   Else
      txtFnNiet01.Text = "__/__/____"
   End If
   If IsDate(wFecNi2) And Format(wFecNi2, "dd/mm/yyyy") > Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnNiet02.Text = Format(wFecNi2, "dd/mm/yyyy")
   Else
      txtFnNiet02.Text = "__/__/____"
   End If
   If IsDate(wFecNi3) And Format(wFecNi3, "dd/mm/yyyy") > Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnNiet03.Text = Format(wFecNi3, "dd/mm/yyyy")
   Else
      txtFnNiet03.Text = "__/__/____"
   End If
   If IsDate(wFecNi4) And Format(wFecNi4, "dd/mm/yyyy") > Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnNiet04.Text = Format(wFecNi4, "dd/mm/yyyy")
   Else
      txtFnNiet04.Text = "__/__/____"
   End If
   If IsDate(wFecNi5) And Format(wFecNi5, "dd/mm/yyyy") > Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnNiet05.Text = Format(wFecNi5, "dd/mm/yyyy")
   Else
      txtFnNiet05.Text = "__/__/____"
   End If
   If IsDate(wFecNi6) And Format(wFecNi6, "dd/mm/yyyy") > Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnNiet06.Text = Format(wFecNi6, "dd/mm/yyyy")
   Else
      txtFnNiet06.Text = "__/__/____"
   End If
   If IsDate(wFecNi7) And Format(wFecNi7, "dd/mm/yyyy") > Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnNiet07.Text = Format(wFecNi7, "dd/mm/yyyy")
   Else
      txtFnNiet07.Text = "__/__/____"
   End If
   If IsDate(wFecNi8) And Format(wFecNi8, "dd/mm/yyyy") > Format("30/12/1899", "dd/mm/yyyy") Then
      txtFnNiet08.Text = Format(wFecNi8, "dd/mm/yyyy")
   Else
      txtFnNiet08.Text = "__/__/____"
   End If

End Sub

Private Sub txtCodSocio_Change()
   Dim aa As Integer

   aa = Leerado6("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(zSocio) + " ")
   If aa > 0 Then
      txtCodigo.Text = ADO6!codigo
      txtIns.Text = ADO6!ins
      txtNombre.Text = ADO6!nombre
      txtGrado.Text = ADO6!grado
      txtESocio.Text = ADO6!e_socio
   End If
End Sub

Private Sub txtDniEsp_GotFocus()
   txtDniEsp.SelStart = 0
   txtDniEsp.SelLength = Len(Trim(txtDniEsp.Text))
End Sub

Private Sub txtDniEsp_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnEsposa.SetFocus
   Case 40
        txtPadre.SetFocus
   End Select
End Sub

Private Sub txtDniEsp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtPadre.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniH01_GotFocus()
   txtDniH01.SelStart = 0
   txtDniH01.SelLength = Len(Trim(txtDniH01.Text))
End Sub

Private Sub txtDniH01_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnHijo01.SetFocus
   Case 40
        txtHijo02.SetFocus
   End Select
End Sub

Private Sub txtDniH01_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtHijo02.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniH02_GotFocus()
   txtDniH02.SelStart = 0
   txtDniH02.SelLength = Len(Trim(txtDniH02.Text))
End Sub

Private Sub txtDniH02_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnHijo02.SetFocus
   Case 40
        txtHijo03.SetFocus
   End Select
End Sub

Private Sub txtDniH02_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtHijo03.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniH03_GotFocus()
   txtDniH03.SelStart = 0
   txtDniH03.SelLength = Len(Trim(txtDniH03.Text))
End Sub

Private Sub txtDniH03_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnHijo03.SetFocus
   Case 40
        txtHijo04.SetFocus
   End Select
End Sub

Private Sub txtDniH03_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtHijo04.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniH04_GotFocus()
   txtDniH04.SelStart = 0
   txtDniH04.SelLength = Len(Trim(txtDniH04.Text))
End Sub

Private Sub txtDniH04_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnHijo04.SetFocus
   Case 40
        txtHijo05.SetFocus
   End Select
End Sub

Private Sub txtDniH04_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtHijo05.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniH05_GotFocus()
   txtDniH05.SelStart = 0
   txtDniH05.SelLength = Len(Trim(txtDniH05.Text))
End Sub

Private Sub txtDniH05_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnHijo05.SetFocus
   Case 40
        txtHijo06.SetFocus
   End Select
End Sub

Private Sub txtDniH05_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtHijo06.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniH06_GotFocus()
   txtDniH06.SelStart = 0
   txtDniH06.SelLength = Len(Trim(txtDniH06.Text))
End Sub

Private Sub txtDniH06_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnHijo06.SetFocus
   Case 40
        txtHijo07.SetFocus
   End Select
End Sub

Private Sub txtDniH06_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtHijo07.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniH07_GotFocus()
   txtDniH07.SelStart = 0
   txtDniH07.SelLength = Len(Trim(txtDniH07.Text))
End Sub

Private Sub txtDniH07_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnHijo07.SetFocus
   Case 40
        txtHijo08.SetFocus
   End Select
End Sub

Private Sub txtDniH07_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtHijo08.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniH08_GotFocus()
   txtDniH08.SelStart = 0
   txtDniH08.SelLength = Len(Trim(txtDniH08.Text))
End Sub

Private Sub txtDniH08_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnHijo08.SetFocus
   Case 40
        txtNiet01.SetFocus
   End Select
End Sub

Private Sub txtDniH08_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtNiet01.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniMad_GotFocus()
   txtDniMad.SelStart = 0
   txtDniMad.SelLength = Len(Trim(txtDniMad.Text))
End Sub

Private Sub txtDniMad_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnMadre.SetFocus
   Case 40
        txtHijo01.SetFocus
   End Select
End Sub

Private Sub txtDniMad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtHijo01.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniN01_GotFocus()
   txtDniN01.SelStart = 0
   txtDniN01.SelLength = Len(Trim(txtDniN01.Text))
End Sub

Private Sub txtDniN01_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnNiet01.SetFocus
   Case 40
        txtNiet02.SetFocus
   End Select
End Sub

Private Sub txtDniN01_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtNiet02.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniN02_GotFocus()
   txtDniN02.SelStart = 0
   txtDniN02.SelLength = Len(Trim(txtDniN02.Text))
End Sub

Private Sub txtDniN02_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnNiet02.SetFocus
   Case 40
        txtNiet03.SetFocus
   End Select
End Sub

Private Sub txtDniN02_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtNiet03.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniN03_GotFocus()
   txtDniN03.SelStart = 0
   txtDniN03.SelLength = Len(Trim(txtDniN03.Text))
End Sub

Private Sub txtDniN03_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnNiet03.SetFocus
   Case 40
        txtNiet04.SetFocus
   End Select
End Sub

Private Sub txtDniN03_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtNiet04.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniN04_GotFocus()
   txtDniN04.SelStart = 0
   txtDniN04.SelLength = Len(Trim(txtDniN04.Text))
End Sub

Private Sub txtDniN04_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnNiet04.SetFocus
   Case 40
        txtNiet05.SetFocus
   End Select
End Sub

Private Sub txtDniN04_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtNiet05.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniN05_GotFocus()
   txtDniN05.SelStart = 0
   txtDniN05.SelLength = Len(Trim(txtDniN05.Text))
End Sub

Private Sub txtDniN05_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnNiet05.SetFocus
   Case 40
        txtNiet06.SetFocus
   End Select
End Sub

Private Sub txtDniN05_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtNiet06.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniN06_GotFocus()
   txtDniN06.SelStart = 0
   txtDniN06.SelLength = Len(Trim(txtDniN06.Text))
End Sub

Private Sub txtDniN06_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnNiet06.SetFocus
   Case 40
        txtNiet07.SetFocus
   End Select
End Sub

Private Sub txtDniN06_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtNiet07.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniN07_GotFocus()
   txtDniN07.SelStart = 0
   txtDniN07.SelLength = Len(Trim(txtDniN07.Text))
End Sub

Private Sub txtDniN07_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnNiet07.SetFocus
   Case 40
        txtNiet08.SetFocus
   End Select
End Sub

Private Sub txtDniN07_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtNiet08.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniN08_GotFocus()
   txtDniN08.SelStart = 0
   txtDniN08.SelLength = Len(Trim(txtDniN08.Text))
End Sub

Private Sub txtDniN08_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnNiet08.SetFocus
   End Select
End Sub

Private Sub txtDniN08_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdGrabar.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtDniPad_GotFocus()
   txtDniPad.SelStart = 0
   txtDniPad.SelLength = Len(Trim(txtDniPad.Text))
End Sub

Private Sub txtDniPad_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtFnPadre.SetFocus
   Case 40
        txtMadre.SetFocus
   End Select
End Sub

Private Sub txtDniPad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtMadre.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtESocio_Change()
   Dim aa As Integer
   aa = Leerado6("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + txtESocio.Text + "' ")
   If aa > 0 Then
      lblESocio.Caption = ADO6!nombre
   Else
      lblESocio.Caption = ""
   End If
   Set ADO8 = Nothing
End Sub

Private Sub txtEsposa_GotFocus()
   txtEsposa.SelStart = 0
   txtEsposa.SelLength = Len(Trim(txtEsposa.Text))
End Sub

Private Sub txtEsposa_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 40
        txtFnEsposa.SetFocus
   End Select
End Sub

Private Sub txtEsposa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnEsposa.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtFnEsposa_GotFocus()
   txtFnEsposa.SelStart = 0
   txtFnEsposa.SelLength = 10
End Sub

Private Sub txtFnEsposa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnEsposa.Text <> "__/__/____" Then
         If Not IsDate(txtFnEsposa.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnEsposa.Text = "__/__/____"
            txtFnEsposa.SetFocus
         End If
      End If
      txtDniEsp.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnHijo01_GotFocus()
   txtFnHijo01.SelStart = 0
   txtFnHijo01.SelLength = 10
End Sub

Private Sub txtFnHijo01_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnHijo01.Text <> "__/__/____" Then
         If Not IsDate(txtFnHijo01.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnHijo01.Text = "__/__/____"
            txtFnHijo01.SetFocus
         End If
      End If
      txtDniH01.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnHijo02_GotFocus()
   txtFnHijo02.SelStart = 0
   txtFnHijo02.SelLength = 10
End Sub

Private Sub txtFnHijo02_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnHijo02.Text <> "__/__/____" Then
         If Not IsDate(txtFnHijo02.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnHijo02.Text = "__/__/____"
            txtFnHijo02.SetFocus
         End If
      End If
      txtDniH02.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnHijo03_GotFocus()
   txtFnHijo03.SelStart = 0
   txtFnHijo03.SelLength = 10
End Sub

Private Sub txtFnHijo03_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnHijo03.Text <> "__/__/____" Then
         If Not IsDate(txtFnHijo03.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnHijo03.Text = "__/__/____"
            txtFnHijo03.SetFocus
         End If
      End If
      txtDniH03.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnHijo04_GotFocus()
   txtFnHijo04.SelStart = 0
   txtFnHijo04.SelLength = 10
End Sub

Private Sub txtFnHijo04_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnHijo04.Text <> "__/__/____" Then
         If Not IsDate(txtFnHijo04.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnHijo04.Text = "__/__/____"
            txtFnHijo04.SetFocus
         End If
      End If
      txtDniH04.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnHijo05_GotFocus()
   txtFnHijo05.SelStart = 0
   txtFnHijo05.SelLength = 10
End Sub

Private Sub txtFnHijo05_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnHijo05.Text <> "__/__/____" Then
         If Not IsDate(txtFnHijo05.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnHijo05.Text = "__/__/____"
            txtFnHijo05.SetFocus
         End If
      End If
      txtDniH05.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnHijo06_GotFocus()
   txtFnHijo06.SelStart = 0
   txtFnHijo06.SelLength = 10
End Sub

Private Sub txtFnHijo06_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnHijo06.Text <> "__/__/____" Then
         If Not IsDate(txtFnHijo06.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnHijo06.Text = "__/__/____"
            txtFnHijo06.SetFocus
         End If
      End If
      txtDniH06.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnHijo07_GotFocus()
   txtFnHijo07.SelStart = 0
   txtFnHijo07.SelLength = 10
End Sub

Private Sub txtFnHijo07_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnHijo07.Text <> "__/__/____" Then
         If Not IsDate(txtFnHijo07.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnHijo07.Text = "__/__/____"
            txtFnHijo07.SetFocus
         End If
      End If
      txtDniH07.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnHijo08_GotFocus()
   txtFnHijo08.SelStart = 0
   txtFnHijo08.SelLength = 10
End Sub

Private Sub txtFnHijo08_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnHijo08.Text <> "__/__/____" Then
         If Not IsDate(txtFnHijo08.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnHijo08.Text = "__/__/____"
            txtFnHijo08.SetFocus
         End If
      End If
      txtDniH08.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnMadre_GotFocus()
   txtFnMadre.SelStart = 0
   txtFnMadre.SelLength = 10
End Sub

Private Sub txtFnMadre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnMadre.Text <> "__/__/____" Then
         If Not IsDate(txtFnMadre.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnMadre.Text = "__/__/____"
            txtFnMadre.SetFocus
         End If
      End If
      txtDniMad.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnNiet01_GotFocus()
   txtFnNiet01.SelStart = 0
   txtFnNiet01.SelLength = 10
End Sub

Private Sub txtFnNiet01_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtNiet01.SetFocus
   Case 40
        txtDniN01.SetFocus
   End Select
End Sub

Private Sub txtFnNiet01_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnNiet01.Text <> "__/__/____" Then
         If Not IsDate(txtFnNiet01.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnNiet01.Text = "__/__/____"
            txtFnNiet01.SetFocus
         End If
      End If
      txtDniN01.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnNiet02_GotFocus()
   txtFnNiet02.SelStart = 0
   txtFnNiet02.SelLength = 10
End Sub

Private Sub txtFnNiet02_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtNiet02.SetFocus
   Case 40
        txtDniN02.SetFocus
   End Select
End Sub

Private Sub txtFnNiet02_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnNiet02.Text <> "__/__/____" Then
         If Not IsDate(txtFnNiet02.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnNiet02.Text = "__/__/____"
            txtFnNiet02.SetFocus
         End If
      End If
      txtDniN02.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnNiet03_GotFocus()
   txtFnNiet03.SelStart = 0
   txtFnNiet03.SelLength = 10
End Sub

Private Sub txtFnNiet03_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtNiet03.SetFocus
   Case 40
        txtDniN03.SetFocus
   End Select
End Sub

Private Sub txtFnNiet03_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnNiet03.Text <> "__/__/____" Then
         If Not IsDate(txtFnNiet03.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnNiet03.Text = "__/__/____"
            txtFnNiet03.SetFocus
         End If
      End If
      txtDniN03.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnNiet04_GotFocus()
   txtFnNiet04.SelStart = 0
   txtFnNiet04.SelLength = 10
End Sub

Private Sub txtFnNiet04_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtNiet04.SetFocus
   Case 40
        txtDniN04.SetFocus
   End Select
End Sub

Private Sub txtFnNiet04_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnNiet04.Text <> "__/__/____" Then
         If Not IsDate(txtFnNiet04.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnNiet04.Text = "__/__/____"
            txtFnNiet04.SetFocus
         End If
      End If
      txtDniN04.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnNiet05_GotFocus()
   txtFnNiet05.SelStart = 0
   txtFnNiet05.SelLength = 10
End Sub

Private Sub txtFnNiet05_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtNiet05.SetFocus
   Case 40
        txtDniN05.SetFocus
   End Select
End Sub

Private Sub txtFnNiet05_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnNiet05.Text <> "__/__/____" Then
         If Not IsDate(txtFnNiet05.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnNiet05.Text = "__/__/____"
            txtFnNiet05.SetFocus
         End If
      End If
      txtDniN05.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnNiet06_GotFocus()
   txtFnNiet06.SelStart = 0
   txtFnNiet06.SelLength = 10
End Sub

Private Sub txtFnNiet06_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtNiet06.SetFocus
   Case 40
        txtDniN06.SetFocus
   End Select
End Sub

Private Sub txtFnNiet06_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnNiet06.Text <> "__/__/____" Then
         If Not IsDate(txtFnNiet06.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnNiet06.Text = "__/__/____"
            txtFnNiet06.SetFocus
         End If
      End If
      txtDniN06.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnNiet07_GotFocus()
   txtFnNiet07.SelStart = 0
   txtFnNiet07.SelLength = 10
End Sub

Private Sub txtFnNiet07_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtNiet07.SetFocus
   Case 40
        txtDniN07.SetFocus
   End Select
End Sub

Private Sub txtFnNiet07_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnNiet07.Text <> "__/__/____" Then
         If Not IsDate(txtFnNiet07.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnNiet07.Text = "__/__/____"
            txtFnNiet07.SetFocus
         End If
      End If
      txtDniN07.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnNiet08_GotFocus()
   txtFnNiet08.SelStart = 0
   txtFnNiet08.SelLength = 10
End Sub

Private Sub txtFnNiet08_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtNiet08.SetFocus
   Case 40
        txtDniN08.SetFocus
   End Select
End Sub

Private Sub txtFnNiet08_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnNiet08.Text <> "__/__/____" Then
         If Not IsDate(txtFnNiet08.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnNiet08.Text = "__/__/____"
            txtFnNiet08.SetFocus
         End If
      End If
      txtDniN08.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtFnPadre_GotFocus()
   txtFnPadre.SelStart = 0
   txtFnPadre.SelLength = 10
End Sub

Private Sub txtFnPadre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtFnPadre.Text <> "__/__/____" Then
         If Not IsDate(txtFnPadre.Text) Then
            MsgBox "Fecha de Nacimiento Digitada Es Invalida", vbExclamation
            txtFnPadre.Text = "__/__/____"
            txtFnPadre.SetFocus
         End If
      End If
      txtDniPad.SetFocus
   Else
      If InStr(1, "0123456789" + Chr(8), Chr(KeyAscii)) = 0 Then
         KeyAscii = 0
      End If
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

Private Sub txtHijo01_GotFocus()
   txtHijo01.SelStart = 0
   txtHijo01.SelLength = Len(Trim(txtHijo01.Text))
End Sub

Private Sub txtHijo01_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniMad.SetFocus
   Case 40
        txtFnHijo01.SetFocus
   End Select
End Sub

Private Sub txtHijo01_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnHijo01.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtHijo02_GotFocus()
   txtHijo02.SelStart = 0
   txtHijo02.SelLength = Len(Trim(txtHijo02.Text))
End Sub

Private Sub txtHijo02_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniH01.SetFocus
   Case 40
        txtFnHijo02.SetFocus
   End Select
End Sub

Private Sub txtHijo02_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnHijo02.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtHijo03_GotFocus()
   txtHijo03.SelStart = 0
   txtHijo03.SelLength = Len(Trim(txtHijo03.Text))
End Sub

Private Sub txtHijo03_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniH02.SetFocus
   Case 40
        txtFnHijo03.SetFocus
   End Select
End Sub

Private Sub txtHijo03_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnHijo03.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtHijo04_GotFocus()
   txtHijo04.SelStart = 0
   txtHijo04.SelLength = Len(Trim(txtHijo04.Text))
End Sub

Private Sub txtHijo04_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniH03.SetFocus
   Case 40
        txtFnHijo04.SetFocus
   End Select
End Sub

Private Sub txtHijo04_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnHijo04.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtHijo05_GotFocus()
   txtHijo05.SelStart = 0
   txtHijo05.SelLength = Len(Trim(txtHijo05.Text))
End Sub

Private Sub txtHijo05_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniH04.SetFocus
   Case 40
        txtFnHijo05.SetFocus
   End Select
End Sub

Private Sub txtHijo05_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnHijo05.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtHijo06_GotFocus()
   txtHijo06.SelStart = 0
   txtHijo06.SelLength = Len(Trim(txtHijo06.Text))
End Sub

Private Sub txtHijo06_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniH05.SetFocus
   Case 40
        txtFnHijo06.SetFocus
   End Select
End Sub

Private Sub txtHijo06_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnHijo06.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtHijo07_GotFocus()
   txtHijo07.SelStart = 0
   txtHijo07.SelLength = Len(Trim(txtHijo07.Text))
End Sub

Private Sub txtHijo07_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniH06.SetFocus
   Case 40
        txtFnHijo07.SetFocus
   End Select
End Sub

Private Sub txtHijo07_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnHijo07.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtHijo08_GotFocus()
   txtHijo08.SelStart = 0
   txtHijo08.SelLength = Len(Trim(txtHijo08.Text))
End Sub

Private Sub txtHijo08_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniH07.SetFocus
   Case 40
        txtFnHijo08.SetFocus
   End Select
End Sub

Private Sub txtHijo08_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnHijo08.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtMadre_GotFocus()
   txtMadre.SelStart = 0
   txtMadre.SelLength = Len(Trim(txtMadre.Text))
End Sub

Private Sub txtMadre_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniPad.SetFocus
   Case 40
        txtFnMadre.SetFocus
   End Select
End Sub

Private Sub txtMadre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnMadre.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtNiet01_GotFocus()
   txtNiet01.SelStart = 0
   txtNiet01.SelLength = Len(Trim(txtNiet01.Text))
End Sub

Private Sub txtNiet01_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniH08.SetFocus
   Case 40
        txtFnNiet01.SetFocus
   End Select
End Sub

Private Sub txtNiet01_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnNiet01.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtNiet02_GotFocus()
   txtNiet02.SelStart = 0
   txtNiet02.SelLength = Len(Trim(txtNiet02.Text))
End Sub

Private Sub txtNiet02_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniN01.SetFocus
   Case 40
        txtFnNiet02.SetFocus
   End Select
End Sub

Private Sub txtNiet02_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnNiet02.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtNiet03_GotFocus()
   txtNiet03.SelStart = 0
   txtNiet03.SelLength = Len(Trim(txtNiet03.Text))
End Sub

Private Sub txtNiet03_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniN02.SetFocus
   Case 40
        txtFnNiet03.SetFocus
   End Select
End Sub

Private Sub txtNiet03_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnNiet03.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtNiet04_GotFocus()
   txtNiet04.SelStart = 0
   txtNiet04.SelLength = Len(Trim(txtNiet04.Text))
End Sub

Private Sub txtNiet04_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniN03.SetFocus
   Case 40
        txtFnNiet04.SetFocus
   End Select
End Sub

Private Sub txtNiet04_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnNiet04.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtNiet05_GotFocus()
   txtNiet05.SelStart = 0
   txtNiet05.SelLength = Len(Trim(txtNiet05.Text))
End Sub

Private Sub txtNiet05_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniN04.SetFocus
   Case 40
        txtFnNiet05.SetFocus
   End Select
End Sub

Private Sub txtNiet05_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnNiet05.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtNiet06_GotFocus()
   txtNiet06.SelStart = 0
   txtNiet06.SelLength = Len(Trim(txtNiet06.Text))
End Sub

Private Sub txtNiet06_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniN05.SetFocus
   Case 40
        txtFnNiet06.SetFocus
   End Select
End Sub

Private Sub txtNiet06_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnNiet06.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtNiet07_GotFocus()
   txtNiet07.SelStart = 0
   txtNiet07.SelLength = Len(Trim(txtNiet07.Text))
End Sub

Private Sub txtNiet07_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniN06.SetFocus
   Case 40
        txtFnNiet07.SetFocus
   End Select
End Sub

Private Sub txtNiet07_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnNiet07.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtNiet08_GotFocus()
   txtNiet08.SelStart = 0
   txtNiet08.SelLength = Len(Trim(txtNiet08.Text))
End Sub

Private Sub txtNiet08_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniN07.SetFocus
   Case 40
        txtFnNiet08.SetFocus
   End Select
End Sub

Private Sub txtNiet08_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnNiet08.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txtPadre_GotFocus()
   txtPadre.SelStart = 0
   txtPadre.SelLength = Len(Trim(txtPadre.Text))
End Sub

Private Sub txtPadre_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 38
        txtDniEsp.SetFocus
   Case 40
        txtFnPadre.SetFocus
   End Select
End Sub

Private Sub txtPadre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFnPadre.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub


