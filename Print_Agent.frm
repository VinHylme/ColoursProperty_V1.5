VERSION 5.00
Begin VB.Form Print_Agent 
   Caption         =   "Print Agent"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line17 
      BorderColor     =   &H00C0C0C0&
      X1              =   4560
      X2              =   11520
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS LINE 2"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS LINE 1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4560
      TabIndex        =   15
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "LAST NAME"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "FIRST NAME"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00C0C0C0&
      X1              =   4560
      X2              =   11520
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00C0C0C0&
      X1              =   4560
      X2              =   11520
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00C0C0C0&
      X1              =   4560
      X2              =   11520
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000003&
      BorderWidth     =   3
      X1              =   0
      X2              =   12360
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00C0C0C0&
      X1              =   120
      X2              =   4200
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "GREENVWAY"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   4440
      Width           =   3255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "LANDLORD COMPANY NAME"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "EMAIL ADDRESS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "PHONE NUMBER"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "CITY"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "COUNTRY"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "POST CODE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS LINE 2"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS LINE 1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LAST NAME"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "FIRST NAME"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000003&
      BorderWidth     =   3
      X1              =   -480
      X2              =   11880
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00C0C0C0&
      X1              =   4560
      X2              =   11520
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00C0C0C0&
      X1              =   4560
      X2              =   11520
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00C0C0C0&
      X1              =   4560
      X2              =   11520
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00C0C0C0&
      X1              =   4560
      X2              =   11520
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C0C0C0&
      X1              =   4560
      X2              =   11520
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C0C0C0&
      X1              =   4560
      X2              =   11520
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      X1              =   4560
      X2              =   11520
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      X1              =   4560
      X2              =   11520
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   4560
      X2              =   11520
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BLUEWORLD"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   120
      X2              =   4200
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Namelebel 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPANY NAME"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "Print_Agent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
