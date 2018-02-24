VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Left            =   8520
      Top             =   2640
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "Brush Script Std"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   7575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
     Timer1.Interval = 10
End Sub

Private Sub Timer1_Timer()
Label2.Caption = DateDiff("d", Now, DateSerial(2013, 11, 3)) & "Days " & CDate(0 - CDbl(Time))
End Sub

