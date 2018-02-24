VERSION 5.00
Begin VB.Form Dashboard 
   Caption         =   "Welcome To The Admin Dashboard"
   ClientHeight    =   9990
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17085
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   9990
   ScaleWidth      =   17085
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   480
      Picture         =   "Form3.frx":14282
      ScaleHeight     =   975
      ScaleWidth      =   3375
      TabIndex        =   12
      Top             =   3120
      Width           =   3375
   End
   Begin VB.PictureBox Notifications_button 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   480
      Picture         =   "Form3.frx":1BD66
      ScaleHeight     =   495
      ScaleWidth      =   3375
      TabIndex        =   11
      Top             =   2160
      Width           =   3375
   End
   Begin VB.PictureBox last_logged_bg 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      Picture         =   "Form3.frx":22379
      ScaleHeight     =   375
      ScaleWidth      =   4095
      TabIndex        =   8
      Top             =   1560
      Width           =   4095
      Begin VB.Label number_visit 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Info_vist 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Visted:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   0
      Picture         =   "Form3.frx":26CE2
      ScaleHeight     =   5775
      ScaleWidth      =   17415
      TabIndex        =   3
      Top             =   4440
      Width           =   17415
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      Picture         =   "Form3.frx":3193C
      ScaleHeight     =   601.754
      ScaleMode       =   0  'User
      ScaleWidth      =   17295
      TabIndex        =   0
      Top             =   0
      Width           =   17295
   End
   Begin VB.PictureBox Add_property 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   4440
      Picture         =   "Form3.frx":3A7E7
      ScaleHeight     =   4215
      ScaleWidth      =   3135
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.PictureBox Manage_agents 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   10680
      Picture         =   "Form3.frx":448F2
      ScaleHeight     =   4215
      ScaleWidth      =   3135
      TabIndex        =   2
      Top             =   480
      Width           =   3135
   End
   Begin VB.PictureBox Manage_landlords 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   7560
      Picture         =   "Form3.frx":4FFDD
      ScaleHeight     =   4215
      ScaleWidth      =   3135
      TabIndex        =   4
      Top             =   480
      Width           =   3135
   End
   Begin VB.PictureBox Other_tools 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   13800
      Picture         =   "Form3.frx":5BBC3
      ScaleHeight     =   4335
      ScaleWidth      =   3375
      TabIndex        =   5
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label email_address 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   5640
      Width           =   4095
   End
   Begin VB.Label your_name 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label welcome_admin 
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "Dashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Add_property_Click()
Unload Me
Add_properties.Show 1
End Sub

Private Sub Form_Activate()
email_address.Caption = Selected_EmailAddress
End Sub

Private Sub Form_Load()
'Dim ViewName As Register_Account_Admin
'Dim ViewNamechannel As Integer
'Dim y As Integer
'y = 1
'ViewNamechannel = FreeFile
'Open Register_account_admin_file For Random As ViewNamechannel Len = Register_account_admin_length
'Get ViewNamechannel, y, ViewName
'Do While Not EOF(ViewNamechannel)
'    If ViewName.email_address = Selected_EmailAddress Then
'        your_name.Caption = ViewName.First_name
'    End If
'    y = y + 1
'    Get ViewNamechannel, y, ViewName
'Loop
'Close ViewNamechannel
End Sub

Private Sub email_address_Change()
Selected_EmailAddress = email_address.Caption
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Manage_agents_Click()
Unload Me
Agents.Show 1
End Sub

Private Sub Manage_landlords_Click()
Unload Me
Landlords.Show 1
End Sub

Private Sub Other_tools_Click()
Unload Me
Tools_use.Show 1
End Sub

