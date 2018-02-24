VERSION 5.00
Begin VB.Form Add_Agent 
   Caption         =   "Add Agent"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Agent_FName 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   14
      Top             =   480
      Width           =   4695
   End
   Begin VB.TextBox Agent_LName 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   13
      Top             =   1080
      Width           =   4695
   End
   Begin VB.TextBox Agent_CName 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   1680
      Width           =   4695
   End
   Begin VB.TextBox Agent_Address2 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   2280
      Width           =   4695
   End
   Begin VB.TextBox Agent_Address1 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   2880
      Width           =   4695
   End
   Begin VB.TextBox Agent_PostCode 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   3480
      Width           =   4695
   End
   Begin VB.TextBox Agent_State 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   4560
      Width           =   4695
   End
   Begin VB.TextBox Agent_City 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   5160
      Width           =   4695
   End
   Begin VB.TextBox Agent_PhoneNo 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   6360
      Width           =   4695
   End
   Begin VB.TextBox Agent_EmailAddress 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   6960
      Width           =   4695
   End
   Begin VB.ComboBox Country_list_property 
      BackColor       =   &H80000004&
      Height          =   315
      ItemData        =   "Add_Agent.frx":0000
      Left            =   2400
      List            =   "Add_Agent.frx":0250
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox Agent_FaxNo 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   5760
      Width           =   4695
   End
   Begin VB.CommandButton Register_Agent 
      Caption         =   "Register Agent"
      Height          =   855
      Left            =   7920
      TabIndex        =   2
      Top             =   8160
      Width           =   1815
   End
   Begin VB.TextBox Agent_RefID 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   7680
      Width           =   1455
   End
   Begin VB.TextBox Agent_Country 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   4680
      TabIndex        =   0
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   600
      TabIndex        =   28
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   600
      TabIndex        =   27
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   600
      TabIndex        =   26
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   600
      TabIndex        =   25
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 1:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   600
      TabIndex        =   24
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 2:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   600
      TabIndex        =   23
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   600
      TabIndex        =   22
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   600
      TabIndex        =   21
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   600
      TabIndex        =   20
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   600
      TabIndex        =   19
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   600
      TabIndex        =   18
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Country:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   600
      TabIndex        =   17
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   600
      TabIndex        =   16
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Ref Code:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   600
      TabIndex        =   15
      Top             =   7800
      Width           =   2175
   End
End
Attribute VB_Name = "Add_Agent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Country_list_property_Click()
Agent_Country.Text = Country_list_property
End Sub
Private Sub GenerateRandomAgentRefID()
Agent_REFID.Text = Int(Rnd * 997) + 1
End Sub
Private Sub CheckAgentRef()
GenerateRandomAgentRefID
Dim ViewAgentRefID As Agent_Record
Dim ViewAgentRefIDChannel As Integer
Dim x As Integer
x = 1
ViewAgentRefIDChannel = FreeFile
Open Agent_File For Random As ViewAgentRefIDChannel Len = Agent_Length
Get ViewAgentRefIDChannel, x, ViewAgentRefID
    Do While Not EOF(ViewAgentRefIDChannel)
        If Trim(ViewAgentRefID.Agent_REFID) = Trim(Agent_REFID.Text) Then
        GenerateRandomAgentRefID
        Agent_File_Pointer = x
        End If
        x = x + 1
        Get ViewAgentRefIDChannel, x, ViewAgentRefID
    Loop
Close ViewAgentRefIDChannel
End Sub

Private Sub Register_Agent_Click()
Dim Add_Agent As Agent_Record
Dim Agentchannel As Integer
CheckAgentRef
    Agentchannel = FreeFile
    Open Agent_File For Random As Agentchannel Len = Agent_Length
            Add_Agent.Agent_Fname = Agent_Fname.Text
            Add_Agent.Agent_LName = Agent_LName.Text
            Add_Agent.Agent_CName = Agent_CName.Text
            Add_Agent.Agent_Address1 = Agent_Address1.Text
            Add_Agent.Agent_Address2 = Agent_Address2.Text
            Add_Agent.Agent_PostCode = Agent_PostCode.Text
            Add_Agent.Agent_CountryList = Agent_Country.Text
            Add_Agent.Agent_State = Agent_State.Text
            Add_Agent.Agent_City = Agent_City.Text
            Add_Agent.Agent_FaxNumber = Agent_FaxNo.Text
            Add_Agent.Agent_PhoneNumber = Agent_PhoneNo.Text
            Add_Agent.Agent_EmailAddress = Agent_EmailAddress.Text
            Add_Agent.Agent_REFID = Agent_REFID.Text
            MsgBox Agent_REFID.Text
            Agent_File_Pointer = Agent_File_Pointer + 1
    Put Agentchannel, Agent_File_Pointer, Add_Agent
    Close Agentchannel
End Sub
