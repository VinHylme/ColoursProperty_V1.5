VERSION 5.00
Begin VB.Form View_Agent 
   Caption         =   "Form1"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   14310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Search All"
      Height          =   495
      Left            =   9720
      TabIndex        =   26
      Top             =   720
      Width           =   1335
   End
   Begin VB.ListBox AgentList_FName 
      Height          =   3570
      Left            =   360
      TabIndex        =   13
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox AgentRefID 
      Height          =   495
      Left            =   3840
      TabIndex        =   12
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   495
      Left            =   8520
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.ListBox AgentList_LName 
      Height          =   3570
      Left            =   2640
      TabIndex        =   10
      Top             =   2040
      Width           =   2175
   End
   Begin VB.ListBox AgentList_CName 
      Height          =   3570
      Left            =   4920
      TabIndex        =   9
      Top             =   2040
      Width           =   2175
   End
   Begin VB.ListBox AgentList_Address2 
      Height          =   3570
      Left            =   2640
      TabIndex        =   8
      Top             =   6240
      Width           =   2175
   End
   Begin VB.ListBox AgentList_PostCode 
      Height          =   3570
      Left            =   4920
      TabIndex        =   7
      Top             =   6240
      Width           =   2175
   End
   Begin VB.ListBox AgentList_Country 
      Height          =   3570
      Left            =   7200
      TabIndex        =   6
      Top             =   6240
      Width           =   2175
   End
   Begin VB.ListBox AgentList_State 
      Height          =   3570
      Left            =   9480
      TabIndex        =   5
      Top             =   6240
      Width           =   2175
   End
   Begin VB.ListBox AgentList_City 
      Height          =   3570
      Left            =   11760
      TabIndex        =   4
      Top             =   6240
      Width           =   2175
   End
   Begin VB.ListBox AgentList_FaxNo 
      Height          =   3570
      Left            =   7200
      TabIndex        =   3
      Top             =   2040
      Width           =   2175
   End
   Begin VB.ListBox AgentList_Address1 
      Height          =   3570
      Left            =   360
      TabIndex        =   2
      Top             =   6240
      Width           =   2175
   End
   Begin VB.ListBox AgentList_PhoneNo 
      Height          =   3570
      Left            =   9480
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin VB.ListBox AgentList_EmailAddress 
      Height          =   3570
      Left            =   11760
      TabIndex        =   0
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label labelFName 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   25
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label LabelLName 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   24
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label LabelCName 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   23
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label LabelFaxNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax Number:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   22
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label LabelPhoneNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   21
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label LabelEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11760
      TabIndex        =   20
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label LabelAddress1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 1:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label LabelAddress2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 2:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   18
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label LabelPostCode 
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label LabelCountry 
      BackStyle       =   0  'Transparent
      Caption         =   "Country:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label LabelState 
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   15
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label LabelCity 
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11760
      TabIndex        =   14
      Top             =   5880
      Width           =   2175
   End
End
Attribute VB_Name = "View_Agent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
ClearListBoxesForAgent
Dim ViewAgent As Agent_Record
Dim ViewAgentChannel As Integer
Dim x As Integer
Dim OnceFoundAgent As Boolean
x = 1
ViewAgentChannel = FreeFile
Open Agent_File For Random As ViewAgentChannel Len = Agent_Length
Get ViewAgentChannel, x, ViewAgent
Do While Not EOF(ViewAgentChannel) And OnceFoundAgent = False
    If Trim(ViewAgent.Agent_RefID) = Trim(AgentRefID.Text) Then
    AgentList_FName.AddItem ViewAgent.Agent_FName
    AgentList_LName.AddItem ViewAgent.Agent_LName
    AgentList_CName.AddItem ViewAgent.Agent_CName
    AgentList_Address1.AddItem ViewAgent.Agent_Address1
    AgentList_Address2.AddItem ViewAgent.Agent_Address2
    AgentList_PostCode.AddItem ViewAgent.Agent_PostCode
    AgentList_Country.AddItem ViewAgent.Agent_CountryList
    AgentList_State.AddItem ViewAgent.Agent_State
    AgentList_City.AddItem ViewAgent.Agent_City
    AgentList_FaxNo.AddItem ViewAgent.Agent_FaxNumber
    AgentList_PhoneNo.AddItem ViewAgent.Agent_PhoneNumber
    AgentList_EmailAddress.AddItem ViewAgent.Agent_EmailAddress
    Agent_File_Pointer = x
    OnceFoundAgent = True
    Else
    MsgBox "This Landlord does not exist"
    x = x + 1
    Get ViewAgentChannel, x, ViewAgent
    End If
Loop
Close ViewAgentChannel
End Sub

Private Sub ClearListBoxesForAgent()
AgentList_FName.Clear
AgentList_LName.Clear
AgentList_CName.Clear
AgentList_Address1.Clear
AgentList_Address2.Clear
AgentList_PostCode.Clear
AgentList_Country.Clear
AgentList_State.Clear
AgentList_City.Clear
AgentList_FaxNo.Clear
AgentList_PhoneNo.Clear
AgentList_EmailAddress.Clear
End Sub

Private Sub Command2_Click()
ClearListBoxesForAgent
Dim ViewAgent As Agent_Record
Dim ViewAgentChannel As Integer
Dim x As Integer
x = 1
ViewAgentChannel = FreeFile
Open Agent_File For Random As ViewAgentChannel Len = Agent_Length
Get ViewAgentChannel, x, ViewAgent
Do While Not EOF(ViewAgentChannel)
    AgentList_FName.AddItem ViewAgent.Agent_FName
    AgentList_LName.AddItem ViewAgent.Agent_LName
    AgentList_CName.AddItem ViewAgent.Agent_CName
    AgentList_Address1.AddItem ViewAgent.Agent_Address1
    AgentList_Address2.AddItem ViewAgent.Agent_Address2
    AgentList_PostCode.AddItem ViewAgent.Agent_PostCode
    AgentList_Country.AddItem ViewAgent.Agent_CountryList
    AgentList_State.AddItem ViewAgent.Agent_State
    AgentList_City.AddItem ViewAgent.Agent_City
    AgentList_FaxNo.AddItem ViewAgent.Agent_FaxNumber
    AgentList_PhoneNo.AddItem ViewAgent.Agent_PhoneNumber
    AgentList_EmailAddress.AddItem ViewAgent.Agent_EmailAddress
    x = x + 1
    Get ViewAgentChannel, x, ViewAgent
Loop
Close ViewAgentChannel
End Sub

