VERSION 5.00
Begin VB.Form Agents 
   Caption         =   "Form1"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   13680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Agent_Country 
      Height          =   285
      Left            =   3840
      TabIndex        =   33
      Top             =   4680
      Width           =   2295
   End
   Begin VB.ListBox Agent_REFID 
      Height          =   2790
      Left            =   240
      TabIndex        =   32
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Agent_FirstName 
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox Agent_LName 
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox Agent_CName 
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox Agent_AddressLine1 
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox Agent_AddressLine2 
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox Agent_PostCode 
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox Agent_Countries 
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox Agent_State 
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox Agent_City 
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox Agent_FaxNumber 
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox Agent_PhoneNumber 
      Height          =   375
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox Agent_EmailAddresses 
      Height          =   375
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox Agent_Fname 
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   4080
      Width           =   3255
   End
   Begin VB.TextBox Agent_LName2 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4680
      Width           =   3255
   End
   Begin VB.TextBox Agent_Cname2 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   5280
      Width           =   3255
   End
   Begin VB.TextBox Agent_AddressLine1_2 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   5880
      Width           =   3255
   End
   Begin VB.TextBox Agent_AddressLine2_2 
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   6480
      Width           =   3255
   End
   Begin VB.TextBox Agent_PostCode2 
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   4080
      Width           =   3255
   End
   Begin VB.TextBox Agent_State2 
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   5280
      Width           =   3255
   End
   Begin VB.TextBox Agent_City2 
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   5880
      Width           =   3255
   End
   Begin VB.TextBox Agent_FaxNo 
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Top             =   6480
      Width           =   3255
   End
   Begin VB.TextBox Agent_PhoneNo 
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   4080
      Width           =   3255
   End
   Begin VB.TextBox AgentRefID 
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Add 
      Caption         =   "ADD"
      Height          =   615
      Left            =   10920
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Delete 
      Caption         =   "DELETE"
      Height          =   615
      Left            =   12360
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Prints 
      Caption         =   "PRINT"
      Height          =   615
      Left            =   10920
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "CANCEL"
      Enabled         =   0   'False
      Height          =   615
      Left            =   12360
      TabIndex        =   5
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Close 
      Caption         =   "CLOSE"
      Height          =   615
      Left            =   10920
      TabIndex        =   4
      Top             =   6360
      Width           =   2655
   End
   Begin VB.CommandButton EdIT 
      Caption         =   "EDIT"
      Height          =   615
      Left            =   12360
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Save 
      Caption         =   "SAVE"
      Enabled         =   0   'False
      Height          =   615
      Left            =   10920
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.ComboBox Country_list_property 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "Agents.frx":0000
      Left            =   3840
      List            =   "Agents.frx":0250
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4680
      Width           =   3255
   End
   Begin VB.TextBox Agent_EmailAddress 
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Label Label27 
      Caption         =   "Label27"
      Height          =   255
      Left            =   7320
      TabIndex        =   60
      Top             =   7080
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent Refrence ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   240
      TabIndex        =   59
      Top             =   120
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   0
      X2              =   13680
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   3120
      TabIndex        =   58
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   3120
      TabIndex        =   57
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   3120
      TabIndex        =   56
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 1:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   3120
      TabIndex        =   55
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 2:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   3120
      TabIndex        =   54
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   6840
      TabIndex        =   53
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Country:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   6840
      TabIndex        =   52
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   6840
      TabIndex        =   51
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   6840
      TabIndex        =   50
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   6840
      TabIndex        =   49
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   10320
      TabIndex        =   48
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   10320
      TabIndex        =   47
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   120
      TabIndex        =   46
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   120
      TabIndex        =   45
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   120
      TabIndex        =   44
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 1:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   120
      TabIndex        =   43
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 2:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   120
      TabIndex        =   42
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   3840
      TabIndex        =   41
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Country:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   3840
      TabIndex        =   40
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   3840
      TabIndex        =   39
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   3840
      TabIndex        =   38
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   3840
      TabIndex        =   37
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   7320
      TabIndex        =   36
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent Refrence ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   7320
      TabIndex        =   35
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Line Line2 
      X1              =   10680
      X2              =   10680
      Y1              =   3600
      Y2              =   7320
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   7320
      TabIndex        =   34
      Top             =   4440
      Width           =   2535
   End
End
Attribute VB_Name = "Agents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Add_Click()
Save_Cancel
Label27.Caption = "ADD"
End Sub

Private Sub Cancel_Click()
CancelFunction
End Sub

Private Sub Close_Click()
Unload Me
End Sub
Private Sub EdIT_Click()
Save_Cancel
Label27.Caption = "EDIT"
End Sub

Private Function Save_Cancel()
Add.Enabled = False
EdIT.Enabled = False
Prints.Enabled = False
Delete.Enabled = False
Save.Enabled = True
Cancel.Enabled = True
Agent_Fname.SetFocus

End Function

Private Function CancelFunction()
Add.Enabled = True
Add.SetFocus
EdIT.Enabled = True
Prints.Enabled = True
Delete.Enabled = True
Save.Enabled = False
Cancel.Enabled = False
End Function

Private Sub Country_list_property_Click()
Agent_Country.Text = Country_list_property
End Sub
Private Sub GenerateRandomAgentRefID()
AgentRefID.Text = Int(Rnd * 997) + 1
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
        If Trim(ViewAgentRefID.Agent_REFID) = Trim(AgentRefID.Text) Then
        GenerateRandomAgentRefID
        Agent_File_Pointer = x
        End If
        x = x + 1
        Get ViewAgentRefIDChannel, x, ViewAgentRefID
    Loop
Close ViewAgentRefIDChannel
End Sub

Private Sub Form_Load()
Dim ViewAgentsID As Agent_Record
Dim ViewAgentsIDChannel As Integer
Dim x As Integer
x = 1
ViewAgentsIDChannel = FreeFile
Open Agent_File For Random As ViewAgentsIDChannel Len = Agent_Length
Get ViewAgentsIDChannel, x, ViewAgentsID
Do While Not EOF(ViewAgentsIDChannel)
    Agent_REFID.AddItem ViewAgentsID.Agent_REFID
    x = x + 1
    Get ViewAgentsIDChannel, x, ViewAgentsID
Loop
Close ViewAgentsIDChannel
End Sub

Private Sub save_Click()
If Label27 = "ADD" Then
Dim Add_Agent As Agent_Record
Dim Agentchannel As Integer
CheckAgentRef
    Agentchannel = FreeFile
    Open Agent_File For Random As Agentchannel Len = Agent_Length
            Add_Agent.Agent_Fname = Agent_Fname.Text
            Add_Agent.Agent_LName = Agent_LName2.Text
            Add_Agent.Agent_CName = Agent_Cname2.Text
            Add_Agent.Agent_Address1 = Agent_AddressLine1_2.Text
            Add_Agent.Agent_Address2 = Agent_AddressLine2_2.Text
            Add_Agent.Agent_PostCode = Agent_PostCode2.Text
            Add_Agent.Agent_CountryList = Agent_Country.Text
            Add_Agent.Agent_State = Agent_State2.Text
            Add_Agent.Agent_City = Agent_City2.Text
            Add_Agent.Agent_FaxNumber = Agent_FaxNo.Text
            Add_Agent.Agent_PhoneNumber = Agent_PhoneNo.Text
            Add_Agent.Agent_EmailAddress = Agent_EmailAddress.Text
            Add_Agent.Agent_REFID = Agent_REFID.Text
            MsgBox AgentRefID.Text
            Agent_File_Pointer = Agent_File_Pointer + 1
    Put Agentchannel, Agent_File_Pointer, Add_Agent
    Close Agentchannel
Else
MsgBox "Something went wrong"
End If
End Sub


