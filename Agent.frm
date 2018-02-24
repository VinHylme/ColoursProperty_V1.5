VERSION 5.00
Begin VB.Form Agent 
   Caption         =   "Manage Agents"
   ClientHeight    =   6945
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19170
   LinkTopic       =   "Form1"
   ScaleHeight     =   7546.9
   ScaleMode       =   0  'User
   ScaleWidth      =   19170
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox Agent_REFID 
      Height          =   2790
      Left            =   120
      TabIndex        =   30
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Agents_Fname 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3960
      Width           =   3255
   End
   Begin VB.TextBox Agents_LName2 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   4560
      Width           =   3255
   End
   Begin VB.TextBox Agents_Cname2 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   5160
      Width           =   3255
   End
   Begin VB.TextBox Agents_AddressLine1_2 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   5760
      Width           =   3255
   End
   Begin VB.TextBox Agents_AddressLine2_2 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   6360
      Width           =   3255
   End
   Begin VB.TextBox Agents_PostCode2 
      Height          =   375
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3960
      Width           =   3255
   End
   Begin VB.TextBox Agents_State2 
      Height          =   375
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   5160
      Width           =   3255
   End
   Begin VB.TextBox Agents_City2 
      Height          =   375
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   5760
      Width           =   3255
   End
   Begin VB.TextBox Agents_PhoneNo 
      Height          =   375
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   6360
      Width           =   3255
   End
   Begin VB.TextBox Agents_RefID 
      Height          =   375
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   22
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Add 
      Caption         =   "ADD"
      Height          =   615
      Left            =   12240
      TabIndex        =   29
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton Delete 
      Caption         =   "DELETE"
      Height          =   615
      Left            =   14880
      TabIndex        =   28
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton Prints 
      Caption         =   "PRINT"
      Height          =   615
      Left            =   12240
      TabIndex        =   27
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "CANCEL"
      Enabled         =   0   'False
      Height          =   615
      Left            =   14880
      TabIndex        =   26
      Top             =   5280
      Width           =   2775
   End
   Begin VB.CommandButton Close 
      Caption         =   "CLOSE"
      Height          =   615
      Left            =   12240
      TabIndex        =   25
      Top             =   6000
      Width           =   5415
   End
   Begin VB.CommandButton EdIT 
      Caption         =   "EDIT"
      Height          =   615
      Left            =   14880
      TabIndex        =   24
      Top             =   3840
      Width           =   2775
   End
   Begin VB.CommandButton Save 
      Caption         =   "SAVE"
      Enabled         =   0   'False
      Height          =   615
      Left            =   12240
      TabIndex        =   23
      Top             =   5280
      Width           =   2535
   End
   Begin VB.ComboBox Country_list_property 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "Agent.frx":0000
      Left            =   3840
      List            =   "Agent.frx":0250
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4560
      Width           =   3255
   End
   Begin VB.TextBox Agents_EmailAddress 
      Height          =   375
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3960
      Width           =   3255
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   2280
      TabIndex        =   10
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   3840
      TabIndex        =   9
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox List3 
      Height          =   2790
      Left            =   5400
      TabIndex        =   8
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox List4 
      Height          =   2790
      Left            =   6960
      TabIndex        =   7
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox List5 
      Height          =   2790
      Left            =   8520
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox List6 
      Height          =   2790
      Left            =   10080
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.ListBox List7 
      Height          =   2790
      Left            =   11280
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox List8 
      Height          =   2790
      Left            =   12840
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox List9 
      Height          =   2790
      Left            =   15960
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox List10 
      Height          =   2790
      Left            =   17520
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox List11 
      Height          =   2790
      Left            =   14400
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Agents_Country 
      Height          =   285
      Left            =   3840
      TabIndex        =   31
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label23 
      Height          =   495
      Left            =   7080
      TabIndex        =   57
      Top             =   9000
      Width           =   2895
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
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   120
      TabIndex        =   56
      Top             =   120
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   0
      X2              =   19200
      Y1              =   3781.6
      Y2              =   3781.6
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
      TabIndex        =   55
      Top             =   3720
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
      Height          =   615
      Left            =   120
      TabIndex        =   54
      Top             =   4320
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
      TabIndex        =   53
      Top             =   4920
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
      TabIndex        =   52
      Top             =   5520
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
      TabIndex        =   51
      Top             =   6120
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
      TabIndex        =   50
      Top             =   3720
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
      TabIndex        =   49
      Top             =   4320
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
      TabIndex        =   48
      Top             =   4920
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
      TabIndex        =   47
      Top             =   5520
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
      Left            =   3840
      TabIndex        =   46
      Top             =   6120
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
      TabIndex        =   45
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Line Line2 
      X1              =   10680
      X2              =   10680
      Y1              =   3781.6
      Y2              =   7824
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
      TabIndex        =   44
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
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
      Height          =   495
      Left            =   2520
      TabIndex        =   43
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
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
      Height          =   495
      Left            =   4200
      TabIndex        =   42
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
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
      Height          =   495
      Left            =   5520
      TabIndex        =   41
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address"
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
      Height          =   495
      Left            =   17640
      TabIndex        =   40
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number"
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
      Height          =   495
      Left            =   16080
      TabIndex        =   39
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
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
      Height          =   495
      Left            =   13320
      TabIndex        =   38
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
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
      Height          =   495
      Left            =   11640
      TabIndex        =   37
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code"
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
      Height          =   495
      Left            =   10200
      TabIndex        =   36
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 2"
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
      Height          =   495
      Left            =   8640
      TabIndex        =   35
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 1"
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
      Height          =   495
      Left            =   7080
      TabIndex        =   34
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
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
      Height          =   495
      Left            =   14880
      TabIndex        =   33
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label7 
      Height          =   375
      Left            =   8040
      TabIndex        =   32
      Top             =   8520
      Width           =   1935
   End
End
Attribute VB_Name = "Agent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Add_Click()
Save_Cancel
Label7.Caption = "ADD"
HideListIndexes
UnlockTexts
End Sub

Private Sub Agent_REFID_Click()
Label23.Caption = Agent_REFID.List(Agent_REFID.ListIndex)
    List1.ListIndex = Agent_REFID.ListIndex
    List2.ListIndex = Agent_REFID.ListIndex
    List3.ListIndex = Agent_REFID.ListIndex
    List4.ListIndex = Agent_REFID.ListIndex
    List5.ListIndex = Agent_REFID.ListIndex
    List6.ListIndex = Agent_REFID.ListIndex
    List7.ListIndex = Agent_REFID.ListIndex
    List8.ListIndex = Agent_REFID.ListIndex
    List9.ListIndex = Agent_REFID.ListIndex
    List10.ListIndex = Agent_REFID.ListIndex
    List11.ListIndex = Agent_REFID.ListIndex
    Agent_File_Pointer = Agent_REFID.ListIndex + 1
End Sub

Private Sub Delete_Click()
Dim tempDeleteAgentchannel As Integer
Dim tempDeleteAgentfile As String
Dim FoundRecord As Boolean
Dim RemoveAgent As Agent_Record
Dim RemoveAgentchannel As Integer
Dim P As Integer
Dim L As Integer
Dim intResponse As Integer
If Agent_REFID.ListIndex = -1 Then
MsgBox ("Please Select a Agent To Delete")
Else
    intResponse = MsgBox("Are you sure you want to delete Agent " & Agent_REFID & "?" _
    & "                                  Once the Agent removed you cannot recover it", _
                         vbYesNo + vbQuestion + vbDefaultButton2, _
                         "Delete")
    If intResponse = vbYes Then
tempDeleteAgentfile = App.Path + "\Saved_Dat\tempfile.tmp"
RemoveAgentchannel = FreeFile
Open Agent_File For Random As RemoveAgentchannel Len = Agent_Length
tempDeleteAgentchannel = FreeFile
Open tempDeleteAgentfile For Random As tempDeleteAgentchannel Len = Agent_Length
P = 1
L = 1
FoundRecord = False
Get RemoveAgentchannel, P, RemoveAgent
Do While Not EOF(RemoveAgentchannel)
        If Agent_REFID.List(Agent_REFID.ListIndex) <> RemoveAgent.Agent_REFID Then
                Put tempDeleteAgentchannel, L, RemoveAgent
                L = L + 1
        Else
                FoundRecord = True
        End If
        P = P + 1
        Get RemoveAgentchannel, P, RemoveAgent
Loop
Close RemoveAgentchannel
Close tempDeleteAgentchannel
    If FoundRecord = True Then
    MsgBox "This Agent Has Been Successfully Deleted"
        Kill Agent_File
        Name tempDeleteAgentfile As Agent_File
        Agent_File_Pointer = Agent_File_Pointer - 1
    Else
   MsgBox "not found"
        Kill tempDeleteTenantfile
    End If
ClearThoseListBoxes
Form_Load
    End If
End If

End Sub

Private Sub Form_Activate()
Label23.Caption = Selected_AgentID
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
    List1.AddItem ViewAgentsID.Agent_FName
    List2.AddItem ViewAgentsID.Agent_LName
    List3.AddItem ViewAgentsID.Agent_CName
    List4.AddItem ViewAgentsID.Agent_Address1
    List5.AddItem ViewAgentsID.Agent_Address2
    List6.AddItem ViewAgentsID.Agent_PostCode
    List7.AddItem ViewAgentsID.Agent_CountryList
    List8.AddItem ViewAgentsID.Agent_State
    List11.AddItem ViewAgentsID.Agent_City
    List9.AddItem ViewAgentsID.Agent_PhoneNumber
    List10.AddItem ViewAgentsID.Agent_EmailAddress
    x = x + 1
    Get ViewAgentsIDChannel, x, ViewAgentsID
Loop
Close ViewAgentsIDChannel
End Sub

Private Sub Country_list_property_Click()
Agents_Country.Text = Country_list_property
End Sub
Private Sub GenerateRandomAgentRefID()
Agents_RefID.Text = Int(Rnd * 997) + 1
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
        If Trim(ViewAgentRefID.Agent_REFID) = Trim(Agents_RefID.Text) Then
        GenerateRandomAgentRefID
        Agent_File_Pointer = x
        End If
        x = x + 1
        Get ViewAgentRefIDChannel, x, ViewAgentRefID
    Loop
Close ViewAgentRefIDChannel
End Sub
Private Function ClearThoseListBoxes()
Agent_REFID.Clear
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
List8.Clear
List9.Clear
List10.Clear
List11.Clear
End Function

Private Function HideListIndexes()
Agent_REFID.ListIndex = -1
List1.ListIndex = -1
List2.ListIndex = -1
List3.ListIndex = -1
List4.ListIndex = -1
List5.ListIndex = -1
List6.ListIndex = -1
List7.ListIndex = -1
List8.ListIndex = -1
List9.ListIndex = -1
List10.ListIndex = -1
List11.ListIndex = -1
End Function
Private Function UnlockTexts()
Agents_Fname.Locked = False
Agents_LName2.Locked = False
Agents_Cname2.Locked = False
Agents_AddressLine1_2.Locked = False
Agents_AddressLine2_2.Locked = False
Agents_PostCode2.Locked = False
Agents_Country.Locked = False
Agents_State2.Locked = False
Agents_City2.Locked = False
Agents_PhoneNo.Locked = False
Agents_EmailAddress.Locked = False
Country_list_property.Locked = False
End Function
Private Function lockTexts()
Agents_Fname.Locked = True
Agents_LName2.Locked = True
Agents_Cname2.Locked = True
Agents_AddressLine1_2.Locked = True
Agents_AddressLine2_2.Locked = True
Agents_PostCode2.Locked = True
Agents_Country.Locked = True
Agents_State2.Locked = True
Agents_City2.Locked = True
Agents_PhoneNo.Locked = True
Agents_EmailAddress.Locked = True
Country_list_property.Locked = True
End Function
Private Function ClearTexts()
Agents_Fname.Text = ""
Agents_LName2.Text = ""
Agents_Cname2.Text = ""
Agents_AddressLine1_2.Text = ""
Agents_AddressLine2_2.Text = ""
Agents_PostCode2.Text = ""
Agents_Country.Text = ""
Agents_State2.Text = ""
Agents_City2.Text = ""
Agents_PhoneNo.Text = ""
Agents_EmailAddress.Text = ""
Agents_RefID.Text = ""
End Function
Private Sub EdIT_Click()
If Agent_REFID.ListIndex = -1 Then
MsgBox "Please Select A Agent Refrence ID"
Else
Agents_RefID.Text = Agent_REFID.List(Agent_REFID.ListIndex)
Agents_RefID.BackColor = &H80FF80
Save_Cancel
Label7.Caption = "EDIT"
UnlockTexts
End If
End Sub

Private Function Save_Cancel()
Add.Enabled = False
EdIT.Enabled = False
Prints.Enabled = False
Delete.Enabled = False
Save.Enabled = True
Cancel.Enabled = True
Agents_Fname.SetFocus
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
Private Sub cancel_Click()
CancelFunction
lockTexts
ClearTexts
End Sub

Private Sub Close_Click()
Unload Me
End Sub

Private Sub Label23_Click()
Selected_AgentID = Label23.Caption
End Sub

Private Sub Save_Click()
Dim Add_Agent As Agent_Record
Dim Agentchannel As Integer
Dim Amend_Agent As Agent_Record
Dim AmendAgentChannel As Integer
        If Label7 = "ADD" Then
CheckAgentRef
    Agentchannel = FreeFile
    Open Agent_File For Random As Agentchannel Len = Agent_Length
            Add_Agent.Agent_FName = Agents_Fname.Text
            Add_Agent.Agent_LName = Agents_LName2.Text
            Add_Agent.Agent_CName = Agents_Cname2.Text
            Add_Agent.Agent_Address1 = Agents_AddressLine1_2.Text
            Add_Agent.Agent_Address2 = Agents_AddressLine2_2.Text
            Add_Agent.Agent_PostCode = Agents_PostCode2.Text
            Add_Agent.Agent_CountryList = Agents_Country.Text
            Add_Agent.Agent_State = Agents_State2.Text
            Add_Agent.Agent_City = Agents_City2.Text
            Add_Agent.Agent_PhoneNumber = Agents_PhoneNo.Text
            Add_Agent.Agent_EmailAddress = Agents_EmailAddress.Text
            Add_Agent.Agent_REFID = Agents_RefID.Text
            MsgBox Agents_RefID.Text
            Agent_File_Pointer = Agent_File_Pointer + 1
    Put Agentchannel, Agent_File_Pointer, Add_Agent
    Close Agentchannel
ClearTexts
    ElseIf Label7.Caption = "EDIT" Then
    AmendAgentChannel = FreeFile
    Open Agent_File For Random As AmendAgentChannel Len = Agent_Length
                Amend_Agent.Agent_REFID = Agents_RefID.Text
                Amend_Agent.Agent_FName = Agents_Fname.Text
                Amend_Agent.Agent_LName = Agents_LName2.Text
                Amend_Agent.Agent_CName = Agents_Cname2.Text
                Amend_Agent.Agent_Address1 = Agents_AddressLine1_2.Text
                Amend_Agent.Agent_Address2 = Agents_AddressLine2_2.Text
                Amend_Agent.Agent_CountryList = Agents_PostCode2.Text
                Amend_Agent.Agent_PostCode = Agents_Country.Text
                Amend_Agent.Agent_City = Agents_State2.Text
                Amend_Agent.Agent_State = Agents_City2.Text
                Amend_Agent.Agent_PhoneNumber = Agents_PhoneNo.Text
                Amend_Agent.Agent_EmailAddress = Agents_EmailAddress.Text
    Put AmendAgentChannel, Agent_File_Pointer, Amend_Agent
    Close AmendAgentChannel
End If
Agents_RefID.BackColor = &H80000005
ClearThoseListBoxes
Form_Load
End Sub
