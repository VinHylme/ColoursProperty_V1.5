VERSION 5.00
Begin VB.Form Properties 
   Caption         =   "Manage Properties"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17025
   LinkTopic       =   "Form2"
   ScaleHeight     =   7425
   ScaleWidth      =   17025
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox RentalLink 
      Height          =   285
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   4080
      Width           =   1455
   End
   Begin VB.ComboBox RentalType 
      Height          =   315
      ItemData        =   "Properties.frx":0000
      Left            =   5160
      List            =   "Properties.frx":000A
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   61
      Top             =   4080
      Width           =   3255
   End
   Begin VB.TextBox Property_RefID 
      Height          =   285
      Left            =   5160
      MaxLength       =   8
      TabIndex        =   22
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CommandButton Add 
      Caption         =   "ADD"
      Height          =   615
      Left            =   10920
      TabIndex        =   51
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton Delete 
      Caption         =   "DELETE"
      Height          =   615
      Left            =   13560
      TabIndex        =   50
      Top             =   4800
      Width           =   2775
   End
   Begin VB.CommandButton Prints 
      Caption         =   "PRINT"
      Height          =   615
      Left            =   10920
      TabIndex        =   49
      Top             =   4800
      Width           =   2535
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "CANCEL"
      Enabled         =   0   'False
      Height          =   615
      Left            =   13560
      TabIndex        =   48
      Top             =   5520
      Width           =   2775
   End
   Begin VB.CommandButton Close 
      Caption         =   "CLOSE"
      Height          =   615
      Left            =   10920
      TabIndex        =   47
      Top             =   6240
      Width           =   5415
   End
   Begin VB.CommandButton EdIT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "EDIT"
      Height          =   615
      Left            =   13560
      TabIndex        =   46
      Top             =   4080
      Width           =   2775
   End
   Begin VB.CommandButton Save 
      Caption         =   "SAVE"
      Enabled         =   0   'False
      Height          =   615
      Left            =   10920
      TabIndex        =   45
      Top             =   5520
      Width           =   2535
   End
   Begin VB.TextBox Address_line_1 
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3960
      Width           =   4695
   End
   Begin VB.TextBox Address_line_2 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   4680
      Width           =   4695
   End
   Begin VB.TextBox Post_code_property 
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   5880
      Width           =   4695
   End
   Begin VB.TextBox RentalPrice 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4680
      Width           =   4695
   End
   Begin VB.ComboBox City_list_property 
      Height          =   315
      ItemData        =   "Properties.frx":001F
      Left            =   120
      List            =   "Properties.frx":00B9
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   5280
      Width           =   2175
   End
   Begin VB.ComboBox Payment_type 
      Height          =   315
      ItemData        =   "Properties.frx":02DE
      Left            =   120
      List            =   "Properties.frx":02E5
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   7080
      Width           =   2295
   End
   Begin VB.ComboBox property_type 
      Height          =   315
      ItemData        =   "Properties.frx":02F2
      Left            =   120
      List            =   "Properties.frx":0314
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   6480
      Width           =   2295
   End
   Begin VB.ComboBox Landlord_list 
      Height          =   315
      Left            =   5160
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   5880
      Width           =   2295
   End
   Begin VB.ComboBox agent_list 
      Height          =   315
      ItemData        =   "Properties.frx":0396
      Left            =   5160
      List            =   "Properties.frx":0398
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   6480
      Width           =   2295
   End
   Begin VB.ComboBox Property_NumberOfBedsLookingForList 
      Height          =   315
      ItemData        =   "Properties.frx":039A
      Left            =   5160
      List            =   "Properties.frx":03B6
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   5280
      Width           =   2295
   End
   Begin VB.ListBox Property_Landlord 
      Height          =   2790
      Left            =   15240
      TabIndex        =   30
      Top             =   600
      Width           =   1695
   End
   Begin VB.ListBox Property_Agent 
      Height          =   2790
      Left            =   13800
      TabIndex        =   29
      Top             =   600
      Width           =   1455
   End
   Begin VB.ListBox Properties_REFID 
      Height          =   2790
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   2055
   End
   Begin VB.ListBox Property_AddressLine1 
      Height          =   2790
      Left            =   2280
      TabIndex        =   7
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox Property_AddressLine2 
      Height          =   2790
      Left            =   3840
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox Property_City 
      Height          =   2790
      Left            =   5400
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox Property_PostCode 
      Height          =   2790
      Left            =   6960
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.ListBox Property_PropertyType 
      Height          =   2790
      Left            =   8040
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.ListBox Property_PaymentType 
      Height          =   2790
      Left            =   9480
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.ListBox Property_Price 
      Height          =   2790
      Left            =   10800
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.ListBox Property_NumberOFBeds 
      Height          =   2790
      Left            =   12240
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox LandlordID 
      Height          =   255
      Left            =   5160
      TabIndex        =   34
      Top             =   5880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox AgentID 
      Height          =   255
      Left            =   5160
      TabIndex        =   35
      Top             =   6480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox City_Link 
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   5280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox PropertyType_Link 
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   6480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox PaymentType_Link 
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   7080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox NumberOfBeds_Link 
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   5280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Landlord_Link 
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   5880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Agent_Link 
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   6480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Price(Monthly):"
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
      Left            =   5160
      TabIndex        =   64
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Rental Type:"
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
      Left            =   5160
      TabIndex        =   63
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   7200
      TabIndex        =   54
      Top             =   9000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label23 
      Caption         =   "Label23"
      Height          =   375
      Left            =   8160
      TabIndex        =   53
      Top             =   8880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Property Refrence ID"
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
      Left            =   5160
      TabIndex        =   52
      Top             =   6840
      Width           =   2535
   End
   Begin VB.Line Line2 
      X1              =   10080
      X2              =   10080
      Y1              =   3480
      Y2              =   7680
   End
   Begin VB.Label Label22 
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
      TabIndex        =   33
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label21 
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
      TabIndex        =   44
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label20 
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
      Left            =   120
      TabIndex        =   43
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code:"
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
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Property Type:"
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
      TabIndex        =   41
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Type:"
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
      TabIndex        =   40
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Price(weekly):"
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
      Left            =   5160
      TabIndex        =   39
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of beds:"
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
      Left            =   5160
      TabIndex        =   38
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent:"
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
      Left            =   5160
      TabIndex        =   37
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Landlord:"
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
      Left            =   5160
      TabIndex        =   36
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   -120
      X2              =   20280
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Landord 
      BackStyle       =   0  'Transparent
      Caption         =   "Landlord Refrence"
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
      Left            =   15360
      TabIndex        =   32
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent Refrence"
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
      Left            =   13920
      TabIndex        =   31
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label PropertyRefID 
      BackStyle       =   0  'Transparent
      Caption         =   "Property Refrence ID:"
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
      TabIndex        =   28
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
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
      Left            =   2520
      TabIndex        =   27
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
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
      Left            =   4080
      TabIndex        =   26
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label City 
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
      Left            =   5880
      TabIndex        =   25
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Paying upto"
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
      Left            =   10920
      TabIndex        =   24
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Type"
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
      Left            =   9600
      TabIndex        =   23
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Property Type"
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
      Left            =   8160
      TabIndex        =   11
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label12 
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
      Left            =   7080
      TabIndex        =   10
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of beds"
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
      Left            =   12360
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const RndString = "01234567890ABCDEFGHIJKLMNOP"
Private Function RndWord(Optional Chrlength As Integer = 8) As String
Dim TempWord As String
Dim LoopVar As Integer
For LoopVar = 1 To Chrlength
    TempWord = TempWord & Mid(RndString, Int(Rnd * 22) + 1, 1)
    Next LoopVar
    RndWord = TempWord
End Function
Private Function Rndchr() As String
Rndchr = Mid(RndString, Int(Rnd * 16) + 1, 1)
End Function
Private Sub GenerateRandomIDProperty()
Property_RefID.Text = RndWord(8)
End Sub
Private Sub CheckPropertyRef()
GenerateRandomIDProperty
Dim ViewPropertyRefID As Property_Record
Dim ViewPropertyRefIDChannel As Integer
Dim x As Integer
x = 1
ViewPropertyRefIDChannel = FreeFile
Open Add_Property_File For Random As ViewPropertyRefIDChannel Len = Add_Property_Length
Get ViewPropertyRefIDChannel, x, ViewPropertyRefID
    Do While Not EOF(ViewPropertyRefIDChannel)
        If Trim(ViewPropertyRefID.Property_RefID) = Trim(Property_RefID.Text) Then
        GenerateRandomIDProperty
        Add_Property_Pointer = x
        End If
        x = x + 1
        Get ViewPropertyRefIDChannel, x, ViewPropertyRefID
    Loop
Close ViewPropertyRefIDChannel
End Sub
Private Sub Add_Click()
Save_Cancel
Label23.Caption = "ADD"
HideListIndexes
UnlockTexts
End Sub

Private Sub AgentID_Click()
Agent_Link.Text = AgentID.List(AgentID.ListIndex)
End Sub

Private Sub cancel_Click()
CancelFunction
lockTexts
ClearTexts
Property_RefID.BackColor = &H80000005
HideListedTextBoxes
End Sub

Private Sub City_list_property_click()
City_Link.Text = City_list_property.List(City_list_property.ListIndex)
End Sub

Private Sub Close_Click()
Unload Me
End Sub
Private Function ShowListedTextBoxes()
City_Link.Visible = True
PropertyType_Link.Visible = True
PaymentType_Link.Visible = True
NumberOfBeds_Link.Visible = True
Landlord_Link.Visible = True
Agent_Link.Visible = True
RentalLink.Visible = True
End Function
Private Function HideListedTextBoxes()
City_Link.Visible = False
PropertyType_Link.Visible = False
PaymentType_Link.Visible = False
NumberOfBeds_Link.Visible = False
Landlord_Link.Visible = False
Agent_Link.Visible = False
RentalLink.Visible = False
End Function

Private Sub Delete_Click()
Dim tempDeletePropertychannel As Integer
Dim tempDeletePropertyfile As String
Dim FoundRecord As Boolean
Dim RemoveProperty As Property_Record
Dim RemovePropertychannel As Integer
Dim P As Integer
Dim L As Integer
Dim intResponse As Integer
If Properties_REFID.ListIndex = -1 Then
MsgBox ("Please Select a Property To Delete")
Else
    intResponse = MsgBox("Are you sure you want to delete property  " & Properties_REFID & "?" _
    & "                                  Once the property removed you cannot recover it", _
                         vbYesNo + vbQuestion + vbDefaultButton2, _
                         "Delete")
    If intResponse = vbYes Then
tempDeletePropertyfile = App.Path + "\Saved_Dat\tempfile.tmp"
RemovePropertychannel = FreeFile
Open Add_Property_File For Random As RemovePropertychannel Len = Add_Property_Length
tempDeletePropertychannel = FreeFile
Open tempDeletePropertyfile For Random As tempDeletePropertychannel Len = Add_Property_Length
P = 1
L = 1
FoundRecord = False
Get RemovePropertychannel, P, RemoveProperty
Do While Not EOF(RemovePropertychannel)
        If Properties_REFID.List(Properties_REFID.ListIndex) <> RemoveProperty.Property_RefID Then
                Put tempDeletePropertychannel, L, RemoveProperty
                L = L + 1
        Else
                FoundRecord = True
        End If
        P = P + 1
        Get RemovePropertychannel, P, RemoveProperty
Loop
Close RemovePropertychannel
Close tempDeletePropertychannel
    If FoundRecord = True Then
    MsgBox "This Property Has Been Successfully Deleted"
        Kill Add_Property_File
        Name tempDeletePropertyfile As Add_Property_File
        Add_Property_Pointer = Add_Property_Pointer - 1
    Else
   MsgBox "not found"
        Kill tempDeletePropertyfile
    End If
ClearThoseListBoxes
Form_Load
        Print "Rows would have been deleted at this point"
    End If
End If
End Sub

Private Sub EdIT_Click()
If Properties_REFID.ListIndex = -1 Then
MsgBox "Please Select A Property Refrence ID"
Else
Property_RefID.Text = Properties_REFID.List(Properties_REFID.ListIndex)
Property_RefID.BackColor = &H80FF80
Save_Cancel
Label23.Caption = "EDIT"
UnlockTexts
End If
Dim ViewProperties As Property_Record
Dim ViewPropertyChannel As Integer
Dim x As Integer
Dim OnceFoundProperty As Boolean
x = 1
ViewPropertyChannel = FreeFile
Open Add_Property_File For Random As ViewPropertyChannel Len = Add_Property_Length
Get ViewPropertyChannel, x, ViewProperties
Do While Not EOF(ViewPropertyChannel) And OnceFoundProperty = False
    If Trim(Properties_REFID.List(Properties_REFID.ListIndex)) = Trim(ViewProperties.Property_RefID) Then
    ShowListedTextBoxes
    Address_line_1.Text = ViewProperties.Address_line_1
    Address_line_2.Text = ViewProperties.Address_line_2
    City_Link.Text = ViewProperties.City
    Post_code_property.Text = ViewProperties.Post_Code
    PropertyType_Link.Text = ViewProperties.property_type
    PaymentType_Link.Text = ViewProperties.Payment_type
    RentalPrice = ViewProperties.Price_Property
    NumberOfBeds_Link.Text = ViewProperties.NumberOfBeds
    Landlord_Link.Text = ViewProperties.Landlord_REFID
    Agent_Link.Text = ViewProperties.Agent_REFID
    Add_Property_Pointer = x
    OnceFoundProperty = True
    End If
    x = x + 1
    Get ViewPropertyChannel, x, ViewProperties
Loop
Close ViewPropertyChannel
End Sub

Private Function Save_Cancel()
Add.Enabled = False
EdIT.Enabled = False
Prints.Enabled = False
Delete.Enabled = False
Save.Enabled = True
Cancel.Enabled = True
Address_line_1.SetFocus
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
Private Function ClearThoseListBoxes()
Properties_REFID.Clear
Property_AddressLine1.Clear
Property_AddressLine2.Clear
Property_City.Clear
Property_PostCode.Clear
Property_PropertyType.Clear
Property_PaymentType.Clear
Property_Price.Clear
Property_NumberOFBeds.Clear
Property_Agent.Clear
Property_Landlord.Clear
End Function

Private Function HideListIndexes()
Properties_REFID.ListIndex = -1
Property_AddressLine1.ListIndex = -1
Property_AddressLine2.ListIndex = -1
Property_City.ListIndex = -1
Property_PostCode.ListIndex = -1
Property_PropertyType.ListIndex = -1
Property_PaymentType.ListIndex = -1
Property_Price.ListIndex = -1
Property_NumberOFBeds.ListIndex = -1
Property_Agent.ListIndex = -1
Property_Landlord.ListIndex = -1
End Function
Private Function UnlockTexts()
Address_line_1.Locked = False
Address_line_2.Locked = False
City_list_property.Locked = False
Post_code_property.Locked = False
property_type.Locked = False
Payment_type.Locked = False
RentalPrice.Locked = False
Property_NumberOfBedsLookingForList.Locked = False
Landlord_list.Locked = False
agent_list.Locked = False
RentalType.Locked = False
End Function
Private Function lockTexts()
Address_line_1.Locked = True
Address_line_2.Locked = True
City_list_property.Locked = True
Post_code_property.Locked = True
property_type.Locked = True
Payment_type.Locked = True
RentalPrice.Locked = True
Property_NumberOfBedsLookingForList.Locked = True
Landlord_list.Locked = True
RentalType.Locked = True
agent_list.Locked = True
End Function
Private Function ClearTexts()
Address_line_1.Text = ""
Address_line_2.Text = ""
City_list_property.ListIndex = -1
Post_code_property.Text = ""
property_type.ListIndex = -1
Payment_type.ListIndex = -1
RentalPrice = ""
Property_NumberOfBedsLookingForList.ListIndex = -1
Landlord_list.ListIndex = 0
agent_list.ListIndex = 0
Property_RefID.Text = ""
End Function

Private Sub Form_Load()
If Landlord_list.ListIndex < 0 Then
LoadAgents
LoadLandlords
End If
Dim ViewProperty As Property_Record
Dim ViewPropertyChannel As Integer
Dim x As Integer
x = 1
ViewPropertyChannel = FreeFile
Open Add_Property_File For Random As ViewPropertyChannel Len = Add_Property_Length
Get ViewPropertyChannel, x, ViewProperty
Do While Not EOF(ViewPropertyChannel)
    Properties_REFID.AddItem ViewProperty.Property_RefID
    Property_AddressLine1.AddItem ViewProperty.Address_line_1
    Property_AddressLine2.AddItem ViewProperty.Address_line_2
    Property_City.AddItem ViewProperty.City
    Property_PostCode.AddItem ViewProperty.Post_Code
    Property_PropertyType.AddItem ViewProperty.property_type
    Property_PaymentType.AddItem ViewProperty.Payment_type
    Property_Price.AddItem ViewProperty.Price_Property
    Property_NumberOFBeds.AddItem ViewProperty.NumberOfBeds
    Property_Agent.AddItem ViewProperty.Agent_REFID
    Property_Landlord.AddItem ViewProperty.Landlord_REFID
    x = x + 1
    Get ViewPropertyChannel, x, ViewProperty
Loop
Close ViewPropertyChannel
End Sub
Private Sub LoadAgents()
Dim ViewAgentCompanyName As Agent_Record
Dim ViewAgentCompanyNameChannel As Integer
Dim x As Integer
x = 1
ViewAgentCompanyNameChannel = FreeFile
Open Agent_File For Random As ViewAgentCompanyNameChannel Len = Agent_Length
 Get ViewAgentCompanyNameChannel, x, ViewAgentCompanyName
    Do While Not EOF(ViewAgentCompanyNameChannel)
    agent_list.AddItem ViewAgentCompanyName.Agent_CName
    AgentID.AddItem ViewAgentCompanyName.Agent_REFID
     x = x + 1
    Get ViewAgentCompanyNameChannel, x, ViewAgentCompanyName
       
    Loop
Close ViewAgentCompanyNameChannel
End Sub
Private Sub LoadLandlords()
Dim ViewLandlordCompanyName As Landlord_Record
Dim ViewLandlordCompanyNameChannel As Integer
Dim x As Integer
x = 1
ViewLandlordCompanyNameChannel = FreeFile
Open Landlord_File For Random As ViewLandlordCompanyNameChannel Len = Landlord_Length
Get ViewLandlordCompanyNameChannel, x, ViewLandlordCompanyName
    Do While Not EOF(ViewLandlordCompanyNameChannel)
    Landlord_list.AddItem ViewLandlordCompanyName.Landlord_CName
    LandlordID.AddItem ViewLandlordCompanyName.Landlord_REFID
     x = x + 1
Get ViewLandlordCompanyNameChannel, x, ViewLandlordCompanyName
    Loop
Close ViewLandlordCompanyNameChannel
End Sub

Private Sub Landlord_list_Click()
LandlordID.ListIndex = Landlord_list.ListIndex
End Sub
Private Sub Agent_list_Click()
AgentID.ListIndex = agent_list.ListIndex
End Sub

Private Sub List1_Click()
City_list_property.ListIndex = List1.ListIndex
End Sub

Private Sub LandlordID_Click()
Landlord_Link.Text = LandlordID.List(LandlordID.ListIndex)
End Sub

Private Sub Payment_type_Click()
PaymentType_Link = Payment_type.List(Payment_type.ListIndex)
End Sub

Private Sub Prints_Click()
Unload Me
FullReport.Show 1
End Sub

Private Sub Properties_REFID_Click()
Label3.Caption = Properties_REFID.List(Properties_REFID.ListIndex)
Properties_REFID.ListIndex = Properties_REFID.ListIndex
Property_AddressLine1.ListIndex = Properties_REFID.ListIndex
Property_AddressLine2.ListIndex = Properties_REFID.ListIndex
Property_City.ListIndex = Properties_REFID.ListIndex
Property_PostCode.ListIndex = Properties_REFID.ListIndex
Property_PropertyType.ListIndex = Properties_REFID.ListIndex
Property_PaymentType.ListIndex = Properties_REFID.ListIndex
Property_Price.ListIndex = Properties_REFID.ListIndex
Property_NumberOFBeds.ListIndex = Properties_REFID.ListIndex
Property_Agent.ListIndex = Properties_REFID.ListIndex
Property_Landlord.ListIndex = Properties_REFID.ListIndex
Add_Property_Pointer = Properties_REFID.ListIndex + 1
End Sub

Private Sub Property_NumberOfBedsLookingForList_Click()
NumberOfBeds_Link.Text = Property_NumberOfBedsLookingForList.List(Property_NumberOfBedsLookingForList.ListIndex)
End Sub

Private Sub property_type_click()
PropertyType_Link = property_type.List(property_type.ListIndex)
End Sub

Private Sub RentalType_Click()
RentalLink = RentalType.List(RentalType.ListIndex)
If RentalType.List(RentalType.ListIndex) = "Weekly" Then
Label13.Visible = False
Label16.Visible = True
ElseIf RentalType.List(RentalType.ListIndex) = "Monthly" Then
Label13.Visible = True
Label16.Visible = False
End If
RentalPrice.Enabled = True
End Sub

Private Sub Save_Click()
Dim Addproperty As Property_Record
Dim PropertyChannel As Integer
If Label23.Caption = "ADD" Then
CheckPropertyRef
PropertyChannel = FreeFile
    Open Add_Property_File For Random As PropertyChannel Len = Add_Property_Length
            Addproperty.Property_RefID = Property_RefID.Text
            Addproperty.Landlord_REFID = LandlordID
            Addproperty.Agent_REFID = AgentID
            Addproperty.Address_line_1 = Address_line_1.Text
            Addproperty.Address_line_2 = Address_line_2.Text
            Addproperty.City = City_list_property
            Addproperty.Post_Code = Post_code_property.Text
            Addproperty.property_type = property_type
            Addproperty.Payment_type = Payment_type
            Addproperty.RentalType = RentalType
            Addproperty.NumberOfBeds = Property_NumberOfBedsLookingForList
            Addproperty.Price_Property = RentalPrice.Text
            Add_Property_Pointer = Add_Property_Pointer + 1
    Put PropertyChannel, Add_Property_Pointer, Addproperty
    Close PropertyChannel
ClearTexts
ElseIf Label23.Caption = "EDIT" Then
Dim AmendPropertyDetails As Property_Record
Dim AmendPropertyChannel As Integer
AmendPropertyChannel = FreeFile
Open Add_Property_File For Random As AmendPropertyChannel Len = Add_Property_Length
            AmendPropertyDetails.Property_RefID = Property_RefID.Text
            AmendPropertyDetails.Landlord_REFID = Landlord_Link.Text
            AmendPropertyDetails.Agent_REFID = Agent_Link.Text
            AmendPropertyDetails.Address_line_1 = Address_line_1.Text
            AmendPropertyDetails.Address_line_2 = Address_line_2.Text
            AmendPropertyDetails.City = City_Link.Text
            AmendPropertyDetails.Post_Code = Post_code_property.Text
            AmendPropertyDetails.property_type = PropertyType_Link.Text
            AmendPropertyDetails.Payment_type = PaymentType_Link.Text
            AmendPropertyDetails.NumberOfBeds = NumberOfBeds_Link.Text
            AmendPropertyDetails.RentalType = RentalLink.Text
            AmendPropertyDetails.Price_Property = RentalPrice.Text
    Put AmendPropertyChannel, Add_Property_Pointer, AmendPropertyDetails
    Close AmendPropertyChannel
End If
HideListedTextBoxes
ClearThoseListBoxes
Property_RefID.BackColor = &H80000005
Form_Load
End Sub
