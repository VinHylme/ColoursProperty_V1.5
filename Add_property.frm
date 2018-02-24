VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Add_properties 
   Caption         =   "Add Property"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   Picture         =   "Add_property.frx":0000
   ScaleHeight     =   8535
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Property_NumberOfBedsLookingForList 
      BackColor       =   &H80000004&
      Height          =   315
      ItemData        =   "Add_property.frx":E4AB
      Left            =   2400
      List            =   "Add_property.frx":E4C7
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   5040
      Width           =   4695
   End
   Begin VB.TextBox Property_RefID 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   24
      Top             =   7080
      Width           =   4695
   End
   Begin VB.ComboBox Agent_list 
      BackColor       =   &H80000004&
      Height          =   315
      ItemData        =   "Add_property.frx":E4E3
      Left            =   2400
      List            =   "Add_property.frx":E4E5
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   6720
      Width           =   4695
   End
   Begin VB.ComboBox Landlord_list 
      BackColor       =   &H80000004&
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   6240
      Width           =   4695
   End
   Begin VB.PictureBox add_this_image 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7200
      Picture         =   "Add_property.frx":E4E7
      ScaleHeight     =   495
      ScaleWidth      =   2295
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox Choose_image 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7200
      Picture         =   "Add_property.frx":12D8D
      ScaleHeight     =   495
      ScaleWidth      =   2295
      TabIndex        =   18
      Top             =   5520
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD PROPERTY"
      Height          =   615
      Left            =   7440
      TabIndex        =   17
      Top             =   7680
      Width           =   2175
   End
   Begin VB.ComboBox property_type 
      BackColor       =   &H80000004&
      Height          =   315
      ItemData        =   "Add_property.frx":16A57
      Left            =   2400
      List            =   "Add_property.frx":16A79
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3360
      Width           =   4695
   End
   Begin VB.ComboBox Payment_type 
      BackColor       =   &H80000004&
      Height          =   315
      ItemData        =   "Add_property.frx":16AFB
      Left            =   2400
      List            =   "Add_property.frx":16B02
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3840
      Width           =   4695
   End
   Begin VB.ComboBox City_list_property 
      BackColor       =   &H80000004&
      Height          =   315
      ItemData        =   "Add_property.frx":16B0F
      Left            =   2400
      List            =   "Add_property.frx":16BA9
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2280
      Width           =   4695
   End
   Begin VB.TextBox add_image 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   5520
      Width           =   4695
   End
   Begin VB.TextBox RentalPrice 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   4320
      Width           =   4695
   End
   Begin VB.TextBox Post_code_property 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2760
      Width           =   4695
   End
   Begin VB.TextBox Address_line_2 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   1680
      Width           =   4695
   End
   Begin VB.TextBox Address_line_1 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   1080
      Width           =   4695
   End
   Begin VB.ListBox LandlordID 
      Height          =   255
      Left            =   2400
      TabIndex        =   27
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox AgentID 
      Height          =   255
      Left            =   2400
      TabIndex        =   28
      Top             =   6600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Property Ref:"
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
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Landlord:"
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
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent:"
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
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Image:"
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
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of beds:"
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
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Price(weekly):"
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
      TabIndex        =   14
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Type:"
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
      TabIndex        =   13
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Property Type:"
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
      TabIndex        =   12
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label5 
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
      TabIndex        =   11
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label3 
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
      TabIndex        =   10
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label2 
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
      TabIndex        =   9
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
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
      TabIndex        =   8
      Top             =   1200
      Width           =   2175
   End
End
Attribute VB_Name = "Add_properties"
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
Private Sub add_this_image_Click()
add_this_image.Visible = False
Choose_image.Visible = True
End Sub
Private Sub Agent_list_Click()
AgentID.ListIndex = agent_list.ListIndex
End Sub
Private Sub GenerateRandomIDProperty()
Property_RefID.Text = RndWord(8)
End Sub

Private Sub Form_Load()
LoadLandlords
LoadAgents
End Sub
Private Sub Choose_image_Click()
CommonDialog1.Filter = "Jpeg Files(*.jpg) | *.jpg"
CommonDialog1.ShowOpen
add_image.Text = CommonDialog1.FileName
Choose_image.Visible = False
add_this_image.Visible = True
End Sub
Private Sub Command1_Click()
'GenerateRandomIDProperty
CheckPropertyRef
Dim Addproperty As Property_Record
Dim PropertyChannel As Integer
PropertyChannel = FreeFile
    Open Add_property_File For Random As PropertyChannel Len = Add_property_Length
            Addproperty.Property_RefID = Property_RefID.Text
            Addproperty.Landlord_REFID = LandlordID.List(LandlordID.ListIndex)
            Addproperty.Agent_REFID = AgentID.List(AgentID.ListIndex)
            Addproperty.Address_line_1 = Address_line_1.Text
            Addproperty.Address_line_2 = Address_line_2.Text
            Addproperty.City = City_list_property
            Addproperty.Post_Code = Post_code_property.Text
            Addproperty.property_type = property_type
            Addproperty.Payment_type = Payment_type
            Addproperty.NumberOfBeds = Property_NumberOfBedsLookingForList
            Addproperty.Price_Property = RentalPrice.Text
            Addproperty.Image_property = add_image.Text
            Add_property_Pointer = Add_property_Pointer + 1
    Put PropertyChannel, Add_property_Pointer, Addproperty
    Close PropertyChannel
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
Private Sub CheckPropertyRef()
GenerateRandomIDProperty
Dim ViewPropertyRefID As Property_Record
Dim ViewPropertyRefIDChannel As Integer
Dim x As Integer
x = 1
ViewPropertyRefIDChannel = FreeFile
Open Add_property_File For Random As ViewPropertyRefIDChannel Len = Add_property_Length
Get ViewPropertyRefIDChannel, x, ViewPropertyRefID
    Do While Not EOF(ViewPropertyRefIDChannel)
        If Trim(ViewPropertyRefID.Property_RefID) = Trim(Property_RefID.Text) Then
        GenerateRandomIDProperty
        Add_property_Pointer = x
        End If
        x = x + 1
        Get ViewPropertyRefIDChannel, x, ViewPropertyRefID
    Loop
Close ViewPropertyRefIDChannel
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
Private Sub Landlord_list_Click()
    LandlordID.ListIndex = Landlord_list.ListIndex
End Sub

Private Sub Landlord_RefID_Select_Click()
IDLandlord.Text = Landlord_RefID_Select
End Sub

