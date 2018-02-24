VERSION 5.00
Begin VB.Form View_Landlord 
   Caption         =   "Display Landlord"
   ClientHeight    =   10635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   ScaleHeight     =   10635
   ScaleWidth      =   14310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton searchall 
      Caption         =   "Search All"
      Height          =   495
      Left            =   9720
      TabIndex        =   26
      Top             =   600
      Width           =   1335
   End
   Begin VB.ListBox LandlordList_EmailAddress 
      Height          =   3570
      Left            =   11760
      TabIndex        =   13
      Top             =   1920
      Width           =   2175
   End
   Begin VB.ListBox LandlordList_PhoneNo 
      Height          =   3570
      Left            =   9480
      TabIndex        =   12
      Top             =   1920
      Width           =   2175
   End
   Begin VB.ListBox LandlordList_Address1 
      Height          =   3570
      Left            =   360
      TabIndex        =   11
      Top             =   6120
      Width           =   2175
   End
   Begin VB.ListBox LandlordList_FaxNo 
      Height          =   3570
      Left            =   7200
      TabIndex        =   10
      Top             =   1920
      Width           =   2175
   End
   Begin VB.ListBox LandlordList_City 
      Height          =   3570
      Left            =   11760
      TabIndex        =   9
      Top             =   6120
      Width           =   2175
   End
   Begin VB.ListBox LandlordList_State 
      Height          =   3570
      Left            =   9480
      TabIndex        =   8
      Top             =   6120
      Width           =   2175
   End
   Begin VB.ListBox LandlordList_Country 
      Height          =   3570
      Left            =   7200
      TabIndex        =   7
      Top             =   6120
      Width           =   2175
   End
   Begin VB.ListBox LandlordList_PostCode 
      Height          =   3570
      Left            =   4920
      TabIndex        =   6
      Top             =   6120
      Width           =   2175
   End
   Begin VB.ListBox LandlordList_Address2 
      Height          =   3570
      Left            =   2640
      TabIndex        =   5
      Top             =   6120
      Width           =   2175
   End
   Begin VB.ListBox LandlordList_CName 
      Height          =   3570
      Left            =   4920
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.ListBox LandlordList_LName 
      Height          =   3570
      Left            =   2640
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   495
      Left            =   8400
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox LandlordRefID 
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
   Begin VB.ListBox LandlordList_FName 
      Height          =   3570
      ItemData        =   "View_Lanlord.frx":0000
      Left            =   360
      List            =   "View_Lanlord.frx":0002
      TabIndex        =   0
      Top             =   1920
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
      TabIndex        =   25
      Top             =   5760
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
      TabIndex        =   24
      Top             =   5760
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
      TabIndex        =   23
      Top             =   5760
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
      TabIndex        =   22
      Top             =   5760
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
      TabIndex        =   21
      Top             =   5760
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
      TabIndex        =   20
      Top             =   5760
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
      TabIndex        =   19
      Top             =   1560
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
      TabIndex        =   18
      Top             =   1560
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
      TabIndex        =   17
      Top             =   1560
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
      TabIndex        =   16
      Top             =   1560
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
      TabIndex        =   15
      Top             =   1560
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
      TabIndex        =   14
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "View_Landlord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
ClearListBoxesForLandlord
Dim ViewLandlord As Landlord_Record
Dim ViewLandlordChannel As Integer
Dim x As Integer
Dim OnceFoundLandlord As Boolean
x = 1
ViewLandlordChannel = FreeFile
Open Landlord_File For Random As ViewLandlordChannel Len = Landlord_Length
Get ViewLandlordChannel, x, ViewLandlord
Do While Not EOF(ViewLandlordChannel) And OnceFoundLandlord = False
    If Trim(ViewLandlord.Landlord_RefID) = Trim(LandlordRefID.Text) Then
    LandlordList_FName.AddItem ViewLandlord.Landlord_FName
    LandlordList_LName.AddItem ViewLandlord.Landlord_LName
    LandlordList_CName.AddItem ViewLandlord.Landlord_CName
    LandlordList_Address1.AddItem ViewLandlord.Landlord_Address1
    LandlordList_Address2.AddItem ViewLandlord.Landlord_Address2
    LandlordList_PostCode.AddItem ViewLandlord.Landlord_PostCode
    LandlordList_Country.AddItem ViewLandlord.Landlord_CountryList
    LandlordList_State.AddItem ViewLandlord.Landlord_State
    LandlordList_City.AddItem ViewLandlord.Landlord_City
    LandlordList_FaxNo.AddItem ViewLandlord.Landlord_FaxNumber
    LandlordList_PhoneNo.AddItem ViewLandlord.Landlord_PhoneNumber
    LandlordList_EmailAddress.AddItem ViewLandlord.Landlord_EmailAddress
    Landlord_File_Pointer = x
    OnceFoundLandlord = True
    Else
    MsgBox "This Landlord does not exist"
    x = x + 1
    Get ViewLandlordChannel, x, ViewLandlord
    End If
Loop
Close ViewLandlordChannel
End Sub

Private Sub ClearListBoxesForLandlord()
LandlordList_FName.Clear
LandlordList_LName.Clear
LandlordList_CName.Clear
LandlordList_Address1.Clear
LandlordList_Address2.Clear
LandlordList_PostCode.Clear
LandlordList_Country.Clear
LandlordList_State.Clear
LandlordList_City.Clear
LandlordList_FaxNo.Clear
LandlordList_PhoneNo.Clear
LandlordList_EmailAddress.Clear
End Sub
Private Sub searchall_Click()
ClearListBoxesForLandlord
Dim ViewTheLandlords As Landlord_Record
Dim TheLandlordChannel As Integer
Dim y As Integer
y = 1
TheLandlordChannel = FreeFile
Open Landlord_File For Random As TheLandlordChannel Len = Landlord_Length
Get TheLandlordChannel, y, ViewTheLandlords
Do While Not EOF(TheLandlordChannel)
    LandlordList_FName.AddItem ViewTheLandlords.Landlord_FName
    LandlordList_LName.AddItem ViewTheLandlords.Landlord_LName
    LandlordList_CName.AddItem ViewTheLandlords.Landlord_CName
    LandlordList_Address1.AddItem ViewTheLandlords.Landlord_Address1
    LandlordList_Address2.AddItem ViewTheLandlords.Landlord_Address2
    LandlordList_PostCode.AddItem ViewTheLandlords.Landlord_PostCode
    LandlordList_Country.AddItem ViewTheLandlords.Landlord_CountryList
    LandlordList_State.AddItem ViewTheLandlords.Landlord_State
    LandlordList_City.AddItem ViewTheLandlords.Landlord_City
    LandlordList_FaxNo.AddItem ViewTheLandlords.Landlord_FaxNumber
    LandlordList_PhoneNo.AddItem ViewTheLandlords.Landlord_PhoneNumber
    LandlordList_EmailAddress.AddItem ViewTheLandlords.Landlord_EmailAddress
y = y + 1
Get TheLandlordChannel, y, ViewTheLandlords
Loop
Close TheLandlordChannel
End Sub
