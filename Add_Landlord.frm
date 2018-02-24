VERSION 5.00
Begin VB.Form Add_Landlord 
   Caption         =   "Add A Landlord"
   ClientHeight    =   9270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Landlord_Country 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   4680
      TabIndex        =   28
      Top             =   4080
      Width           =   2415
   End
   Begin VB.TextBox LandlordRefID 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Register_Landlord 
      Caption         =   "Register Landlord"
      Height          =   855
      Left            =   6960
      TabIndex        =   25
      Top             =   8280
      Width           =   1815
   End
   Begin VB.TextBox Landlord_FaxNo 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   23
      Top             =   5760
      Width           =   4695
   End
   Begin VB.ComboBox Country_list_property 
      BackColor       =   &H80000004&
      Height          =   315
      ItemData        =   "Add_Landlord.frx":0000
      Left            =   2400
      List            =   "Add_Landlord.frx":0250
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox Landlord_EmailAddress 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   18
      Top             =   6960
      Width           =   4695
   End
   Begin VB.TextBox Landlord_PhoneNo 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   16
      Top             =   6360
      Width           =   4695
   End
   Begin VB.TextBox Landlord_City 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   14
      Top             =   5160
      Width           =   4695
   End
   Begin VB.TextBox Landlord_State 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   4560
      Width           =   4695
   End
   Begin VB.TextBox Landlord_PostCode 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   3480
      Width           =   4695
   End
   Begin VB.TextBox Landlord_Address1 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   2880
      Width           =   4695
   End
   Begin VB.TextBox Landlord_Address2 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   2280
      Width           =   4695
   End
   Begin VB.TextBox Landlord_CName 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   1680
      Width           =   4695
   End
   Begin VB.TextBox Landlord_LName 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   4695
   End
   Begin VB.TextBox Landlord_FName 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   4695
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
      TabIndex        =   27
      Top             =   7800
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
      TabIndex        =   24
      Top             =   5880
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
      TabIndex        =   22
      Top             =   4080
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
      TabIndex        =   20
      Top             =   7200
      Width           =   2175
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
      TabIndex        =   19
      Top             =   7080
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
      TabIndex        =   17
      Top             =   6480
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
      TabIndex        =   15
      Top             =   5280
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
      TabIndex        =   13
      Top             =   4680
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
      TabIndex        =   11
      Top             =   3600
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
      TabIndex        =   9
      Top             =   3000
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
      TabIndex        =   7
      Top             =   2400
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
      TabIndex        =   5
      Top             =   1800
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
      TabIndex        =   3
      Top             =   1200
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
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "Add_Landlord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Country_list_property_Click()
Landlord_Country = Country_list_property
End Sub
Private Sub GenerateRandomLandlordRefID()
LandlordRefID.Text = Int(Rnd * 999) + 1
End Sub
Private Sub CheckLandlordRef()
GenerateRandomLandlordRefID
Dim ViewLandlord As Landlord_Record
Dim ViewLandlordChannel As Integer
Dim x As Integer
x = 1
ViewLandlordChannel = FreeFile
Open Landlord_File For Random As ViewLandlordChannel Len = Landlord_Length
Get ViewLandlordChannel, x, ViewLandlord
    Do While Not EOF(ViewLandlordChannel)
        If Trim(ViewLandlord.Landlord_RefID) = Trim(LandlordRefID.Text) Then
        GenerateRandomLandlordRefID
        Landlord_File_Pointer = x
        End If
        x = x + 1
        Get ViewLandlordChannel, x, ViewLandlord
    Loop
Close ViewLandlordChannel
End Sub

Private Sub Register_Landlord_Click()
Dim Add_Landlord As Landlord_Record
Dim LandLordchannel As Integer
CheckLandlordRef
    LandLordchannel = FreeFile
    Open Landlord_File For Random As LandLordchannel Len = Landlord_Length
            Add_Landlord.Landlord_FName = Landlord_FName.Text
            Add_Landlord.Landlord_LName = Landlord_LName.Text
            Add_Landlord.Landlord_CName = Landlord_CName.Text
            Add_Landlord.Landlord_Address1 = Landlord_Address1.Text
            Add_Landlord.Landlord_Address2 = Landlord_Address2.Text
            Add_Landlord.Landlord_PostCode = Landlord_PostCode.Text
            Add_Landlord.Landlord_CountryList = Landlord_Country.Text
            Add_Landlord.Landlord_State = Landlord_State.Text
            Add_Landlord.Landlord_City = Landlord_City.Text
            Add_Landlord.Landlord_FaxNumber = Landlord_FaxNo.Text
            Add_Landlord.Landlord_PhoneNumber = Landlord_PhoneNo.Text
            Add_Landlord.Landlord_EmailAddress = Landlord_EmailAddress.Text
            Add_Landlord.Landlord_RefID = LandlordRefID.Text
            MsgBox LandlordRefID
            Landlord_File_Pointer = Landlord_File_Pointer + 1
    Put LandLordchannel, Landlord_File_Pointer, Add_Landlord
    Close LandLordchannel
End Sub
