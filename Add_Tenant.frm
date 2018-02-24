VERSION 5.00
Begin VB.Form Add_Tenant 
   Caption         =   "Add Tenant"
   ClientHeight    =   9375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Tenant_NumberOfBedsLookingFor 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   4800
      TabIndex        =   31
      Top             =   7080
      Width           =   2415
   End
   Begin VB.TextBox Tenant_PropertyLocation 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   4800
      TabIndex        =   30
      Top             =   6480
      Width           =   2415
   End
   Begin VB.ComboBox PropertyLocationCitiesList 
      BackColor       =   &H80000004&
      Height          =   315
      ItemData        =   "Add_Tenant.frx":0000
      Left            =   2520
      List            =   "Add_Tenant.frx":009A
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   6480
      Width           =   2295
   End
   Begin VB.TextBox Tenant_RentalPrice 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2520
      TabIndex        =   27
      Top             =   7560
      Width           =   4695
   End
   Begin VB.ComboBox Tenant_NumberOfBedsLookingForList 
      BackColor       =   &H80000004&
      Height          =   315
      ItemData        =   "Add_Tenant.frx":02BF
      Left            =   2520
      List            =   "Add_Tenant.frx":02DB
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   7080
      Width           =   2295
   End
   Begin VB.TextBox Tenant_EmailAddress 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2520
      TabIndex        =   22
      Top             =   5640
      Width           =   4695
   End
   Begin VB.TextBox Tenant_FName 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   360
      Width           =   4695
   End
   Begin VB.TextBox Tenant_LName 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   960
      Width           =   4695
   End
   Begin VB.TextBox Tenant_Address2 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   1560
      Width           =   4695
   End
   Begin VB.TextBox Tenant_Address1 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   2160
      Width           =   4695
   End
   Begin VB.TextBox Tenant_PostCode 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   2760
      Width           =   4695
   End
   Begin VB.TextBox Tenant_State 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   3840
      Width           =   4695
   End
   Begin VB.TextBox Tenant_City 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   4440
      Width           =   4695
   End
   Begin VB.TextBox Tenant_PhoneNo 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   5040
      Width           =   4695
   End
   Begin VB.ComboBox Country_list_property 
      BackColor       =   &H80000004&
      Height          =   315
      ItemData        =   "Add_Tenant.frx":02F7
      Left            =   2520
      List            =   "Add_Tenant.frx":0547
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton Register_Tenant 
      Caption         =   "Register Tenant"
      Height          =   855
      Left            =   7440
      TabIndex        =   2
      Top             =   8400
      Width           =   1815
   End
   Begin VB.TextBox TenantRefID 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   8160
      Width           =   1455
   End
   Begin VB.TextBox Tenant_Country 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   4800
      TabIndex        =   0
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9360
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Location of Property:"
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
      Height          =   615
      Left            =   600
      TabIndex        =   29
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Rental  Price:"
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
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Number Of Beds:"
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
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label3 
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
      TabIndex        =   23
      Top             =   5760
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
      TabIndex        =   21
      Top             =   480
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
      TabIndex        =   20
      Top             =   1080
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
      TabIndex        =   19
      Top             =   1680
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
      TabIndex        =   18
      Top             =   2280
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
      TabIndex        =   17
      Top             =   2880
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
      TabIndex        =   16
      Top             =   3960
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
      Top             =   4560
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
      TabIndex        =   14
      Top             =   5160
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
      TabIndex        =   13
      Top             =   3360
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
      TabIndex        =   12
      Top             =   8280
      Width           =   2175
   End
End
Attribute VB_Name = "Add_Tenant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Tenant_NumberOfBedsLookingForList_Click()
Tenant_NumberOfBedsLookingFor.Text = Tenant_NumberOfBedsLookingForList
End Sub
Private Sub Country_list_property_Click()
Tenant_Country = Country_list_property
End Sub
Private Sub PropertyLocationCitiesList_click()
Tenant_PropertyLocation.Text = PropertyLocationCitiesList
End Sub
Private Sub GenerateRandomTenantRefID()
TenantRefID.Text = Int(Rnd * 996) + 1
End Sub
Private Sub CheckTenantRef()
GenerateRandomTenantRefID
Dim ViewTenantRefID As Tenant_Record
Dim ViewTenantRefIDChannelz As Integer
Dim x As Integer
x = 1
ViewTenantRefIDChannelz = FreeFile
Open Tenant_File For Random As ViewTenantRefIDChannelz Len = Tenant_Length
Get ViewTenantRefIDChannelz, x, ViewTenantRefID
    Do While Not EOF(ViewTenantRefIDChannelz)
        If Trim(ViewTenantRefID.Tenant_REFID) = Trim(TenantRefID.Text) Then
        GenerateRandomTenantRefID
        Tenant_File_Pointer = x
        End If
        x = x + 1
Get ViewTenantRefIDhannelz, x, ViewTenantRefID
    Loop
Close ViewTenantRefIDChannelz
End Sub
Private Sub Register_Tenant_Click()
Dim Add_Tenant As Tenant_Record
Dim TenantChannel As Integer
CheckTenantRef
TenantChannel = FreeFile
Open Tenant_File For Random As TenantChannel Len = Tenant_Length
            Add_Tenant.Tenant_FName = Tenant_FName.Text
            Add_Tenant.Tenant_LName = Tenant_LName.Text
            Add_Tenant.Tenant_PropetyLocation = Tenant_PropertyLocation.Text
            Add_Tenant.Tenant_Address1 = Tenant_Address1.Text
            Add_Tenant.Tenant_Address2 = Tenant_Address2.Text
            Add_Tenant.Tenant_PostCode = Tenant_PostCode.Text
            Add_Tenant.Tenant_CountryList = Tenant_Country.Text
            Add_Tenant.Tenant_State = Tenant_State.Text
            Add_Tenant.Tenant_City = Tenant_City.Text
            Add_Tenant.Tenant_RentalPrice = Tenant_RentalPrice.Text
            Add_Tenant.Tenant_NumberOfBeds = Tenant_NumberOfBedsLookingFor.Text
            Add_Tenant.Tenant_PhoneNumber = Tenant_PhoneNo.Text
            Add_Tenant.Tenant_EmailAddress = Tenant_EmailAddress.Text
            Add_Tenant.Tenant_REFID = TenantRefID.Text
            MsgBox TenantRefID.Text
            Tenant_File_Pointer = Tenant_File_Pointer + 1
    Put TenantChannel, Tenant_File_Pointer, Add_Tenant
    Close TenantChannel
End Sub



