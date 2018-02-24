VERSION 5.00
Begin VB.Form View_Tenant 
   Caption         =   "Form1"
   ClientHeight    =   9765
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   ScaleHeight     =   9765
   ScaleWidth      =   14100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Search All"
      Height          =   495
      Left            =   9720
      TabIndex        =   26
      Top             =   360
      Width           =   1455
   End
   Begin VB.ListBox TenantList_FName 
      Height          =   3570
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox LandlordRefID 
      Height          =   495
      Left            =   3720
      TabIndex        =   12
      Top             =   360
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   495
      Left            =   8280
      TabIndex        =   11
      Top             =   360
      Width           =   1455
   End
   Begin VB.ListBox TenantList_LName 
      Height          =   3570
      Left            =   2520
      TabIndex        =   10
      Top             =   1680
      Width           =   2175
   End
   Begin VB.ListBox TenantList_locationofproperty 
      Height          =   3570
      Left            =   4800
      TabIndex        =   9
      Top             =   1680
      Width           =   2175
   End
   Begin VB.ListBox TenantList_Address2 
      Height          =   3570
      Left            =   2520
      TabIndex        =   8
      Top             =   5880
      Width           =   2175
   End
   Begin VB.ListBox TenantList_PostCode 
      Height          =   3570
      Left            =   4800
      TabIndex        =   7
      Top             =   5880
      Width           =   2175
   End
   Begin VB.ListBox TenantList_Country 
      Height          =   3570
      Left            =   7080
      TabIndex        =   6
      Top             =   5880
      Width           =   2175
   End
   Begin VB.ListBox TenantList_State 
      Height          =   3570
      Left            =   9360
      TabIndex        =   5
      Top             =   5880
      Width           =   2175
   End
   Begin VB.ListBox TenantList_City 
      Height          =   3570
      Left            =   11640
      TabIndex        =   4
      Top             =   5880
      Width           =   2175
   End
   Begin VB.ListBox TenantList_Numberofbeds 
      Height          =   3570
      Left            =   7080
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
   Begin VB.ListBox TenantList_Address1 
      Height          =   3570
      Left            =   240
      TabIndex        =   2
      Top             =   5880
      Width           =   2175
   End
   Begin VB.ListBox TenantList_PhoneNo 
      Height          =   3570
      Left            =   9360
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.ListBox TenantList_EmailAddress 
      Height          =   3570
      Left            =   11640
      TabIndex        =   0
      Top             =   1680
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
      Left            =   240
      TabIndex        =   25
      Top             =   1320
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
      Left            =   2520
      TabIndex        =   24
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label LabePropertylocation 
      BackStyle       =   0  'Transparent
      Caption         =   "Location of Property:"
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
      Left            =   4800
      TabIndex        =   23
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Labelnumberofbeds 
      BackStyle       =   0  'Transparent
      Caption         =   "Number Of Beds:"
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
      Left            =   7080
      TabIndex        =   22
      Top             =   1320
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
      Left            =   9360
      TabIndex        =   21
      Top             =   1320
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
      Left            =   11640
      TabIndex        =   20
      Top             =   1320
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
      Left            =   240
      TabIndex        =   19
      Top             =   5520
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
      Left            =   2520
      TabIndex        =   18
      Top             =   5520
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
      Left            =   4800
      TabIndex        =   17
      Top             =   5520
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
      Left            =   7080
      TabIndex        =   16
      Top             =   5520
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
      Left            =   9360
      TabIndex        =   15
      Top             =   5520
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
      Left            =   11640
      TabIndex        =   14
      Top             =   5520
      Width           =   2175
   End
End
Attribute VB_Name = "View_Tenant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
ClearListBoxesForTenant
Dim ViewTenant As Tenant_Record
Dim ViewTenantChannel As Integer
Dim x As Integer
Dim OnceFoundTenant As Boolean
x = 1
ViewTenantChannel = FreeFile
Open Tenant_File For Random As ViewTenantChannel Len = Tenant_Length
Get ViewTenantChannel, x, ViewTenant
Do While Not EOF(ViewTenantChannel) And OnceFoundTenant = False
    If Trim(ViewTenant.Tenant_RefID) = Trim(TenantRefID.Text) Then
    TenantList_FName.AddItem ViewTenant.Agent_FName
    TenantList_LName.AddItem ViewTenant.Agent_LName
    TenantList_TenantList_locationofproperty.AddItem ViewTenant.Tenant_PropetyLocation
    TenantList_Address1.AddItem ViewTenant.Tenant_Address1
    TenantList_Address2.AddItem ViewTenant.Tenant_Address2
    TenantList_PostCode.AddItem ViewTenant.Tenant_PostCode
    TenantList_Country.AddItem ViewTenant.Tenant_CountryList
    TenantList_State.AddItem ViewTenant.Tenant_State
    TenantList_City.AddItem ViewTenant.Tenant_City
    TenantList_FaxNo.AddItem ViewTenant.Tenant_NumberOfBeds
    TenantList_PhoneNo.AddItem ViewTenant.Tenant_PhoneNumber
    TenantList_EmailAddress.AddItem ViewTenant.Tenant_EmailAddress
    Tenant_File_Pointer = x
    OnceFoundTenant = True
    Else
    MsgBox "This Landlord does not exist"
    x = x + 1
    Get ViewTenantChannel, x, ViewTenant
    End If
Loop
Close ViewTenantChannel
End Sub

Private Sub ClearListBoxesForTenant()

End Sub

Private Sub Command2_Click()
ClearListBoxesForTenant
Dim ViewTenant As Tenant_Record
Dim ViewTenantChannel As Integer
Dim x As Integer
x = 1
ViewTenantChannel = FreeFile
Open Tenant_File For Random As ViewTenantChannel Len = Tenant_Length
Get ViewTenantChannel, x, ViewTenant
Do While Not EOF(ViewTenantChannel)
    TenantList_FName.AddItem ViewTenant.Tenant_FName
    TenantList_LName.AddItem ViewTenant.Tenant_LName
    TenantList_locationofproperty.AddItem ViewTenant.Tenant_PropetyLocation
    TenantList_Address1.AddItem ViewTenant.Tenant_Address1
    TenantList_Address2.AddItem ViewTenant.Tenant_Address2
    TenantList_PostCode.AddItem ViewTenant.Tenant_PostCode
    TenantList_Country.AddItem ViewTenant.Tenant_CountryList
    TenantList_State.AddItem ViewTenant.Tenant_State
    TenantList_City.AddItem ViewTenant.Tenant_City
    TenantList_Numberofbeds.AddItem ViewTenant.Tenant_NumberOfBeds
    TenantList_PhoneNo.AddItem ViewTenant.Tenant_PhoneNumber
    TenantList_EmailAddress.AddItem ViewTenant.Tenant_EmailAddress
    x = x + 1
    Get ViewTenantChannel, x, ViewTenant
Loop
Close ViewTenantChannel
End Sub

