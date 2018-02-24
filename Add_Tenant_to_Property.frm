VERSION 5.00
Begin VB.Form Add_Tenant_to_Property 
   Caption         =   "Add Tenant To Property"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   12825
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox StartDate 
      Height          =   495
      Left            =   8400
      TabIndex        =   54
      Top             =   5520
      Width           =   2535
   End
   Begin VB.TextBox Tenant_RentalPrice 
      Height          =   285
      Left            =   8400
      TabIndex        =   51
      Top             =   5040
      Width           =   2535
   End
   Begin VB.TextBox Tenant_NoofBeds 
      Height          =   285
      Left            =   8400
      TabIndex        =   50
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox Tenant_LocationOfProperty 
      Height          =   285
      Left            =   8400
      TabIndex        =   36
      Top             =   4320
      Width           =   2535
   End
   Begin VB.TextBox Tenant_EmailAddress 
      Height          =   285
      Left            =   8400
      TabIndex        =   35
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox Tenant_PhoneNumber 
      Height          =   285
      Left            =   8400
      TabIndex        =   34
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox Tenant_City 
      Height          =   285
      Left            =   8400
      TabIndex        =   33
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox Tenant_State 
      Height          =   285
      Left            =   8400
      TabIndex        =   32
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox Tenant_Country 
      Height          =   285
      Left            =   8400
      TabIndex        =   31
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox Tenant_PostCode 
      Height          =   285
      Left            =   8400
      TabIndex        =   30
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Tenant_Address2 
      Height          =   285
      Left            =   8400
      TabIndex        =   29
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox Tenant_Address1 
      Height          =   285
      Left            =   8400
      TabIndex        =   28
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox Tenant_LName 
      Height          =   285
      Left            =   8400
      TabIndex        =   27
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Tenant_FName 
      Height          =   285
      Left            =   8400
      TabIndex        =   26
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Property_AgentID 
      Height          =   285
      Left            =   2040
      TabIndex        =   14
      Top             =   4320
      Width           =   2535
   End
   Begin VB.TextBox Property_LandlordID 
      Height          =   285
      Left            =   2040
      TabIndex        =   13
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox Property_Image 
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox Property_NoofBeds 
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox Property_RentalPrice 
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox Property_PaymentType 
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox Property_PropertyType 
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Property_PostCode 
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox Property_City 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox Property_Address2 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Property_Address1 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   720
      Width           =   2535
   End
   Begin VB.ListBox Tenant_Property_Detail 
      Height          =   450
      Left            =   240
      TabIndex        =   3
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton Regs 
      Caption         =   "Register Tenant To Property"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   5640
      Width           =   2175
   End
   Begin VB.ListBox TenantID 
      Height          =   4740
      Left            =   6480
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.ListBox PropertyID 
      Height          =   3960
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   6360
      X2              =   6360
      Y1              =   0
      Y2              =   6960
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11040
      TabIndex        =   55
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Tenant"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   53
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Property "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   52
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address"
      Height          =   375
      Left            =   11040
      TabIndex        =   49
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   375
      Left            =   11040
      TabIndex        =   48
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
      Height          =   375
      Left            =   11040
      TabIndex        =   47
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      Height          =   375
      Left            =   11040
      TabIndex        =   46
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   375
      Left            =   11040
      TabIndex        =   45
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number"
      Height          =   375
      Left            =   11040
      TabIndex        =   44
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Number Of Beds"
      Height          =   375
      Left            =   11040
      TabIndex        =   43
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Rental Price"
      Height          =   375
      Left            =   11040
      TabIndex        =   42
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Location Of Propeperty"
      Height          =   375
      Left            =   11040
      TabIndex        =   41
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code"
      Height          =   375
      Left            =   11040
      TabIndex        =   40
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   375
      Left            =   11040
      TabIndex        =   39
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 2"
      Height          =   375
      Left            =   11040
      TabIndex        =   38
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 1"
      Height          =   375
      Left            =   11040
      TabIndex        =   37
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent Reference ID"
      Height          =   375
      Left            =   4680
      TabIndex        =   25
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Landlord Reference ID"
      Height          =   375
      Left            =   4680
      TabIndex        =   24
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Image"
      Height          =   375
      Left            =   4680
      TabIndex        =   23
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Number Of Beds"
      Height          =   375
      Left            =   4680
      TabIndex        =   22
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Rental Price"
      Height          =   375
      Left            =   4680
      TabIndex        =   21
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Type"
      Height          =   375
      Left            =   4680
      TabIndex        =   20
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Property Type"
      Height          =   375
      Left            =   4680
      TabIndex        =   19
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code"
      Height          =   375
      Left            =   4680
      TabIndex        =   18
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   375
      Left            =   4680
      TabIndex        =   17
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 2"
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 1"
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "Add_Tenant_to_Property"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
LoadTenants
LoadProperties
End Sub

Private Sub LoadTenants()
Dim ViewTenant As Tenant_Record
Dim ViewTenantChannel As Integer
Dim x As Integer
x = 1
ViewTenantChannel = FreeFile
Open Tenant_File For Random As ViewTenantChannel Len = Tenant_Length
Get ViewTenantChannel, x, ViewTenant
Do While Not EOF(ViewTenantChannel)
    TenantID.AddItem ViewTenant.Tenant_REFID
    x = x + 1
    Get ViewTenantChannel, x, ViewTenant
Loop
Close ViewTenantChannel
End Sub

Private Sub LoadProperties()
Dim ViewProperties As Property_Record
Dim ViewPropertyChannel As Integer
Dim x As Integer
x = 1
ViewPropertyChannel = FreeFile
    Open Add_Property_File For Random As ViewPropertyChannel Len = Add_Property_Length
    Get ViewPropertyChannel, x, ViewProperties
Do While Not EOF(ViewPropertyChannel)
    PropertyID.AddItem ViewProperties.Property_RefID
    x = x + 1
    Get ViewPropertyChannel, x, ViewProperties
Loop
Close ViewPropertyChannel
End Sub

Private Sub PropertyID_Click()
Dim ViewProperties As Property_Record
Dim ViewPropertyChannel As Integer
Dim x As Integer
Dim OnceFoundProperty As Boolean
x = 1
ViewPropertyChannel = FreeFile
Open Add_Property_File For Random As ViewPropertyChannel Len = Add_Property_Length
Get ViewPropertyChannel, x, ViewProperties
Do While Not EOF(ViewPropertyChannel) And OnceFoundProperty = False
    If Trim(PropertyID.List(PropertyID.ListIndex)) = Trim(ViewProperties.Property_RefID) Then
    Property_Address1.Text = ViewProperties.Address_line_1
    Property_Address2.Text = ViewProperties.Address_line_2
    Property_City.Text = ViewProperties.City
    Property_PostCode.Text = ViewProperties.Post_Code
    Property_PropertyType = ViewProperties.property_type
    Property_PaymentType.Text = ViewProperties.Payment_type
    Property_NoofBeds.Text = ViewProperties.NumberOfBeds
    Property_RentalPrice.Text = ViewProperties.Price_Property
    Property_LandlordID.Text = ViewProperties.Landlord_REFID
    Property_AgentID.Text = ViewProperties.Agent_REFID
    Add_Property_Pointer = x
    OnceFoundProperty = True
    End If
    x = x + 1
    Get ViewPropertyChannel, x, ViewProperties
Loop
Close ViewPropertyChannel
End Sub

Private Sub Regs_Click()
Dim AddTenantToProperty As TenantProperty_Record
Dim TenantPropertyChannel As Integer
Dim ViewTenantProperty As TenantProperty_Record
Dim ViewTenantPropertyChannel As Integer
Dim OnceFoundTenantProperty As Boolean
Dim x As Integer
ChangeNumberOfBed
StartDate.Text = Format$(Now + 7, "short Date")
Tenant_Property_Detail.Clear
TenantPropertyChannel = FreeFile
Open TenantProperty_File For Random As TenantPropertyChannel Len = TenantProperty_Length
            AddTenantToProperty.Property_RefID = PropertyID.List(PropertyID.ListIndex)
            AddTenantToProperty.Tenant_REFID = TenantID.List(TenantID.ListIndex)
            AddTenantToProperty.StartDate = StartDate.Text
            AddTenantToProperty.PaymentDueDate = StartDate.Text
            AddTenantToProperty.Add7days = StartDate.Text
            MsgBox ("The Start Date For The Tenant's Property Is: " & StartDate.Text)
            TenantProperty_File_Pointer = TenantProperty_File_Pointer + 1
Put TenantPropertyChannel, TenantProperty_File_Pointer, AddTenantToProperty
Close TenantPropertyChannel
x = 1
ViewTenantPropertyChannel = FreeFile
Open TenantProperty_File For Random As ViewTenantPropertyChannel Len = TenantProperty_Length
Get ViewTenantPropertyChannel, x, ViewTenantProperty
Do While Not EOF(ViewTenantPropertyChannel) And OnceFoundTenantProperty = False
    If Trim(ViewTenantProperty.Tenant_REFID) = Trim(TenantID.List(TenantID.ListIndex)) And Trim(ViewTenantProperty.Property_RefID) = Trim(PropertyID.List(PropertyID.ListIndex)) Then
    Tenant_Property_Detail.AddItem ViewTenantProperty.Property_RefID
    Tenant_Property_Detail.AddItem ViewTenantProperty.Tenant_REFID
    TenantProperty_File_Pointer = x
    OnceFoundTenantProperty = True
    End If
    x = x + 1
    Get ViewTenantPropertyChannel, x, ViewTenantProperty
Loop
Close ViewTenantPropertyChannel
End Sub

Private Sub Tenant_Property_Detail_Click()
Dim ViewTenantP As Tenant_Record, Property_Record
Dim ViewTenantPChannel As Integer
Dim x As Integer
Dim OnceFoundTenantP As Boolean
x = 1
ViewTenantPChannel = FreeFile
Open (TenantProperty_File And Tenant_File And Add_Property_File) For Random As ViewTenantPChannel Len = (TenantProperty_Length And Tenant_Length And Add_Property_Length)
Get ViewTenantPChannel, x, ViewTenantP
Do While Not EOF(ViewTenantPChannel) And OnceFoundTenantP = False
    If Trim(TenantID.List(TenantID.ListIndex)) = Trim(Tenant_Property_Detail.List(Tenant_Property_Detail.ListIndex)) Then
    List1.AddItem ViewTenantP.Tenant_City
    List1.AddItem ViewTenantP.Tenant_Address1
    List1.AddItem ViewTenantP.Tenant_Address2
    TenantProperty_File_Pointer = x
    OnceFoundTenantP = True
    End If
    x = x + 1
    Get ViewTenantPChannel, x, ViewTenantP
Loop
Close ViewTenantPChannel
End Sub

Private Sub TenantID_Click()
Dim ViewTenant As Tenant_Record
Dim ViewTenantChannel As Integer
Dim x As Integer
Dim OnceFoundTenant As Boolean
x = 1
ViewTenantChannel = FreeFile
Open Tenant_File For Random As ViewTenantChannel Len = Tenant_Length
Get ViewTenantChannel, x, ViewTenant
Do While Not EOF(ViewTenantChannel) And OnceFoundTenant = False
    If Trim(TenantID.List(TenantID.ListIndex)) = Trim(ViewTenant.Tenant_REFID) Then
    Tenant_Fname.Text = ViewTenant.Tenant_Fname
    Tenant_LName.Text = ViewTenant.Tenant_LName
    Tenant_LocationOfProperty.Text = ViewTenant.Tenant_PropetyLocation
    Tenant_Address1.Text = ViewTenant.Tenant_Address1
    Tenant_Address2.Text = ViewTenant.Tenant_Address2
    Tenant_PostCode.Text = ViewTenant.Tenant_PostCode
    Tenant_RentalPrice.Text = ViewTenant.Tenant_RentalPrice
    Tenant_Country.Text = ViewTenant.Tenant_CountryList
    Tenant_State.Text = ViewTenant.Tenant_State
    Tenant_City.Text = ViewTenant.Tenant_City
    Tenant_NoofBeds.Text = ViewTenant.Tenant_NumberOfBeds
    Tenant_PhoneNumber.Text = ViewTenant.Tenant_PhoneNumber
    Tenant_EmailAddress.Text = ViewTenant.Tenant_EmailAddress
    Tenant_File_PointerR = x
    OnceFoundTenant = True
    End If
    x = x + 1
    Get ViewTenantChannel, x, ViewTenant
Loop
Close ViewTenantChannel
End Sub

Private Function ClearTexts()
Property_Address1 = ""
Property_Address2 = ""
Property_City = ""
Property_PostCode = ""
Property_PropertyType = ""
Property_PaymentType = ""
Property_NoofBeds = ""
Property_RentalPrice = ""
Property_LandlordID = ""
Property_AgentID = ""
Tenant_Fname = ""
Tenant_LName = ""
Tenant_Address1 = ""
Tenant_Address2 = ""
Tenant_PostCode = ""
Tenant_RentalPrice = ""
Tenant_Country = ""
Tenant_State = ""
Tenant_City = ""
Tenant_NoofBeds = ""
Tenant_PhoneNumber = ""
Tenant_EmailAddress = ""
Tenant_LocationOfProperty = ""
StartDate.Text = ""
End Function

Private Function ChangeNumberOfBed()
Dim ChangeNumberOfBeds As Property_Record
Dim ChangeNumberofBedsChannel As Integer
Dim AmendDone As Boolean
Dim Y As Integer
If Property_NoofBeds.Text <= 0 Then
MsgBox "There is no bedrooms left for this property, please select another property"
Else
Y = 1
ChangeNumberofBedsChannel = FreeFile
    Open Add_Property_File For Random As ChangeNumberofBedsChannel Len = Add_Property_Length
        Get ChangeNumberofBedsChannel, Y, ChangeNumberOfBeds
Do While Not EOF(ChangeNumberofBedsChannel) And AmendDone = False
         If Trim(ChangeNumberOfBeds.NumberOfBeds) = Trim(Property_NoofBeds.Text) Then
            Property_NoofBeds.Text = Property_NoofBeds.Text - 1
            ChangeNumberOfBeds.NumberOfBeds = Property_NoofBeds.Text
            Add_Property_Pointer = Y
            AmendDone = True
        Put ChangeNumberofBedsChannel, Add_Property_Pointer, ChangeNumberOfBeds
        End If
        Y = Y + 1
        Get ChangeNumberofBedsChannel, Y, ChangeNumberOfBeds
        
Loop
Close ChangeNumberofBedsChannel
End If
End Function
