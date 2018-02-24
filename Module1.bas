Attribute VB_Name = "AColoursPropertyModule"
Option Explicit
Type Landlord_Record
        Landlord_REFID As String * 3
        Tenant_REFID As String * 3
        Landlord_Fname As String * 15
        Landlord_LName As String * 15
        Landlord_CName As String * 25
        Landlord_Address1 As String * 30
        Landlord_Address2 As String * 30
        Landlord_CountryList As String * 45
        Landlord_PostCode As String * 10
        Landlord_City As String * 30
        Landlord_State As String * 45
        Landlord_PhoneNumber As String * 15
        Landlord_EmailAddress As String * 30
End Type
Global Landlord_File_Pointer As Integer
Global Landlord_File As String
Global Landlord_Length As Integer

Type Agent_Record
        Agent_REFID As String * 3
        Tenant_REFID As String * 3
        Agent_FName As String * 15
        Agent_LName As String * 15
        Agent_CName As String * 25
        Agent_Address1 As String * 30
        Agent_Address2 As String * 30
        Agent_CountryList As String * 45
        Agent_PostCode As String * 10
        Agent_City As String * 30
        Agent_State As String * 45
        Agent_PhoneNumber As String * 15
        Agent_EmailAddress As String * 30
End Type
Global Agent_File_Pointer As Integer
Global Agent_File As String
Global Agent_Length As Integer

Type Tenant_Record
        Tenant_REFID As String * 20
        Tenant_Fname As String * 15
        Tenant_LName As String * 15
        Tenant_Address1 As String * 30
        Tenant_Address2 As String * 30
        Tenant_CountryList As String * 45
        Tenant_PostCode As String * 10
        Tenant_City As String * 30
        Tenant_State As String * 45
        Tenant_PhoneNumber As String * 15
        Tenant_EmailAddress As String * 30
        Tenant_PropetyLocation As String * 35
        Tenant_NumberOfBeds As Integer
        Tenant_RentalPrice As Currency
        Tenant_Photo As String * 255
End Type
Global Tenant_File_PointerR As Integer
Global Tenant_File As String
Global Tenant_Length As Integer

Type Property_Record
        Property_RefID As String * 8
        Agent_REFID As String * 25
        Landlord_REFID As String * 25
        Address_line_1 As String * 35
        Address_line_2 As String * 35
        City As String * 40
        Post_Code As String * 6
        property_type As String * 34
        Payment_type As String * 15
        NumberOfBeds As Integer
        Price_Property As Currency
        RentalType As String * 9
        Image_property As String * 200
End Type
Global Add_Property_Pointer As Integer
Global Add_Property_File As String
Global Add_Property_Length As Integer

Type TenantProperty_Record
    Property_RefID As String * 20
    Tenant_REFID As String * 20
    
    StartDate As Date
    PaymentDueDate As Date
    Add7days As Date
    UpdateDateDaily As Date
    PaymentMade As Boolean
    LatestPayment As Date
    TenantEndDate As Date
    TotalRentalPricePayed As Currency
    OverDue01 As Date
    OverDue02 As Date
    OverDue03 As Date
    PaymentPriceOverDue As Currency
    End Type
Global TenantProperty_File_Pointer As Integer
Global TenantProperty_File As String
Global TenantProperty_Length As Integer
'/////////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\'
'----------------------------------------ADMIN ACCOUNT--------------------------------------------'
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\////////////////////////////////////////////////////'
Type Register_Account_Admin
        ID As Integer
        email_address As String * 30
        Password As String * 40
        First_name As String * 25
        Last_name As String * 25
        State As String * 30
        Country As String * 45
        City As String * 40
        Address As String * 50
        Telephone As String * 20
End Type
Global Register_account_admin_pointer As Integer
Global Register_account_admin_file As String
Global Register_account_admin_length As Integer
Type Login_Account_Admin
        ID As Integer
        email_address As String * 30
        Password As String * 40
End Type
Global Login_account_admin_pointer As Integer
Global Login_account_admin_file As String
Global Login_account_admin_length As Integer
'/////////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\'
'-------------------------------------------INCLUDE-----------------------------------------------'
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\////////////////////////////////////////////////////'
Global Selected_EmailAddress As String
Global Selected_LandlordID As String
Global Selected_AgentID As String
Global Selected_TenantID As String
Global Where_Record_Was As Integer
