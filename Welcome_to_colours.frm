VERSION 5.00
Begin VB.Form Welcome_to_colours 
   Caption         =   "Welcome To Colours Property"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   Picture         =   "Welcome_to_colours.frx":0000
   ScaleHeight     =   6570
   ScaleWidth      =   13230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   11280
      TabIndex        =   16
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   255
      Left            =   11280
      TabIndex        =   15
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   11160
      TabIndex        =   14
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   11160
      TabIndex        =   13
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Command11"
      Height          =   555
      Left            =   11160
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   615
      Left            =   11160
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   615
      Left            =   11280
      TabIndex        =   10
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   615
      Left            =   11280
      TabIndex        =   9
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Timer SetTimerForSignUpButton 
      Interval        =   1500
      Left            =   4320
      Top             =   4560
   End
   Begin VB.Timer SetTimerForLoginButton 
      Interval        =   1500
      Left            =   9480
      Top             =   4560
   End
   Begin VB.PictureBox Register_button 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4800
      Picture         =   "Welcome_to_colours.frx":27B9D
      ScaleHeight     =   495
      ScaleWidth      =   2175
      TabIndex        =   7
      Top             =   4560
      Width           =   2175
   End
   Begin VB.PictureBox Login_button 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7200
      Picture         =   "Welcome_to_colours.frx":2B847
      ScaleHeight     =   495
      ScaleWidth      =   2175
      TabIndex        =   6
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox Password_main 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   4800
      MaxLength       =   35
      TabIndex        =   4
      Top             =   3360
      Width           =   4455
   End
   Begin VB.TextBox Email_address_main 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   4800
      MaxLength       =   35
      TabIndex        =   2
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label password_forgot 
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot your password?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label password_state 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label email_state 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Login_state 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your email and password to sign in"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   1920
      Width           =   5295
   End
   Begin VB.Label welcome_state 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to our login center "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   0
      Top             =   1560
      Width           =   8895
   End
End
Attribute VB_Name = "Welcome_to_colours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
Tenants.Show 1
End Sub

Private Sub Command10_Click()
Landlords.Show 1
End Sub

Private Sub Command11_Click()
Agent.Show 1
End Sub

Private Sub Command2_Click()
Properties.Show 1
End Sub

Private Sub Command5_Click()
Add_Tenant.Show 1
End Sub

Private Sub Command6_Click()
View_Tenant.Show 1
End Sub



Private Sub Command7_Click()

End Sub

Private Sub Command3_Click()
RemoveTenantFromProperty.Show 1
End Sub

Private Sub Command8_Click()
Dashboard.Show 1
End Sub

Private Sub Command9_Click()
Add_Tenant_to_Property.Show 1
End Sub

Private Sub Email_address_main_Change()
Selected_EmailAddress = Email_address_main.Text
End Sub

Private Sub Form_Activate()
Email_address_main.Text = Selected_EmailAddress
End Sub

Private Sub Form_Load()
Dim RegisterAccountAdmin As Register_Account_Admin
Dim RegisterAccountAdminchannel As Integer
        Register_account_admin_file = App.Path + "\Saved_Dat\Register_AdminAccount.dat"
        Register_account_admin_length = Len(RegisterAccountAdmin)
        RegisterAccountAdminchannel = FreeFile
        Open Register_account_admin_file For Random As RegisterAccountAdminchannel Len = Register_account_admin_length
        Register_account_admin_pointer = FileLen(Register_account_admin_file) / Register_account_admin_length
        Close RegisterAccountAdminchannel
        
Dim PropertyRecord As Property_Record
Dim PropertyChannel As Integer
        Add_Property_File = App.Path + "\Saved_Dat\add_property.dat"
        Add_Property_Length = Len(PropertyRecord)
        PropertyChannel = FreeFile
        Open Add_Property_File For Random As PropertyChannel Len = Add_Property_Length
        Add_Property_Pointer = FileLen(Add_Property_File) / Add_Property_Length
        Close PropertyChannel
        
Dim AddLandlord As Landlord_Record
Dim LandLordchannel As Integer
        Landlord_File = App.Path + "\Saved_Dat\Add_Landlords.dat"
        Landlord_Length = Len(AddLandlord)
        LandLordchannel = FreeFile
        Open Landlord_File For Random As LandLordchannel Len = Landlord_Length
        Landlord_File_Pointer = FileLen(Landlord_File) / Landlord_Length
        Close LandLordchannel
        
Dim AddAgent As Agent_Record
Dim Agentchannel As Integer
        Agent_File = App.Path + "\Saved_Dat\Add_Agents.dat"
        Agent_Length = Len(AddAgent)
        Agentchannel = FreeFile
        Open Agent_File For Random As Agentchannel Len = Agent_Length
        Agent_File_Pointer = FileLen(Agent_File) / Agent_Length
        Close Agentchannel
        
Dim AddTenant As Tenant_Record
Dim TenantChannel As Integer
        Tenant_File = App.Path + "\Saved_Dat\Add_Tenants.dat"
        Tenant_Length = Len(AddTenant)
        TenantChannel = FreeFile
        Open Tenant_File For Random As Agentchannel Len = Tenant_Length
        Tenant_File_PointerR = FileLen(Tenant_File) / Tenant_Length
        Close TenantChannel

Dim TenantProperty As TenantProperty_Record
Dim TenantPropertyChannel As Integer
    TenantProperty_File = App.Path + "\saved_dat\TenantProperty.dat"
    TenantProperty_Length = Len(TenantProperty)
    TenantPropertyChannel = FreeFile
    Open TenantProperty_File For Random As TenantPropertyChannel Len = TenantProperty_Length
    TenantProperty_File_Pointer = FileLen(TenantProperty_File) / TenantProperty_Length
    Close TenantPropertyChannel
End Sub

Private Sub Login_button_Click()
'Login_button.Picture = LoadPicture("Images\" & Login_button.Tag & "login_button_rollover.jpg")
Dim LoginAccountAdmin As Register_Account_Admin
Dim LoginAdminchannel As Integer
Dim x As Integer
Dim WhenLoginDone As Boolean
WhenLoginDone = False
x = 1
LoginAdminchannel = FreeFile
Open Register_account_admin_file For Random As LoginAdminchannel Len = Register_account_admin_length
Get LoginAdminchannel, x, LoginAccountAdmin
Do While Not EOF(LoginAdminchannel) And WhenLoginDone = False
        If Trim(LoginAccountAdmin.email_address) = Trim(Email_address_main.Text) And Trim(LoginAccountAdmin.Password) = Trim(Password_main.Text) Then
        Email_address_main.Text = LoginAccountAdmin.email_address
        Password_main.Text = LoginAccountAdmin.Password
        Register_account_admin_pointer = x
        WhenLoginDone = True
        Unload Me
        Dashboard.Show 1
        Else
        x = x + 1
        Get LoginAdminchannel, x, LoginAccountAdmin
        End If
Loop
        MsgBox ("Please enter the correct information")
Email_address_main = ""
Password_main = ""
End Sub



Private Sub Register_button_Click()
'Register_button.Picture = LoadPicture("Images\" & Register_button.Tag & "signup_hover_button.jpg")
Unload Me
Register_account.Show 1
End Sub

Private Sub SetTimerForLoginButton_Timer()
'Login_button.Picture = LoadPicture("Images\" & Login_button.Tag & "login_button.jpg")
End Sub

Private Sub SetTimerForSignUpButton_Timer()
'Register_button.Picture = LoadPicture("Images\" & Register_button.Tag & "signup_button.jpg")
End Sub



