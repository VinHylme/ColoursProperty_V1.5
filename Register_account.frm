VERSION 5.00
Begin VB.Form Register_account 
   Caption         =   "Register Account"
   ClientHeight    =   10260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   Picture         =   "Register_account.frx":0000
   ScaleHeight     =   10260
   ScaleWidth      =   13155
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Statee 
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
      Left            =   5400
      MaxLength       =   35
      TabIndex        =   31
      Top             =   4800
      Width           =   4455
   End
   Begin VB.PictureBox email_enter 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   9960
      Picture         =   "Register_account.frx":27B9D
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   30
      Top             =   7680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox password_enter 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   9960
      Picture         =   "Register_account.frx":2CB68
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   29
      Top             =   8400
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox telephone_enter 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   9960
      Picture         =   "Register_account.frx":31B33
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   28
      Top             =   6960
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox city_enter 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   9960
      Picture         =   "Register_account.frx":36AFE
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   27
      Top             =   6240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox address_enter 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   9960
      Picture         =   "Register_account.frx":3BAC9
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   26
      Top             =   5520
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   9960
      Picture         =   "Register_account.frx":40A94
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   24
      Top             =   4800
      Visible         =   0   'False
      Width           =   3015
      Begin VB.PictureBox Statee_enter 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   0
         Picture         =   "Register_account.frx":45A5F
         ScaleHeight     =   615
         ScaleWidth      =   3015
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin VB.PictureBox Country_enter 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   9960
      Picture         =   "Register_account.frx":4AA2A
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   23
      Top             =   4080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox Lastname_enter 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   9960
      Picture         =   "Register_account.frx":4F8AE
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   22
      Top             =   3360
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox Firstname_enter 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   9960
      Picture         =   "Register_account.frx":54879
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   21
      Top             =   2640
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox Email_exists_sticker 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   9960
      Picture         =   "Register_account.frx":599CB
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   20
      Top             =   7680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Timer SetTimerForRegisterbutton 
      Interval        =   1500
      Left            =   9840
      Top             =   9240
   End
   Begin VB.ComboBox Country_list 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      ItemData        =   "Register_account.frx":6048A
      Left            =   5400
      List            =   "Register_account.frx":606DA
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   4200
      Width           =   4455
   End
   Begin VB.PictureBox Register_complete 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7680
      Picture         =   "Register_account.frx":60F16
      ScaleHeight     =   495
      ScaleWidth      =   2175
      TabIndex        =   18
      Top             =   9240
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
      Left            =   5400
      MaxLength       =   35
      TabIndex        =   8
      Top             =   8400
      Width           =   4455
   End
   Begin VB.TextBox Email_address 
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
      Left            =   5400
      MaxLength       =   35
      TabIndex        =   7
      Top             =   7680
      Width           =   4455
   End
   Begin VB.TextBox Telephone_number 
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
      Left            =   5400
      MaxLength       =   35
      TabIndex        =   6
      Top             =   6960
      Width           =   4455
   End
   Begin VB.TextBox City_main 
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
      Left            =   5400
      MaxLength       =   35
      TabIndex        =   5
      Top             =   6240
      Width           =   4455
   End
   Begin VB.TextBox Address_main 
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
      Left            =   5400
      MaxLength       =   35
      TabIndex        =   4
      Top             =   5520
      Width           =   4455
   End
   Begin VB.TextBox Last_name 
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
      Left            =   5400
      MaxLength       =   35
      TabIndex        =   3
      Top             =   3360
      Width           =   4455
   End
   Begin VB.TextBox First_name 
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
      Left            =   5400
      MaxLength       =   35
      TabIndex        =   2
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label8 
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
      Left            =   3480
      TabIndex        =   17
      Top             =   8520
      Width           =   2415
   End
   Begin VB.Label Label7 
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
      Left            =   3480
      TabIndex        =   16
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone:"
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
      Left            =   3480
      TabIndex        =   15
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
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
      Left            =   3480
      TabIndex        =   14
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
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
      Left            =   3480
      TabIndex        =   13
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
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
      Left            =   3480
      TabIndex        =   12
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Country:"
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
      Left            =   3480
      TabIndex        =   11
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
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
      Left            =   3480
      TabIndex        =   10
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label email_state 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
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
      Left            =   3480
      TabIndex        =   9
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label welcome_state 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to our Register center "
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
      Left            =   5160
      TabIndex        =   1
      Top             =   1680
      Width           =   8895
   End
   Begin VB.Label Register_state 
      BackStyle       =   0  'Transparent
      Caption         =   "Please complete the form bellow to register an account"
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
      Left            =   3960
      TabIndex        =   0
      Top             =   2040
      Width           =   6135
   End
End
Attribute VB_Name = "Register_account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Register_complete_Click()
'Register_complete.Picture = LoadPicture("Images\" & Register_complete.Tag & "register_button_hover.jpg")


Dim RegisterAdminAccount As Register_Account_Admin
Dim RegisterAdminAccountchannel As Integer
Dim Check_for_exist As Boolean
Dim Check_To_Continue As Boolean
Dim X As Integer
Check_for_exist = False
X = 1
RegisterAdminAccountchannel = FreeFile
Open Register_account_admin_file For Random As RegisterAdminAccountchannel Len = Register_account_admin_length
Get RegisterAdminAccountchannel, X, RegisterAdminAccount
Do While Not EOF(RegisterAdminAccountchannel) And Check_for_exist = False
    If Trim(RegisterAdminAccount.email_address) = Trim(email_address.Text) Then
        Check_for_exist = True
        Check_To_Continue = True
        Email_exists_sticker.Visible = True
    Else
        Email_exists_sticker.Visible = False
        Check_To_Continue = False
        Check_for_exist = False
        X = X + 1
        Get RegisterAdminAccountchannel, X, RegisterAdminAccount
    End If
Loop
If Check_To_Continue = False Then
        RegisterAdminAccountchannel = FreeFile
        Open Register_account_admin_file For Random As RegisterAdminAccountchannel Len = Register_account_admin_length
        RegisterAdminAccount.First_name = First_name.Text
        If First_name.Text = "" Then
            Check_for_exist = False
            Firstname_enter.Visible = True
        Else: Firstname_enter.Visible = False
            Check_for_exist = True
        End If
        RegisterAdminAccount.Last_name = Last_name.Text
        If Last_name.Text = "" Then
            Check_for_exist = False
            Lastname_enter.Visible = True
        Else: Lastname_enter.Visible = False
            Check_for_exist = True
        End If
        Country_list.AddItem RegisterAdminAccount.Country
        If Country_list = "" Then
            Check_for_exist = False
            Country_enter.Visible = True
        Else: Country_enter.Visible = False
            Check_for_exist = True
        End If
        RegisterAdminAccount.Address = Address_main.Text
        If Address_main.Text = "" Then
            Check_for_exist = False
            address_enter.Visible = True
        Else: address_enter.Visible = False
            Check_for_exist = True
        End If
        RegisterAdminAccount.City = City_main.Text
        If City_main.Text = "" Then
            Check_for_exist = False
            city_enter.Visible = True
        Else: city_enter.Visible = False
            Check_for_exist = True
        End If
        RegisterAdminAccount.Telephone = Telephone_number.Text
        If Telephone_number.Text = "" Then
            Check_for_exist = False
            telephone_enter.Visible = True
        Else: telephone_enter.Visible = False
            Check_for_exist = True
        End If
        RegisterAdminAccount.State = Statee.Text
        If Statee.Text = "" Then
            Check_for_exist = False
            Statee_enter.Visible = True
        Else: Statee_enter.Visible = False
            Check_for_exist = True
        End If
        RegisterAdminAccount.email_address = email_address.Text
        If email_address.Text = "" Then
            Check_for_exist = False
            email_enter.Visible = True
        Else: email_enter.Visible = False
            Check_for_exist = True
        End If
        RegisterAdminAccount.Password = Password_main.Text
        If Password_main.Text = "" Then
            Check_for_exist = False
            password_enter.Visible = True
        Else: password_enter.Visible = False
            Check_for_exist = True
        End If
    If Check_for_exist = True Then
        Register_account_admin_pointer = Register_account_admin_pointer + 1
        Put RegisterAdminAccountchannel, Register_account_admin_pointer, RegisterAdminAccount
        Close RegisterAdminAccountchannel
        MsgBox ("Admin Account Has Been successfully Created")
        Unload Me
        Welcome_to_colours.Show 1
    End If
End If
End Sub

Private Sub SetTimerForRegisterbutton_Timer()
'Register_complete.Picture = LoadPicture("Images\" & Register_complete.Tag & "register_button.jpg")
End Sub

Private Sub State_Change()

End Sub

