VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Tenants 
   Caption         =   "Manage Tenants"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   ScaleHeight     =   16507.87
   ScaleMode       =   0  'User
   ScaleWidth      =   13665
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9480
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton BrowseImage 
      Caption         =   "Browse "
      Height          =   375
      Left            =   8520
      TabIndex        =   72
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox Tenant_Image 
      Height          =   375
      Left            =   7200
      TabIndex        =   70
      Top             =   6000
      Width           =   1335
   End
   Begin VB.ComboBox LocationOfproperty 
      Height          =   315
      ItemData        =   "Tenants.frx":0000
      Left            =   7200
      List            =   "Tenants.frx":009A
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   69
      Top             =   4800
      Width           =   3255
   End
   Begin VB.ComboBox NumberOfBedslist 
      Height          =   315
      ItemData        =   "Tenants.frx":02BF
      Left            =   7200
      List            =   "Tenants.frx":02DE
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   68
      Top             =   4200
      Width           =   3255
   End
   Begin VB.CommandButton Add 
      Caption         =   "ADD"
      Height          =   615
      Left            =   10680
      TabIndex        =   65
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Delete 
      Caption         =   "DELETE"
      Height          =   615
      Left            =   12120
      TabIndex        =   64
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Prints 
      Caption         =   "PRINT"
      Height          =   615
      Left            =   10680
      TabIndex        =   63
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "CANCEL"
      Enabled         =   0   'False
      Height          =   615
      Left            =   12120
      TabIndex        =   62
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton EdIT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "EDIT"
      Height          =   615
      Left            =   12120
      TabIndex        =   61
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Save 
      Caption         =   "SAVE"
      Enabled         =   0   'False
      Height          =   615
      Left            =   10680
      TabIndex        =   60
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox Tenant_RentalPrice 
      Height          =   375
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   5400
      Width           =   3255
   End
   Begin VB.TextBox Tenant_Location 
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   4800
      Width           =   3255
   End
   Begin VB.TextBox Tenant_NB 
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   4200
      Width           =   3255
   End
   Begin VB.TextBox Tenant_Fname 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   4200
      Width           =   3255
   End
   Begin VB.TextBox Tenant_LName 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   4800
      Width           =   3255
   End
   Begin VB.TextBox Tenant_AddressLine1 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   5400
      Width           =   3255
   End
   Begin VB.TextBox Tenant_AddressLine2 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   6000
      Width           =   3255
   End
   Begin VB.TextBox Tenant_PostCode 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   6600
      Width           =   3255
   End
   Begin VB.TextBox Tenant_State 
      Height          =   375
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   4800
      Width           =   3255
   End
   Begin VB.TextBox Tenant_City 
      Height          =   375
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   5400
      Width           =   3255
   End
   Begin VB.TextBox Tenant_PhoneNo 
      Height          =   375
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   6000
      Width           =   3255
   End
   Begin VB.TextBox TenantRefID 
      Height          =   375
      Left            =   7200
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   42
      Top             =   6600
      Width           =   2295
   End
   Begin VB.ComboBox Country_list_property 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "Tenants.frx":02FD
      Left            =   3720
      List            =   "Tenants.frx":054D
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   4200
      Width           =   3255
   End
   Begin VB.TextBox Tenant_EmailAddress 
      Height          =   375
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   6600
      Width           =   3255
   End
   Begin VB.ListBox Tenant_List_RentalPrice 
      Height          =   1425
      Left            =   12000
      TabIndex        =   24
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox Tenant_List_LocationProperty 
      Height          =   1425
      Left            =   10080
      TabIndex        =   23
      Top             =   480
      Width           =   1935
   End
   Begin VB.ListBox Tenant_List_NumberBeds 
      Height          =   1425
      Left            =   8520
      TabIndex        =   22
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox Tenant_REFID 
      Height          =   3180
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   2055
   End
   Begin VB.ListBox Tenant_List_FName 
      Height          =   1425
      Left            =   2280
      TabIndex        =   9
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox Tenant_List_LName 
      Height          =   1425
      Left            =   3840
      TabIndex        =   8
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox Tenant_List_Address1 
      Height          =   1425
      Left            =   2280
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin VB.ListBox Tenant_List_Address2 
      Height          =   1425
      Left            =   3840
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.ListBox Tenant_List_PostCode 
      Height          =   1425
      Left            =   8520
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.ListBox Tenant_List_Country 
      Height          =   1425
      Left            =   10080
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin VB.ListBox Tenant_List_State 
      Height          =   1425
      Left            =   5400
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.ListBox Tenant_List_PhoneNumber 
      Height          =   1425
      Left            =   5400
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox Tenant_List_EmailAddress 
      Height          =   1425
      Left            =   6960
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox Tenant_List_City 
      Height          =   1425
      Left            =   6960
      TabIndex        =   0
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Tenant_Country 
      Height          =   285
      Left            =   3720
      TabIndex        =   45
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Tenant Image"
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
      Left            =   7200
      TabIndex        =   71
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label33 
      Caption         =   "Label33"
      Height          =   495
      Left            =   6240
      TabIndex        =   67
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label28 
      Caption         =   "Label28"
      Height          =   495
      Left            =   8280
      TabIndex        =   66
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Line Line2 
      X1              =   10560
      X2              =   10560
      Y1              =   8729.107
      Y2              =   16613.46
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Rental Price:"
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
      Left            =   7200
      TabIndex        =   59
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Location Of Property:"
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
      Left            =   7200
      TabIndex        =   58
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Number Of Beds"
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
      Left            =   7200
      TabIndex        =   57
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
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
      Left            =   240
      TabIndex        =   56
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
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
      Left            =   240
      TabIndex        =   55
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label27 
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
      Left            =   240
      TabIndex        =   54
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label23 
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
      Left            =   240
      TabIndex        =   53
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code"
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
      Left            =   240
      TabIndex        =   52
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Country:"
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
      Left            =   3720
      TabIndex        =   51
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
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
      Left            =   3720
      TabIndex        =   50
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label22 
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
      Left            =   3720
      TabIndex        =   49
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
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
      Left            =   3720
      TabIndex        =   48
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Tenant Refrence ID:"
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
      Left            =   7200
      TabIndex        =   47
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address:"
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
      Left            =   3720
      TabIndex        =   46
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Country:"
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
      Left            =   3720
      TabIndex        =   44
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
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
      Left            =   3720
      TabIndex        =   29
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
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
      Left            =   3720
      TabIndex        =   28
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   13800
      Y1              =   8729.107
      Y2              =   8729.107
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Number Of beds"
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
      Left            =   8640
      TabIndex        =   27
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Location Of property"
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
      Left            =   10200
      TabIndex        =   26
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Rental Price"
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
      Left            =   12240
      TabIndex        =   25
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Tenant Refrence ID:"
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
      TabIndex        =   21
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
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
      Left            =   2640
      TabIndex        =   20
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
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
      Left            =   4200
      TabIndex        =   19
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address"
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
      TabIndex        =   18
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number"
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
      Left            =   5640
      TabIndex        =   17
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
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
      TabIndex        =   16
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
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
      Left            =   10440
      TabIndex        =   15
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label11 
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
      Left            =   8760
      TabIndex        =   14
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label12 
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
      Left            =   3960
      TabIndex        =   13
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label13 
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
      Left            =   2400
      TabIndex        =   12
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label4 
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
      Left            =   7440
      TabIndex        =   11
      Top             =   1920
      Width           =   735
   End
   Begin VB.Menu Closetenat 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Tenants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BrowseImage_Click()
CommonDialog1.Filter = "Jpeg Files(*.jpg) | *.jpg"
CommonDialog1.ShowOpen
Tenant_Image.Text = CommonDialog1.FileName
End Sub

Private Sub Delete_Click()
Dim tempDeleteTenantchannel As Integer
Dim tempDeleteTenantfile As String
Dim FoundRecord As Boolean
Dim RemoveTenant As Tenant_Record
Dim RemoveTenantchannel As Integer
Dim P As Integer
Dim L As Integer
Dim intResponse As Integer
If Tenant_REFID.ListIndex = -1 Then
MsgBox ("Please Select a Tenant To Delete")
Else
    intResponse = MsgBox("Are you sure you want to delete Tenant  " & Tenant_REFID & "?" _
    & "                                  Once the Tenant removed you cannot recover it", _
                         vbYesNo + vbQuestion + vbDefaultButton2, _
                         "Delete")
    If intResponse = vbYes Then
tempDeleteTenantfile = App.Path + "\Saved_Dat\tempfile.tmp"
RemoveTenantchannel = FreeFile
Open Tenant_File For Random As RemoveTenantchannel Len = Tenant_Length
tempDeleteTenantchannel = FreeFile
Open tempDeleteTenantfile For Random As tempDeleteTenantchannel Len = Tenant_Length
P = 1
L = 1
FoundRecord = False
Get RemoveTenantchannel, P, RemoveTenant
Do While Not EOF(RemoveTenantchannel)
        If Tenant_REFID.List(Tenant_REFID.ListIndex) <> RemoveTenant.Tenant_REFID Then
                Put tempDeleteTenantchannel, L, RemoveTenant
                L = L + 1
        Else
                FoundRecord = True
        End If
        P = P + 1
        Get RemoveTenantchannel, P, RemoveTenant
Loop
Close RemoveTenantchannel
Close tempDeleteTenantchannel
    If FoundRecord = True Then
    MsgBox "This Tenant Has Been Successfully Deleted"
        Kill Tenant_File
        Name tempDeleteTenantfile As Tenant_File
        Tenant_File_PointerR = Tenant_File_PointerR - 1
    Else
   MsgBox "not found"
        Kill tempDeleteTenantfile
    End If
ClearThoseListBoxes
Form_Load
    End If
End If
End Sub

Private Sub EdIT_Click()
If Tenant_REFID.ListIndex = -1 Then
MsgBox "Please Select A Tenant Refrence ID"
Else
TenantRefID.Text = Tenant_REFID.List(Tenant_REFID.ListIndex)
TenantRefID.BackColor = &H80FF80
Save_Cancel
Label28.Caption = "EDIT"
UnlockTexts
End If
End Sub

Private Function Save_Cancel()
Add.Enabled = False
EdIT.Enabled = False
Prints.Enabled = False
Delete.Enabled = False
Save.Enabled = True
Cancel.Enabled = True
Tenant_Fname.SetFocus
End Function
Private Sub Add_Click()
Save_Cancel
Label28.Caption = "ADD"
HideListIndexes
UnlockTexts
End Sub

Private Sub cancel_Click()
CancelFunction
lockTexts
ClearTexts
End Sub

Private Sub Close_Click()
Unload Me
End Sub
Private Function CancelFunction()
Add.Enabled = True
Add.SetFocus
EdIT.Enabled = True
Prints.Enabled = True
Delete.Enabled = True
Save.Enabled = False
Cancel.Enabled = False
End Function

Private Sub NumberOfBedslist_Click()
Tenant_NB.Text = NumberOfBedslist
End Sub
Private Sub Country_list_property_Click()
Tenant_Country = Country_list_property
End Sub
Private Sub LocationOfproperty_click()
Tenant_Location.Text = LocationOfproperty
End Sub
Private Sub GenerateRandomTenantRefID()
TenantRefID.Text = Int(Rnd * 996) + 1
End Sub
Private Sub CheckTenantRef()
GenerateRandomTenantRefID
Dim CheckTenantID As Tenant_Record
Dim CheckTenantIDChannel As Integer
Dim O As Integer
O = 1
CheckTenantIDChannel = FreeFile
Open Tenant_File For Random As CheckTenantIDChannel Len = Tenant_Length
Get CheckTenantIDChannel, O, CheckTenantID
    Do While Not EOF(CheckTenantIDChannel)
        If Trim(CheckTenantID.Tenant_REFID) = Trim(TenantRefID.Text) Then
            GenerateRandomTenantRefID
            Landlord_File_Pointer = O
        End If
        O = O + 1
        Get CheckTenantIDChannel, O, CheckTenantID
    Loop
Close CheckTenantIDChannel
End Sub
Private Sub Form_Activate()
Label33.Caption = Selected_TenantID
End Sub

Private Sub Form_Load()
Dim ViewTenantDetails As Tenant_Record
Dim ViewTenantChannel As Integer
Dim X As Integer
X = 1
ViewTenantChannel = FreeFile
Open Tenant_File For Random As ViewTenantChannel Len = Tenant_Length
Get ViewTenantChannel, X, ViewTenantDetails
Do While Not EOF(ViewTenantChannel)
    Tenant_REFID.AddItem ViewTenantDetails.Tenant_REFID
    Tenant_List_FName.AddItem ViewTenantDetails.Tenant_Fname
    Tenant_List_LName.AddItem ViewTenantDetails.Tenant_LName
    Tenant_List_PhoneNumber.AddItem ViewTenantDetails.Tenant_PhoneNumber
    Tenant_List_EmailAddress.AddItem ViewTenantDetails.Tenant_EmailAddress
    Tenant_List_NumberBeds.AddItem ViewTenantDetails.Tenant_NumberOfBeds
    Tenant_List_RentalPrice.AddItem ViewTenantDetails.Tenant_RentalPrice
    Tenant_List_Address1.AddItem ViewTenantDetails.Tenant_Address1
    Tenant_List_Address2.AddItem ViewTenantDetails.Tenant_Address2
    Tenant_List_State.AddItem ViewTenantDetails.Tenant_State
    Tenant_List_City.AddItem ViewTenantDetails.Tenant_City
    Tenant_List_PostCode.AddItem ViewTenantDetails.Tenant_PostCode
    Tenant_List_Country.AddItem ViewTenantDetails.Tenant_CountryList
    Tenant_List_LocationProperty.AddItem ViewTenantDetails.Tenant_PropetyLocation
    tenant_file_pointer = Tenant_REFID.ListIndex + 1
    X = X + 1
    Get ViewTenantChannel, X, ViewTenantDetails
Loop
Close ViewTenantChannel
End Sub

Private Sub Label33_change()
Selected_TenantID = Label33.Caption
End Sub

Private Sub Tenant_REFID_Click()
Label33.Caption = Tenant_REFID.List(Tenant_REFID.ListIndex)
    Tenant_List_FName.ListIndex = Tenant_REFID.ListIndex
    Tenant_List_LName.ListIndex = Tenant_REFID.ListIndex
    Tenant_List_PhoneNumber.ListIndex = Tenant_REFID.ListIndex
    Tenant_List_EmailAddress.ListIndex = Tenant_REFID.ListIndex
    Tenant_List_NumberBeds.ListIndex = Tenant_REFID.ListIndex
    Tenant_List_LocationProperty.ListIndex = Tenant_REFID.ListIndex
    Tenant_List_RentalPrice.ListIndex = Tenant_REFID.ListIndex
    Tenant_List_Address1.ListIndex = Tenant_REFID.ListIndex
    Tenant_List_Address2.ListIndex = Tenant_REFID.ListIndex
    Tenant_List_State.ListIndex = Tenant_REFID.ListIndex
    Tenant_List_City.ListIndex = Tenant_REFID.ListIndex
    Tenant_List_PostCode.ListIndex = Tenant_REFID.ListIndex
    Tenant_List_Country.ListIndex = Tenant_REFID.ListIndex
    tenant_file_pointer = Tenant_REFID.ListIndex + 1
End Sub
Private Sub Prints_Click()
If Tenant_REFID.ListIndex = -1 Then
MsgBox "Please Select A Tenant Ref ID To Access This"
Else
Print_Tenant.Show 1
End If
End Sub

Private Sub Save_Click()
If Label28.Caption = "ADD" Then
Dim Add_Tenant As Tenant_Record
Dim TenantChannel As Integer
CheckTenantRef
TenantChannel = FreeFile
Open Tenant_File For Random As TenantChannel Len = Tenant_Length
            Add_Tenant.Tenant_Fname = Tenant_Fname.Text
            Add_Tenant.Tenant_LName = Tenant_LName.Text
            Add_Tenant.Tenant_Address1 = Tenant_AddressLine1.Text
            Add_Tenant.Tenant_Address2 = Tenant_AddressLine2.Text
            Add_Tenant.Tenant_PostCode = Tenant_PostCode.Text
            Add_Tenant.Tenant_CountryList = Tenant_Country.Text
            Add_Tenant.Tenant_State = Tenant_State.Text
            Add_Tenant.Tenant_City = Tenant_City.Text
            Add_Tenant.Tenant_RentalPrice = Tenant_RentalPrice.Text
            Add_Tenant.Tenant_PropetyLocation = Tenant_Location.Text
            Add_Tenant.Tenant_NumberOfBeds = Tenant_NB.Text
            Add_Tenant.Tenant_PhoneNumber = Tenant_PhoneNo.Text
            Add_Tenant.Tenant_EmailAddress = Tenant_EmailAddress.Text
            Add_Tenant.Tenant_Photo = Tenant_Image.Text
            Add_Tenant.Tenant_REFID = TenantRefID.Text
            MsgBox TenantRefID.Text
            Tenant_File_PointerR = Tenant_File_PointerR + 1
    Put TenantChannel, Tenant_File_PointerR, Add_Tenant
    Close TenantChannel
ClearTexts
ElseIf Label28.Caption = "EDIT" Then
Dim Amend_Landlord As Landlord_Record
Dim AmendLandlordchannel As Integer
        AmendTenantChannel = FreeFile
        Open Tenant_File For Random As AmendTenantChannel Len = Tenant_Length
            Add_Tenant.Tenant_Fname = Tenant_Fname.Text
            Add_Tenant.Tenant_LName = Tenant_LName.Text
            Add_Tenant.Tenant_Address1 = Tenant_AddressLine1.Text
            Add_Tenant.Tenant_Address2 = Tenant_AddressLine2.Text
            Add_Tenant.Tenant_PostCode = Tenant_PostCode.Text
            Add_Tenant.Tenant_CountryList = Tenant_Country.Text
           Add_Tenant.Tenant_State = Tenant_State.Text
            Add_Tenant.Tenant_City = Tenant_City.Text
            Add_Tenant.Tenant_RentalPrice = Tenant_RentalPrice.Text
            Add_Tenant.Tenant_PropetyLocation = Tenant_Location.Text
            Add_Tenant.Tenant_NumberOfBeds = Tenant_NB.Text
            Add_Tenant.Tenant_PhoneNumber = Tenant_PhoneNo.Text
            Add_Tenant.Tenant_EmailAddress = Tenant_EmailAddress.Text
            Add_Tenant.Tenant_REFID = TenantRefID.Text
        Put AmendTenantChannel, tenant_file_pointer, Amend_Tenant
        Close AmendTenantChannel
End If
ClearThoseListBoxes
TenantRefID.BackColor = &H80000005
Form_Load
End Sub
Private Function ClearThoseListBoxes()
Tenant_REFID.Clear
Tenant_List_LocationProperty.Clear
Tenant_List_FName.Clear
Tenant_List_LName.Clear
Tenant_List_PhoneNumber.Clear
Tenant_List_EmailAddress.Clear
Tenant_List_NumberBeds.Clear
Tenant_List_RentalPrice.Clear
Tenant_List_Address1.Clear
Tenant_List_Address2.Clear
Tenant_List_State.Clear
Tenant_List_City.Clear
Tenant_List_PostCode.Clear
Tenant_List_Country.Clear
End Function

Private Function HideListIndexes()
Tenant_REFID.ListIndex = -1
Tenant_List_FName.ListIndex = -1
Tenant_List_LName.ListIndex = -1
Tenant_List_PhoneNumber.ListIndex = -1
Tenant_List_EmailAddress.ListIndex = -1
Tenant_List_NumberBeds.ListIndex = -1
Tenant_List_RentalPrice.ListIndex = -1
Tenant_List_Address1.ListIndex = -1
Tenant_List_Address2.ListIndex = -1
Tenant_List_State.ListIndex = -1
Tenant_List_City.ListIndex = -1
Tenant_List_PostCode.ListIndex = -1
Tenant_List_Country.ListIndex = -1
End Function
Private Function UnlockTexts()
NumberOfBedslist.Locked = False
LocationOfproperty.Locked = False
Tenant_Fname.Locked = False
Tenant_LName.Locked = False
Tenant_AddressLine1.Locked = False
Tenant_AddressLine2.Locked = False
Tenant_PostCode.Locked = False
Tenant_Country.Locked = False
Tenant_State.Locked = False
Tenant_City.Locked = False
Tenant_PhoneNo.Locked = False
Tenant_EmailAddress.Locked = False
Country_list_property.Locked = False
Tenant_NB.Locked = False
Tenant_RentalPrice.Locked = False
Tenant_Location.Locked = False
End Function
Private Function lockTexts()
NumberOfBedslist.Locked = True
LocationOfproperty.Locked = True
Tenant_Fname.Locked = True
Tenant_LName.Locked = True
Tenant_AddressLine1.Locked = True
Tenant_AddressLine2.Locked = True
Tenant_PostCode.Locked = True
Tenant_Country.Locked = True
Tenant_State.Locked = True
Tenant_City.Locked = True
Tenant_PhoneNo.Locked = True
Tenant_EmailAddress.Locked = True
Country_list_property.Locked = True
Tenant_NB.Locked = True
Tenant_RentalPrice.Locked = True
Tenant_Location.Locked = True
End Function
Private Function ClearTexts()
Tenant_Fname = ""
Tenant_LName = ""
Tenant_AddressLine1 = ""
Tenant_AddressLine2 = ""
Tenant_PostCode = ""
Tenant_Country = ""
Tenant_State = ""
Tenant_City = ""
Tenant_PhoneNo = ""
Tenant_EmailAddress = ""
Tenant_Country = ""
Tenant_NB = ""
Tenant_RentalPrice = ""
Tenant_Location = ""
End Function

