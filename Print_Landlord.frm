VERSION 5.00
Begin VB.Form Print_Landlord 
   BackColor       =   &H00404040&
   Caption         =   "Report"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18285
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   18285
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TotalPayedTenant 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8160
      TabIndex        =   67
      Top             =   3360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox TotalPayedByTenant 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4920
      TabIndex        =   66
      Top             =   6960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ListBox List14 
      BackColor       =   &H00404040&
      ForeColor       =   &H80000005&
      Height          =   1815
      Left            =   6120
      TabIndex        =   65
      Top             =   7680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox List26 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   14880
      TabIndex        =   53
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ListBox List25 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   13200
      TabIndex        =   52
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ListBox List24 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   11520
      TabIndex        =   51
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ListBox List23 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   9840
      TabIndex        =   50
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ListBox List22 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   8520
      TabIndex        =   49
      Top             =   4920
      Width           =   1215
   End
   Begin VB.ListBox List21 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   6840
      TabIndex        =   48
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ListBox List20 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   5160
      TabIndex        =   47
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ListBox List19 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   3480
      TabIndex        =   46
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ListBox List18 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1800
      TabIndex        =   45
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ListBox List17 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   120
      TabIndex        =   44
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ListBox List16 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   16560
      TabIndex        =   43
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox Startit 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   8520
      TabIndex        =   42
      Top             =   7200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox SaveDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10440
      TabIndex        =   41
      Top             =   7320
      Width           =   2055
   End
   Begin VB.TextBox Tenant_PayDueDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10440
      TabIndex        =   40
      Top             =   6960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   12840
      Top             =   6960
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Property Details:"
      ForeColor       =   &H8000000E&
      Height          =   2655
      Left            =   120
      TabIndex        =   24
      Top             =   1920
      Width           =   18135
      Begin VB.ListBox List12 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   1785
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   1575
      End
      Begin VB.ListBox List13 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         ItemData        =   "Print_Landlord.frx":0000
         Left            =   1800
         List            =   "Print_Landlord.frx":0002
         TabIndex        =   30
         Top             =   600
         Width           =   2415
      End
      Begin VB.ListBox List15 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   4320
         TabIndex        =   29
         Top             =   600
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox RentalPriceWeek 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8040
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox Tenant_StartDate 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   11040
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   11040
         TabIndex        =   26
         Top             =   1320
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox PaymentNextDue 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   14040
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label16 
         Caption         =   "Get rented properties, current and between specific dates"
         Height          =   255
         Left            =   3000
         TabIndex        =   68
         Top             =   120
         Width           =   4815
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Properties "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   1800
         TabIndex        =   38
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   4320
         TabIndex        =   37
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Rental Price (Weekly)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   8040
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Price Payed by Tenant"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   8040
         TabIndex        =   35
         Top             =   1080
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant Date Started"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   11040
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant Date Ended"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   11040
         TabIndex        =   33
         Top             =   1080
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Next Due"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   14040
         TabIndex        =   32
         Top             =   360
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin VB.ListBox List11 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   16560
      TabIndex        =   22
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1800
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   3480
      TabIndex        =   8
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   5160
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ListBox List5 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   6840
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ListBox List6 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   8520
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ListBox List7 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   9840
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ListBox List8 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   11520
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ListBox List9 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   13200
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ListBox List10 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   14880
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label34 
      Caption         =   "OverDue Payment"
      Height          =   495
      Left            =   4920
      TabIndex        =   69
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Label Label33 
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   13200
      TabIndex        =   64
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label32 
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   5160
      TabIndex        =   63
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label31 
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   6840
      TabIndex        =   62
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label30 
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   8520
      TabIndex        =   61
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label29 
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   9840
      TabIndex        =   60
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label28 
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   11520
      TabIndex        =   59
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label27 
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   14880
      TabIndex        =   58
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label26 
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   16560
      TabIndex        =   57
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   3480
      TabIndex        =   56
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   1800
      TabIndex        =   55
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label23 
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   120
      TabIndex        =   54
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   2640
      TabIndex        =   23
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label7 
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   1800
      TabIndex        =   20
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   3480
      TabIndex        =   19
      Top             =   840
      Width           =   1455
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   16560
      TabIndex        =   18
      Top             =   840
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   14880
      TabIndex        =   17
      Top             =   840
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   11520
      TabIndex        =   16
      Top             =   840
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   9840
      TabIndex        =   15
      Top             =   840
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   8520
      TabIndex        =   14
      Top             =   840
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   6840
      TabIndex        =   13
      Top             =   840
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   5160
      TabIndex        =   12
      Top             =   840
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
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   13200
      TabIndex        =   11
      Top             =   840
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   19440
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "REPORT FOR "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Print_Landlord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
Dim s As TenantProperty_Record
Dim channel As Integer
Dim x As Integer
Dim O As Boolean
x = 1
channel = FreeFile
Open TenantProperty_File For Random As channel Len = TenantProperty_Length
Get channel, x, s
Do While Not EOF(channel) And O = False
    List16.AddItem s.PaymentDueDate
    TenantProperty_File_Pointer = x
    O = True
    x = x + 1
    Get channel, x, s
Loop
End Sub
Private Function UpdateDetailss()
    If Text2.Text = "Ongoing" Then
Dim UpdateDetails As TenantProperty_Record
Dim UpdateDetailsChannel As Integer
Dim x As Integer
Dim OnceUpdated As Boolean
Dim AddRentalPrice As Currency
x = 1
        If SaveDate.Text = PaymentNextDue.Text Then
AddRentalPrice = RentalPriceWeek.Text
UpdateDetailsChannel = FreeFile
Open TenantProperty_File For Random As UpdateDetailsChannel Len = TenantProperty_Length
Get UpdateDetailsChannel, x, UpdateDetails
        Do While Not EOF(UpdateDetailsChannel) And OnceUpdated = False
        If Trim(UpdateDetails.PaymentDueDate) = Trim(Tenant_PayDueDate.Text) And Trim(UpdateDetails.TotalRentalPricePayed) = Trim(TotalPayedByTenant.Text) Then
        UpdateDetails.PaymentDueDate = PaymentNextDue.Text
        TotalPayedByTenant.Text = TotalPayedByTenant + AddRentalPrice
        UpdateDetails.TotalRentalPricePayed = TotalPayedByTenant.Text
        TenantProperty_File_Pointer = x
        OnceUpdated = True
Put UpdateDetailsChannel, TenantProperty_File_Pointer, UpdateDetails
        End If
        x = x + 1
        Get UpdateDetailsChannel, x, UpdateDetails
        Loop
        Else
        End If
    Else
    MsgBox "Tenant Has been left the property"
    End If
End Function
Private Function ShowInfo()
Label18.Visible = True
RentalPriceWeek.Visible = True
Label19.Visible = True
TotalPayedTenant.Visible = True
Label20.Visible = True
Tenant_StartDate.Visible = True
Label21.Visible = True
Text2.Visible = True
Label22.Visible = True
PaymentNextDue.Visible = True
End Function

Private Sub Form_Activate()
Label10.Caption = Selected_LandlordID
PaymentNextDue.Text = Tenant_PayDueDate.Text
Startit.Text = List13.List(List13.ListIndex)
End Sub

Private Function ClearSecondListBoxes()
List16.Clear
List17.Clear
List18.Clear
List19.Clear
List20.Clear
List21.Clear
List21.Clear
List22.Clear
List23.Clear
List24.Clear
List25.Clear
List26.Clear
End Function

Private Sub Label10_Change()
Selected_LandlordID = Label10.Caption
End Sub
Private Sub Form_Load()
Timer1_Timer
Dim ViewLandlordID As Landlord_Record
Dim ViewLandlordIDChannel As Integer
Dim OnceFoundLandlord As Boolean
Dim x As Integer
x = 1
ViewLandlordIDChannel = FreeFile
Open Landlord_File For Random As ViewLandlordIDChannel Len = Landlord_Length
Get ViewLandlordIDChannel, x, ViewLandlordID
Do While Not EOF(ViewLandlordIDChannel) And OnceFoundLandlord = False
    If Trim(ViewLandlordID.Landlord_REFID) = Trim(Selected_LandlordID) Then
    List1.AddItem ViewLandlordID.Landlord_Fname
    List2.AddItem ViewLandlordID.Landlord_LName
    List3.AddItem ViewLandlordID.Landlord_CName
    List4.AddItem ViewLandlordID.Landlord_Address1
    List5.AddItem ViewLandlordID.Landlord_Address2
    List6.AddItem ViewLandlordID.Landlord_PostCode
    List7.AddItem ViewLandlordID.Landlord_CountryList
    List8.AddItem ViewLandlordID.Landlord_State
    List11.AddItem ViewLandlordID.Landlord_City
    List9.AddItem ViewLandlordID.Landlord_PhoneNumber
    List10.AddItem ViewLandlordID.Landlord_EmailAddress
    Landlord_File_Pointer = x
    OnceFoundLandlord = True
    Else
    x = x + 1
    Get ViewLandlordIDChannel, x, ViewLandlordID
End If
Loop
Close ViewLandlordIDChannel
End Sub

Private Sub List13_Click()
Dim ShowID As TenantProperty_Record
Dim RentalPriceOfProperty As Property_Record
Dim RentalPriceChannel As Integer
Dim ShowChannel As Integer
Dim f As Integer
Dim T As Integer

Startit.Text = List13.List(List13.ListIndex)
List15.Clear
f = 1
ShowChannel = FreeFile
Open TenantProperty_File For Random As ShowChannel Len = TenantProperty_Length
    Get ShowChannel, f, ShowID
    Do While Not EOF(ShowChannel)
        If Trim(Startit.Text) = Trim(ShowID.Property_RefID) Then
            List15.AddItem ShowID.Tenant_REFID
            Label17.Visible = True
            List15.Visible = True
        End If
        f = f + 1
        Get ShowChannel, f, ShowID
    Loop
Close ShowChannel
T = 1
RentalPriceChannel = FreeFile
Open Add_Property_File For Random As RentalPriceChannel Len = Add_Property_Length
Get RentalPriceChannel, T, RentalPriceOfProperty
Do While Not EOF(RentalPriceChannel)
    If Trim(List13.List(List13.ListIndex)) = Trim(RentalPriceOfProperty.Property_RefID) Then
    RentalPriceWeek.Text = RentalPriceOfProperty.Price_Property
    End If
   T = T + 1
    Get RentalPriceChannel, T, RentalPriceOfProperty
Loop
Close RentalPriceChannel
End Sub

Private Sub List14_Click()
Dim RentalPriceOfProperty As Property_Record
Dim RentalPriceChannel As Integer
Dim T As Integer
T = 1
RentalPriceChannel = FreeFile
Open Add_Property_File For Random As RentalPriceChannel Len = Add_Property_Length
Get RentalPriceChannel, T, RentalPriceOfProperty
Do While Not EOF(RentalPriceChannel)
    If Trim(List14.List(List14.ListIndex)) = Trim(RentalPriceOfProperty.Property_RefID) Then
    RentalPriceWeek.Text = RentalPriceOfProperty.Price_Property
    End If
    T = T + 1
    Get RentalPriceChannel, T, RentalPriceOfProperty
Loop
Close RentalPriceChannel
End Sub

Private Sub List15_Click()
ClearSecondListBoxes
Dim ViewTenant As Tenant_Record
Dim ViewTenantChannel As Integer
Dim R As Integer
Dim OnceFoundTenant As Boolean
R = 1
ViewTenantChannel = FreeFile
Open Tenant_File For Random As ViewTenantChannel Len = Tenant_Length
Get ViewTenantChannel, R, ViewTenant
Do While Not EOF(ViewTenantChannel) And OnceFoundTenant = False
    If Trim(List15.List(List15.ListIndex)) = Trim(ViewTenant.Tenant_REFID) Then
    List17.AddItem ViewTenant.Tenant_Fname
    List18.AddItem ViewTenant.Tenant_LName
    List19.AddItem ViewTenant.Tenant_Address1
    List20.AddItem ViewTenant.Tenant_Address2
    List21.AddItem ViewTenant.Tenant_PostCode
    List22.AddItem ViewTenant.Tenant_CountryList
    List23.AddItem ViewTenant.Tenant_State
    List24.AddItem ViewTenant.Tenant_City
    List25.AddItem ViewTenant.Tenant_PropetyLocation
    List26.AddItem ViewTenant.Tenant_PhoneNumber
    List16.AddItem ViewTenant.Tenant_EmailAddress
    Add_Property_Pointer = R
    OnceFoundTenant = True
    End If
    R = R + 1
    Get ViewTenantChannel, R, ViewTenant
Loop
Close ViewTenantChannel
Dim ShowCurrentProperties As TenantProperty_Record
Dim ShowCurrentPropertiesChannel As Integer
Dim V As Integer
V = 1
ShowCurrentPropertiesChannel = FreeFile
Open TenantProperty_File For Random As ShowCurrentPropertiesChannel Len = TenantProperty_Length
Get ShowCurrentPropertiesChannel, V, ShowCurrentProperties
Do While Not EOF(ShowCurrentPropertiesChannel)
    If Trim(List15.List(List15.ListIndex)) = Trim(ShowCurrentProperties.Tenant_REFID) And Trim(List13.List(List13.ListIndex)) = Trim(ShowCurrentProperties.Property_RefID) Then
        List14.AddItem ShowCurrentProperties.Property_RefID
    End If
    V = V + 1
    Get ShowCurrentPropertiesChannel, V, ShowCurrentProperties
Loop
Close ShowCurrentPropertiesChannel
Dim ShowPropertyDetail As TenantProperty_Record
Dim ShowPropertyChannel As Integer
Dim A As Integer
A = 1
ShowPropertyChannel = FreeFile
Open TenantProperty_File For Random As ShowPropertyChannel Len = TenantProperty_Length
Get ShowPropertyChannel, A, ShowPropertyDetail
Do While Not EOF(ShowPropertyChannel)
    If Trim(List15.List(List15.ListIndex)) = Trim(ShowPropertyDetail.Tenant_REFID) Then
       Tenant_StartDate.Text = ShowPropertyDetail.StartDate
       Tenant_PayDueDate.Text = ShowPropertyDetail.PaymentDueDate
       TotalPayedByTenant.Text = ShowPropertyDetail.TotalRentalPricePayed
       Text2.Text = ShowPropertyDetail.TenantEndDate
    End If
    A = A + 1
    Get ShowPropertyChannel, A, ShowPropertyDetail
Loop
        If Text2.Text = "00:00:00" Then
        Text2.Text = "Ongoing"
        Else
         End If
Close ShowPropertyChannel
SaveDate.Text = Format$(Now, "short Date")
ShowInfo
UpdateDetailss
       
End Sub

Private Sub RentalPriceWeek_Change()
RentalPriceWeek.Text = Format(RentalPriceWeek.Text, "£0.00")
End Sub

Private Sub Tenant_PayDueDate_Change()
PaymentNextDue.Text = Tenant_PayDueDate.Text
End Sub

Private Sub PaymentNextDue_Change()
Dim Days As Date
If IsDate(Tenant_PayDueDate.Text) Then
Days = Tenant_PayDueDate.Text
Days = DateAdd("d", 7, Days)
PaymentNextDue.Text = Days
Else
MsgBox ("Please input a proper date value!")
End If
End Sub

Private Sub Timer1_Timer()
If Startit.Visible = False Then
Timer1.Interval = 10
Timer1.Enabled = True
Dim PropertyDetails As Property_Record
Dim PropertyChannel As Integer
Dim C As Integer
C = 1
PropertyChannel = FreeFile
Open Add_Property_File For Random As PropertyChannel Len = Add_Property_Length
Get PropertyChannel, C, PropertyDetails
Do While Not EOF(PropertyChannel)
    If Trim(PropertyDetails.Landlord_REFID) = Trim(Label10.Caption) Then
     List12.AddItem PropertyDetails.Agent_REFID
     List13.AddItem PropertyDetails.Property_RefID
     Timer1.Enabled = False
End If
    C = C + 1
    Get PropertyChannel, C, PropertyDetails
Loop
Close PropertyChannel
End If
End Sub

Private Sub TotalPayedByTenant_Change()
TotalPayedTenant = TotalPayedByTenant.Text
End Sub

Private Sub TotalPayedTenant_Change()
TotalPayedTenant = Format(TotalPayedTenant, "£0.00")
End Sub
