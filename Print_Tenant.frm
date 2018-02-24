VERSION 5.00
Begin VB.Form Print_Tenant 
   Caption         =   "REPORT TENANT"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   14970
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox PhoneNumber 
      Height          =   285
      Left            =   2400
      TabIndex        =   72
      Top             =   8280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox EmailAddress 
      Height          =   285
      Left            =   2400
      TabIndex        =   71
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Types 
      Height          =   285
      Left            =   4560
      TabIndex        =   70
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Status 
      Height          =   285
      Left            =   2880
      TabIndex        =   69
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox FeeTYPE 
      Height          =   285
      Left            =   960
      TabIndex        =   68
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Overdue3 
      Height          =   285
      Left            =   8040
      TabIndex        =   67
      Top             =   7200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Overdue2 
      Height          =   285
      Left            =   8040
      TabIndex        =   66
      Top             =   6840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Overdue1 
      Height          =   285
      Left            =   8040
      TabIndex        =   65
      Top             =   6480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox MoveOUTDATE 
      Height          =   285
      Left            =   9720
      TabIndex        =   64
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Moveindate 
      Height          =   285
      Left            =   6960
      TabIndex        =   63
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox LatestPayment 
      Height          =   285
      Left            =   13440
      TabIndex        =   62
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox NextPayment 
      Height          =   285
      Left            =   13440
      TabIndex        =   61
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TotalRent 
      Height          =   285
      Left            =   9480
      TabIndex        =   60
      Top             =   6120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox RentalPrice 
      Height          =   285
      Left            =   6240
      TabIndex        =   59
      Top             =   6120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Paddress 
      Height          =   285
      Left            =   2400
      TabIndex        =   58
      Top             =   6120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox LandlordsCname 
      Height          =   285
      Left            =   2400
      TabIndex        =   57
      Top             =   6480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox LandlordsFname 
      Height          =   285
      Left            =   2400
      TabIndex        =   56
      Top             =   6840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox City 
      Height          =   285
      Left            =   2400
      TabIndex        =   55
      Top             =   7200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Post_Code 
      Height          =   285
      Left            =   2400
      TabIndex        =   54
      Top             =   7560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton HistoryInfo 
      Caption         =   "History"
      Height          =   375
      Left            =   3840
      TabIndex        =   47
      Top             =   5040
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3675
      ScaleWidth      =   4755
      TabIndex        =   46
      Top             =   960
      Width           =   4815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   4320
      Top             =   0
   End
   Begin VB.CommandButton AgentInfo 
      Caption         =   "See Agent Report"
      Height          =   375
      Left            =   7080
      TabIndex        =   31
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton ContactInfo 
      Caption         =   "Contact Details"
      Height          =   375
      Left            =   2520
      TabIndex        =   30
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton LandlordInfo 
      Caption         =   "See Landlord Report"
      Height          =   375
      Left            =   5160
      TabIndex        =   16
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton PaymentInfo 
      Caption         =   "Payments"
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton GeneralInfo 
      Caption         =   "General"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label LandlordID 
      Caption         =   "Label23"
      Height          =   255
      Left            =   8520
      TabIndex        =   74
      Top             =   8640
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label PropertyID 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   12240
      TabIndex        =   73
      Top             =   9000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   53
      Top             =   7920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   52
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Latest Payment Made:"
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
      Height          =   375
      Left            =   11280
      TabIndex        =   51
      Top             =   6480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment OverDue Date #3:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5760
      TabIndex        =   50
      Top             =   7200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment OverDue Date #2:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5760
      TabIndex        =   49
      Top             =   6840
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment OverDue Date #1:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5760
      TabIndex        =   48
      Top             =   6480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label TenantP_Country 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   9120
      TabIndex        =   45
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label TenantP_FirstName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   9120
      TabIndex        =   44
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label TenantP_LastName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   9120
      TabIndex        =   43
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label TenantP_Address1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   9120
      TabIndex        =   42
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label TenantP_Address2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   9120
      TabIndex        =   41
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label TenantP_PostCode 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   9120
      TabIndex        =   40
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label TenantP_City 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   9120
      TabIndex        =   39
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label TenantP_PhoneNumber 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   9120
      TabIndex        =   38
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label TenantP_EmailAddress 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   9120
      TabIndex        =   37
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label TenantP_Pay 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   9120
      TabIndex        =   36
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label TenantP_LocationProp 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   9120
      TabIndex        =   35
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label TenantP_NumberBeds 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   9120
      TabIndex        =   34
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   5520
      X2              =   12480
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label FirstNameLb 
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
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   480
      TabIndex        =   33
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label LastNameLb 
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
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   2280
      TabIndex        =   32
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Next Payment Due Date:"
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
      Height          =   375
      Left            =   11280
      TabIndex        =   29
      Top             =   6120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Move Out Date:"
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
      Height          =   375
      Left            =   8400
      TabIndex        =   28
      Top             =   7560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Move in date:"
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
      Height          =   375
      Left            =   5760
      TabIndex        =   27
      Top             =   7560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Rent Payed:"
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
      Height          =   375
      Left            =   8040
      TabIndex        =   26
      Top             =   6120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Rent:"
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
      Height          =   375
      Left            =   5760
      TabIndex        =   25
      Top             =   6120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   7560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "City :"
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
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   7200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Landlords First Name:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   6840
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Landlords Company Name:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4080
      TabIndex        =   20
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Property Address:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   6120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Fee Type:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Line Line15 
      BorderColor     =   &H80000003&
      BorderWidth     =   3
      X1              =   0
      X2              =   15000
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00C0C0C0&
      X1              =   5520
      X2              =   12480
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "NUMBER OF BEDROOMS "
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
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "PREFERED LOCATION OF PROPERTY"
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
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "LOOKING TO PAY"
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
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00C0C0C0&
      X1              =   5520
      X2              =   12480
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00C0C0C0&
      X1              =   5520
      X2              =   12480
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00C0C0C0&
      X1              =   5520
      X2              =   12480
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "EMAIL ADDRESS"
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
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "PHONE NUMBER"
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
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "CITY"
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
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "COUNTRY"
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
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "POST CODE"
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
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS LINE 2"
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
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS LINE 1"
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
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LAST NAME"
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
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "FIRST NAME"
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
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00C0C0C0&
      X1              =   5520
      X2              =   12480
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00C0C0C0&
      X1              =   5520
      X2              =   12480
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00C0C0C0&
      X1              =   5520
      X2              =   12480
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C0C0C0&
      X1              =   5520
      X2              =   12480
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C0C0C0&
      X1              =   5520
      X2              =   12480
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      X1              =   5520
      X2              =   12480
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      X1              =   5520
      X2              =   12480
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Top             =   8640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   120
      X2              =   4920
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Namelebel 
      BackStyle       =   0  'Transparent
      Caption         =   "TENANT NAME"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Print_Tenant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AgentInfo_Click()
Print_Agent.Show 1
End Sub

Private Sub ContactInfo_Click()
ShowContactDetails
Dim ShowEmailAndPhone As Tenant_Record
Dim ShowEPChannel As Integer
Dim x As Integer
Dim OnceFoundEP As Boolean

x = 1
ShowEPChannel = FreeFile
Open Tenant_File For Random As ShowEPChannel Len = Tenant_Length
Get ShowEPChannel, x, ShowEmailAndPhone
Do While Not EOF(ShowEPChannel) And OnceFoundEP = False
    If Label1.Caption = ShowEmailAndPhone.Tenant_REFID Then
    EmailAddress.Text = ShowEmailAndPhone.Tenant_EmailAddress
    PhoneNumber.Text = ShowEmailAndPhone.Tenant_PhoneNumber
    Tenant_File_PointerR = x
    OnceFoundEP = True
    End If
    x = x + 1
    Get ShowEPChannel, x, ShowEmailAndPhone
Loop
Close ShowEPChannel
End Sub

Private Sub Form_Activate()
Label1.Caption = Selected_TenantID
End Sub

Private Sub Form_Load()
Timer1_Timer
End Sub

Private Sub GeneralInfo_Click()
ShowGeneralInfo
Dim GeneralPropertyINFO As TenantProperty_Record
Dim GeneralPropertyChannel As Integer
Dim ShowPropertyDetails As Property_Record
Dim ShowPropertyChannel As Integer
Dim ShowLandlordINFO As Landlord_Record
Dim ShowLandlordChannel As Integer
Dim OnceRecordFound As Boolean
Dim x As Integer
Dim Y As Integer
Dim z As Integer
ShowGeneralInfo
x = 1
GeneralPropertyChannel = FreeFile
Open TenantProperty_File For Random As GeneralPropertyChannel Len = TenantProperty_Length
Get GeneralPropertyChannel, x, GeneralPropertyINFO
Do While Not EOF(GeneralPropertyChannel) And OnceRecordFound = False
    If Label1.Caption = GeneralPropertyINFO.Tenant_REFID Then
        PropertyID.Caption = GeneralPropertyINFO.Property_RefID
        NextPayment.Text = GeneralPropertyINFO.PaymentDueDate
        TotalRent.Text = GeneralPropertyINFO.TotalRentalPricePayed
        Moveindate.Text = GeneralPropertyINFO.StartDate
        MoveOUTDATE.Text = GeneralPropertyINFO.TenantEndDate
        Overdue1.Text = GeneralPropertyINFO.OverDue01
        Overdue2.Text = GeneralPropertyINFO.OverDue02
        Overdue3.Text = GeneralPropertyINFO.OverDue03
        LatestPayment.Text = GeneralPropertyINFO.LatestPayment
        TenantProperty_File_Pointer = x
            z = 1
                If MoveOUTDATE.Text = "00:00:00" Then
                Status.Text = "Renting"
                Else
                Status.Text = "Left Property"
                End If
            ShowPropertyChannel = FreeFile
            Open Add_Property_File For Random As ShowPropertyChannel Len = Add_Property_Length
            Get ShowPropertyChannel, z, ShowPropertyDetails
            Do While Not EOF(ShowPropertyChannel) And OnceRecordFound = False
            If Trim(ShowPropertyDetails.Property_RefID) = Trim(PropertyID.Caption) And Trim(Label1.Caption) = Trim(GeneralPropertyINFO.Tenant_REFID) Then
            Paddress.Text = ShowPropertyDetails.Address_line_1
            FeeTYPE.Text = ShowPropertyDetails.Payment_type
            Types.Text = ShowPropertyDetails.RentalType
            RentalPrice.Text = ShowPropertyDetails.Price_Property
            LandlordID.Caption = ShowPropertyDetails.Landlord_REFID
            Add_Property_Pointer = z
            OnceFoundRecord = True
            End If
            z = z + 1
            Get ShowPropertyChannel, z, ShowPropertyDetails
            Loop
            Y = 1
            ShowLandlordChannel = FreeFile
            Open Landlord_File For Random As ShowLandlordChannel Len = Landlord_Length
            Get ShowLandlordChannel, Y, ShowLandlordINFO
            Do While Not EOF(ShowLandlordChannel) And OnceRecordFound = False
            If Trim(ShowLandlordINFO.Landlord_REFID) = Trim(LandlordID.Caption) Then
                    LandlordsCname.Text = ShowLandlordINFO.Landlord_CName
                    LandlordsFname.Text = ShowLandlordINFO.Landlord_Fname
                    City.Text = ShowLandlordINFO.Landlord_City
                    Post_Code.Text = ShowLandlordINFO.Landlord_PostCode
                    Landlord_File_Pointer = Y
            OnceFoundRecord = True
            End If
            Y = Y + 1
            Get ShowLandlordChannel, Y, ShowLandlordINFO
            Loop
        OnceRecordFound = True
        End If
        x = x + 1
        Get GeneralPropertyChannel, x, GeneralPropertyINFO
Loop
Close GeneralPropertyChannel
Close ShowPropertyChannel
            
End Sub

Private Sub Label1_Click()
Selected_TenantID = Label1.Caption
End Sub

Private Sub Label29_Click()
Selected_TenantLastName = Label29.Caption
End Sub

Private Function ShowFirstNameAndSecondName()
Dim ShowFirstLastName As Tenant_Record
Dim SFLNameChannel As Integer
Dim x As Integer
Dim OnceFoundRecord As Boolean
x = 1
SFLNameChannel = FreeFile
Open Tenant_File For Random As SFLNameChannel Len = Tenant_Length
Get SFLNameChannel, x, ShowFirstLastName
Do While Not EOF(SFLNameChannel) And OnceFoundRecord = False
    If Trim(Label1.Caption) = Trim(ShowFirstLastName.Tenant_REFID) Then
    FirstNameLb.Caption = ShowFirstLastName.Tenant_Fname
    LastNameLb.Caption = ShowFirstLastName.Tenant_LName
    Tenant_File_PointerR = x
    Timer1.Enabled = False
    OnceFoundRecord = True
    End If
    x = x + 1
    Get SFLNameChannel, x, ShowFirstLastName
Loop
Close SFLNameChannel
End Function
Private Function ShowOtherTenantInfo()
Dim ShowOtherInfo As Tenant_Record
Dim SOIChannel As Integer
Dim Y As Integer
Dim OnceFoundRecord As Boolean
Y = 1
SOIChannel = FreeFile
Open Tenant_File For Random As SOIChannel Len = Tenant_Length
Get SOIChannel, Y, ShowOtherInfo
Do While Not EOF(SOIChannel) And OnceFoundRecord = False
If Trim(Label1.Caption) = Trim(ShowOtherInfo.Tenant_REFID) Then
    TenantP_FirstName.Caption = ShowOtherInfo.Tenant_Fname
    TenantP_LastName.Caption = ShowOtherInfo.Tenant_LName
    TenantP_Address1.Caption = ShowOtherInfo.Tenant_Address1
    TenantP_Address2.Caption = ShowOtherInfo.Tenant_Address2
    TenantP_PostCode.Caption = ShowOtherInfo.Tenant_PostCode
    TenantP_Country.Caption = ShowOtherInfo.Tenant_CountryList
    TenantP_City.Caption = ShowOtherInfo.Tenant_City
    TenantP_PhoneNumber.Caption = ShowOtherInfo.Tenant_PhoneNumber
    TenantP_EmailAddress.Caption = ShowOtherInfo.Tenant_EmailAddress
    TenantP_Pay.Caption = ShowOtherInfo.Tenant_RentalPrice
    TenantP_LocationProp.Caption = ShowOtherInfo.Tenant_PropetyLocation
    TenantP_NumberBeds.Caption = ShowOtherInfo.Tenant_NumberOfBeds
    Dim Pic As Picture
    Picture1.AutoRedraw = True
    Set Pic = LoadPicture(ShowOtherInfo.Tenant_Photo)
    Picture1.PaintPicture Pic, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
    Set Picture1.Picture = Picture1.Image
Tenant_File_PointerR = Y
Timer1.Enabled = False
OnceFoundRecord = True
End If
Y = Y + 1
Get SOIChannel, Y, ShowOtherInfo
Loop
Close siochannel
End Function

Private Sub LandlordInfo_Click()
Print_Landlord.Show 1
End Sub

Private Sub PaymentInfo_Click()
ShowPaymentInfo
End Sub

Private Sub Timer1_Timer()
If FirstNameLb.Visible = True Then
Timer1.Interval = 10
Timer1.Enabled = True
ShowFirstNameAndSecondName
ShowOtherTenantInfo
End If
End Sub
Private Function ShowContactDetails()
RentalPrice.Visible = False
TotalRent.Visible = False
NextPayment.Visible = False
LatestPayment.Visible = False
Overdue1.Visible = False
Overdue2.Visible = False
Overdue3.Visible = False
Moveindate.Visible = False
MoveOUTDATE.Visible = False
Types.Visible = False
EmailAddress.Visible = False
PhoneNumber.Visible = False
FeeTYPE.Visible = False
Status.Visible = False
Types.Visible = False
Paddress.Visible = False
LandlordsCname.Visible = False
LandlordsFname.Visible = False
City.Visible = False
Post_Code.Visible = False
EmailAddress.Visible = False
PhoneNumber.Visible = False
Label12.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = True
EmailAddress.Visible = True
PhoneNumber.Visible = True
Label33.Visible = True
End Function

Private Function ShowGeneralInfo()
Label12.Visible = True
Label15.Visible = True
Label16.Visible = True
Label17.Visible = True
Label18.Visible = True
Label19.Visible = True
Label20.Visible = True
Label21.Visible = True
Label22.Visible = True
Label24.Visible = True
Label25.Visible = True
Label26.Visible = True
Label27.Visible = True
Label28.Visible = True
Label29.Visible = True
Label30.Visible = True
Label31.Visible = True
Label32.Visible = False
Label33.Visible = False
FeeTYPE.Visible = True
Status.Visible = True
Paddress.Visible = True
LandlordsCname.Visible = True
LandlordsFname.Visible = True
City.Visible = True
Post_Code.Visible = True
RentalPrice.Visible = True
TotalRent.Visible = True
NextPayment.Visible = True
LatestPayment.Visible = True
Overdue1.Visible = True
Overdue2.Visible = True
Overdue3.Visible = True
Moveindate.Visible = True
MoveOUTDATE.Visible = True
Types.Visible = True
EmailAddress.Visible = False
PhoneNumber.Visible = False
End Function

Private Function ShowPaymentInfo()
FeeTYPE.Visible = False
Status.Visible = False
Types.Visible = False
Paddress.Visible = False
LandlordsCname.Visible = False
LandlordsFname.Visible = False
City.Visible = False
Post_Code.Visible = False
EmailAddress.Visible = False
PhoneNumber.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label32.Visible = False
Label33.Visible = False
Label24.Visible = True
Label25.Visible = True
Label28.Visible = True
Label12.Visible = True
Label29.Visible = True
Label30.Visible = True
Label31.Visible = True
Label26.Visible = True
Label27.Visible = True
RentalPrice.Visible = True
TotalRent.Visible = True
NextPayment.Visible = True
LatestPayment.Visible = True
Overdue1.Visible = True
Overdue2.Visible = True
Overdue3.Visible = True
Moveindate.Visible = True
MoveOUTDATE.Visible = True
End Function

