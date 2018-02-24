VERSION 5.00
Begin VB.Form FullReport 
   Caption         =   "REPORT"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15555
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   15555
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame HistoricalEvents 
      Caption         =   "History"
      Height          =   6615
      Left            =   120
      TabIndex        =   75
      Top             =   1680
      Width           =   6135
   End
   Begin VB.TextBox SaveDate 
      Height          =   285
      Left            =   2160
      TabIndex        =   73
      Top             =   8640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Overduepayment 
      Caption         =   "Take Overdue Payment"
      Height          =   495
      Left            =   9120
      TabIndex        =   72
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Calculate3 
      Height          =   285
      Left            =   4080
      TabIndex        =   70
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Calculate2 
      Height          =   285
      Left            =   4080
      TabIndex        =   69
      Top             =   8640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Calculate 
      Height          =   285
      Left            =   4080
      TabIndex        =   68
      Top             =   8400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TotalOverDue 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8640
      TabIndex        =   66
      Top             =   7080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton SaveChanges 
      Caption         =   "Save Changes"
      Height          =   495
      Left            =   9120
      TabIndex        =   64
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Cancelbutton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6720
      TabIndex        =   63
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton TakePayment 
      Caption         =   "Take Rental Payment"
      Height          =   495
      Left            =   7920
      TabIndex        =   62
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TheTotalRental 
      Height          =   285
      Left            =   8640
      TabIndex        =   60
      Top             =   7440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TakeRentalPayment 
      Height          =   285
      Left            =   8640
      TabIndex        =   58
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Add7days 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   8400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search"
      Height          =   1455
      Left            =   120
      TabIndex        =   47
      Top             =   120
      Width           =   6135
      Begin VB.ComboBox PropertyID 
         Height          =   315
         Left            =   360
         TabIndex        =   51
         Top             =   480
         Width           =   3375
      End
      Begin VB.ComboBox TenantIDs 
         Height          =   315
         Left            =   360
         TabIndex        =   50
         Top             =   1080
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox LandlordRID 
         Height          =   285
         Left            =   3960
         TabIndex        =   49
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox AgentRID 
         Height          =   285
         Left            =   5040
         TabIndex        =   48
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label PID 
         BackStyle       =   0  'Transparent
         Caption         =   "Property Reference ID"
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
         Left            =   360
         TabIndex        =   55
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant Living at this Property"
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
         Left            =   360
         TabIndex        =   54
         Top             =   840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Landlord ID"
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
         Left            =   3960
         TabIndex        =   53
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent ID"
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
         Left            =   5040
         TabIndex        =   52
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.TextBox TP_Feetype 
      Height          =   285
      Left            =   9000
      TabIndex        =   46
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TP_Status 
      Height          =   285
      Left            =   9000
      TabIndex        =   44
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TP_Moveindate 
      Height          =   285
      Left            =   9000
      TabIndex        =   37
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TP_MoveOUTDATE 
      Height          =   285
      Left            =   9000
      TabIndex        =   36
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TP_Overdue1 
      Height          =   285
      Left            =   9000
      TabIndex        =   35
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TP_Overdue2 
      Height          =   285
      Left            =   9000
      TabIndex        =   34
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TP_Overdue3 
      Height          =   285
      Left            =   9000
      TabIndex        =   33
      Top             =   3480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TP_Latest_PayMade 
      Height          =   285
      Left            =   9000
      TabIndex        =   32
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TP_Next_DUE 
      Height          =   285
      Left            =   9000
      TabIndex        =   31
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TP_TotalRental 
      Height          =   285
      Left            =   9000
      TabIndex        =   30
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TP_RentalPrice 
      Height          =   285
      Left            =   9000
      TabIndex        =   29
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   11160
      ScaleHeight     =   3315
      ScaleWidth      =   4155
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Frame TenantCheckFrame 
      Height          =   975
      Left            =   6840
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   3615
      Begin VB.OptionButton ClearOption 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   74
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton NoOption 
         Caption         =   "No"
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton YesOption 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.Label PayedDate 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1920
         TabIndex        =   56
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tenant Payment For:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.TextBox DateCalculator 
      Height          =   285
      Left            =   2160
      TabIndex        =   71
      Top             =   8880
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   6720
      TabIndex        =   67
      Top             =   7440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label inform 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6840
      TabIndex        =   65
      Top             =   6840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "OverDue Total:"
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
      Left            =   6720
      TabIndex        =   61
      Top             =   7080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label23 
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
      Left            =   6720
      TabIndex        =   59
      Top             =   6720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "TAKE TENANT'S PAYMENT"
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
      Left            =   7560
      TabIndex        =   57
      Top             =   5280
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Line Line14 
      Visible         =   0   'False
      X1              =   6480
      X2              =   6480
      Y1              =   3480
      Y2              =   8880
   End
   Begin VB.Shape StatusMovedOut 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape StatusRenting 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line Line13 
      Visible         =   0   'False
      X1              =   6480
      X2              =   11280
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label Label21 
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
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6600
      TabIndex        =   45
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label20 
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
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6600
      TabIndex        =   43
      Top             =   960
      Visible         =   0   'False
      Width           =   735
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
      Left            =   6600
      TabIndex        =   42
      Top             =   3840
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
      Left            =   6600
      TabIndex        =   41
      Top             =   4200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label19 
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
      Left            =   6600
      TabIndex        =   40
      Top             =   2760
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
      Left            =   6600
      TabIndex        =   39
      Top             =   3120
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
      Left            =   6600
      TabIndex        =   38
      Top             =   3480
      Visible         =   0   'False
      Width           =   2895
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
      Left            =   6600
      TabIndex        =   28
      Top             =   2400
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label14 
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
      Left            =   6600
      TabIndex        =   27
      Top             =   2040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label label13 
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
      Left            =   6600
      TabIndex        =   26
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label12 
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
      Left            =   6600
      TabIndex        =   25
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "TENANT'S PAYMENT INFROMATION"
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
      Left            =   7200
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Line Line12 
      Visible         =   0   'False
      X1              =   6480
      X2              =   6480
      Y1              =   -120
      Y2              =   5280
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00C0C0C0&
      Visible         =   0   'False
      X1              =   11160
      X2              =   15360
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label18 
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
      Left            =   11160
      TabIndex        =   23
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label17 
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
      Left            =   11160
      TabIndex        =   22
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00C0C0C0&
      Visible         =   0   'False
      X1              =   11160
      X2              =   15360
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00C0C0C0&
      Visible         =   0   'False
      X1              =   11160
      X2              =   15360
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Label Label16 
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
      Left            =   11160
      TabIndex        =   21
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label15 
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
      Left            =   11160
      TabIndex        =   20
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C0C0C0&
      Visible         =   0   'False
      X1              =   11160
      X2              =   15360
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label TP_Address1 
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
      Left            =   13080
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label TP_Address2 
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
      Left            =   13080
      TabIndex        =   18
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label TP_PostCode 
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
      Left            =   13080
      TabIndex        =   17
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label TP_City 
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
      Left            =   13080
      TabIndex        =   16
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TENANT'S PROPERTY DETAILS"
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
      Left            =   11160
      TabIndex        =   15
      Top             =   5280
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Line Line5 
      Visible         =   0   'False
      X1              =   10800
      X2              =   10800
      Y1              =   3120
      Y2              =   8400
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "TENANT'S PERSONAL PERSONAL INFROMATION"
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
      Left            =   11280
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label T_EmailAddress 
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
      Left            =   13080
      TabIndex        =   13
      Top             =   4800
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label T_PhoneNumber 
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
      Left            =   13080
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label T_LastName 
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
      Left            =   13080
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label T_FirstName 
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
      Left            =   13080
      TabIndex        =   10
      Top             =   3720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Line Line4 
      Visible         =   0   'False
      X1              =   10800
      X2              =   15600
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00C0C0C0&
      Visible         =   0   'False
      X1              =   11160
      X2              =   15360
      Y1              =   4680
      Y2              =   4680
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
      Left            =   11160
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   11160
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00C0C0C0&
      Visible         =   0   'False
      X1              =   11160
      X2              =   15360
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      Visible         =   0   'False
      X1              =   11160
      X2              =   15360
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label6 
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
      Left            =   11160
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label5 
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
      Left            =   11160
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      Visible         =   0   'False
      X1              =   11160
      X2              =   15360
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   10800
      X2              =   10800
      Y1              =   0
      Y2              =   5280
   End
End
Attribute VB_Name = "FullReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
 UpdateDUEDATE
End Sub

Private Sub Command1_Click()
Dim times As Date
'times = DateAdd("d", 7, times)
'TP_Next_DUE.Text = times
times = TP_Next_DUE.Text
TP_Next_DUE.Text = Format$(times + 7, "short Date")
End Sub

Private Sub Cancelbutton_Click()
TotalOverDue.Text = ""
TheTotalRental.Text = ""
TakeRentalPayment.Visible = False
TotalOverDue.Visible = False
TheTotalRental.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
inform.Visible = False
ClearOption(1).SetFocus
TakePayment.Visible = False
Overduepayment.Visible = False
Cancelbutton.Visible = False
SaveChanges.Visible = False
TP_Overdue1.Text = "00:00:00"
TP_Overdue1.BackColor = &HFFFFFF
End Sub

Private Sub Form_Activate()
TP_Next_DUE.Text = PayedDate.Caption
End Sub

Private Sub Form_Load()
Dim ShowPropertyID As Property_Record
Dim ShowPropertyIDChannels As Integer
Dim x As Integer

x = 1
ShowPropertyIDChannels = FreeFile
Open Add_Property_File For Random As ShowPropertyIDChannels Len = Add_Property_Length
Get ShowPropertyIDChannels, x, ShowPropertyID
Do While Not EOF(ShowPropertyIDChannels) And oncefound = False
    PropertyID.AddItem ShowPropertyID.Property_RefID
        x = x + 1
        Get ShowPropertyIDChannels, x, ShowPropertyID
Loop
Close ShowPropertyIDChannels
End Sub

Private Sub NoOption_Click()
If TP_Overdue1.Text = "00:00:00" Then
TP_Overdue1.Text = PayedDate.Caption
TP_Overdue1.BackColor = &H80FF80
inform.Caption = "Payment Overdue 1 has been set "
ElseIf TP_Overdue2.Text = "00:00:00" Then
TP_Overdue2.Text = PayedDate.Caption
TP_Overdue2.BackColor = &H80FF80
inform.Caption = "Payment Overdue 2 has been set "
ElseIf TP_Overdue3.Text = "00:00:00" Then
TP_Overdue3.Text = PayedDate.Caption
TP_Overdue3.BackColor = &H80FF80
inform.Caption = "Payment Overdue 3 has been set "
End If
TakeRentalPayment.Visible = False
TheTotalRental.Visible = False
Label23.Visible = False
TotalOverDue.Visible = False
Label25.Visible = False
Label24.Visible = False
TakePayment.Visible = False
Cancelbutton.Visible = True
SaveChanges.Visible = True
inform.Visible = True
Overduepayment.Visible = False
End Sub

Private Sub Overduepayment_Click()
If TP_Status.Text = "Renting" Then
Dim UpdateOverDueDetails As TenantProperty_Record
Dim UpdateOverDueDetailsChannel As Integer
Dim x As Integer
Dim OnceUpdated As Boolean
Dim AddRentalPrice As Currency
x = 1
AddOverDuePrice = TotalOverDue.Text
UpdateOverDueDetailsChannel = FreeFile
Open TenantProperty_File For Random As UpdateOverDueDetailsChannel Len = TenantProperty_Length
Get UpdateOverDueDetailsChannel, x, UpdateOverDueDetails
        Do While Not EOF(UpdateOverDueDetailsChannel) And OnceUpdated = False
            If Trim(UpdateOverDueDetails.PaymentDueDate) = Trim(TP_Next_DUE.Text) And Trim(UpdateOverDueDetails.PaymentPriceOverDue) = Trim(TotalOverDue.Text) Then
            TheTotalRental.Text = UpdateOverDueDetails.TotalRentalPricePayed + AddOverDuePrice
            UpdateOverDueDetails.TotalRentalPricePayed = TheTotalRental.Text
            TotalOverDue.Text = "0"
            TP_Overdue1.Text = "00:00:00"
            TP_Overdue2.Text = "00:00:00"
            TP_Overdue3.Text = "00:00:00"
            UpdateOverDueDetails.PaymentPriceOverDue = TotalOverDue.Text
            UpdateOverDueDetails.OverDue01 = TP_Overdue1.Text
            UpdateOverDueDetails.OverDue02 = TP_Overdue2.Text
            UpdateOverDueDetails.OverDue03 = TP_Overdue3.Text
            TenantProperty_File_Pointer = x
            OnceUpdated = True
    Put UpdateOverDueDetailsChannel, TenantProperty_File_Pointer, UpdateOverDueDetails
            End If
        x = x + 1
        Get UpdateOverDueDetailsChannel, x, UpdateOverDueDetails
        Loop
Else
MsgBox "Tenant Has left the property"
End If
End Sub

Private Sub PropertyID_Click()
HideTInfo
HidePInfo
HideTPInfo
SHOWLANDLORDAGENTID
ShowTenantProperty
End Sub

Private Function ShowTenantProperty()
Dim TenantPropertyID As TenantProperty_Record
Dim TenantPropertyIDChannel As Integer
Dim C As Integer
Dim RecordFound As Boolean
C = 1
TenantIDs.Visible = True
Label2.Visible = True
TenantIDs.Clear
TenantPropertyIDChannel = FreeFile
    Open TenantProperty_File For Random As TenantPropertyIDChannel Len = TenantProperty_Length
    Get TenantPropertyIDChannel, C, TenantPropertyID
        Do While Not EOF(TenantPropertyIDChannel)
            If Trim(PropertyID.List(PropertyID.ListIndex)) = Trim(TenantPropertyID.Property_RefID) Then
            TenantIDs.AddItem TenantPropertyID.Tenant_REFID
            TenantProperty_File_Pointer = C
            End If
            C = C + 1
        Get TenantPropertyIDChannel, C, TenantPropertyID
        Loop
    Close TenantPropertyIDChannel
End Function

Private Function SHOWLANDLORDAGENTID()
Dim ShowLandlordID As Property_Record
Dim ShowLandlordIDChannel As Integer
Dim z As Integer
Dim oncefound As Boolean
z = 1
ShowLandlordIDChannel = FreeFile
Open Add_Property_File For Random As ShowLandlordIDChannel Len = Add_Property_Length
Get ShowLandlordIDChannel, z, ShowLandlordID
Do While Not EOF(ShowLandlordIDChannel) And oncefound = False
    If Trim(PropertyID.List(PropertyID.ListIndex)) = Trim(ShowLandlordID.Property_RefID) Then
        LandlordRID.Visible = True
        Label3.Visible = True
        Label4.Visible = True
        AgentRID.Visible = True
       LandlordRID.Text = ShowLandlordID.Landlord_REFID
       AgentRID.Text = ShowLandlordID.Agent_REFID
        Add_Property_Pointer = z
        oncefound = True
   Else
    z = z + 1
    Get ShowLandlordIDChannel, z, ShowLandlordID
    End If
    Loop
    Close ShowLandlordIDChannel
End Function

Private Sub SaveChanges_Click()
Dim Addoverdue As TenantProperty_Record
Dim Overduechannel As Integer
Dim x As Integer
Dim oncefound As Boolean
Dim time As Date
Dim intResponse As Integer
time = TP_Next_DUE.Text
x = 1
Overduechannel = FreeFile
Open TenantProperty_File For Random As Overduechannel Len = TenantProperty_Length
Get Overduechannel, x, Addoverdue
Do While Not EOF(Overduechannel) And oncefound = False
If Trim(TenantIDs.List(TenantIDs.ListIndex)) = Trim(Addoverdue.Tenant_REFID) And Addoverdue.OverDue01 = "00:00:00" Then
    Addoverdue.PaymentMade = False
    TP_Next_DUE.Text = Format$(time + 7, "short Date")
    Addoverdue.PaymentDueDate = TP_Next_DUE.Text
    Addoverdue.OverDue01 = TP_Overdue1.Text
    Calculate.Text = TP_RentalPrice.Text
    Addoverdue.PaymentPriceOverDue = Calculate.Text
    inform.Visible = False
    TP_Overdue1.BackColor = &H8000000F
    SaveChanges.Visible = False
    TenantIDs.SetFocus
    TenantProperty_File_Pointer = x
    oncefound = True
    Put Overduechannel, x, Addoverdue
ElseIf Trim(TenantIDs.List(TenantIDs.ListIndex)) = Trim(Addoverdue.Tenant_REFID) And Addoverdue.OverDue02 = "00:00:00" Then
    Addoverdue.PaymentMade = False
    TP_Next_DUE.Text = Format$(time + 7, "short Date")
    Addoverdue.PaymentDueDate = TP_Next_DUE.Text
    Addoverdue.OverDue02 = TP_Overdue2.Text
    Calculate2.Text = (TP_RentalPrice.Text + Addoverdue.PaymentPriceOverDue)
    Addoverdue.PaymentPriceOverDue = Calculate2.Text
    time = TP_Next_DUE.Text
    inform.Visible = False
    TP_Overdue2.BackColor = &H8000000F
    SaveChanges.Visible = False
    TenantIDs.SetFocus
    TenantProperty_File_Pointer = x
    oncefound = True
    Put Overduechannel, x, Addoverdue
ElseIf Trim(TenantIDs.List(TenantIDs.ListIndex)) = Trim(Addoverdue.Tenant_REFID) And Addoverdue.OverDue03 = "00:00:00" Then
    Addoverdue.PaymentMade = False
    TP_Next_DUE.Text = Format$(time + 7, "short Date")
    Addoverdue.PaymentDueDate = TP_Next_DUE.Text
    Addoverdue.OverDue03 = TP_Overdue3.Text
    Calculate3.Text = (TP_RentalPrice.Text + Addoverdue.PaymentPriceOverDue)
    Addoverdue.PaymentPriceOverDue = Calculate3.Text
    inform.Visible = False
    TP_Overdue3.BackColor = &H8000000F
    SaveChanges.Visible = False
    TenantIDs.SetFocus
    TenantProperty_File_Pointer = x
    oncefound = True
    Put Overduechannel, x, Addoverdue
Else
    intResponse = MsgBox("This Tenant has not payed for 3 bills and has reached the maximum overdue payment " _
                       & "Send Report?", _
                         vbYesNo + vbQuestion + vbDefaultButton2, _
                         "Send Report?")
    If intResponse = vbYes Then
    Print "Something goes here"
    End If
x = x + 1
Get Overduechannel, x, Addoverdue
End If
Loop
Close Overduechannel
End Sub

Private Sub TakePayment_Click()
If TP_Status.Text = "Renting" Then
Dim UpdateDetails As TenantProperty_Record
Dim UpdateDetailsChannel As Integer
Dim x As Integer
Dim OnceUpdated As Boolean
Dim AddRentalPrice As Currency
Dim time As Date
time = TP_Next_DUE.Text
x = 1
        If SaveDate.Text = TP_Next_DUE.Text Then
AddRentalPrice = TakeRentalPayment.Text
UpdateDetailsChannel = FreeFile
Open TenantProperty_File For Random As UpdateDetailsChannel Len = TenantProperty_Length
Get UpdateDetailsChannel, x, UpdateDetails
        Do While Not EOF(UpdateDetailsChannel) And OnceUpdated = False
            If Trim(UpdateDetails.PaymentDueDate) = Trim(TP_Next_DUE.Text) And Trim(UpdateDetails.TotalRentalPricePayed) = Trim(TheTotalRental.Text) Then
            TP_Next_DUE.Text = Format$(time + 7, "short Date")
            UpdateDetails.PaymentDueDate = TP_Next_DUE.Text
            TheTotalRental.Text = UpdateDetails.TotalRentalPricePayed + AddRentalPrice
            TP_Latest_PayMade.Text = SaveDate.Text
            UpdateDetails.TotalRentalPricePayed = TheTotalRental.Text
            UpdateDetails.LatestPayment = TP_Latest_PayMade.Text
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
MsgBox "Tenant Has left the property"
End If
End Sub

Private Sub TenantIDs_Click()
Dim ShowPaymentsInfo As TenantProperty_Record
Dim ShowPaymentsChannel As Integer
Dim ShowTenantPropertyInfo As Property_Record
Dim ShowTenantPropertyChannel As Integer
Dim ShowTenantInfo As Tenant_Record
Dim ShowTenantInfoChannel As Integer
Dim x As Integer
Dim Y As Integer
Dim z As Integer
Dim time As Date
Dim OnceRecordFound As Boolean
Dim Pic As Picture
x = 1

ShowTenantInfoChannel = FreeFile
Open Tenant_File For Random As ShowTenantInfoChannel Len = Tenant_Length
Get ShowTenantInfoChannel, x, ShowTenantInfo
Do While Not EOF(ShowTenantInfoChannel) And OnceRecordFound = False
    If Trim(TenantIDs.List(TenantIDs.ListIndex)) = Trim(ShowTenantInfo.Tenant_REFID) Then
        ShowTInfo
        T_FirstName.Caption = ShowTenantInfo.Tenant_Fname
        T_LastName.Caption = ShowTenantInfo.Tenant_LName
        T_PhoneNumber.Caption = ShowTenantInfo.Tenant_PhoneNumber
        T_EmailAddress.Caption = ShowTenantInfo.Tenant_EmailAddress
        Picture1.AutoRedraw = True
        Set Pic = LoadPicture(ShowTenantInfo.Tenant_Photo)
        Picture1.PaintPicture Pic, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
        Set Picture1.Picture = Picture1.Image
        Tenant_File_PointerR = x
        Y = 1
        ShowTenantPropertyChannel = FreeFile
        Open Add_Property_File For Random As ShowTenantPropertyChannel Len = Add_Property_Length
        Get ShowTenantPropertyChannel, Y, ShowTenantPropertyInfo
        Do While Not EOF(ShowTenantPropertyChannel) And OnceRecordFound = False
            If Trim(PropertyID.List(PropertyID.ListIndex)) = Trim(ShowTenantPropertyInfo.Property_RefID) And Trim(TenantIDs.List(TenantIDs.ListIndex)) = Trim(ShowTenantInfo.Tenant_REFID) Then
            ShowPInfo
            TP_Address1.Caption = ShowTenantPropertyInfo.Address_line_1
            TP_Address2.Caption = ShowTenantPropertyInfo.Address_line_2
            TP_PostCode.Caption = ShowTenantPropertyInfo.Post_Code
            TP_City.Caption = ShowTenantPropertyInfo.City
            TP_RentalPrice.Text = ShowTenantPropertyInfo.Price_Property
            TakeRentalPayment.Text = ShowTenantPropertyInfo.Price_Property
            TP_Feetype.Text = ShowTenantPropertyInfo.Payment_type
            Add_Property_Pointer = Y
                z = 1
                ShowPaymentsChannel = FreeFile
                Open TenantProperty_File For Random As ShowPaymentsChannel Len = TenantProperty_Length
                Get ShowPaymentsChannel, z, ShowPaymentsInfo
                Do While Not EOF(ShowPaymentsChannel) And OnceRecordFound = False
                If Trim(PropertyID.List(PropertyID.ListIndex)) = Trim(ShowTenantPropertyInfo.Property_RefID) And Trim(TenantIDs.List(TenantIDs.ListIndex)) = Trim(ShowTenantInfo.Tenant_REFID) Then
                ShowTPInfo
                TP_TotalRental.Text = ShowPaymentsInfo.TotalRentalPricePayed
                TheTotalRental.Text = ShowPaymentsInfo.TotalRentalPricePayed
                TotalOverDue.Text = ShowPaymentsInfo.PaymentPriceOverDue
                Add7days.Text = ShowPaymentsInfo.Add7days
                TP_Next_DUE.Text = ShowPaymentsInfo.PaymentDueDate
                TP_Latest_PayMade.Text = ShowPaymentsInfo.LatestPayment
                time = TP_Next_DUE.Text
                TP_Overdue1.Text = ShowPaymentsInfo.OverDue01
                TP_Overdue2.Text = ShowPaymentsInfo.OverDue02
                TP_Overdue3.Text = ShowPaymentsInfo.OverDue03
                UpdateDUEDATE
                TP_Latest_PayMade.Text = ShowPaymentsInfo.LatestPayment
                TP_MoveOUTDATE.Text = ShowPaymentsInfo.TenantEndDate
                TP_Moveindate.Text = ShowPaymentsInfo.StartDate
                    If TP_Next_DUE.Text = TP_Moveindate.Text And Trim(TenantIDs.List(TenantIDs.ListIndex)) = Trim(ShowTenantInfo.Tenant_REFID) Then
                    TP_Next_DUE.Text = Format$(time + 7, "short Date")
                    ShowPaymentsInfo.PaymentDueDate = TP_Next_DUE.Text
                    TenantProperty_File_Pointer = z
                    OnceRecordFound = True
                    Put ShowPaymentsChannel, z, ShowPaymentsInfo
                    End If
                    OnceRecordFound = False
                                        If TP_MoveOUTDATE.Text = "00:00:00" Then
                                        TP_Status.Text = "Renting"
                                        Else
                                        TP_Status.Text = "Left Property"
                                        End If
                                            If TP_Status.Text = "Renting" Then
                                            StatusRenting.Visible = True
                                            ElseIf TP_Status.Text = "Left Property" Then
                                            StatusMovedOut.Visible = True
                                            End If
                TenantProperty_File_Pointer = z
                OnceRecordFound = True
                End If
                z = z + 1
                Get ShowPaymentsChannel, z, ShowPaymentsInfo
                Loop
            OnceRecordFound = True
            End If
            Y = Y + 1
            Get ShowTenantPropertyChannel, Y, ShowTenantPropertyInfo
        Loop
    OnceRecordFound = True
    End If
    x = x + 1
    Get ShowTenantInfoChannel, x, ShowTenantInfo
Loop
Close ShowTenantInfoChannel
Close ShowPaymentsChannel
Close ShowTenantPropertyChannel
Dim s As TenantProperty_Record
Dim channell As Integer
Dim f As Integer
Dim Done As Boolean
f = 1
channell = FreeFile
Open TenantProperty_File For Random As channell Len = TenantProperty_Length
Get channell, f, s
Do While Not EOF(channell) And Done = False
If Trim(TenantIDs.List(TenantIDs.ListIndex)) = Trim(s.Tenant_REFID) Then
PayedDate.Caption = s.PaymentDueDate
TenantProperty_File_Pointer = f
Done = True
End If
f = f + 1
Get channell, f, s
Loop
SaveDate.Text = Format$(Now, "short Date")
End Sub
Private Function UpdateDUEDATE()
Dim UpdateDate As TenantProperty_Record
Dim UpdateDateChannel As Integer
Dim M As Integer
Dim OnceAmendDone
M = 1
UpdateDateChannel = FreeFile
Open TenantProperty_File For Random As UpdateDateChannel Len = TenantProperty_Length
Get UpdateDateChannel, M, UpdateDate
    Do While Not EOF(UpdateDateChannel) And OnceAmendDone = False
        If Trim(TenantIDs.List(TenantIDs.ListIndex)) = Trim(UpdateDate.Tenant_REFID) Then
        UpdateDate.PaymentDueDate = TP_Next_DUE.Text
        TenantProperty_File_Pointer = M
        OnceAmendDone = True
        Put UpdateDateChannel, M, UpdateDate
        End If
        M = M + 1
    Get UpdateDateChannel, M, UpdateDate
    Loop
End Function
Private Function ShowTInfo()
Line1.Visible = True
Line2.Visible = True
Line3.Visible = True
Line9.Visible = True
Line10.Visible = True
Line4.Visible = True
Picture1.Visible = True
Label10.Visible = True
Label9.Visible = True
Label6.Visible = True
Label7.Visible = True
Label5.Visible = True
T_FirstName.Visible = True
T_LastName.Visible = True
T_EmailAddress.Visible = True
T_PhoneNumber.Visible = True
End Function
Private Function HideTInfo()
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line9.Visible = False
Line10.Visible = False
Label7.Visible = False
Line4.Visible = False
Picture1.Visible = False
Label10.Visible = False
Label9.Visible = False
Label6.Visible = False
Label5.Visible = False
T_FirstName.Visible = False
T_LastName.Visible = False
T_EmailAddress.Visible = False
T_PhoneNumber.Visible = False
End Function

Private Function ShowPInfo()
Label8.Visible = True
Label17.Visible = True
Label18.Visible = True
Label15.Visible = True
Label16.Visible = True
Line5.Visible = True
Line11.Visible = True
Line8.Visible = True
Line7.Visible = True
Line6.Visible = True
TP_Address1.Visible = True
TP_Address2.Visible = True
TP_PostCode.Visible = True
TP_City.Visible = True
End Function

Private Function HidePInfo()
Label8.Visible = False
Label17.Visible = False
Label18.Visible = False
Label15.Visible = False
Label16.Visible = False
Line5.Visible = False
Line11.Visible = False
Line8.Visible = False
Line7.Visible = False
Line6.Visible = False
TP_Address1.Visible = False
TP_Address2.Visible = False
TP_PostCode.Visible = False
TP_City.Visible = False
End Function

Private Function ShowTPInfo()
Label22.Visible = True
Line12.Visible = True
Line13.Visible = True
Label14.Visible = True
Line14.Visible = True
Label20.Visible = True
Label21.Visible = True
Label12.Visible = True
Label13.Visible = True
Label14.Visible = True
Label31.Visible = True
Label19.Visible = True
Label29.Visible = True
Label30.Visible = True
Label27.Visible = True
Label26.Visible = True
TenantCheckFrame.Visible = True
TP_Feetype.Visible = True
TP_Status.Visible = True
TP_RentalPrice.Visible = True
TP_TotalRental.Visible = True
TP_Next_DUE.Visible = True
TP_Latest_PayMade.Visible = True
TP_Overdue1.Visible = True
TP_Overdue2.Visible = True
TP_Overdue3.Visible = True
TP_MoveOUTDATE.Visible = True
TP_Moveindate.Visible = True
TenantCheckFrame.Visible = True

End Function
Private Function HideTPInfo()
Line12.Visible = False
Line13.Visible = False
Label20.Visible = False
Line14.Visible = False
Label21.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label31.Visible = False
Label19.Visible = False
Label29.Visible = False
Label30.Visible = False
Label27.Visible = False
Label26.Visible = False
TP_Feetype.Visible = False
TP_Status.Visible = False
TP_RentalPrice.Visible = False
TP_TotalRental.Visible = False
TP_Next_DUE.Visible = False
TP_Latest_PayMade.Visible = False
TP_Overdue1.Visible = False
TP_Overdue2.Visible = False
TP_Overdue3.Visible = False
TP_MoveOUTDATE.Visible = False
TP_Moveindate.Visible = False
StatusRenting.Visible = False
StatusMovedOut.Visible = False
TenantCheckFrame.Visible = False
TakeRentalPayment.Visible = False
TheTotalRental.Visible = False
Label23.Visible = False
Label24.Visible = False
TakePayment.Visible = False
Cancelbutton.Visible = False
SaveChanges.Visible = False
TotalOverDue.Visible = False
Label25.Visible = False
End Function

Private Sub TP_Next_DUE_Change()
Dim Days As Date
If IsDate(Add7days.Text) Then
Days = Add7days.Text
Days = DateAdd("d", 7, Days)
DateCalculator.Text = Days
'DateCalculator = Days
Else
MsgBox ("Please input a proper date value!")
End If
End Sub

Private Sub YesOption_Click(index As Integer)
Dim Overduechannel As Integer
Dim Addoverdue As TenantProperty_Record
Dim PreviewRentalInfo As TenantProperty_Record
Dim PreviewRentalInfoChannel As Integer
Dim OnceRecordFound As Boolean
Dim x As Integer
Dim oncefound As Boolean
    ShowAddInfo
x = 1
PreviewRentalInfoChannel = FreeFile
Open TenantProperty_File For Random As PreviewRentalInfoChannel Len = TenantProperty_Length
Get PreviewRentalInfoChannel, x, PreviewRentalInfo
    Do While Not EOF(PreviewRentalInfoChannel) And OnceRecordFound = False
        If Trim(TenantIDs.List(TenantIDs.ListIndex)) = Trim(PreviewRentalInfo.Tenant_REFID) Then
        TheTotalRental.Text = PreviewRentalInfo.TotalRentalPricePayed
        TotalOverDue.Text = PreviewRentalInfo.PaymentPriceOverDue
        TenantProperty_File_Pointer = x
        OnceRecordFound = True
        End If
        x = x + 1
        Get PreviewRentalInfoChannel, x, PreviewRentalInfo
    Loop
End Sub

Private Function ShowAddInfo()
    TakeRentalPayment.Visible = True
    TheTotalRental.Visible = True
    Overduepayment.Visible = True
    Label23.Visible = True
    Label24.Visible = True
    TakePayment.Visible = True
    Cancelbutton.Visible = True
    TotalOverDue.Visible = True
    Label25.Visible = True
    inform.Visible = False
End Function
