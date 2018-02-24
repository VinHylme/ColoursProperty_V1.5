VERSION 5.00
Begin VB.Form View_Properties 
   Caption         =   "View Properties"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Close 
      Caption         =   "Close"
      Height          =   735
      Left            =   7440
      TabIndex        =   26
      Top             =   7680
      Width           =   2535
   End
   Begin VB.ListBox PropertyRefIDList 
      BackColor       =   &H80000004&
      Height          =   6495
      Left            =   480
      TabIndex        =   24
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox Property_Address1 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   5280
      TabIndex        =   23
      Top             =   240
      Width           =   4695
   End
   Begin VB.TextBox Property_PayType 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   5280
      TabIndex        =   22
      Top             =   3840
      Width           =   4695
   End
   Begin VB.TextBox Property_PropertyType 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   5280
      TabIndex        =   21
      Top             =   3240
      Width           =   4695
   End
   Begin VB.TextBox Property_Country 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   5280
      TabIndex        =   20
      Top             =   1440
      Width           =   4695
   End
   Begin VB.TextBox LandlordRefID 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   5280
      TabIndex        =   19
      Top             =   5640
      Width           =   4695
   End
   Begin VB.TextBox AgentRefID 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   5280
      TabIndex        =   18
      Top             =   6240
      Width           =   4695
   End
   Begin VB.TextBox Property_Address2 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox Property_State 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   2040
      Width           =   4695
   End
   Begin VB.TextBox Property_PostCode 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   2640
      Width           =   4695
   End
   Begin VB.TextBox Property_RentalPrice 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   4440
      Width           =   4695
   End
   Begin VB.TextBox Property_NumberofBeds 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   5040
      Width           =   4695
   End
   Begin VB.TextBox PropertyRefID 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   6840
      Width           =   4695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Property Refrence ID:"
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
      Left            =   480
      TabIndex        =   25
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
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
      Left            =   3480
      TabIndex        =   17
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label2 
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
      Left            =   3480
      TabIndex        =   16
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label3 
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
      Left            =   3480
      TabIndex        =   15
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label4 
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
      Left            =   3480
      TabIndex        =   14
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label5 
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
      Left            =   3480
      TabIndex        =   13
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Property Type:"
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
      Left            =   3480
      TabIndex        =   12
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Type:"
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
      Left            =   3480
      TabIndex        =   11
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Price(weekly):"
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
      Left            =   3480
      TabIndex        =   10
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of beds:"
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
      Left            =   3480
      TabIndex        =   9
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Landlord ID:"
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
      Left            =   3480
      TabIndex        =   8
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent ID:"
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
      Left            =   3480
      TabIndex        =   7
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Property Ref:"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   6960
      Width           =   2175
   End
End
Attribute VB_Name = "View_Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
Dim ViewProperties As Property_Record
Dim ViewPropertyChannel As Integer
Dim x As Integer
x = 1
ViewPropertyChannel = FreeFile
Open Add_property_file For Random As ViewPropertyChannel Len = Add_property_length
Get ViewPropertyChannel, x, ViewProperties
Do While Not EOF(ViewPropertyChannel)
    PropertyRefIDList.AddItem ViewProperties.Property_RefID
    x = x + 1
    Get ViewPropertyChannel, x, ViewProperties
Loop
Close ViewPropertyChannel
End Sub

Private Sub PropertyRefIDList_Click()
Dim ViewProperties As Property_Record
Dim ViewPropertyChannel As Integer
Dim x As Integer
Dim OnceFoundProperty As Boolean
x = 1
ViewPropertyChannel = FreeFile
Open Add_property_file For Random As ViewPropertyChannel Len = Add_property_length
Get ViewPropertyChannel, x, ViewProperties
Do While Not EOF(ViewPropertyChannel) And OnceFoundProperty = False
    If Trim(PropertyRefIDList.List(PropertyRefIDList.ListIndex)) = Trim(ViewProperties.Property_RefID) Then
    Property_Address1.Text = ViewProperties.Address_line_1
    Property_Address2.Text = ViewProperties.Address_line_2
    Property_Country.Text = ViewProperties.City
    Property_PostCode.Text = ViewProperties.Post_Code
    Property_PropertyType.Text = ViewProperties.property_type
    Property_PayType.Text = ViewProperties.Payment_type
    Property_NumberofBeds.Text = ViewProperties.NumberOfBeds
    Property_RentalPrice.Text = ViewProperties.Price_Property
    LandlordRefID.Text = ViewProperties.Landlord_REFID
    AgentRefID.Text = ViewProperties.Agent_REFID
    PropertyRefID.Text = ViewProperties.Property_RefID
    Add_property_Pointer = x
    OnceFoundProperty = True
    End If
    x = x + 1
    Get ViewPropertyChannel, x, ViewProperties
Loop
Close ViewPropertyChannel
End Sub

