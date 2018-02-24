VERSION 5.00
Begin VB.Form Landlords 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Manage Landlord"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   19170
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   19170
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List11 
      Height          =   2790
      Left            =   14400
      TabIndex        =   55
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox List10 
      Height          =   2790
      Left            =   17520
      TabIndex        =   44
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox List9 
      Height          =   2790
      Left            =   15960
      TabIndex        =   43
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox List8 
      Height          =   2790
      Left            =   12840
      TabIndex        =   42
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox List7 
      Height          =   2790
      Left            =   11280
      TabIndex        =   41
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox List6 
      Height          =   2790
      Left            =   10080
      TabIndex        =   40
      Top             =   600
      Width           =   1215
   End
   Begin VB.ListBox List5 
      Height          =   2790
      Left            =   8520
      TabIndex        =   39
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox List4 
      Height          =   2790
      Left            =   6960
      TabIndex        =   38
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox List3 
      Height          =   2790
      Left            =   5400
      TabIndex        =   37
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   3840
      TabIndex        =   36
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   2280
      TabIndex        =   35
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Landlord_EmailAddress 
      Height          =   375
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   4200
      Width           =   3255
   End
   Begin VB.ComboBox Country_list_property 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "Landlords.frx":0000
      Left            =   3840
      List            =   "Landlords.frx":0250
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   4800
      Width           =   3255
   End
   Begin VB.CommandButton Save 
      Caption         =   "SAVE"
      Enabled         =   0   'False
      Height          =   615
      Left            =   12240
      TabIndex        =   29
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CommandButton EdIT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "EDIT"
      Height          =   615
      Left            =   14880
      TabIndex        =   28
      Top             =   4080
      Width           =   2775
   End
   Begin VB.CommandButton Close 
      Caption         =   "CLOSE"
      Height          =   615
      Left            =   12240
      TabIndex        =   27
      Top             =   6240
      Width           =   5415
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "CANCEL"
      Enabled         =   0   'False
      Height          =   615
      Left            =   14880
      TabIndex        =   26
      Top             =   5520
      Width           =   2775
   End
   Begin VB.CommandButton Prints 
      Caption         =   "PRINT"
      Height          =   615
      Left            =   12240
      TabIndex        =   25
      Top             =   4800
      Width           =   2535
   End
   Begin VB.CommandButton Delete 
      Caption         =   "DELETE"
      Height          =   615
      Left            =   14880
      TabIndex        =   24
      Top             =   4800
      Width           =   2775
   End
   Begin VB.CommandButton Add 
      Caption         =   "ADD"
      Height          =   615
      Left            =   12240
      TabIndex        =   23
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox LandlordRefID 
      Height          =   375
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   11
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Landlord_PhoneNo 
      Height          =   375
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   6600
      Width           =   3255
   End
   Begin VB.TextBox Landlord_City2 
      Height          =   375
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   6000
      Width           =   3255
   End
   Begin VB.TextBox Landlord_State2 
      Height          =   375
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   5400
      Width           =   3255
   End
   Begin VB.TextBox Landlord_PostCode2 
      Height          =   375
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4200
      Width           =   3255
   End
   Begin VB.TextBox Landlord_AddressLine2_2 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   6600
      Width           =   3255
   End
   Begin VB.TextBox Landlord_AddressLine1_2 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   6000
      Width           =   3255
   End
   Begin VB.TextBox Landlord_Cname2 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   5400
      Width           =   3255
   End
   Begin VB.TextBox Landlord_LName2 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4800
      Width           =   3255
   End
   Begin VB.TextBox Landlord_Fname 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4200
      Width           =   3255
   End
   Begin VB.ListBox Landlord_REFID 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Landlord_Country 
      Height          =   285
      Left            =   3840
      TabIndex        =   31
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label7 
      Height          =   375
      Left            =   7440
      TabIndex        =   57
      Top             =   7320
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   14880
      TabIndex        =   56
      Top             =   360
      Width           =   735
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
      Left            =   7080
      TabIndex        =   54
      Top             =   360
      Width           =   1815
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
      Left            =   8640
      TabIndex        =   53
      Top             =   360
      Width           =   1935
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
      Left            =   10200
      TabIndex        =   52
      Top             =   360
      Width           =   1455
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
      Left            =   11640
      TabIndex        =   51
      Top             =   360
      Width           =   1095
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
      Left            =   13320
      TabIndex        =   50
      Top             =   360
      Width           =   735
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
      Left            =   16080
      TabIndex        =   49
      Top             =   360
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
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   17640
      TabIndex        =   48
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   5520
      TabIndex        =   47
      Top             =   360
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
      TabIndex        =   46
      Top             =   360
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
      TabIndex        =   45
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label27 
      Caption         =   "Label27"
      Height          =   375
      Left            =   7320
      TabIndex        =   34
      Top             =   8040
      Width           =   2655
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
      Left            =   7320
      TabIndex        =   33
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Line Line2 
      X1              =   10680
      X2              =   10680
      Y1              =   3720
      Y2              =   7440
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Landlord Refrence ID:"
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
      Left            =   7320
      TabIndex        =   22
      Top             =   4680
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
      Left            =   3840
      TabIndex        =   21
      Top             =   6360
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
      Left            =   3840
      TabIndex        =   20
      Top             =   5760
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
      Left            =   3840
      TabIndex        =   19
      Top             =   5160
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
      Left            =   3840
      TabIndex        =   18
      Top             =   4560
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
      Left            =   3840
      TabIndex        =   17
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label18 
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
      Left            =   120
      TabIndex        =   16
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Label Label17 
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
      Left            =   120
      TabIndex        =   15
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name:"
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
      Left            =   120
      TabIndex        =   14
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label15 
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
      Left            =   120
      TabIndex        =   13
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label14 
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
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   0
      X2              =   19200
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Landlord Refrence ID:"
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
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.Menu FileOption 
      Caption         =   "File"
      Begin VB.Menu previous 
         Caption         =   "Go back"
      End
   End
   Begin VB.Menu helpform 
      Caption         =   "Help"
   End
   Begin VB.Menu exitlandlord 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Landlords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Delete_Click()
Dim tempDeleteLandlordchannel As Integer
Dim tempDeleteLandlordfile As String
Dim FoundRecord As Boolean
Dim RemoveLandlord As Landlord_Record
Dim RemoveLandlordchannel As Integer
Dim P As Integer
Dim L As Integer
Dim intResponse As Integer
If Landlord_REFID.ListIndex = -1 Then
MsgBox ("Please Select a Landlord To Delete")
Else
    intResponse = MsgBox("Are you sure you want to delete Landlord  " & Landlord_REFID & "?" _
    & "                                  Once the Landlord removed you cannot recover it", _
                         vbYesNo + vbQuestion + vbDefaultButton2, _
                         "Delete")
    If intResponse = vbYes Then
tempDeleteLandlordfile = App.Path + "\Saved_Dat\tempfile.tmp"
RemoveLandlordchannel = FreeFile
Open Landlord_File For Random As RemoveLandlordchannel Len = Landlord_Length
tempDeleteLandlordchannel = FreeFile
Open tempDeleteLandlordfile For Random As tempDeleteLandlordchannel Len = Landlord_Length
P = 1
L = 1
FoundRecord = False
Get RemoveLandlordchannel, P, RemoveLandlord
Do While Not EOF(RemoveLandlordchannel)
        If Landlord_REFID.List(Landlord_REFID.ListIndex) <> RemoveLandlord.Landlord_REFID Then
                Put tempDeleteLandlordchannel, L, RemoveLandlord
                L = L + 1
        Else
                FoundRecord = True
        End If
        P = P + 1
        Get RemoveLandlordchannel, P, RemoveLandlord
Loop
Close RemoveLandlordchannel
Close tempDeleteLandlordchannel
    If FoundRecord = True Then
    MsgBox "This Landlord Has Been Successfully Deleted"
        Kill Landlord_File
        Name tempDeleteLandlordfile As Landlord_File
        Landlord_File_Pointer = Landlord_File_Pointer - 1
    Else
   MsgBox "not found"
        Kill tempDeleteLandlordfile
    End If
ClearThoseListBoxes
Form_Load
    End If
End If
End Sub

Private Sub EdIT_Click()
If Landlord_REFID.ListIndex = -1 Then
MsgBox "Please Select A Landlord Refrence ID"
Else
LandlordRefID.Text = Landlord_REFID.List(Landlord_REFID.ListIndex)
LandlordRefID.BackColor = &H80FF80
Save_Cancel
Label27.Caption = "EDIT"
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
Landlord_Fname.SetFocus
End Function
Private Sub Add_Click()
Save_Cancel
Label27.Caption = "ADD"
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

Private Sub Country_list_property_Click()
Landlord_Country = Country_list_property
End Sub
Private Sub GenerateRandomLandlordRefID()
LandlordRefID.Text = Int(Rnd * 999) + 1
End Sub
Private Sub CheckLandlordRef()
GenerateRandomLandlordRefID
Dim ViewLandlord As Landlord_Record
Dim ViewLandlordChannel As Integer
Dim x As Integer
x = 1
ViewLandlordChannel = FreeFile
Open Landlord_File For Random As ViewLandlordChannel Len = Landlord_Length
Get ViewLandlordChannel, x, ViewLandlord
    Do While Not EOF(ViewLandlordChannel)
        If Trim(ViewLandlord.Landlord_REFID) = Trim(LandlordRefID.Text) Then
            GenerateRandomLandlordRefID
            Landlord_File_Pointer = x
        End If
        x = x + 1
        Get ViewLandlordChannel, x, ViewLandlord
    Loop
Close ViewLandlordChannel
End Sub

Private Sub Form_Activate()
Label7.Caption = Selected_LandlordID
End Sub

Private Sub Form_Load()
Dim ViewLandlordID As Landlord_Record
Dim ViewLandlordIDChannel As Integer
Dim x As Integer
x = 1
ViewLandlordIDChannel = FreeFile
Open Landlord_File For Random As ViewLandlordIDChannel Len = Landlord_Length
Get ViewLandlordIDChannel, x, ViewLandlordID
Do While Not EOF(ViewLandlordIDChannel)
    Landlord_REFID.AddItem ViewLandlordID.Landlord_REFID
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
    x = x + 1
    Get ViewLandlordIDChannel, x, ViewLandlordID
Loop
Close ViewLandlordIDChannel
End Sub

Private Sub Label7_change()
Selected_LandlordID = Label7.Caption
End Sub

Private Sub Landlord_REFID_Click()
Label7.Caption = Landlord_REFID.List(Landlord_REFID.ListIndex)
    List1.ListIndex = Landlord_REFID.ListIndex
    List2.ListIndex = Landlord_REFID.ListIndex
    List3.ListIndex = Landlord_REFID.ListIndex
    List4.ListIndex = Landlord_REFID.ListIndex
    List5.ListIndex = Landlord_REFID.ListIndex
    List6.ListIndex = Landlord_REFID.ListIndex
    List7.ListIndex = Landlord_REFID.ListIndex
    List8.ListIndex = Landlord_REFID.ListIndex
    List9.ListIndex = Landlord_REFID.ListIndex
    List10.ListIndex = Landlord_REFID.ListIndex
    List11.ListIndex = Landlord_REFID.ListIndex
    Landlord_File_Pointer = Landlord_REFID.ListIndex + 1
End Sub

Private Sub Prints_Click()
Print_Landlord.Show 1
End Sub

Private Sub Save_Click()
If Label27.Caption = "ADD" Then
Dim Add_Landlord As Landlord_Record
Dim LandLordchannel As Integer
CheckLandlordRef
    LandLordchannel = FreeFile
    Open Landlord_File For Random As LandLordchannel Len = Landlord_Length
            Add_Landlord.Landlord_Fname = Landlord_Fname.Text
            Add_Landlord.Landlord_LName = Landlord_LName2.Text
            Add_Landlord.Landlord_CName = Landlord_Cname2.Text
            Add_Landlord.Landlord_Address1 = Landlord_AddressLine1_2.Text
            Add_Landlord.Landlord_Address2 = Landlord_AddressLine2_2.Text
            Add_Landlord.Landlord_PostCode = Landlord_PostCode2.Text
            Add_Landlord.Landlord_CountryList = Landlord_Country.Text
            Add_Landlord.Landlord_State = Landlord_State2.Text
            Add_Landlord.Landlord_City = Landlord_City2.Text
            Add_Landlord.Landlord_PhoneNumber = Landlord_PhoneNo.Text
            Add_Landlord.Landlord_EmailAddress = Landlord_EmailAddress.Text
            Add_Landlord.Landlord_REFID = LandlordRefID.Text
            MsgBox LandlordRefID
            Landlord_File_Pointer = Landlord_File_Pointer + 1
    Put LandLordchannel, Landlord_File_Pointer, Add_Landlord
    Close LandLordchannel
ClearTexts
ElseIf Label27.Caption = "EDIT" Then
    Dim Amend_Landlord As Landlord_Record
    Dim AmendLandlordchannel As Integer
        AmendLandlordchannel = FreeFile
        Open Landlord_File For Random As AmendLandlordchannel Len = Landlord_Length
                    Amend_Landlord.Landlord_REFID = LandlordRefID.Text
                    Amend_Landlord.Landlord_Fname = Landlord_Fname.Text
                    Amend_Landlord.Landlord_LName = Landlord_LName2.Text
                    Amend_Landlord.Landlord_CName = Landlord_Cname2.Text
                    Amend_Landlord.Landlord_Address1 = Landlord_AddressLine1_2.Text
                    Amend_Landlord.Landlord_Address2 = Landlord_AddressLine2_2.Text
                    Amend_Landlord.Landlord_PostCode = Landlord_PostCode2.Text
                    Amend_Landlord.Landlord_CountryList = Landlord_Country.Text
                    Amend_Landlord.Landlord_State = Landlord_State2.Text
                    Amend_Landlord.Landlord_City = Landlord_City2.Text
                    Amend_Landlord.Landlord_PhoneNumber = Landlord_PhoneNo.Text
                    Amend_Landlord.Landlord_EmailAddress = Landlord_EmailAddress.Text
            Put AmendLandlordchannel, Landlord_File_Pointer, Amend_Landlord
            Close LandLordchannel
End If
ClearThoseListBoxes
LandlordRefID.BackColor = &H80000005
Form_Load
End Sub

Private Function ClearThoseListBoxes()
Landlord_REFID.Clear
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
List8.Clear
List9.Clear
List10.Clear
List11.Clear
End Function

Private Function HideListIndexes()
Landlord_REFID.ListIndex = -1
List1.ListIndex = -1
List2.ListIndex = -1
List3.ListIndex = -1
List4.ListIndex = -1
List5.ListIndex = -1
List6.ListIndex = -1
List7.ListIndex = -1
List8.ListIndex = -1
List9.ListIndex = -1
List10.ListIndex = -1
List11.ListIndex = -1
End Function
Private Function UnlockTexts()
Landlord_Fname.Locked = False
Landlord_LName2.Locked = False
Landlord_Cname2.Locked = False
Landlord_AddressLine1_2.Locked = False
Landlord_AddressLine2_2.Locked = False
Landlord_PostCode2.Locked = False
Landlord_Country.Locked = False
Landlord_State2.Locked = False
Landlord_City2.Locked = False
Landlord_PhoneNo.Locked = False
Landlord_EmailAddress.Locked = False
Country_list_property.Locked = False
End Function
Private Function lockTexts()
Landlord_Fname.Locked = True
Landlord_LName2.Locked = True
Landlord_Cname2.Locked = True
Landlord_AddressLine1_2.Locked = True
Landlord_AddressLine2_2.Locked = True
Landlord_PostCode2.Locked = True
Landlord_Country.Locked = True
Landlord_State2.Locked = True
Landlord_City2.Locked = True
Landlord_PhoneNo.Locked = True
Landlord_EmailAddress.Locked = True
Country_list_property.Locked = True
End Function
Private Function ClearTexts()
Landlord_Fname.Text = ""
Landlord_LName2.Text = ""
Landlord_Cname2.Text = ""
Landlord_AddressLine1_2.Text = ""
Landlord_AddressLine2_2.Text = ""
Landlord_PostCode2.Text = ""
Landlord_Country.Text = ""
Landlord_State2.Text = ""
Landlord_City2.Text = ""
Landlord_PhoneNo.Text = ""
Landlord_EmailAddress.Text = ""
LandlordRefID.Text = ""
End Function
