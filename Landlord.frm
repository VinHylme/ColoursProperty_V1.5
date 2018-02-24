VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Landlord 
   Caption         =   "Landlord"
   ClientHeight    =   9090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13770
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   13770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   9960
      TabIndex        =   42
      Top             =   6240
      Width           =   3855
      Begin VB.ComboBox Combo5 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   1  'Simple Combo
         TabIndex        =   45
         Top             =   600
         Width           =   2535
      End
      Begin VB.ComboBox Combo6 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Search On"
         Height          =   735
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10320
      Top             =   5640
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Landlord.frx":0000
      OLEDBString     =   $"Landlord.frx":0092
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3735
      Left            =   0
      TabIndex        =   6
      Top             =   5280
      Width           =   9975
      Begin VB.ComboBox Landlord_ID 
         Height          =   315
         Left            =   6720
         Style           =   1  'Simple Combo
         TabIndex        =   39
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox Landlord_EmailAd 
         Height          =   315
         Left            =   6720
         Style           =   1  'Simple Combo
         TabIndex        =   38
         Top             =   480
         Width           =   3135
      End
      Begin VB.ComboBox Landlord_PhoneNum 
         Height          =   315
         Left            =   6720
         Style           =   1  'Simple Combo
         TabIndex        =   37
         Top             =   120
         Width           =   3135
      End
      Begin VB.ComboBox Landlord_FaxNo 
         Height          =   315
         Left            =   1800
         Style           =   1  'Simple Combo
         TabIndex        =   36
         Top             =   3360
         Width           =   3135
      End
      Begin VB.ComboBox Landlord_City 
         Height          =   315
         Left            =   1800
         Style           =   1  'Simple Combo
         TabIndex        =   35
         Top             =   3000
         Width           =   3135
      End
      Begin VB.ComboBox Landlord_State 
         Height          =   315
         Left            =   1800
         Style           =   1  'Simple Combo
         TabIndex        =   34
         Top             =   2640
         Width           =   3135
      End
      Begin VB.ComboBox Landlord_PostCode 
         Height          =   315
         Left            =   1800
         Style           =   1  'Simple Combo
         TabIndex        =   33
         Top             =   1920
         Width           =   3135
      End
      Begin VB.ComboBox Landlord_CName 
         Height          =   315
         Left            =   1800
         Style           =   1  'Simple Combo
         TabIndex        =   32
         Top             =   1560
         Width           =   3135
      End
      Begin VB.ComboBox Landlord_Address2 
         Height          =   315
         Left            =   1800
         Style           =   1  'Simple Combo
         TabIndex        =   31
         Top             =   1200
         Width           =   3135
      End
      Begin VB.ComboBox Landlord_Address1 
         Height          =   315
         Left            =   1800
         Style           =   1  'Simple Combo
         TabIndex        =   30
         Top             =   840
         Width           =   3135
      End
      Begin VB.ComboBox Landlord_LName 
         Height          =   315
         Left            =   1800
         Style           =   1  'Simple Combo
         TabIndex        =   29
         Top             =   480
         Width           =   3135
      End
      Begin VB.ComboBox Landlord_FName 
         Height          =   315
         Left            =   1800
         Style           =   1  'Simple Combo
         TabIndex        =   28
         Top             =   120
         Width           =   3135
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Print"
         Height          =   495
         Left            =   8040
         TabIndex        =   27
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6720
         TabIndex        =   26
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5400
         TabIndex        =   25
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   8040
         TabIndex        =   24
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   6720
         TabIndex        =   23
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton command1 
         Caption         =   "&Add"
         Height          =   495
         Left            =   5400
         TabIndex        =   22
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox Country_list_property 
         Height          =   315
         ItemData        =   "Landlord.frx":0124
         Left            =   1800
         List            =   "Landlord.frx":0374
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2280
         Width           =   3135
      End
      Begin VB.ComboBox Landlord_Country 
         Height          =   315
         Left            =   1800
         Style           =   1  'Simple Combo
         TabIndex        =   40
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number:"
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
         Left            =   5040
         TabIndex        =   21
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address:"
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
         Left            =   5040
         TabIndex        =   20
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
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
         Left            =   5760
         TabIndex        =   19
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Ref Code:"
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
         Left            =   5040
         TabIndex        =   18
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
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
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name:"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label5 
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
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label6 
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
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label8 
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
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "City:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label7 
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
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax Number:"
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
         Left            =   120
         TabIndex        =   8
         Top             =   3360
         Width           =   2175
      End
   End
   Begin VB.TextBox Landlord_PhoneNo 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      Top             =   11280
      Width           =   4695
   End
   Begin VB.TextBox Landlord_EmailAddress 
      BackColor       =   &H80000004&
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   11880
      Width           =   4695
   End
   Begin MSComctlLib.ListView LandlordTable 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "FirstName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "LastName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "AddressLine 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "AddressLine 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "CompanyName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "PostCode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Country"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "State"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "City"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "FaxNumber"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "EmailAddress"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Height          =   375
      Left            =   11280
      TabIndex        =   41
      Top             =   9120
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
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
      Left            =   5760
      TabIndex        =   5
      Top             =   11400
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address:"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   12000
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
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
      Left            =   5760
      TabIndex        =   3
      Top             =   12120
      Width           =   2175
   End
End
Attribute VB_Name = "Landlord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CN As ADODB.Connection
Private RS As ADODB.Recordset
Private CMD As ADODB.Command
Dim SQL As String, ID As Long, NEWID As Long, _
list As ListItem, strLFName As String, strLLName As String, strLaddress1 As String, _
strLaddress2 As String, strLCName As String, strLpostcode As String, _
strLCountry As String, strLstate As String, strLCity As String, _
strLFaxNo As String, strLPhoneNum As String, strLEmailad As String

Sub ConnDB()
Set CN = New ADODB.Connection
CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DataBase.mdb" & ";Persist Security Info=False"
CN.Open
Set CMD = New ADODB.Command
Set CMD.ActiveConnection = CN
CMD.CommandType = adCmdText
End Sub
Sub LoadLandlordTable()
SQL = "Select * from LANDLORDTABLE"
CMD.CommandText = SQL
Set RS = CMD.Execute
LandlordTable.ListItems.clear
With RS
    Do Until .EOF
Set list = LandlordTable.ListItems.Add(, , !FirstName & "")
list.SubItems(2) = !LastName & ""
list.SubItems(3) = !AddressLine1 & ""
list.SubItems(4) = !AddressLine2 & ""
list.SubItems(5) = !PostCode & ""
list.SubItems(6) = !Country & ""
list.SubItems(7) = !State & ""
list.SubItems(8) = !City & ""
list.SubItems(9) = !FaxNumber & ""
list.SubItems(10) = !PhoneNumber & ""
list.SubItems(11) = !EmailAdress & ""
list.SubItems(1) = CStr(!ID)
.MoveNext
    Loop
End With
With LandlordTable
If .ListItems.Count > 0 Then
Set .SelectedItem = .ListItems(1)
LandlordTable_ItemClick .SelectedItem
End If
End With
Set list = Nothing
Set RS = Nothing
End Sub
Private Function GetNextID() As Long
CMD.CommandText = "SELECT  MAX(ID) AS MAXID FROM LANDLORDTABLE"
Set RS = CMD.Execute
If RS.EOF Then
GetNextID = 1
ElseIf IsNull(RS!MAXID) Then
GetNextID = 1
Else
GetNextID = RS!MAXID + 1
End If
Set RS = Nothing
End Function

Private Sub Add_Landlord_Click()
Label18.Caption = "ADD"
ClearText
Add_Edit
Landlord_FName.SetFocus
End Sub

Private Sub Command2_Click()
Label18.Caption = "EDIT"
Add_Edit
Landlord_FName.SetFocus
End Sub

Private Sub Command3_Click()
With LandlordTable.SelectedItem
strLFName = .Text
strLLName = .SubItems(2)
strLaddress1 = .SubItems(3)
strLaddress2 = .SubItems(4)
strLCName = .SubItems(5)
strLpostcode = .SubItems(6)
strLstate = .SubItems(7)
strLCity = .SubItems(8)
strLFaxNo = .SubItems(9)
strLPhoneNum = .SubItems(10)
strLEmailad = .SubItems(11)
ID = CLng(.SubItems(1))
End With
CMD.CommandText = "DELETE FROM LANDLORDTABLE WHERE ID = " & ID
CMD.Execute
With LandlordTable
If .SelectedItem.Index = .ListItems.Count Then
NEWID = .ListItems.Count - 1
Else
NEWID = .SelectedItem.Index
End If
.ListItems.Remove .SelectedItem.Index
If .ListItems.Count > 0 Then
Set .SelectedItem = .ListItems(NEWID)
LandlordTable_ItemClick .SelectedItem
End If
End With
LandlordTable.SetFocus
End Sub

Private Sub Command4_Click()
Save_Cancel
Select Case Label18.Caption
Case "ADD"
ID = GetNextID()
SQL = "INSERT INTO LANDLORDTABLE VALUES ("
SQL = SQL & ID
SQL = SQL & ", '" & Replace(Landlord_FName.Text, "'", "''") & "'"
SQL = SQL & ", '" & Replace(Landlord_LName.Text, "'", "''") & "'"
SQL = SQL & ", '" & Replace(Landlord_Address1.Text, "'", "''") & "'"
SQL = SQL & ", '" & Replace(Landlord_Address2.Text, "'", "''") & "'"
SQL = SQL & ", '" & Replace(Landlord_CName.Text, "'", "''") & "'"
SQL = SQL & ", '" & Replace(Landlord_PostCode.Text, "'", "''") & "'"
SQL = SQL & ", '" & Replace(Landlord_Country.Text, "'", "''") & "'"
SQL = SQL & ", '" & Replace(Landlord_State.Text, "'", "''") & "'"
SQL = SQL & ", '" & Replace(Landlord_City.Text, "'", "''") & "'"
SQL = SQL & ", '" & Replace(Landlord_FaxNo.Text, "'", "''") & "'"
SQL = SQL & ", '" & Replace(Landlord_PhoneNum.Text, "'", "''") & "'"
SQL = SQL & ", '" & Replace(Landlord_EmailAd.Text, "'", "''") & "'"
SQL = SQL & ")"
Set list = LandlordTable.ListItems.Add(, , Landlord_FName.Text)
PopItem list
With list
.SubItems(10) = CStr(ID)
.EnsureVisible
End With
Set LandlordTable.SelectedItem = list
Set list = Nothing
Case Else
ID = CLng(LandlordTable.SelectedItem.SubItems(3))
SQL = "UPDATE LANDLORDTABLE SET "
SQL = SQL & " FirstName = '" & Replace(Landlord_FName.Text, "'", "''") & "'"
SQL = SQL & ", LastName = '" & Replace(Landlord_LName.Text, "'", "''") & "'"
SQL = SQL & ", Addressline1 = '" & Replace(Landlord_Address1.Text, "'", "''") & "'"
SQL = SQL & " AddressLine2 = '" & Replace(Landlord_Address2.Text, "'", "''") & "'"
SQL = SQL & ", CompanyName = '" & Replace(Landlord_CName.Text, "'", "''") & "'"
SQL = SQL & ", PostCode = '" & Replace(Landlord_PostCode.Text, "'", "''") & "'"
SQL = SQL & " Country = '" & Replace(Landlord_Country.Text, "'", "''") & "'"
SQL = SQL & ", State = '" & Replace(Landlord_State.Text, "'", "''") & "'"
SQL = SQL & ", City = '" & Replace(Landlord_City.Text, "'", "''") & "'"
SQL = SQL & " FaxNo = '" & Replace(Landlord_FaxNo.Text, "'", "''") & "'"
SQL = SQL & ", PhoneNum = '" & Replace(Landlord_PhoneNum.Text, "'", "''") & "'"
SQL = SQL & ", EmailAddress = '" & Replace(Landlord_EmailAd.Text, "'", "''") & "'"
SQL = SQL & " Where ID = " & ID
LandlordTable.SelectedItem.Text = Landlord_FName.Text
PopItem LandlordTable.SelectedItem
End Select
CMD.CommandText = SQL
CMD.Execute
LandlordTable.SetFocus
End Sub

Private Sub Command5_Click()
Save_Cancel
LandlordTable_ItemClick LandlordTable.SelectedItem
LandlordTable.SetFocus
End Sub

Private Sub Form_Load()
LoadForSearch
ConnDB
LoadLandlordTable
End Sub

Private Sub LandlordTable_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
    Landlord_FName.Text = .Text
    Landlord_LName.Text = .SubItems(2)
    Landlord_Address1.Text = .SubItems(3)
    Landlord_Address2.Text = .SubItems(4)
    Landlord_CName.Text = .SubItems(5)
    Landlord_PostCode.Text = .SubItems(6)
    Landlord_Country.Text = .SubItems(7)
    Landlord_State.Text = .SubItems(8)
    Landlord_City.Text = .SubItems(9)
    Landlord_FaxNo.Text = .SubItems(10)
    Landlord_PhoneNum.Text = .SubItems(11)
    'Landlord_EmailAd.Text = .SubItems(12)
End With
End Sub
Sub ClearText()
Landlord_FName.Text = ""
Landlord_LName.Text = ""
Landlord_Address1.Text = ""
Landlord_Address2.Text = ""
Landlord_CName.Text = ""
Landlord_PostCode.Text = ""
Landlord_Country.Text = ""
Landlord_State.Text = ""
Landlord_City.Text = ""
Landlord_FaxNo.Text = ""
Landlord_PhoneNum.Text = ""
Landlord_ID.Text = ""
End Sub
Sub Add_Edit()
command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command6.Enabled = False
LandlordTable.Enabled = False
Command4.Enabled = True
Command5.Enabled = True
Frame1.Enabled = True
End Sub
Sub Save_Cancel()
command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command6.Enabled = True
LandlordTable.Enabled = True
Command4.Enabled = False
Command5.Enabled = False
Frame1.Enabled = False
End Sub
Sub PopItem(Item As ListItem)
With Item
    .Text = Landlord_FName.Text
    .SubItems(2) = Landlord_LName.Text
    .SubItems(3) = Landlord_Address1.Text
    .SubItems(4) = Landlord_Address2.Text
    .SubItems(5) = Landlord_CName.Text
    .SubItems(6) = Landlord_PostCode.Text
    .SubItems(7) = Landlord_Country.Text
    .SubItems(8) = Landlord_State.Text
    .SubItems(9) = Landlord_City.Text
    .SubItems(10) = Landlord_FaxNo.Text
    .SubItems(11) = Landlord_PhoneNum.Text
    .SubItems(1) = Landlord_EmailAd.Text
End With
End Sub
Sub LoadForSearch()
Combo6.AddItem "Name"
Combo6.AddItem "Address"
Combo6.AddItem "Post Code"
Combo6.ListIndex = 0
End Sub

