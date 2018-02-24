VERSION 5.00
Begin VB.Form RemoveTenantFromProperty 
   Caption         =   "Remove Tenant From Property"
   ClientHeight    =   3750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Remove_Tenant 
      Caption         =   "Remove Tenant"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox ENDDATE 
      Height          =   405
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.ListBox TenantRef 
      Height          =   3375
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.ListBox PropertyRef 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "RemoveTenantFromProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
PropertyRef.Visible = True
TenantRef.Visible = False
PropertyLinked.Visible = False
Unlink.Visible = False
Cancel.Visible = False
End Sub

Private Sub Command1_Click()
ENDDATE.Text = Format$(Now, "short Date")
End Sub

Private Sub Form_Load()
Dim ShowTenantProperty As TenantProperty_Record
Dim ShowTenantPropertyCHANNEL As Integer
Dim s As Integer
s = 1
ShowTenantPropertyCHANNEL = FreeFile
Open TenantProperty_File For Random As ShowTenantPropertyCHANNEL Len = TenantProperty_Length
Get ShowTenantPropertyCHANNEL, s, ShowTenantProperty
Do While Not EOF(ShowTenantPropertyCHANNEL)
    PropertyRef.AddItem ShowTenantProperty.Property_RefID
    s = s + 1
Get ShowTenantPropertyCHANNEL, s, ShowTenantProperty
Loop
Close ShowTenantPropertyCHANNEL
End Sub

Private Sub PropertyRef_Click()
Dim SelectTenant As TenantProperty_Record
Dim SelectTenantChannel As Integer
Dim c As Integer
Dim Found As Boolean
ClearListBox
c = 1
SelectTenantChannel = FreeFile
Open TenantProperty_File For Random As SelectTenantChannel Len = TenantProperty_Length
Get SelectTenantChannel, c, SelectTenant
Do While Not EOF(SelectTenantChannel) And Found = False
    If Trim(PropertyRef.List(PropertyRef.ListIndex)) = Trim(SelectTenant.Property_RefID) Then
    TenantRef.List(TenantRef.ListIndex) = SelectTenant.Tenant_REFID
    'PropertyLinked.List(PropertyLinked.ListIndex) = SelectTenant.Property_RefID
    'TenantProperty_File_Pointer = c
    'Found = True
    End If
    c = c + 1
    Get SelectTenantChannel, c, SelectTenant
Loop
Close SelectTenantChannel
'PropertyRef.Visible = False
'TenantRef.Visible = True
'PropertyLinked.Visible = True
'Unlink.Visible = True
'cancel.Visible = True
'closee.Visible = True
End Sub

Private Function ClearListBox()
TenantRef.Clear
'PropertyLinked.Clear
End Function

Private Sub Unlink_Click()
Dim P As Integer
Dim FoundRecord As Boolean
Dim intResponse As Integer
Dim RemoveTenantProperty As TenantProperty_Record
Dim RemoveTenantPropertyChannel As Integer
    intResponse = MsgBox("Are you sure you want to Remove This Tenant," & TenantRef & "From this Property?" _
    & "                                  Once the Tenant has been removed you cannot recover it", _
                         vbYesNo + vbQuestion + vbDefaultButton2, _
                         "Delete")
    If intResponse = vbYes Then

P = 1
RemoveTenantPropertyChannel = FreeFile
    Open TenantProperty_File For Random As RemoveTenantPropertyChannel Len = TenantProperty_Length
        Get RemoveTenantPropertyChannel, P, RemoveTenantProperty
Do While Not EOF(RemoveTenantPropertyChannel) And FoundRecord = False
         If Trim(RemoveTenantProperty.TenantEndDate) = Trim(ENDDATE.Text) Then
            RemoveTenantProperty.TenantEndDate = ENDDATE.Text
            TenantProperty_File_Pointer = P
            AmendDone = True
        Put RemoveTenantPropertyChannel, TenantProperty_File_Pointer, RemoveTenantProperty
        End If
        P = P + 1
        Get RemoveTenantPropertyChannel, P, RemoveTenantProperty
Loop
Close RemoveTenantPropertyChannel
    End If
End Sub

Private Function RecordTenantEndDate()
Dim ENDdATEs As TenantProperty_Record
Dim EndateCHANNEL As Integer
EndateCHANNEL = FreeFile
Open TenantProperty_File For Random As EndateCHANNEL Len = TenantProperty_Length
        ENDdATEs.TenantEndDate = ENDDATE.Text
        TenantProperty_File_Pointer = TenantProperty_File_Pointer + 1
        Put EndateCHANNEL, TenantProperty_File_Pointer, ENDdATEs
Close EndateCHANNEL
End Function

Private Sub Remove_Tenant_Click()
Dim P As Integer
Dim FoundRecord As Boolean
Dim intResponse As Integer
Dim RemoveTenantProperty As TenantProperty_Record
Dim RemoveTenantPropertyChannel As Integer
   intResponse = MsgBox("Are you sure you want to Remove This Tenant," & TenantRef & "From this Property?" _
    & " Once the Tenant has been removed you cannot recover it", _
                         vbYesNo + vbQuestion + vbDefaultButton2, _
                         "Delete")
    If intResponse = vbYes Then
P = 1
RemoveTenantPropertyChannel = FreeFile
    Open TenantProperty_File For Random As RemoveTenantPropertyChannel Len = TenantProperty_Length
        Get RemoveTenantPropertyChannel, P, RemoveTenantProperty
Do While Not EOF(RemoveTenantPropertyChannel) And FoundRecord = False
         If Trim(RemoveTenantProperty.Tenant_REFID) = Trim(TenantRef.List(TenantRef.ListIndex)) Then
            RemoveTenantProperty.TenantEndDate = ENDDATE.Text
            TenantProperty_File_Pointer = P
            AmendDone = True
            MsgBox "woop woop"
        Put RemoveTenantPropertyChannel, TenantProperty_File_Pointer, RemoveTenantProperty
        End If
        P = P + 1
        Get RemoveTenantPropertyChannel, P, RemoveTenantProperty
Loop
Close RemoveTenantPropertyChannel
    End If
End Sub

Private Sub TenantRef_Click()
ENDDATE.Text = Format$(Now, "short Date")
End Sub
