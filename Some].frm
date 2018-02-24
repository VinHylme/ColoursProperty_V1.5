VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim s As TenantProperty_Record
Dim schannel As Integer
Dim x As Integer

x = 1
schannel = FreeFile
Open TenantProperty_File For Random As schannel Len = TenantProperty_Length
Get schannel, x, s
Do While Not EOF(schannel)
    List1.AddItem s.TenantEndDate
    x = x + 1
Get schannel, x, s
Loop
Close schannel
End Sub

Private Sub Text1_Change()

End Sub
