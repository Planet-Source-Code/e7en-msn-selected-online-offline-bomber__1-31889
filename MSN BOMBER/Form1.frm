VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MSN Online/Offline Bomb"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Text            =   "4"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bomb!"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Number of times to Bomb:"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Email Adress:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'======================================================================'
'| Created By: §e7eN                                                  |'
'| Description: This will Online/Offline Bomb Someone.                |'
'|              Many Thanks to Lesesne For the block code             |'
'|              and thanks to the dude who made the wait Function.    |'
'|                                                                    |'
'| Contact: hate_114@hotmail.com                                      |'
'|                                                                    |'
'| *If you wish to use this in one of your Programs please E-mail me* |'
'======================================================================


Public WithEvents oMSN As MessengerAPI.Messenger 'Assign to the Messanger API Libaray
Attribute oMSN.VB_VarHelpID = -1

Private Sub Command1_Click()
For z = 0 To List1.ListCount - 1 'Loop for all people Added to the List
Bomber (List1.List(z)) 'Start Bombing
Next
End Sub

Private Sub Command2_Click()
List1.AddItem Text1.Text 'Add Email Adress to the List
Text1.Text = ""
End Sub

Private Sub Command3_Click()

If List1.Text = "" Then Exit Sub
List1.RemoveItem List1.Text 'Remove Item From The List

End Sub

Private Sub Form_Load()
Set oMSN = New MessengerAPI.Messenger 'Set the Messanger API

End Sub

Sub Bomber(EmailAddy As String)

Dim oGroup As IMessengerGroup 'Assign the Groups
Dim oContact As IMessengerContact 'Assign the Contacts

For Each oGroup In oMSN.MyGroups 'Loop for Every Group

        For Each oContact In oGroup.Contacts 'Loop For every Contact In the group
        
        If EmailAddy = oContact.SigninName Then ' Check if E-mail adresses Match
        
            For x = 1 To Text2.Text * 2 'Bomb them how Evermany times.
                                'We use Text2.Text * 2 because in every loop it only Makes them offline or online
                                
         If oContact.Blocked = False Then oContact.Blocked = True Else oContact.Blocked = False
            'Make them blocked or Unbock them
                 Wait 1 'Wait for status to change
            Next 'Do it all again
    
            oContact.Blocked = False ' Finally Unblock them
        End If
        Next oContact 'Goto next contact in the group
    
Next oGroup 'Goto next group

End Sub

'Thanks to who ever made this wait code it rocks!
'Please post your name undercomments so we can give you credit

Public Function Wait(ByVal TimeToWait As Long) 'Time In seconds
    Dim EndTime As Long
    EndTime = GetTickCount + TimeToWait * 500 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds

    Do Until GetTickCount > EndTime

        DoEvents
        Loop
    End Function

