VERSION 5.00
Begin VB.Form frmStringPractice 
   Caption         =   "String Practice"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtString 
      Height          =   1575
      Left            =   960
      TabIndex        =   2
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "How many E's?"
      Height          =   1215
      Left            =   6120
      TabIndex        =   1
      Top             =   2400
      Width           =   3735
   End
   Begin VB.TextBox txtPhrase 
      Height          =   1575
      Left            =   960
      TabIndex        =   0
      Top             =   3960
      Width           =   3015
   End
End
Attribute VB_Name = "frmStringPractice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNumber_Click()
Dim a As Integer
Dim c As Integer
c = 1
For a = 1 To Len(txtString)
If Mid(txtString, a, 1) = " " Then

c = c + 1

End If

Next a
txtPhrase = c
End Sub
