VERSION 5.00
Begin VB.Form frmRearrange 
   Caption         =   "Rearrange to A"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClick 
      Caption         =   "click"
      Height          =   1935
      Left            =   6600
      TabIndex        =   2
      Top             =   3480
      Width           =   3735
   End
   Begin VB.TextBox txtWords 
      Height          =   2055
      Left            =   1800
      TabIndex        =   1
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox txtWord 
      Height          =   2055
      Left            =   1800
      TabIndex        =   0
      Top             =   1560
      Width           =   2535
   End
End
Attribute VB_Name = "frmRearrange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClick_Click()
Dim a As Integer
Dim c As Integer

For a = 1 To Len(txtWord)
If Mid(txtWord, a, 1) = "a" Then



c = c + 1

End If

Next a
txtWords = c


End Sub
