VERSION 5.00
Begin VB.Form frmLoop 
   Caption         =   "Looping Practice"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraLoops 
      Caption         =   "Conditional Loops"
      Height          =   7575
      Left            =   2760
      TabIndex        =   7
      Top             =   480
      Width           =   1695
      Begin VB.CommandButton cmdConditonalLoop2 
         Caption         =   "Conditional Loop 2"
         Height          =   735
         Left            =   0
         TabIndex        =   9
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton cmdConditionalLoop1 
         Caption         =   "Conditional Loop 1"
         Height          =   735
         Left            =   0
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdLoop4 
      Caption         =   "Loop 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7800
      TabIndex        =   6
      Top             =   5880
      Width           =   3135
   End
   Begin VB.TextBox txtStop 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      TabIndex        =   5
      Text            =   "0"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox txtStart 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      TabIndex        =   4
      Text            =   "0"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoop3 
      Caption         =   "Loop 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4560
      TabIndex        =   3
      Top             =   5880
      Width           =   3135
   End
   Begin VB.CommandButton cmdLoop2 
      Caption         =   "Loop 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4560
      TabIndex        =   2
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdLoop1 
      Caption         =   "Loop 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4560
      TabIndex        =   0
      Top             =   3840
      Width           =   3135
   End
End
Attribute VB_Name = "frmLoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConditionalLoop1_Click()
Dim intNum As Integer

Do While intNum < 10
txtOut = txtOut & intNum & vbCrLf
intNum = intNum + 1


Loop
End Sub

Private Sub cmdConditonalLoop2_Click()
Dim intNum As Integer

Do
txtOut = txtOut & intNum & vbCrLf
intNum = intNum + 1


Loop While intNum < 10

End Sub

Private Sub cmdLoop1_Click()
txtOut.Text = ""
Dim a As Integer
For a = 1 To 10
txtOut.Text = txtOut.Text & a & vbCrLf


Next a

End Sub

Private Sub cmdLoop2_Click()
txtOut.Text = ""
Dim a As Integer
For a = 15 To 1 Step -1
'txtOut.Text = txtOut.Text & a & vbCrLf
txtOut.Text = txtOut.Text & "hello" & vbCrLf '

Next a 'goes back up to For A'


End Sub

Private Sub cmdLoop3_Click()
txtOut.Text = ""
Dim a As Integer
For a = 1 To 10
txtOut.Text = txtOut.Text & a & "x" & "6" & "=" & a * 6 & vbCrLf

Next a

End Sub

Private Sub cmdLoop4_Click()
txtOut.Text = ""
Dim a As Integer
Dim intStart, intStop As Integer
intStart = txtStart.Text 'assignment statement'
intStop = txtStop.Text
For a = intStart To intStop
txtOut = txtOut & a & vbCrLf
Next a


End Sub
