VERSION 5.00
Begin VB.Form mainpage 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Welcome"
   ClientHeight    =   2625
   ClientLeft      =   7830
   ClientTop       =   3900
   ClientWidth     =   4800
   LinkTopic       =   "Form3"
   ScaleHeight     =   2625
   ScaleWidth      =   4800
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "My Zone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Log Out"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Flight List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cancel Ticket "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ticket Booking"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "mainpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
bookticketfront.Show
Unload Me
End Sub

Private Sub Command2_Click()
passlist.Show
Unload Me
End Sub


Private Sub Command4_Click()
DataReport2.Show
End Sub

Private Sub Command5_Click()
Dim a As Integer
a = MsgBox("Are You Sure ?", vbOKCancel + vbQuestion, "Confirmation")
If a = 1 Then
MsgBox "You have successfully logged out..", vbInformation, "Log Out"
Unload Me
frmSplash1.Show
End If
End Sub

Private Sub Command6_Click()
passwaitlist.Show
Unload Me
End Sub

