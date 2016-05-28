VERSION 5.00
Begin VB.Form adminspace 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Administrator's Workspace"
   ClientHeight    =   3270
   ClientLeft      =   7635
   ClientTop       =   3330
   ClientWidth     =   4590
   LinkTopic       =   "Form2"
   ScaleHeight     =   3270
   ScaleWidth      =   4590
   Begin VB.CommandButton newuser 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Manage Users"
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
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Hello Admin !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Passenger List"
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Manage Flights"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
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
         Height          =   615
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Welcome to the Admin Portal. "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "adminspace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
flightedit.Show
Unload Me
End Sub

Private Sub Command2_Click()
DataReport1.Show
End Sub

Private Sub Command3_Click()
Dim a As Integer
a = MsgBox("Are You Sure ?", vbOKCancel + vbQuestion, "Confirmation")
If a = 1 Then
MsgBox "You have successfully logged out..", vbInformation, "Log Out"
Form1.Show
Unload Me
End If
End Sub


Private Sub Form_Load()
Dim dtmdate As Date
dtmdate = DateValue(Now)
Dim dtmtime As Date
dtmtime = TimeValue(Now)
Label1.Caption = dtmdate
Label2.Caption = dtmtime


End Sub

Private Sub newuser_Click()
Unload Me
userlist.Show
End Sub
