VERSION 5.00
Begin VB.Form bookticketfront 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Book a Ticket"
   ClientHeight    =   2640
   ClientLeft      =   8580
   ClientTop       =   4485
   ClientWidth     =   4695
   LinkTopic       =   "Form2"
   ScaleHeight     =   2640
   ScaleWidth      =   4695
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Select Your Choice"
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
      Width           =   4455
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Book a Ticket"
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
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Back"
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
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Search Flight "
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
         TabIndex        =   1
         Top             =   1320
         Width           =   1455
      End
   End
End
Attribute VB_Name = "bookticketfront"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
bookticket1.Show
Unload Me
End Sub

Private Sub Command3_Click()
mainpage.Show
Unload Me
End Sub

Private Sub Command4_Click()
flightsearch.Show
Unload Me
End Sub

